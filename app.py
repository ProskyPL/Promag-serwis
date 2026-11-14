# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import datetime
import os
import json
import base64
import math
import re
import smtplib
from email.message import EmailMessage
from io import BytesIO
from PIL import Image
from streamlit_drawable_canvas import st_canvas
import gspread
from google.oauth2.service_account import Credentials

# Import zewnętrznej Bazy Wiedzy (z fallbackiem, gdyby plik jeszcze nie był wgrany)
try:
    from baza_wiedzy import (
        CENNIK_BAZOWY,
        DOMYSLNE_NARZEDZIA,
        DOMYSLNE_MATERIALY,
        DIAGNOSTYKA_BAZA,
        INSTRUKCJE_KONFIGURACJI_TEXT,
        POSTAL_PREFIX_COORDS
    )
except ImportError:
    CENNIK_BAZOWY = {"Przegląd/Konserwacja PC": 650.0, "Wymiana glikolu + odpowietrzenie": 1120.0}
    DOMYSLNE_NARZEDZIA = [{"nazwa": "Refraktometr", "ilosc": 1}]
    DOMYSLNE_MATERIALY = [{"nazwa": "Glikol 20L", "ilosc": 2, "koszt_zakupu": 180.0, "cena_klienta": 320.0, "slowo_klucz": "glikol"}]
    DIAGNOSTYKA_BAZA = {"Brak pliku baza_wiedzy.py": {"Wgraj plik baza_wiedzy.py na GitHub": "Wgraj plik baza_wiedzy.py do tego samego katalogu co app.py."}}
    INSTRUKCJE_KONFIGURACJI_TEXT = "Wgraj plik baza_wiedzy.py."
    POSTAL_PREFIX_COORDS = {"05": (52.08, 20.85)}

# ==========================================
# 1. KONFIGURACJA STRONY I RESPONSYWNY CSS
# ==========================================
st.set_page_config(
    page_title="PROMAG OZE - System Serwisowy",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    .stButton > button {
        width: 100%;
        min-height: 48px;
        font-size: 15px !important;
        font-weight: 600 !important;
        border-radius: 10px !important;
        margin-bottom: 6px;
    }
    .kpi-card {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        color: #ffffff;
        padding: 15px;
        border-radius: 12px;
        border-left: 6px solid #10b981;
        margin-bottom: 10px;
    }
    .kpi-card-warn { border-left: 6px solid #f59e0b; }
    .kpi-card-danger {
        background: linear-gradient(135deg, #7f1d1d 0%, #450a0a 100%);
        color: #ffffff;
        padding: 15px;
        border-radius: 12px;
        border-left: 6px solid #ef4444;
        margin-bottom: 12px;
    }
    .kpi-title { font-size: 12px; text-transform: uppercase; color: #94a3b8; }
    .kpi-value { font-size: 21px; font-weight: 800; color: #f8fafc; }
    .kpi-sub { font-size: 13px; color: #38bdf8; margin-top: 4px; }
    .day-card {
        background-color: #1e293b;
        color: #f8fafc;
        border-radius: 12px;
        padding: 15px;
        margin-bottom: 10px;
        border-left: 6px solid #38bdf8;
    }
    .day-card-today {
        border-left: 6px solid #10b981;
        background: linear-gradient(135deg, #064e3b 0%, #0f172a 100%);
    }
    .client-box {
        background-color: #f8fafc;
        border: 2px solid #cbd5e1;
        border-radius: 12px;
        padding: 18px;
        margin-top: 12px;
        margin-bottom: 12px;
        color: #0f172a;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. STAŁE OPERACYJNE I KOSZTOWE
# ==========================================
GOOGLE_SHEET_NAME = "Serwisowa posortowana"
LOCAL_EXCEL_FILE = "Serwisowa posortowana.xlsx"
PROFILES_FILE = "profiles.json"
LOGO_FILE = "logo.png"

PIN_KOORDYNATOR = "1234"
PIN_ZARZAD = "9999"

DEFAULT_HOURLY_RATE = 70.0   # 70 zł / godzina pracy serwisanta
DEFAULT_FUEL_PRICE = 8.0     # 8 zł / litr paliwa
DEFAULT_ONSITE_HOURS = 1.5   # 1.5 godziny średnio u klienta
DEFAULT_BURN_L_100KM = 9.0   # domyślne spalanie busa 9 l / 100 km
DEFAULT_SPEED_KMH = 60.0     # średnia prędkość przelotowa 60 km/h
ROAD_TORTUOSITY = 1.3        # współczynnik krętości dróg względem linii prostej GPS
TARGET_NET_PROFIT = 1200.0   # docelowy czysty zysk z trasy

COMMISSION_TIER_1 = 15000.0
COMMISSION_RATE_1 = 0.03
COMMISSION_TIER_2 = 30000.0
COMMISSION_RATE_2 = 0.04

def get_default_tech_profile(name, pin, role_desc, umowa, base_sal):
    return {
        "pin": pin,
        "pelne_dane": f"{name} - {role_desc}",
        "telefon": "",
        "umowa": umowa,
        "wynagrodzenie_bazowe": base_sal,
        "bhp_data": "2026-01-15",
        "bhp_protokol": "Ważne",
        "medycyna_data": "2026-01-10",
        "medycyna_zakres": "Praca pow. 3m, SEP E+D, kat. B",
        "urlopy_l4": [],
        "notatki_poufne": "",
        "narzedzia": [dict(x) for x in DOMYSLNE_NARZEDZIA],
        "materialy": [dict(x) for x in DOMYSLNE_MATERIALY],
        "flota": {
            "nr_rej": "WPR 00000",
            "przebieg_km": 120000,
            "limit_olej_km": 130000,
            "data_oc_przeglad": "2027-05-01"
        }
    }

DEFAULT_TECHS = {
    "Leszek": get_default_tech_profile("Leszek", "4829", "Serwisant OZE", "Umowa o pracę", 6500.0),
    "Mikołaj": get_default_tech_profile("Mikołaj", "7315", "Serwisant OZE", "B2B", 7000.0),
    "Wojtek": get_default_tech_profile("Wojtek", "5942", "Serwisant OZE", "Umowa o pracę", 6000.0),
    "Serwisant 4": get_default_tech_profile("Serwisant 4", "8164", "Wakat", "Umowa zlecenie", 6000.0)
}

DEFAULT_EMAIL_CONFIG = {
    "smtp_server": "smtp.gmail.com",
    "smtp_port": 587,
    "sender_email": "",
    "sender_password": "",
    "email_montaz": "montaz@twojadomena.pl",
    "email_technik": "technik@twojadomena.pl",
    "handlowcy": [
        {"nazwa": "Handlowiec 1 (Pompy / PV)", "email": "handlowiec1@twojadomena.pl", "rejon": "Cała Polska"},
        {"nazwa": "Handlowiec 2 (Magazyny / Modernizacje)", "email": "handlowiec2@twojadomena.pl", "rejon": "Mazowieckie"}
    ]
}

# ==========================================
# 3. FUNKCJE POMOCNICZE (TELEFONY, GPS, KORYTARZ, GSHEETS, EMAIL)
# ==========================================
def extract_phone_numbers(raw_phone_str, raw_client_str=""):
    combined = f"{raw_phone_str} {raw_client_str}"
    candidates = re.findall(r'(?:\+48)?\s*\d{3}\s*\d{3}\s*\d{3}|\d{9,11}|\d{2}\s*\d{3}\s*\d{2}\s*\d{2}', combined)
    phones = []
    for c in candidates:
        digits = re.sub(r'\D', '', c)
        if len(digits) >= 9:
            short_num = digits[-9:]
            if short_num not in phones:
                phones.append(short_num)
    if not phones:
        digits_all = re.findall(r'\d{9}', re.sub(r'\s+', '', combined))
        for d in digits_all:
            if d not in phones:
                phones.append(d)
    return phones

def extract_coords_from_row(row):
    raw_text = f"{row.get('Kod_Pocztowy', '')} {row.get('Adres', '')} {row.get('Miasto', '')}"
    match = re.search(r'\b(\d{2})-\d{3}\b', str(raw_text))
    if match:
        prefix = match.group(1)
        if prefix in POSTAL_PREFIX_COORDS:
            full_zip = re.sub(r'\D', '', match.group(0))
            offset = (int(full_zip[-2:]) - 50) * 0.003
            base_lat, base_lon = POSTAL_PREFIX_COORDS[prefix]
            return base_lat + offset, base_lon + offset
    return 52.08, 20.85

def get_coords_for_postal(postal_str):
    match = re.search(r'\b(\d{2})\b', str(postal_str))
    if match:
        prefix = match.group(1)
        if prefix in POSTAL_PREFIX_COORDS:
            full_digits = re.sub(r'\D', '', str(postal_str))
            offset = (int(full_digits[-2:]) - 50) * 0.003 if len(full_digits) >= 5 else 0.0
            base_lat, base_lon = POSTAL_PREFIX_COORDS[prefix]
            return base_lat + offset, base_lon + offset
    return 52.08, 20.85

def haversine_km(lat1, lon1, lat2, lon2):
    R = 6371.0
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlambda = math.radians(lon2 - lon1)
    a = math.sin(dphi / 2.0)**2 + math.cos(phi1) * math.cos(phi2) * math.sin(dlambda / 2.0)**2
    return 2 * R * math.atan2(math.sqrt(a), math.sqrt(1 - a))

def distance_from_segment_km(p_lat, p_lon, a_lat, a_lon, b_lat, b_lon):
    """Oblicza odległość punktu P od odcinka trasy A -> B w kilometrach (Korytarz trasy)."""
    lat_scale = 111.32
    lon_scale = 111.32 * math.cos(math.radians((a_lat + b_lat) / 2.0))
    px, py = (p_lon - a_lon) * lon_scale, (p_lat - a_lat) * lat_scale
    bx, by = (b_lon - a_lon) * lon_scale, (b_lat - a_lat) * lat_scale
    seg_len_sq = bx * bx + by * by
    if seg_len_sq == 0:
        return math.hypot(px, py)
    t = max(0.0, min(1.0, (px * bx + py * by) / seg_len_sq))
    proj_x, proj_y = t * bx, t * by
    return math.hypot(px - proj_x, py - proj_y)

def estimate_route_economics(points_coords, base_coords, gross_revenue, parts_cost=0.0,
                             hourly_rate=DEFAULT_HOURLY_RATE, fuel_price=DEFAULT_FUEL_PRICE,
                             burn_l_100km=DEFAULT_BURN_L_100KM, onsite_hours_per_client=DEFAULT_ONSITE_HOURS):
    """Liczy pełną pętlę trasy (Baza -> Punkty -> Baza), paliwo, czas jazdy, czas pracy i Czysty Zysk."""
    if not points_coords:
        return {
            "road_km": 0.0, "drive_hours": 0.0, "onsite_hours": 0.0,
            "fuel_cost": 0.0, "labor_drive_cost": 0.0, "labor_onsite_cost": 0.0,
            "total_cost": 0.0, "net_profit": 0.0
        }
    remaining = list(points_coords)
    curr_lat, curr_lon = base_coords
    straight_km = 0.0
    while remaining:
        nearest_idx = min(range(len(remaining)), key=lambda i: haversine_km(curr_lat, curr_lon, remaining[i][0], remaining[i][1]))
        nxt_lat, nxt_lon = remaining.pop(nearest_idx)
        straight_km += haversine_km(curr_lat, curr_lon, nxt_lat, nxt_lon)
        curr_lat, curr_lon = nxt_lat, nxt_lon
    straight_km += haversine_km(curr_lat, curr_lon, base_coords[0], base_coords[1])

    road_km = straight_km * ROAD_TORTUOSITY
    drive_hours = road_km / DEFAULT_SPEED_KMH
    onsite_hours = len(points_coords) * onsite_hours_per_client

    fuel_cost = (road_km / 100.0) * burn_l_100km * fuel_price
    labor_drive_cost = drive_hours * hourly_rate
    labor_onsite_cost = onsite_hours * hourly_rate
    total_cost = fuel_cost + labor_drive_cost + labor_onsite_cost + parts_cost
    net_profit = gross_revenue - total_cost

    return {
        "road_km": round(road_km, 1),
        "drive_hours": round(drive_hours, 1),
        "onsite_hours": round(onsite_hours, 1),
        "fuel_cost": round(fuel_cost, 2),
        "labor_drive_cost": round(labor_drive_cost, 2),
        "labor_onsite_cost": round(labor_onsite_cost, 2),
        "total_cost": round(total_cost, 2),
        "net_profit": round(net_profit, 2)
    }

def send_system_email(email_cfg, to_emails, subject, body_text, attachments=None):
    sender = email_cfg.get("sender_email", "").strip()
    password = email_cfg.get("sender_password", "").strip()
    server_host = email_cfg.get("smtp_server", "smtp.gmail.com").strip()
    port = int(email_cfg.get("smtp_port", 587))

    if not sender or not password or not to_emails:
        return False, "Brak skonfigurowanego konta nadawcy SMTP lub adresu odbiorcy w zakładce Ustawienia E-mail."

    if isinstance(to_emails, str):
        to_list = [x.strip() for x in to_emails.split(",") if x.strip()]
    else:
        to_list = [x.strip() for x in to_emails if str(x).strip()]

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = ", ".join(to_list)
    msg.set_content(body_text)

    if attachments:
        for att in attachments:
            try:
                file_bytes = att.getvalue()
                file_name = att.name
                msg.add_attachment(file_bytes, maintype="application", subtype="octet-stream", filename=file_name)
            except Exception:
                pass

    try:
        if port == 465:
            with smtplib.SMTP_SSL(server_host, port, timeout=10) as server:
                server.login(sender, password)
                server.send_message(msg)
        else:
            with smtplib.SMTP(server_host, port, timeout=10) as server:
                server.starttls()
                server.login(sender, password)
                server.send_message(msg)
        return True, f"Wysłano e-mail do: {', '.join(to_list)}"
    except Exception as e:
        return False, f"Błąd wysyłki SMTP: {e}"

def get_gsheet_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    try:
        if "gcp_service_account" in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
            creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
            return gspread.authorize(creds)
        elif os.path.exists("credentials.json"):
            creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
            return gspread.authorize(creds)
    except Exception:
        return None
    return None

def ensure_tech_profile_fields(profiles):
    for t_name, p in profiles.items():
        if "narzedzia" not in p or not isinstance(p["narzedzia"], list):
            p["narzedzia"] = [dict(x) for x in DOMYSLNE_NARZEDZIA]
        if "materialy" not in p or not isinstance(p["materialy"], list):
            p["materialy"] = [dict(x) for x in DOMYSLNE_MATERIALY]
        if "flota" not in p or not isinstance(p["flota"], dict):
            p["flota"] = {
                "nr_rej": "WPR 00000",
                "przebieg_km": 120000,
                "limit_olej_km": 130000,
                "data_oc_przeglad": "2027-05-01"
            }
    return profiles

def load_config_and_profiles():
    profiles = DEFAULT_TECHS.copy()
    logo_b64 = None
    email_cfg = DEFAULT_EMAIL_CONFIG.copy()
    if os.path.exists(PROFILES_FILE):
        try:
            with open(PROFILES_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
                profiles = saved.get("techs", DEFAULT_TECHS)
                logo_b64 = saved.get("logo_b64", None)
                email_cfg = saved.get("email_cfg", DEFAULT_EMAIL_CONFIG)
        except Exception:
            pass
    client = get_gsheet_client()
    if client:
        try:
            sh = client.open(GOOGLE_SHEET_NAME)
            try:
                ws = sh.worksheet("Config_System")
                raw_json = ws.acell("A1").value
                raw_logo = ws.acell("B1").value
                raw_email = ws.acell("C1").value
                if raw_json:
                    profiles = json.loads(raw_json)
                if raw_logo:
                    logo_b64 = raw_logo
                if raw_email:
                    email_cfg = json.loads(raw_email)
            except gspread.WorksheetNotFound:
                ws = sh.add_worksheet(title="Config_System", rows="10", cols="5")
                ws.update_acell("A1", json.dumps(profiles, ensure_ascii=False))
                ws.update_acell("C1", json.dumps(email_cfg, ensure_ascii=False))
        except Exception:
            pass
    return ensure_tech_profile_fields(profiles), logo_b64, email_cfg

def save_config_and_profiles(profiles, logo_b64=None, email_cfg=None):
    if logo_b64 is None:
        logo_b64 = st.session_state.get("logo_b64", None)
    if email_cfg is None:
        email_cfg = st.session_state.get("email_cfg", DEFAULT_EMAIL_CONFIG)
    profiles = ensure_tech_profile_fields(profiles)
    payload = {"techs": profiles, "logo_b64": logo_b64, "email_cfg": email_cfg}
    try:
        with open(PROFILES_FILE, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
    client = get_gsheet_client()
    if client:
        try:
            sh = client.open(GOOGLE_SHEET_NAME)
            try:
                ws = sh.worksheet("Config_System")
            except gspread.WorksheetNotFound:
                ws = sh.add_worksheet(title="Config_System", rows="10", cols="5")
            ws.update_acell("A1", json.dumps(profiles, ensure_ascii=False))
            if logo_b64:
                ws.update_acell("B1", logo_b64)
            if email_cfg:
                ws.update_acell("C1", json.dumps(email_cfg, ensure_ascii=False))
        except Exception:
            pass

def render_brand_logo(max_width=240, centered=False):
    logo_b64 = st.session_state.get("logo_b64", None)
    if logo_b64:
        align_style = "display: block; margin-left: auto; margin-right: auto;" if centered else ""
        st.markdown(
            f'<img src="data:image/png;base64,{logo_b64}" style="max-width:{max_width}px; max-height:95px; object-fit:contain; {align_style} margin-bottom:10px;" alt="Logo Firmy">',
            unsafe_allow_html=True
        )
    elif os.path.exists(LOGO_FILE):
        st.image(LOGO_FILE, width=max_width)
    else:
        if centered:
            st.markdown("<h3 style='text-align:center; color:#0ea5e9; margin-bottom:5px;'>⚡ PROMAG SERWIS OZE</h3>", unsafe_allow_html=True)
        else:
            st.markdown("<h4 style='color:#0ea5e9; margin-bottom:5px;'>⚡ PROMAG SERWIS OZE</h4>", unsafe_allow_html=True)

REQUIRED_COLS = [
    "ID_Zgloszenia", "Klient", "Adres", "Kod_Pocztowy", "Miasto", "Telefon",
    "Typ_Instalacji", "Wymagany_Przeglad", "Opis_Zgloszenia", "Priorytet_Krytyczny", "Typ_Serwisu",
    "Serwisant", "Data_Planowana", "Status", "Naprawione_1_Wizyta",
    "Kwota_Przewidywana", "Zgoda_Klienta_Kwota", "Kwota_Koncowa", "Koszt_Materialow",
    "Forma_Platnosci", "Gotowka_Rozliczona", "Czas_Pracy_Min",
    "Wyniki_Przegladu", "Zdjecia_Potwierdzone", "Zalecenia_Klienta",
    "Notatka_Serwisanta", "Data_Zamkniecia", "Podpis_Zlozony",
    "Do_Oddzwonienia", "Data_Oddzwonienia", "Potwierdzony_Tel", "Trudny_Klient",
    "Potrzebne_Czesci", "Lead_Handlowy", "Oddelegowano_Dzial", "Oddelegowano_Email",
    "Data_Oddelegowania", "Notatka_Dla_Dzialu", "Otrzymano_Oferte_Termin", "Ostatnie_Przypomnienie"
]

NUMERIC_COLS = ["Kwota_Przewidywana", "Zgoda_Klienta_Kwota", "Kwota_Koncowa", "Koszt_Materialow", "Czas_Pracy_Min"]

def ensure_columns(df):
    if df is None or df.empty:
        empty_df = pd.DataFrame(columns=REQUIRED_COLS)
        for nc in NUMERIC_COLS:
            empty_df[nc] = empty_df[nc].astype(float)
        return empty_df

    clean_cols = []
    cols_to_keep = []
    for idx, c in enumerate(df.columns):
        c_str = str(c).strip()
        if c_str != "" and not c_str.lower().startswith("unnamed"):
            cols_to_keep.append(idx)
            clean_cols.append(c_str)

    df = df.iloc[:, cols_to_keep].copy()
    df.columns = clean_cols
    df = df.loc[:, ~df.columns.duplicated()].copy()

    col_map = {}
    used_targets = set(df.columns)

    for c in df.columns:
        cl = c.lower()
        target = None
        if ("klient" in cl or "nazwisko" in cl or "imie" in cl or "imię" in cl or cl == "dane"):
            target = "Klient"
        elif ("adres" in cl or "ulica" in cl):
            target = "Adres"
        elif ("kod" in cl and "poczt" in cl):
            target = "Kod_Pocztowy"
        elif ("miasto" in cl or "miejscowo" in cl):
            target = "Miasto"
        elif ("tel" in cl):
            target = "Telefon"
        elif ("opis_od_klienta" in cl or "opis od" in cl or "usterk" in cl):
            target = "Opis_Zgloszenia"
        elif (cl == "kwota"):
            target = "Kwota_Przewidywana"
        elif ("data_serwisu" in cl or "data ser" in cl):
            target = "Data_Planowana"
        elif ("platnosc" in cl or "płatność" in cl):
            target = "Forma_Platnosci"

        if target and target not in used_targets:
            col_map[c] = target
            used_targets.add(target)

    if col_map:
        df = df.rename(columns=col_map)

    df = df.loc[:, ~df.columns.duplicated()].copy()
    df = df.reset_index(drop=True)

    for rc in REQUIRED_COLS:
        if rc not in df.columns:
            if rc == "ID_Zgloszenia":
                df[rc] = [f"ZGL-{i+1001}" for i in range(len(df))]
            elif rc == "Status":
                df[rc] = "Nowe"
            elif rc in [
                "Priorytet_Krytyczny", "Wymagany_Przeglad", "Zdjecia_Potwierdzone",
                "Podpis_Zlozony", "Do_Oddzwonienia", "Potwierdzony_Tel",
                "Trudny_Klient", "Otrzymano_Oferte_Termin"
            ]:
                df[rc] = "Nie"
            elif rc in ["Naprawione_1_Wizyta", "Gotowka_Rozliczona"]:
                df[rc] = "Tak"
            elif rc == "Typ_Serwisu":
                df[rc] = "Płatny"
            elif rc == "Forma_Platnosci":
                df[rc] = "Przelew"
            elif rc == "Kwota_Przewidywana":
                df[rc] = 650.0
            elif rc in NUMERIC_COLS:
                df[rc] = 0.0
            else:
                df[rc] = ""

    extra_cols = [c for c in df.columns if c not in REQUIRED_COLS]
    df = df[REQUIRED_COLS + extra_cols].copy()

    for col in df.columns:
        if col in NUMERIC_COLS:
            default_val = 650.0 if col == "Kwota_Przewidywana" else 0.0
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(default_val).astype(float)
        else:
            df[col] = df[col].fillna("").astype(str).replace(["None", "nan", "NaT", "<NA>"], "")

    for idx, r in df.iterrows():
        adr = str(r["Adres"])
        if str(r["Kod_Pocztowy"]).strip() == "":
            m_zip = re.search(r'\b(\d{2}-\d{3})\b', adr)
            if m_zip:
                df.at[idx, "Kod_Pocztowy"] = m_zip.group(1)
        if str(r["Telefon"]).strip() == "":
            found_p = extract_phone_numbers("", str(r["Klient"]))
            if found_p:
                df.at[idx, "Telefon"] = " / ".join(found_p)
        st_val = str(r["Status"]).strip().lower()
        if "zako" in st_val:
            df.at[idx, "Status"] = "Zakończone"
        elif "częśc" in st_val or "czesc" in st_val:
            df.at[idx, "Status"] = "Do zamówienia części"
        elif "oddeleg" in st_val:
            df.at[idx, "Status"] = "Oddelegowane"
        elif st_val == "" or st_val == "brak":
            df.at[idx, "Status"] = "Nowe"

    return df.reset_index(drop=True)

def load_tickets():
    client = get_gsheet_client()
    df = pd.DataFrame()
    if client:
        try:
            sh = client.open(GOOGLE_SHEET_NAME)
            ws = sh.get_worksheet(0)
            raw_values = ws.get_all_values()
            if raw_values and len(raw_values) > 1:
                headers = raw_values[0]
                rows = raw_values[1:]
                df = pd.DataFrame(rows, columns=headers)
        except Exception:
            pass

    if df.empty and os.path.exists(LOCAL_EXCEL_FILE):
        try:
            df = pd.read_excel(LOCAL_EXCEL_FILE)
        except Exception:
            pass

    if df.empty:
        df = pd.DataFrame([
            {
                "ID_Zgloszenia": "ZGL-1001", "Klient": "Jan Kowalski", "Adres": "ul. Słoneczna 12",
                "Kod_Pocztowy": "05-800", "Miasto": "Pruszków", "Telefon": "500100200 / 600700800",
                "Typ_Instalacji": "Pompa Ciepła ARGOM", "Wymagany_Przeglad": "Tak",
                "Opis_Zgloszenia": "Brak grzania CO - błąd przepływu AL007 + wymiana glikolu i anody",
                "Priorytet_Krytyczny": "Tak", "Typ_Serwisu": "Płatny", "Serwisant": "Leszek",
                "Data_Planowana": str(datetime.date.today()), "Status": "Zaplanowane",
                "Naprawione_1_Wizyta": "Tak", "Kwota_Przewidywana": 1120.0, "Zgoda_Klienta_Kwota": 1120.0,
                "Kwota_Koncowa": 0.0, "Koszt_Materialow": 0.0, "Forma_Platnosci": "Gotówka pobrana",
                "Gotowka_Rozliczona": "Tak", "Czas_Pracy_Min": 0.0, "Wyniki_Przegladu": "",
                "Zdjecia_Potwierdzone": "Nie", "Zalecenia_Klienta": "", "Notatka_Serwisanta": "",
                "Data_Zamkniecia": "", "Podpis_Zlozony": "Nie", "Do_Oddzwonienia": "Nie",
                "Data_Oddzwonienia": "", "Potwierdzony_Tel": "Tak", "Trudny_Klient": "Nie",
                "Potrzebne_Czesci": "", "Lead_Handlowy": "", "Oddelegowano_Dzial": "",
                "Oddelegowano_Email": "", "Data_Oddelegowania": "", "Notatka_Dla_Dzialu": "",
                "Otrzymano_Oferte_Termin": "Nie", "Ostatnie_Przypomnienie": ""
            }
        ])

    return ensure_columns(df)

def save_tickets(df):
    clean_df = ensure_columns(df)
    st.session_state["df_tickets"] = clean_df.copy()
    try:
        clean_df.to_excel(LOCAL_EXCEL_FILE, index=False)
    except Exception:
        pass
    client = get_gsheet_client()
    if client:
        try:
            sh = client.open(GOOGLE_SHEET_NAME)
            ws = sh.get_worksheet(0)
            ws.clear()
            df_str = clean_df.fillna("").astype(str)
            ws.update([df_str.columns.values.tolist()] + df_str.values.tolist())
        except Exception as e:
            st.warning(f"Zapisano lokalnie. Synchronizacja z Google Sheets napotkała problem: {e}")

# ==========================================
# 4. WATCHDOG 3-DNIOWY DLA ODDELEGOWANYCH ZLECEŃ
# ==========================================
def check_and_send_3day_reminders(df, email_cfg):
    today = datetime.date.today()
    reminders_sent = 0
    overdue_ids = []
    changed = False

    for idx, r in df.iterrows():
        dzial = str(r.get("Oddelegowano_Dzial", "")).strip()
        odbiorca = str(r.get("Oddelegowano_Email", "")).strip()
        otrzymano = str(r.get("Otrzymano_Oferte_Termin", "Nie")).strip()
        data_odd_str = str(r.get("Data_Oddelegowania", "")).strip()[:10]
        ost_przyp_str = str(r.get("Ostatnie_Przypomnienie", "")).strip()[:10]

        if dzial and odbiorca and otrzymano != "Tak" and data_odd_str:
            try:
                d_odd = datetime.date.fromisoformat(data_odd_str)
                days_diff = (today - d_odd).days
                if days_diff >= 3:
                    overdue_ids.append(r["ID_Zgloszenia"])
                    should_send = True
                    if ost_przyp_str:
                        try:
                            d_ost = datetime.date.fromisoformat(ost_przyp_str)
                            if (today - d_ost).days < 3:
                                should_send = False
                        except Exception:
                            pass
                    if should_send and email_cfg.get("sender_email") and email_cfg.get("sender_password"):
                        subj = f"🚨 PRZYPOMNIENIE (Minęło {days_diff} dni): Brak oferty/terminu - {r['ID_Zgloszenia']} ({r['Klient']})"
                        body_lines = [
                            "Dzień dobry,",
                            "",
                            f"System PROMAG OZE przypomina o oddelegowanym zgłoszeniu z dnia {data_odd_str} ({dzial}), dla którego w systemie nadal nie odnotowano przygotowania oferty lub ustalenia terminu.",
                            "",
                            f"- ID Zgłoszenia: {r['ID_Zgloszenia']}",
                            f"- Klient: {r['Klient']} (Tel: {r['Telefon']})",
                            f"- Adres: {r['Kod_Pocztowy']} {r['Miasto']}, {r['Adres']}",
                            f"- Urządzenie: {r['Typ_Instalacji']}",
                            f"- Opis usterki: {r['Opis_Zgloszenia']}",
                            f"- Wytyczne od Koordynatora: {r['Notatka_Dla_Dzialu']}",
                            "",
                            "Prosimy o pilny kontakt z klientem i informację zwrotną do Koordynatora Serwisu."
                        ]
                        body = "\n".join(body_lines)
                        ok, _ = send_system_email(email_cfg, odbiorca, subj, body)
                        if ok:
                            df.at[idx, "Ostatnie_Przypomnienie"] = str(today)
                            reminders_sent += 1
                            changed = True
            except Exception:
                pass

    if changed:
        save_tickets(df)
    return overdue_ids, reminders_sent

def calculate_tech_metrics(df, tech_name):
    tech_df = df[df["Serwisant"] == tech_name]
    closed_df = tech_df[tech_df["Status"] == "Zakończone"].copy()
    total_closed = len(closed_df)
    ftfr_pct = round((len(closed_df[closed_df["Naprawione_1_Wizyta"] == "Tak"]) / total_closed) * 100.0, 1) if total_closed > 0 else 100.0

    paid_count = len(closed_df[closed_df["Typ_Serwisu"] == "Płatny"])
    warranty_count = len(closed_df[closed_df["Typ_Serwisu"] == "Gwarancyjny"])
    leads_count = len(tech_df[tech_df["Lead_Handlowy"].astype(str).str.strip() != ""])

    unsettled_cash_df = closed_df[(closed_df["Forma_Platnosci"] == "Gotówka pobrana") & (closed_df["Gotowka_Rozliczona"] != "Tak")]
    cash_in_hand = float(unsettled_cash_df["Kwota_Koncowa"].sum())

    current_ym = datetime.date.today().strftime("%Y-%m")
    closed_df["ym"] = closed_df["Data_Zamkniecia"].astype(str).str.slice(0, 7)
    monthly_df = closed_df[(closed_df["ym"] == current_ym) | (closed_df["ym"] == "")]
    monthly_revenue = float(monthly_df["Kwota_Koncowa"].sum())
    monthly_parts_cost = float(monthly_df["Koszt_Materialow"].sum())

    missing_tier_1 = max(0.0, COMMISSION_TIER_1 - monthly_revenue)
    missing_tier_2 = max(0.0, COMMISSION_TIER_2 - monthly_revenue)

    if monthly_revenue >= COMMISSION_TIER_2:
        commission_pln = monthly_revenue * COMMISSION_RATE_2
        active_rate_str = "4% (Próg II zdobyty!)"
    elif monthly_revenue >= COMMISSION_TIER_1:
        commission_pln = monthly_revenue * COMMISSION_RATE_1
        active_rate_str = "3% (Próg I zdobyty!)"
    else:
        commission_pln = 0.0
        active_rate_str = "0% (W drodze do Progu I)"

    return {
        "total_closed": total_closed,
        "ftfr_pct": ftfr_pct,
        "paid_count": paid_count,
        "warranty_count": warranty_count,
        "leads_count": leads_count,
        "cash_in_hand": cash_in_hand,
        "monthly_revenue": monthly_revenue,
        "monthly_parts_cost": monthly_parts_cost,
        "missing_tier_1": missing_tier_1,
        "missing_tier_2": missing_tier_2,
        "commission_pln": commission_pln,
        "active_rate_str": active_rate_str
    }

def get_low_stock_items(tech_profile):
    low = []
    for item in tech_profile.get("materialy", []):
        qty = int(item.get("ilosc", 0))
        if qty <= 1:
            low.append(f"{item.get('nazwa')} (zostało: {qty} szt.)")
    return low

# ==========================================
# 5. INICJALIZACJA SESJI I LOGOWANIE Z AUTO-PIN
# ==========================================
if "profiles" not in st.session_state or "logo_b64" not in st.session_state or "email_cfg" not in st.session_state:
    p_loaded, l_loaded, e_loaded = load_config_and_profiles()
    st.session_state["profiles"] = p_loaded
    st.session_state["logo_b64"] = l_loaded
    st.session_state["email_cfg"] = e_loaded

if "df_tickets" not in st.session_state:
    st.session_state["df_tickets"] = load_tickets()
else:
    st.session_state["df_tickets"] = ensure_columns(st.session_state["df_tickets"])

if "logged_role" not in st.session_state:
    st.session_state["logged_role"] = None
if "logged_user" not in st.session_state:
    st.session_state["logged_user"] = None
if "active_day" not in st.session_state:
    st.session_state["active_day"] = None
if "active_ticket_id" not in st.session_state:
    st.session_state["active_ticket_id"] = None

if st.session_state["logged_role"] is None:
    qp_user = st.query_params.get("user", None)
    qp_token = st.query_params.get("token", None)
    if qp_user == "Koordynator" and qp_token == PIN_KOORDYNATOR:
        st.session_state["logged_role"] = "koordynator"
        st.session_state["logged_user"] = "Koordynator"
    elif qp_user == "Zarząd" and qp_token == PIN_ZARZAD:
        st.session_state["logged_role"] = "zarzad"
        st.session_state["logged_user"] = "Zarząd"
    elif qp_user in st.session_state["profiles"]:
        expected_pin = str(st.session_state["profiles"][qp_user].get("pin", ""))
        if qp_token == expected_pin:
            st.session_state["logged_role"] = "serwisant"
            st.session_state["logged_user"] = qp_user

if st.session_state["logged_role"] is None:
    render_brand_logo(max_width=280, centered=True)
    st.markdown("<p style='text-align:center; color:#64748b;'>Mobilny System Zarządzania Serwisem OZE (Telefon / Tablet)</p>", unsafe_allow_html=True)
    st.divider()

    col_login, col_onboard = st.columns([1.2, 1])
    with col_login:
        st.subheader("🔐 Logowanie do systemu")
        tech_list = list(st.session_state["profiles"].keys())
        user_choice = st.selectbox(
            "Wybierz swój profil:",
            ["--- Wybierz z listy ---"] + tech_list + ["Koordynator Serwisu", "Zarząd (Raporty)"]
        )
        pin_input = st.text_input("Wpisz 4-cyfrowy kod PIN:", type="password", max_chars=8)
        remember_me = st.checkbox("🔒 Zapamiętaj mnie na tym urządzeniu (nie wylogowuj po nocy)", value=True)

        if st.button("🚀 Zaloguj się", type="primary"):
            if user_choice == "--- Wybierz z listy ---":
                st.error("Wybierz swój profil z listy powyżej.")
            elif user_choice == "Koordynator Serwisu":
                if pin_input == PIN_KOORDYNATOR:
                    st.session_state["logged_role"] = "koordynator"
                    st.session_state["logged_user"] = "Koordynator"
                    if remember_me:
                        st.query_params["user"] = "Koordynator"
                        st.query_params["token"] = PIN_KOORDYNATOR
                    st.rerun()
                else:
                    st.error("Nieprawidłowy kod PIN Koordynatora.")
            elif user_choice == "Zarząd (Raporty)":
                if pin_input == PIN_ZARZAD:
                    st.session_state["logged_role"] = "zarzad"
                    st.session_state["logged_user"] = "Zarząd"
                    if remember_me:
                        st.query_params["user"] = "Zarząd"
                        st.query_params["token"] = PIN_ZARZAD
                    st.rerun()
                else:
                    st.error("Nieprawidłowy kod PIN Zarządu.")
            elif user_choice in st.session_state["profiles"]:
                expected_pin = str(st.session_state["profiles"][user_choice].get("pin", ""))
                if pin_input == expected_pin:
                    st.session_state["logged_role"] = "serwisant"
                    st.session_state["logged_user"] = user_choice
                    st.session_state["active_day"] = None
                    st.session_state["active_ticket_id"] = None
                    if remember_me:
                        st.query_params["user"] = user_choice
                        st.query_params["token"] = expected_pin
                    st.rerun()
                else:
                    st.error("Nieprawidłowy kod PIN serwisanta.")

    with col_onboard:
        st.subheader("🎓 Standardy Pracy i Baza Wiedzy")
        st.info("Kody błędów Deye (F01-F64), GoodWe, Argom (AL001-AL174) oraz instrukcje Wi-Fi są dostępne bezpośrednio z ekranu każdego zlecenia.")
    st.stop()

# Nagłówek po zalogowaniu
top_c1, top_c2, top_c3 = st.columns([2, 2, 1])
with top_c1:
    render_brand_logo(max_width=200, centered=False)
with top_c2:
    st.markdown(f"**Zalogowany jako:** `{st.session_state['logged_user']}`")
with top_c3:
    if st.button("🚪 Wyloguj"):
        st.query_params.clear()
        st.session_state["logged_role"] = None
        st.session_state["logged_user"] = None
        st.session_state["active_day"] = None
        st.session_state["active_ticket_id"] = None
        st.rerun()

st.divider()

# ==========================================
# 6. PANEL SERWISANTA (Z MAGAZYNEM NA AUCIE, ALERTAMI <=1 SZT, FLOTĄ I LEADAMI)
# ==========================================
if st.session_state["logged_role"] == "serwisant":
    tech_name = st.session_state["logged_user"]
    df = st.session_state["df_tickets"]
    profiles = st.session_state["profiles"]
    my_prof = profiles[tech_name]
    metrics = calculate_tech_metrics(df, tech_name)

    # 1. ALERTY MAGAZYNOWE NA SAMOCHODZIE (<= 1 szt.)
    low_stock = get_low_stock_items(my_prof)
    if low_stock:
        st.markdown(f"""
        <div class="kpi-card-danger">
            <div style="font-size:15px; font-weight:800;">🚨 PILNE: UZUPEŁNIJ ZAPASY MATERIAŁOWE NA AUCIE!</div>
            <div style="font-size:14px; margin-top:4px;">Kończą się pozycje (<= 1 szt.): <b>{", ".join(low_stock)}</b></div>
        </div>
        """, unsafe_allow_html=True)

    # 2. KARTY KPI SERWISANTA
    m_col1, m_col2, m_col3 = st.columns([1.4, 1.4, 1])
    with m_col1:
        if metrics["missing_tier_1"] > 0:
            headline = f"Brakuje: {metrics['missing_tier_1']:,.0f} zł do prowizji 3%"
            subline = f"Próg I (15 000 zł) | Do 4% brakuje: {metrics['missing_tier_2']:,.0f} zł"
        elif metrics["missing_tier_2"] > 0:
            headline = f"🔥 Masz 3%! Brakuje: {metrics['missing_tier_2']:,.0f} zł do 4%"
            subline = f"Obrót: {metrics['monthly_revenue']:,.0f} zł (Premia: {metrics['commission_pln']:,.0f} zł)"
        else:
            headline = "🏆 Maksymalny próg 4% zdobyty!"
            subline = f"Obrót: {metrics['monthly_revenue']:,.0f} zł (Premia: {metrics['commission_pln']:,.0f} zł)"

        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-title">🎯 Cel Prowizyjny (Miesiąc)</div>
            <div class="kpi-value">{headline}</div>
            <div class="kpi-sub">{subline}</div>
        </div>
        """, unsafe_allow_html=True)

    with m_col2:
        st.markdown(f"""
        <div class="kpi-card kpi-card-warn">
            <div class="kpi-title">⚡ Skuteczność 1. Wizyty & Kasa</div>
            <div class="kpi-value">{metrics['ftfr_pct']}% | Gotówka przy sobie: {metrics['cash_in_hand']:,.0f} zł</div>
            <div class="kpi-sub">Obsłużone: {metrics['total_closed']} | Leady handlowe: {metrics['leads_count']}</div>
        </div>
        """, unsafe_allow_html=True)

    with m_col3:
        if st.button("🩺 KODY BŁĘDÓW / WI-FI"):
            st.session_state["show_kb"] = not st.session_state.get("show_kb", False)
        if st.button("🚐 MÓJ MAGAZYN / AUTO"):
            st.session_state["show_van"] = not st.session_state.get("show_van", False)

    if st.session_state.get("show_kb", False):
        with st.container(border=True):
            st.subheader("🩺 Interaktywna Baza Wiedzy (Deye / GoodWe / Argom / Magazyny)")
            kb_tab1, kb_tab2 = st.tabs(["🔍 Kody Błędów i Naprawa", "📶 Konfiguracja Wi-Fi i Konserwacja"])
            with kb_tab1:
                kb_cat = st.selectbox("Wybierz urządzenie:", list(DIAGNOSTYKA_BAZA.keys()), key="kb_cat_top")
                kb_err = st.selectbox("Wybierz kod błędu:", list(DIAGNOSTYKA_BAZA[kb_cat].keys()), key="kb_err_top")
                st.info(DIAGNOSTYKA_BAZA[kb_cat][kb_err])
            with kb_tab2:
                st.markdown(INSTRUKCJE_KONFIGURACJI_TEXT)

    if st.session_state.get("show_van", False):
        with st.container(border=True):
            st.subheader(f"🚐 Wyposażenie i Stan Magazynowy na Aucie ({tech_name})")
            v_c1, v_c2, v_c3 = st.columns([1.2, 1, 1])
            with v_c1:
                st.markdown("**📦 Materiały zużywalne na aucie:**")
                st.dataframe(pd.DataFrame(my_prof.get("materialy", []))[["nazwa", "ilosc", "cena_klienta"]], use_container_width=True, hide_index=True)
            with v_c2:
                st.markdown("**🧰 Przypisane narzędzia:**")
                st.dataframe(pd.DataFrame(my_prof.get("narzedzia", [])), use_container_width=True, hide_index=True)
            with v_c3:
                st.markdown("**🚗 Karta Floty (Stan licznika):**")
                flota = my_prof.get("flota", {})
                new_odo = st.number_input("Aktualny przebieg z licznika [km]:", min_value=0, step=100, value=int(flota.get("przebieg_km", 120000)))
                if st.button("💾 Zapisz stan licznika"):
                    my_prof["flota"]["przebieg_km"] = int(new_odo)
                    save_config_and_profiles(profiles)
                    st.success("Zaktualizowano przebieg pojazdu!")
                if int(flota.get("limit_olej_km", 130000)) - new_odo <= 1000:
                    st.error("🚨 Zbliża się wymiana oleju w busie (< 1000 km)!")

    st.divider()

    my_tickets = df[(df["Serwisant"] == tech_name) & (df["Status"].isin(["Zaplanowane", "W trakcie"]))].copy()

    if my_tickets.empty:
        st.success("🎉 Wszystkie zaplanowane dla Ciebie zlecenia zostały zrealizowane!")
        st.session_state["active_day"] = None
        st.session_state["active_ticket_id"] = None
    else:
        today_str = str(datetime.date.today())
        my_tickets["Data_Grupa"] = my_tickets["Data_Planowana"].astype(str).str.strip()
        my_tickets.loc[my_tickets["Data_Grupa"] == "", "Data_Grupa"] = today_str

        if st.session_state["active_day"] is None:
            st.subheader("📅 Twoje zaplanowane dni wyjazdowe (Trasy)")
            unique_days = sorted(my_tickets["Data_Grupa"].unique())
            for day_str in unique_days:
                day_cluster = my_tickets[my_tickets["Data_Grupa"] == day_str]
                pts_count = len(day_cluster)
                cities_list = ", ".join([m for m in day_cluster["Miasto"].astype(str).unique() if m.strip()]) or "Trasa lokalna"
                est_rev = float(day_cluster["Kwota_Przewidywana"].sum())
                crit_cnt = len(day_cluster[day_cluster["Priorytet_Krytyczny"] == "Tak"])

                # Inteligentna Checklista Załadunku rano na podstawie opisów zleceń z tego dnia
                opisy_dnia = " ".join(day_cluster["Opis_Zgloszenia"].astype(str).tolist()).lower()
                co_spakowac = []
                for m_item in my_prof.get("materialy", []):
                    klucz = str(m_item.get("slowo_klucz", "")).lower()
                    if klucz and klucz in opisy_dnia:
                        co_spakowac.append(f"{m_item['nazwa']} (masz na aucie: {m_item['ilosc']} szt.)")

                is_today = (day_str == today_str)
                card_class = "day-card day-card-today" if is_today else "day-card"
                badge_today = "🟢 DZISIAJ" if is_today else "🗓️ ZAPLANOWANA TRASA"

                with st.container(border=True):
                    c_day_info, c_day_btn = st.columns([2.2, 1])
                    with c_day_info:
                        st.markdown(f"""
                        <div class="{card_class}">
                            <div style="font-size:13px; color:#38bdf8; font-weight:700;">{badge_today}</div>
                            <div style="font-size:21px; font-weight:800;">📅 Data trasy: {day_str}</div>
                            <div style="font-size:15px; margin-top:4px;">📍 <b>Lokalizacje:</b> {cities_list}</div>
                            <div style="font-size:14px; color:#cbd5e1; margin-top:4px;">
                                🛠️ Punktów: <b>{pts_count}</b> | 💰 Szacowany obrót: <b>{est_rev:,.0f} zł</b> | 🔥 Krytyczne: <b>{crit_cnt}</b>
                            </div>
                        </div>
                        """, unsafe_allow_html=True)
                        if co_spakowac:
                            st.warning("📦 **Checklista załadunku przed wyjazdem (z opisów usterek):** " + ", ".join(co_spakowac))
                    with c_day_btn:
                        st.write("")
                        if st.button(f"🚀 ROZPOCZNIJ DZIEŃ ({day_str})", key=f"start_day_{day_str}", type="primary"):
                            st.session_state["active_day"] = day_str
                            st.session_state["active_ticket_id"] = None
                            st.rerun()

        elif st.session_state["active_day"] is not None and st.session_state["active_ticket_id"] is None:
            act_day = st.session_state["active_day"]
            day_tickets = my_tickets[my_tickets["Data_Grupa"] == act_day].copy()

            head_c1, head_c2 = st.columns([2.2, 1])
            with head_c1:
                st.subheader(f"🚚 Aktywny dzień pracy: {act_day}")
            with head_c2:
                if st.button("⬅️ Wróć do wyboru dnia"):
                    st.session_state["active_day"] = None
                    st.session_state["active_ticket_id"] = None
                    st.rerun()

            for idx_card, (t_idx, t_row) in enumerate(day_tickets.iterrows(), 1):
                t_id = t_row["ID_Zgloszenia"]
                is_crit = (str(t_row["Priorytet_Krytyczny"]) == "Tak")
                is_trudny = (str(t_row.get("Trudny_Klient", "Nie")) == "Tak")
                zgoda_kw = float(t_row.get("Zgoda_Klienta_Kwota", 0.0))

                with st.container(border=True):
                    if is_trudny:
                        st.error("🚩 **UWAGA: TRUDNY KLIENT / WYMAGANA PRZEDPŁATA LUB GOTÓWKA PRZED PRACĄ!**")
                    st.markdown(f"#### Punkt #{idx_card}: **{t_row['Klient']}** ({t_id}) {'🔥 AWARIA' if is_crit else ''}")
                    col_d1, col_d2, col_d3 = st.columns([1.4, 1.3, 1.1])
                    with col_d1:
                        st.markdown(f"📍 **Adres:** `{t_row['Kod_Pocztowy']} {t_row['Miasto']}, {t_row['Adres']}`")
                        st.markdown(f"🔧 **Urządzenie:** {t_row['Typ_Instalacji']} ({t_row['Typ_Serwisu']})")
                        st.markdown(f"📝 **Opis:** {t_row['Opis_Zgloszenia']}")
                        if zgoda_kw > 0:
                            st.success(f"✅ **Klient zaakceptował telefonicznie kwotę do: {zgoda_kw:,.0f} zł**")
                    with col_d2:
                        phones = extract_phone_numbers(str(t_row["Telefon"]), str(t_row["Klient"]))
                        for p_num in phones:
                            pc1, pc2 = st.columns(2)
                            with pc1:
                                st.link_button(f"📞 Zadzwoń ({p_num})", f"tel:{p_num}")
                            with pc2:
                                sms_txt = f"Dzien dobry, tu serwis PROMAG OZE. Dzisiaj realizujemy zgloszenie pod adresem {t_row['Adres']}. O ktorej godzinie bedzie Pan/Pani w domu?"
                                st.link_button(f"💬 SMS ({p_num})", f"sms:{p_num}?body={sms_txt.replace(' ', '%20')}")
                        maps_q = f"{t_row['Adres']} {t_row['Kod_Pocztowy']} {t_row['Miasto']}".replace(" ", "+")
                        st.link_button("🗺️ Nawiguj w Google Maps", f"https://www.google.com/maps/search/?api=1&query={maps_q}")
                    with col_d3:
                        st.markdown(f"💰 **Wycena wstępna:** `{float(t_row['Kwota_Przewidywana']):,.0f} zł`")
                        if st.button(f"🛠️ ROZPOCZNIJ ZLECENIE ({t_id})", key=f"start_t_{t_id}", type="primary"):
                            df.loc[t_idx, "Status"] = "W trakcie"
                            save_tickets(df)
                            st.session_state["active_ticket_id"] = t_id
                            st.rerun()

        elif st.session_state["active_ticket_id"] is not None:
            selected_id = st.session_state["active_ticket_id"]
            act_day = st.session_state["active_day"]
            matching_rows = df[df["ID_Zgloszenia"] == selected_id]
            if matching_rows.empty:
                st.session_state["active_ticket_id"] = None
                st.rerun()
            t_idx = matching_rows.index[0]
            t_row = df.loc[t_idx]

            nav_c1, nav_c2 = st.columns([2.2, 1])
            with nav_c1:
                st.subheader(f"🛠️ Obsługa zlecenia: {selected_id} ({t_row['Klient']})")
            with nav_c2:
                if st.button("⬅️ Wróć do listy klientów z tego dnia"):
                    st.session_state["active_ticket_id"] = None
                    st.rerun()

            if float(t_row.get("Zgoda_Klienta_Kwota", 0.0)) > 0:
                st.info(f"💡 **Ustalenia z biura:** Klient wyraził telefonicznie zgodę na koszt naprawy do kwoty **{float(t_row['Zgoda_Klienta_Kwota']):,.0f} zł**.")

            with st.expander("🩺 Ścieżka Diagnostyczna (Kody Błędów Deye / GoodWe / Argom)", expanded=False):
                diag_cat = st.selectbox("Wybierz sprzęt:", list(DIAGNOSTYKA_BAZA.keys()), key="diag_cat_t")
                diag_err = st.selectbox("Wybierz kod błędu:", list(DIAGNOSTYKA_BAZA[diag_cat].keys()), key="diag_err_t")
                st.warning(DIAGNOSTYKA_BAZA[diag_cat][diag_err])

            # 1. PROTOKÓŁ PRZEGLĄDU I POMIARÓW
            przeglad_summary_lines = []
            need_insp = (str(t_row.get("Wymagany_Przeglad", "Nie")) == "Tak" or "przegl" in str(t_row["Opis_Zgloszenia"]).lower())
            enable_inspection = st.checkbox("🔍 Wypełnij Protokół Przeglądu i Walidator Norm (Glikol / pH / Riso)", value=need_insp)
            if enable_inspection:
                with st.container(border=True):
                    insp_type = st.radio("Typ instalacji:", ["Pompa Ciepła / Solar (Hydraulika + Glikol)", "Fotowoltaika PV / Magazyn (DC/AC)"], horizontal=True)
                    if insp_type.startswith("Pompa"):
                        p_c1, p_c2 = st.columns(2)
                        with p_c1:
                            cisnienie_co = st.number_input("Ciśnienie statyczne [bar] (Norma: 1.2-2.2):", 0.0, 6.0, 1.5, 0.1)
                            delta_t = st.number_input("Delta T Zasilanie-Powrót [st.C] (Norma: 4-8):", 0.0, 30.0, 5.5, 0.5)
                            filtr_czystosc = st.selectbox("Filtr siatkowy i magnetyczny:", ["Wyczyszczono", "Silny szlam - wyczyszczono"])
                        with p_c2:
                            glikol_temp = st.number_input("Temp. krzepnięcia glikolu [st.C] (Norma: -15 do -28):", -50.0, 0.0, -20.0, 1.0)
                            glikol_ph = st.number_input("Odczyn pH zładu (Norma: 7.5-9.0):", 3.0, 12.0, 8.0, 0.2)
                        if glikol_temp > -15.0:
                            st.error("🚨 Za słaby glikol (> -15 st.C)! Ryzyko zamarznięcia wymiennika.")
                        if glikol_ph < 7.0:
                            st.error("🚨 Zakwaszony glikol (pH < 7.0)! Konieczna wymiana zładu.")
                        przeglad_summary_lines.append(f"PC/SOLAR: Ciśnienie={cisnienie_co}bar, DeltaT={delta_t}C, Glikol={glikol_temp}C, pH={glikol_ph}, Filtr={filtr_czystosc}")
                    else:
                        pv_c1, pv_c2 = st.columns(2)
                        with pv_c1:
                            riso_val = st.number_input("Rezystancja izolacji Riso [MOhm] (Norma >= 1.0):", 0.0, 500.0, 50.0, 0.5)
                            u_max_ac = st.number_input("Napięcie AC [V] (Norma 207-253V):", 180.0, 280.0, 238.0, 1.0)
                        with pv_c2:
                            uziemienie_val = st.number_input("Uziemienie [Ohm] (Norma <= 10):", 0.1, 100.0, 4.5, 0.5)
                        przeglad_summary_lines.append(f"PV: Riso={riso_val}MOhm, Vac={u_max_ac}V, Uziemienie={uziemienie_val}Ohm")

            st.divider()

            # 2. ZUŻYCIE MATERIAŁÓW Z BUSA (AUTOMATYCZNE ZEJŚCIE ZE STANU I DOLICZENIE DO RACHUNKU)
            st.subheader("🚐 Materiały zużyte z Twojego samochodu")
            st.caption("Wpisz w kolumnie **Zużyto_szt** ile sztuk materiału zużyłeś u klienta. Po zamknięciu zlecenia stan na busie automatycznie się zmniejszy, a cena doliczy do rachunku.")
            van_materials = my_prof.get("materialy", [])
            mat_df = pd.DataFrame(van_materials)
            if not mat_df.empty:
                mat_df["Zużyto_szt"] = 0
                edited_mat_df = st.data_editor(
                    mat_df[["nazwa", "ilosc", "cena_klienta", "koszt_zakupu", "Zużyto_szt"]],
                    column_config={
                        "nazwa": st.column_config.TextColumn("Materiał na aucie", disabled=True),
                        "ilosc": st.column_config.NumberColumn("Stan w busie (szt.)", disabled=True),
                        "cena_klienta": st.column_config.NumberColumn("Cena dla klienta (zł)", disabled=True),
                        "koszt_zakupu": st.column_config.NumberColumn("Koszt zakupu (zł)", disabled=True),
                        "Zużyto_szt": st.column_config.NumberColumn("Zużyto u klienta (szt.)", min_value=0, max_value=20, step=1)
                    },
                    use_container_width=True,
                    hide_index=True,
                    key=f"van_mat_use_{selected_id}"
                )
                used_materials_price = float((edited_mat_df["Zużyto_szt"] * edited_mat_df["cena_klienta"]).sum())
                used_materials_cost = float((edited_mat_df["Zużyto_szt"] * edited_mat_df["koszt_zakupu"]).sum())
                used_items_desc = [
                    f"{r['nazwa']} x{int(r['Zużyto_szt'])}"
                    for _, r in edited_mat_df.iterrows() if int(r["Zużyto_szt"]) > 0
                ]
            else:
                used_materials_price = 0.0
                used_materials_cost = 0.0
                used_items_desc = []

            st.divider()

            # 3. OBOWIĄZKOWE ZDJĘCIA
            st.subheader("📸 Obowiązkowa dokumentacja zdjęciowa (Tabliczka + Instalacja)")
            uploaded_photos = st.file_uploader("Załącz zdjęcia (min. 1 zdjęcie wymagane):", type=["png", "jpg", "jpeg"], accept_multiple_files=True, key=f"ph_{selected_id}")
            confirm_photos = st.checkbox("✅ Potwierdzam wykonanie i załączenie zdjęcia tabliczki znamionowej oraz instalacji.", key=f"cph_{selected_id}")

            st.divider()

            # 4. ROZLICZENIE, GOTÓWKA, CZĘŚCI DO ZAMÓWIENIA I LEAD HANDLOWY
            st.subheader("🧾 Rozliczenie wizyty, Gotówka i Przekazanie dalej")
            r_c1, r_c2 = st.columns(2)
            with r_c1:
                selected_services = st.multiselect("Wykonane usługi z cennika:", options=list(CENNIK_BAZOWY.keys()))
                extra_parts_manual = st.number_input("Inne części spoza magazynu busa (zł):", min_value=0.0, step=50.0, value=0.0)
                forma_plat = st.selectbox("Forma płatności:", ["Gotówka pobrana", "Przelew (Faktura)", "BLIK / Terminal", "Gwarancja (0 zł)"])
            with r_c2:
                work_minutes = st.number_input("Czas pracy na miejscu (minuty):", min_value=15, step=15, value=90)
                typ_serwisu_real = st.selectbox("Kwalifikacja:", ["Płatny", "Gwarancyjny"], index=0 if t_row["Typ_Serwisu"] != "Gwarancyjny" else 1)
                wynik_wizyty = st.radio(
                    "Status zakończenia prac:",
                    [
                        "✅ Naprawione podczas 1. wizyty (Zamknij zlecenie)",
                        "📦 Wymagane zamówienie części (Przekaż do Koordynatora)"
                    ]
                )

            potrzebne_czesci_txt = ""
            if wynik_wizyty.startswith("📦"):
                potrzebne_czesci_txt = st.text_input("🛠️ Wpisz dokładnie jaką część trzeba zamówić (model, S/N z tabliczki):", value=str(t_row.get("Potrzebne_Czesci", "")))

            czy_lead = st.checkbox("🔥 Klient zainteresowany ofertą handlową / rozbudową (Przekaż Lead do Handlowca)")
            lead_opis = ""
            if czy_lead:
                lead_opis = st.text_input("Czego dotyczy oferta dla klienta? (np. Magazyn energii 10kWh, nowa Pompa Ciepła):", value=str(t_row.get("Lead_Handlowy", "")))

            services_sum = sum(CENNIK_BAZOWY[s] for s in selected_services)
            if services_sum == 0 and typ_serwisu_real == "Płatny" and used_materials_price == 0:
                services_sum = float(t_row["Kwota_Przewidywana"])
            auto_total = 0.0 if (typ_serwisu_real == "Gwarancyjny" or forma_plat == "Gwarancja (0 zł)") else (services_sum + used_materials_price + extra_parts_manual)

            final_price = st.number_input("💰 Kwota końcowa do zapłaty przez klienta (zł brutto):", min_value=0.0, step=50.0, value=float(auto_total))
            internal_note = st.text_area("Notatka wewnętrzna dla Koordynatora:", value=str(t_row["Notatka_Serwisanta"]))

            # Protokół dla klienta + Podpis
            st.markdown('<div class="client-box">', unsafe_allow_html=True)
            render_brand_logo(max_width=200, centered=False)
            st.markdown("### 📄 Protokół Odbioru Serwisu")
            st.markdown(f"**Klient:** {t_row['Klient']} | **Adres:** {t_row['Kod_Pocztowy']} {t_row['Miasto']}, {t_row['Adres']}")
            if used_items_desc:
                st.markdown(f"**Zużyte materiały:** {', '.join(used_items_desc)}")
            client_recommendations = st.text_area("Zalecenia po-serwisowe dla Klienta:", value=str(t_row["Zalecenia_Klienta"]))
            st.markdown(f"#### Kwota rozliczenia: **{final_price:,.2f} zł** ({forma_plat})")
            canvas_result = st_canvas(
                fill_color="rgba(255, 255, 255, 0)", stroke_width=3, stroke_color="#0f172a",
                background_color="#ffffff", height=160, width=600, drawing_mode="freedraw", key=f"canv_{selected_id}"
            )
            st.markdown('</div>', unsafe_allow_html=True)

            if st.button("✅ ZATWIERDŹ WIZYTĘ I ZAKTUALIZUJ MAGAZYN BUSA", type="primary"):
                if not confirm_photos or not uploaded_photos:
                    st.error("🚨 BLOKADA: Załącz zdjęcie tabliczki znamionowej / instalacji i zaznacz potwierdzenie zdjęć!")
                elif wynik_wizyty.startswith("📦") and not potrzebne_czesci_txt.strip():
                    st.error("🚨 Wpisz dokładnie, jaką część należy zamówić!")
                else:
                    # Odejmowanie zużytych materiałów ze stanu na aucie
                    if not mat_df.empty:
                        for i_m, r_m in edited_mat_df.iterrows():
                            used_q = int(r_m["Zużyto_szt"])
                            if used_q > 0:
                                my_prof["materialy"][i_m]["ilosc"] = max(0, int(my_prof["materialy"][i_m]["ilosc"]) - used_q)
                        save_config_and_profiles(profiles)

                    if wynik_wizyty.startswith("📦"):
                        df.loc[t_idx, "Status"] = "Do zamówienia części"
                        df.loc[t_idx, "Naprawione_1_Wizyta"] = "Nie"
                        df.loc[t_idx, "Potrzebne_Czesci"] = potrzebne_czesci_txt.strip()
                    else:
                        df.loc[t_idx, "Status"] = "Zakończone"
                        df.loc[t_idx, "Naprawione_1_Wizyta"] = "Tak"

                    df.loc[t_idx, "Typ_Serwisu"] = typ_serwisu_real
                    df.loc[t_idx, "Forma_Platnosci"] = forma_plat
                    df.loc[t_idx, "Gotowka_Rozliczona"] = "Nie" if (forma_plat == "Gotówka pobrana" and final_price > 0) else "Tak"
                    df.loc[t_idx, "Kwota_Koncowa"] = float(final_price)
                    df.loc[t_idx, "Koszt_Materialow"] = float(used_materials_cost)
                    df.loc[t_idx, "Czas_Pracy_Min"] = float(work_minutes)
                    df.loc[t_idx, "Wyniki_Przegladu"] = "; ".join(przeglad_summary_lines + used_items_desc)
                    df.loc[t_idx, "Zdjecia_Potwierdzone"] = f"Tak ({len(uploaded_photos)} zdj.)"
                    df.loc[t_idx, "Zalecenia_Klienta"] = str(client_recommendations)
                    df.loc[t_idx, "Notatka_Serwisanta"] = str(internal_note)
                    df.loc[t_idx, "Lead_Handlowy"] = str(lead_opis).strip() if czy_lead else str(t_row.get("Lead_Handlowy", ""))
                    df.loc[t_idx, "Data_Zamkniecia"] = str(datetime.date.today())
                    df.loc[t_idx, "Podpis_Zlozony"] = "Tak"

                    save_tickets(df)
                    st.session_state["active_ticket_id"] = None
                    st.success("Zapisano wizytę i zaktualizowano stan magazynowy na samochodzie!")
                    st.rerun()

# ==========================================
# 7. PANEL KOORDYNATORA
# ==========================================
elif st.session_state["logged_role"] == "koordynator":
    df = st.session_state["df_tickets"]
    profiles = st.session_state["profiles"]
    email_cfg = st.session_state["email_cfg"]

    # 1. AUTOMATYCZNY WATCHDOG 3-DNIOWY (PRZYPOMNIENIA EMAIL)
    overdue_delegated, sent_cnt = check_and_send_3day_reminders(df, email_cfg)
    if sent_cnt > 0:
        st.toast(f"📧 Wysłano automatyczne przypomnienia e-mail dla {sent_cnt} przeterminowanych zleceń (>3 dni)!", icon="📧")
    if overdue_delegated:
        st.error(f"🚨 **ALARM ODDELEGOWANIA (>3 DNI BEZ OFERTY/TERMINU):** Zgłoszenia `{', '.join(overdue_delegated)}` nie mają potwierdzonej oferty lub terminu od działu!")

    # 2. ALERTY NISKICH STANÓW MAGAZYNOWYCH NA BUSACH SERWISANTÓW
    for t_n, t_p in profiles.items():
        low_s = get_low_stock_items(t_p)
        if low_s:
            st.warning(f"🚐 **Braki na aucie serwisanta {t_n} (<=1 szt.):** {', '.join(low_s)}")

    # 3. SYSTEM 15:30 - LISTA KLIENTÓW DO ODDZWONIENIA (WISI DO SKUTKU!)
    now_time = datetime.datetime.now().time()
    today_iso = str(datetime.date.today())
    callbacks_df = df[df["Do_Oddzwonienia"] == "Tak"].copy()

    is_after_1530 = (now_time >= datetime.time(15, 30))
    if not callbacks_df.empty:
        header_cb = "🔔 KLIENCI DO ODDZWONIENIA (PO 15:30 - OBOWIĄZKOWA LISTA DO SKUTKU!)" if is_after_1530 else f"📞 Klienci oznaczeni do oddzwonienia ({len(callbacks_df)} poz. - główny blok po 15:30)"
        with st.expander(header_cb, expanded=(is_after_1530 or len(callbacks_df) > 0)):
            for cb_idx, cb_row in callbacks_df.iterrows():
                cb_id = cb_row["ID_Zgloszenia"]
                cb_phones = extract_phone_numbers(str(cb_row["Telefon"]), str(cb_row["Klient"]))
                cb_phone_main = cb_phones[0] if cb_phones else ""
                ost_wpis = cb_row["Notatka_Serwisanta"] or cb_row["Opis_Zgloszenia"]

                with st.container(border=True):
                    c_cb1, c_cb2, c_cb3 = st.columns([1.6, 1.2, 1.2])
                    with c_cb1:
                        st.markdown(f"🔴 **{cb_row['Klient']}** (`{cb_id}`) – 📅 Termin kontaktu: **{cb_row['Data_Oddzwonienia'] or today_iso}**")
                        st.markdown(f"📍 `{cb_row['Kod_Pocztowy']} {cb_row['Miasto']}, {cb_row['Adres']}` | 🔧 {cb_row['Typ_Instalacji']}")
                        st.markdown(f"💬 **Ostatni wpis:** *{ost_wpis}*")
                    with c_cb2:
                        if cb_phone_main:
                            st.link_button(f"📞 Zadzwoń teraz: {cb_phone_main}", f"tel:{cb_phone_main}")
                            sms_missed = "Dzien dobry, kontaktowalismy sie z Dzialu Serwisu PROMAG OZE w sprawie zgloszenia serwisowego. Prosimy o kontakt zwrotny lub informacje, w jakich godzinach mozemy zadzwonic."
                            st.link_button("📲 Wyślij SMS (Nie odbiera)", f"sms:{cb_phone_main}?body={sms_missed.replace(' ', '%20')}")
                        nowa_notatka_cb = st.text_input("Notatka z rozmowy:", key=f"cb_note_{cb_id}", placeholder="Ustalenia z klientem...")
                    with c_cb3:
                        if st.button("✅ Załatwione (Zdejmij alert)", key=f"cb_done_{cb_id}", type="primary"):
                            df.loc[cb_idx, "Do_Oddzwonienia"] = "Nie"
                            df.loc[cb_idx, "Potwierdzony_Tel"] = "Tak"
                            if nowa_notatka_cb.strip():
                                df.loc[cb_idx, "Notatka_Serwisanta"] = f"[{today_iso} Tel OK]: {nowa_notatka_cb.strip()} | " + str(df.loc[cb_idx, "Notatka_Serwisanta"])
                            save_tickets(df)
                            st.rerun()
                        if st.button("⏳ Nie odbiera (Zostaw do skutku)", key=f"cb_wait_{cb_id}"):
                            stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
                            df.loc[cb_idx, "Notatka_Serwisanta"] = f"[Próba kontaktu {stamp} - nie odbiera] " + str(df.loc[cb_idx, "Notatka_Serwisanta"])
                            save_tickets(df)
                            st.rerun()

    # ZAKŁADKI KOORDYNATORA - START OD ZAKŁADKI Z KLIENTAMI!
    tab_klienci, tab_plan, tab_oddeleguj, tab_kadry, tab_ustawienia = st.tabs([
        "📋 1. Klienci i Zgłoszenia (Start + Kalkulator)",
        "🗺️ 2. Planer Korytarzowy i Czysty Zysk (1200 zł)",
        "📤 3. Oddelegowania, Części i Leady",
        "👷 4. Kadry, Narzędzia, Magazyn Busa i Gotówka",
        "⚙️ 5. Ustawienia E-mail (SMTP), Logo i Excel"
    ])

    # ----------------------------------------------------
    # ZAKŁADKA 1: KLIENCI, KALKULATOR WYCENY, DODAWANIE Z HISTORIĄ ADRESU
    # ----------------------------------------------------
    with tab_klienci:
        col_new_t, col_calc_t = st.columns([1.3, 1])

        with col_calc_t:
            with st.container(border=True):
                st.subheader("🧮 Kalkulator Szybkiej Wyceny (Na telefon)")
                st.caption("Zaznacz usługi i materiały podczas rozmowy z klientem, aby od razu podać mu widełki cenowe i wpisać kwotę zgody.")
                calc_srv = st.multiselect("Usługi z cennika:", list(CENNIK_BAZOWY.keys()), key="coord_calc_srv")
                calc_mat = st.multiselect("Materiały z katalogu:", [f"{m['nazwa']} ({m['cena_klienta']} zł)" for m in DOMYSLNE_MATERIALY], key="coord_calc_mat")
                calc_extra_km = st.number_input("Dojazd ponad ryczałt [km]:", 0, 500, 0, 10, key="coord_calc_km")

                sum_srv_c = sum(CENNIK_BAZOWY[s] for s in calc_srv)
                sum_mat_c = 0.0
                for cm in calc_mat:
                    for dm in DOMYSLNE_MATERIALY:
                        if cm.startswith(dm["nazwa"]):
                            sum_mat_c += dm["cena_klienta"]
                total_quote_calc = sum_srv_c + sum_mat_c + (calc_extra_km * 2.5)
                st.markdown(f"### 💰 Wyliczona kwota dla klienta: **{total_quote_calc:,.0f} zł brutto**")

        with col_new_t:
            with st.container(border=True):
                st.subheader("➕ Wprowadź nowe zgłoszenie klienta")
                q_check_tel = st.text_input("🔍 Wpisz telefon lub fragment adresu (sprawdzam historię wizyt i czarną listę):", placeholder="np. 500100200 lub Słoneczna")
                if len(q_check_tel.strip()) >= 4:
                    q_clean = q_check_tel.strip()
                    hist_matches = df[
                        df["Telefon"].astype(str).str.contains(q_clean, case=False, na=False) |
                        df["Adres"].astype(str).str.contains(q_clean, case=False, na=False) |
                        df["Klient"].astype(str).str.contains(q_clean, case=False, na=False)
                    ]
                    if not hist_matches.empty:
                        st.warning(f"⚠️ **HISTORIA ADRESU / KLIENTA:** Znaleziono {len(hist_matches)} wcześniejszych zgłoszeń w bazie!")
                        for _, hr in hist_matches.tail(3).iterrows():
                            trudny_b = " 🚩 [TRUDNY KLIENT]" if hr.get("Trudny_Klient") == "Tak" else ""
                            st.caption(f"• `{hr['ID_Zgloszenia']}` | Data: {hr['Data_Planowana'] or hr['Data_Zamkniecia']} | Serwisant: **{hr['Serwisant']}** | Kwota: {hr['Kwota_Koncowa'] or hr['Kwota_Przewidywana']} zł | Opis: *{hr['Opis_Zgloszenia']}*{trudny_b}")

                with st.form("form_add_ticket_coord"):
                    fa1, fa2 = st.columns(2)
                    with fa1:
                        n_klient = st.text_input("Imię i nazwisko klienta:")
                        n_tel = st.text_input("Telefon (można 2 po /):", value=q_check_tel if re.search(r'\d{6}', q_check_tel) else "")
                        n_kod = st.text_input("Kod pocztowy (XX-XXX):", placeholder="05-800")
                        n_miasto = st.text_input("Miejscowość:")
                        n_adres = st.text_input("Ulica i nr:")
                    with fa2:
                        n_typ_inst = st.selectbox("Urządzenie:", [
                            "Pompa Ciepła ARGOM", "Falownik Deye (PV / Hybryda)", "Falownik GoodWe (PV / Hybryda)",
                            "Kolektory Słoneczne (Solar)", "Pompa Ciepła + PV", "Magazyn Energii", "Inne"
                        ])
                        n_przeglad = st.selectbox("Wymagany pełny przegląd z protokołem?", ["Nie", "Tak"])
                        n_kryt = st.selectbox("Priorytet krytyczny (Awaria)?", ["Nie", "Tak"])
                        n_trudny = st.selectbox("🚩 Trudny klient / Wymagana przedpłata?", ["Nie", "Tak"])
                        n_typ_serw = st.selectbox("Rodzaj serwisu:", ["Płatny", "Gwarancyjny"])

                    fb1, fb2, fb3, fb4 = st.columns(4)
                    with fb1:
                        n_kwota = st.number_input("Szacowana kwota (zł):", 0.0, 50000.0, float(total_quote_calc) if total_quote_calc > 0 else 650.0, 50.0)
                    with fb2:
                        n_zgoda = st.number_input("✅ Zgoda klienta do kwoty (zł):", 0.0, 50000.0, float(total_quote_calc) if total_quote_calc > 0 else 0.0, 50.0)
                    with fb3:
                        n_oddzwonic = st.checkbox("📞 Oddzwonić do klienta?", value=False)
                    with fb4:
                        n_data_oddzw = st.date_input("Data oddzwonienia:", value=datetime.date.today())

                    n_opis = st.text_area("Opis usterki / ustalenia:")

                    if st.form_submit_button("➕ Zapisz zgłoszenie w bazie", type="primary"):
                        new_id = f"ZGL-{len(df) + 1001}"
                        new_row = {
                            "ID_Zgloszenia": new_id, "Klient": n_klient, "Adres": n_adres,
                            "Kod_Pocztowy": n_kod, "Miasto": n_miasto, "Telefon": n_tel,
                            "Typ_Instalacji": n_typ_inst, "Wymagany_Przeglad": n_przeglad,
                            "Opis_Zgloszenia": n_opis, "Priorytet_Krytyczny": n_kryt,
                            "Typ_Serwisu": n_typ_serw, "Serwisant": "", "Data_Planowana": "",
                            "Status": "Nowe", "Naprawione_1_Wizyta": "Tak",
                            "Kwota_Przewidywana": float(n_kwota), "Zgoda_Klienta_Kwota": float(n_zgoda),
                            "Kwota_Koncowa": 0.0, "Koszt_Materialow": 0.0, "Forma_Platnosci": "Przelew",
                            "Gotowka_Rozliczona": "Tak", "Czas_Pracy_Min": 0.0, "Wyniki_Przegladu": "",
                            "Zdjecia_Potwierdzone": "Nie", "Zalecenia_Klienta": "", "Notatka_Serwisanta": "",
                            "Data_Zamkniecia": "", "Podpis_Zlozony": "Nie",
                            "Do_Oddzwonienia": "Tak" if n_oddzwonic else "Nie",
                            "Data_Oddzwonienia": str(n_data_oddzw) if n_oddzwonic else "",
                            "Potwierdzony_Tel": "Nie" if n_oddzwonic else "Tak",
                            "Trudny_Klient": n_trudny,
                            "Potrzebne_Czesci": "", "Lead_Handlowy": "", "Oddelegowano_Dzial": "",
                            "Oddelegowano_Email": "", "Data_Oddelegowania": "", "Notatka_Dla_Dzialu": "",
                            "Otrzymano_Oferte_Termin": "Nie", "Ostatnie_Przypomnienie": ""
                        }
                        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                        save_tickets(df)
                        st.success(f"Dodano zgłoszenie {new_id}!")
                        st.rerun()

        st.divider()

        # Szybka edycja i Kopalnia Zleceń
        sub_t1, sub_t2 = st.tabs(["📋 Pełna Baza Zgłoszeń i Szybka Edycja", "💰 Kopalnia Zleceń (Minął rok od przeglądu/serwisu)"])
        with sub_t1:
            f_c1, f_c2, f_c3 = st.columns(3)
            with f_c1:
                f_stat = st.multiselect("Status:", ["Nowe", "Zaplanowane", "W trakcie", "Do zamówienia części", "Oddelegowane", "Zakończone"], default=["Nowe", "Zaplanowane", "W trakcie", "Do zamówienia części", "Oddelegowane"])
            with f_c2:
                f_tech = st.multiselect("Serwisant:", [""] + list(profiles.keys()))
            with f_c3:
                f_q = st.text_input("Szukaj w bazie:")

            v_df = df.copy()
            if f_stat:
                v_df = v_df[v_df["Status"].isin(f_stat)]
            if f_tech:
                v_df = v_df[v_df["Serwisant"].isin(f_tech)]
            if f_q:
                m_q = (
                    v_df["Klient"].astype(str).str.contains(f_q, case=False, na=False) |
                    v_df["Miasto"].astype(str).str.contains(f_q, case=False, na=False) |
                    v_df["Adres"].astype(str).str.contains(f_q, case=False, na=False) |
                    v_df["Kod_Pocztowy"].astype(str).str.contains(f_q, case=False, na=False)
                )
                v_df = v_df[m_q]

            st.dataframe(v_df, use_container_width=True, hide_index=True)

            with st.expander("✏️ Edytuj wybrane zgłoszenie (Zmień status / Włącz Oddzwonić)"):
                if not df.empty:
                    e_opts = [f"{r['ID_Zgloszenia']} | {r['Klient']} ({r['Miasto']})" for _, r in df.iterrows()]
                    s_ed = st.selectbox("Wybierz zgłoszenie:", e_opts)
                    s_ed_id = s_ed.split(" | ")[0]
                    e_i = df[df["ID_Zgloszenia"] == s_ed_id].index[0]
                    er = df.loc[e_i]
                    ec1, ec2, ec3, ec4 = st.columns(4)
                    with ec1:
                        ed_stat = st.selectbox("Status:", ["Nowe", "Zaplanowane", "W trakcie", "Do zamówienia części", "Oddelegowane", "Zakończone"], index=0)
                        ed_tech = st.selectbox("Serwisant:", [""] + list(profiles.keys()))
                    with ec2:
                        ed_oddzw = st.selectbox("Do oddzwonienia?", ["Nie", "Tak"], index=1 if er["Do_Oddzwonienia"] == "Tak" else 0)
                        ed_data_oddzw = st.text_input("Data oddzwonienia:", value=str(er["Data_Oddzwonienia"] or today_iso))
                    with ec3:
                        ed_zgoda = st.number_input("Zgoda klienta do (zł):", 0.0, 50000.0, float(er["Zgoda_Klienta_Kwota"]), 50.0)
                        ed_trudny = st.selectbox("Trudny klient?", ["Nie", "Tak"], index=1 if er["Trudny_Klient"] == "Tak" else 0)
                    with ec4:
                        ed_kwota = st.number_input("Kwota przewidywana:", 0.0, 50000.0, float(er["Kwota_Przewidywana"]), 50.0)
                    ed_opis = st.text_area("Opis zgłoszenia:", value=str(er["Opis_Zgloszenia"]))
                    if st.button("💾 Zapisz zmiany w zgłoszeniu"):
                        df.loc[e_i, "Status"] = ed_stat
                        if ed_tech:
                            df.loc[e_i, "Serwisant"] = ed_tech
                        df.loc[e_i, "Do_Oddzwonienia"] = ed_oddzw
                        df.loc[e_i, "Data_Oddzwonienia"] = ed_data_oddzw
                        df.loc[e_i, "Zgoda_Klienta_Kwota"] = float(ed_zgoda)
                        df.loc[e_i, "Trudny_Klient"] = ed_trudny
                        df.loc[e_i, "Kwota_Przewidywana"] = float(ed_kwota)
                        df.loc[e_i, "Opis_Zgloszenia"] = ed_opis
                        save_tickets(df)
                        st.success("Zapisano!")
                        st.rerun()

        with sub_t2:
            st.markdown("#### 💰 Klienci z bazy historycznej do corocznego przeglądu / wymiany glikolu")
            old_closed = df[df["Status"] == "Zakończone"].copy()
            if not old_closed.empty:
                st.dataframe(old_closed[["ID_Zgloszenia", "Klient", "Telefon", "Kod_Pocztowy", "Miasto", "Adres", "Typ_Instalacji", "Data_Planowana", "Kwota_Przewidywana"]].head(50), use_container_width=True, hide_index=True)

    # ----------------------------------------------------
    # ZAKŁADKA 2: PLANER KORYTARZOWY I KALKULATOR CZYSTEGO ZYSKU (1200 ZŁ NETTO)
    # ----------------------------------------------------
    with tab_plan:
        st.subheader("🗺️ Planer Korytarzowy („Po Drodze”) + Kalkulator Czystego Zysku (Cel: min. 1200 zł na czysto)")
        st.markdown(
            f"**Parametry kosztowe w algorytmie:** Godzina pracy serwisanta = **`{DEFAULT_HOURLY_RATE:.0f} zł/h`** | "
            f"Paliwo = **`{DEFAULT_FUEL_PRICE:.2f} zł/l`** | Średni czas u 1 klienta = **`{DEFAULT_ONSITE_HOURS} h` (105 zł)** | "
            f"Współczynnik krętości dróg = **`{ROAD_TORTUOSITY}x GPS`**."
        )

        pc1, pc2, pc3, pc4, pc5 = st.columns(5)
        with pc1:
            sel_tech_plan = st.selectbox("Serwisant:", list(profiles.keys()), key="cor_tech")
        with pc2:
            start_postal = st.text_input("Punkt Startowy (Baza):", value="05-552", help="Domyślnie 05-552 Łazy / Pruszków")
        with pc3:
            target_postal = st.text_input("Kod Docelowy (Cel):", value="05-800")
        with pc4:
            search_mode = st.selectbox("Tryb szukania:", ["Korytarz trasy (Po drodze + Okolica celu)", "Koło wokół kodu docelowego"])
        with pc5:
            corridor_width_km = st.slider("Szerokość korytarza / Promień [km]:", 10, 150, 40, 5)

        p_c_extra1, p_c_extra2, p_c_extra3 = st.columns(3)
        with p_c_extra1:
            plan_date = st.date_input("Data wyjazdu:", value=datetime.date.today() + datetime.timedelta(days=1))
        with p_c_extra2:
            burn_rate = st.number_input("Spalanie busa [l/100km]:", 5.0, 20.0, DEFAULT_BURN_L_100KM, 0.5)
        with p_c_extra3:
            only_confirmed = st.checkbox("Tylko potwierdzone telefonicznie", value=False)

        start_lat, start_lon = get_coords_for_postal(start_postal)
        dest_lat, dest_lon = get_coords_for_postal(target_postal)

        pending_df = df[df["Status"].isin(["Nowe", "Do zaplanowania", ""])].copy()
        if only_confirmed:
            pending_df = pending_df[pending_df["Potwierdzony_Tel"] == "Tak"]

        if not pending_df.empty:
            dist_from_start = []
            dist_from_dest = []
            dist_from_corridor = []
            coords_list = []

            for _, r in pending_df.iterrows():
                rlat, rlon = extract_coords_from_row(r)
                coords_list.append((rlat, rlon))
                d_s = haversine_km(start_lat, start_lon, rlat, rlon)
                d_d = haversine_km(dest_lat, dest_lon, rlat, rlon)
                d_c = distance_from_segment_km(rlat, rlon, start_lat, start_lon, dest_lat, dest_lon)
                dist_from_start.append(round(d_s * ROAD_TORTUOSITY, 1))
                dist_from_dest.append(round(d_d, 1))
                dist_from_corridor.append(round(d_c, 1))

            pending_df["Km_od_Bazy"] = dist_from_start
            pending_df["Km_od_Celu_GPS"] = dist_from_dest
            pending_df["Km_od_Osi_Trasy"] = dist_from_corridor
            pending_df["_coords"] = coords_list

            if search_mode.startswith("Korytarz"):
                filtered_plan = pending_df[
                    (pending_df["Km_od_Osi_Trasy"] <= corridor_width_km) |
                    (pending_df["Km_od_Celu_GPS"] <= corridor_width_km)
                ].sort_values(by=["Priorytet_Krytyczny", "Km_od_Osi_Trasy"], ascending=[False, True]).copy()
            else:
                filtered_plan = pending_df[pending_df["Km_od_Celu_GPS"] <= corridor_width_km].sort_values(by="Km_od_Celu_GPS").copy()

            st.markdown(f"#### 📍 Znalezione zlecenia na trasie `{start_postal} ➔ {target_postal}` ({len(filtered_plan)} poz.):")
            st.caption("Zaznacz zlecenia w kolumnie **Wybierz**, a kalkulator poniżej na żywo policzy pełną pętlę kilometrów, paliwo, roboczogodziny i **Czysty Zysk**.")

            if not filtered_plan.empty:
                disp_cols = [
                    "ID_Zgloszenia", "Km_od_Osi_Trasy", "Km_od_Bazy", "Klient", "Kod_Pocztowy",
                    "Miasto", "Adres", "Typ_Instalacji", "Priorytet_Krytyczny", "Kwota_Przewidywana", "Opis_Zgloszenia"
                ]
                editor_df = filtered_plan[disp_cols].copy()
                editor_df.insert(0, "Wybierz", False)

                edited_route = st.data_editor(
                    editor_df,
                    use_container_width=True,
                    hide_index=True,
                    key="corridor_editor"
                )

                chosen_ids = edited_route[edited_route["Wybierz"] == True]["ID_Zgloszenia"].tolist()
                chosen_rows = filtered_plan[filtered_plan["ID_Zgloszenia"].isin(chosen_ids)]

                chosen_coords = chosen_rows["_coords"].tolist()
                gross_rev = float(chosen_rows["Kwota_Przewidywana"].sum())

                econ = estimate_route_economics(
                    chosen_coords, (start_lat, start_lon), gross_rev, parts_cost=0.0,
                    hourly_rate=DEFAULT_HOURLY_RATE, fuel_price=DEFAULT_FUEL_PRICE,
                    burn_l_100km=burn_rate, onsite_hours_per_client=DEFAULT_ONSITE_HOURS
                )

                # Baner ekonomiki trasy na żywo
                st.markdown("### 📊 Kalkulacja Rentowności Wybranej Trasy (Na Żywo)")
                ek1, ek2, ek3, ek4, ek5 = st.columns(5)
                ek1.metric("Dystans pętli (z powrotem)", f"{econ['road_km']} km", f"Jazda: {econ['drive_hours']} h")
                ek2.metric("Przychód ze zleceń", f"{gross_rev:,.0f} zł", f"Punkty: {len(chosen_ids)}")
                ek3.metric("Koszt paliwa (8 zł/l)", f"-{econ['fuel_cost']:,.0f} zł", f"Spalanie: {burn_rate}l/100km")
                ek4.metric("Koszt pracy (Jazda + Klient)", f"-{(econ['labor_drive_cost'] + econ['labor_onsite_cost']):,.0f} zł", f"Łącznie: {econ['drive_hours'] + econ['onsite_hours']:.1f} h x 70 zł")
                ek5.metric("💎 CZYSTY ZYSK (Cel >=1200 zł)", f"{econ['net_profit']:,.0f} zł", delta=f"{econ['net_profit'] - TARGET_NET_PROFIT:,.0f} zł vs cel")

                if len(chosen_ids) > 0:
                    if econ["net_profit"] >= TARGET_NET_PROFIT:
                        st.success(f"✅ **TRASA RENTOWNA!** Czysty zysk po odjęciu paliwa i czasu pracy wynosi **{econ['net_profit']:,.2f} zł** (przekracza próg 1200 zł).")
                    else:
                        brakuje_netto = TARGET_NET_PROFIT - econ["net_profit"]
                        st.warning(f"⚠️ **UWAGA:** Do założonego czystego zysku 1200 zł brakuje jeszcze **{brakuje_netto:,.0f} zł**. Dobierz jeszcze 1 zlecenie po drodze lub dosprzedaj przegląd/czyszczenie.")

                    # Podpowiedź Hotel vs Powrót przy dalekich trasach (> 220 km w jedną stronę)
                    if econ["road_km"] >= 420:
                        return_cost = (econ["fuel_cost"] / 2.0) + (econ["labor_drive_cost"] / 2.0)
                        st.info(f"🏨 **Kalkulator Hotel vs. Powrót:** Trasa ma {econ['road_km']} km. Sam powrót do bazy i ponowny dojazd jutro kosztuje ok. **{return_cost*2:,.0f} zł** (paliwo + godziny jazdy). Nocleg serwisanta w hotelu (~200 zł) pozwoli zaoszczędzić ok. **{max(0, return_cost*2 - 200):,.0f} zł**!")

                    if st.button(f"🚀 ZATWIERDŹ I WYŚLIJ TRASĘ ({len(chosen_ids)} pkt) DO: {sel_tech_plan} NA DZIEŃ {plan_date}", type="primary"):
                        for cid in chosen_ids:
                            idx_m = df[df["ID_Zgloszenia"] == cid].index[0]
                            df.loc[idx_m, "Serwisant"] = sel_tech_plan
                            df.loc[idx_m, "Data_Planowana"] = str(plan_date)
                            df.loc[idx_m, "Status"] = "Zaplanowane"
                        save_tickets(df)
                        st.success(f"Przypisano {len(chosen_ids)} zleceń dla {sel_tech_plan} na {plan_date}!")
                        st.rerun()
            else:
                st.warning("Brak zleceń w zadanym korytarzu/promieniu. Zwiększ szerokość korytarza suwakiem powyżej.")

    # ----------------------------------------------------
    # ZAKŁADKA 3: ODDELEGOWANIA DO DZIAŁÓW (E-MAIL + 3 DNI WATCHDOG), CZĘŚCI I LEADY
    # ----------------------------------------------------
    with tab_oddeleguj:
        st.subheader("📤 Oddelegowanie Zlecenia do Działu (Handlowy / Montaż / Specjalista Techniczny)")
        st.caption("Wybierz zlecenie, wskaż dział lub konkretnego handlowca, dopisz notatkę i wyślij oficjalny e-mail. Jeśli w ciągu 3 dni nie zaznaczysz otrzymania oferty/terminu, system sam wyśle przypomnienie.")

        del_c1, del_c2 = st.columns([1.2, 1])
        with del_c1:
            if not df.empty:
                del_opts = [f"{r['ID_Zgloszenia']} | {r['Klient']} ({r['Miasto']} - {r['Typ_Instalacji']})" for _, r in df.iterrows()]
                sel_del_str = st.selectbox("Wybierz zgłoszenie do przekazania:", del_opts, key="del_ticket_sel")
                sel_del_id = sel_del_str.split(" | ")[0]
                d_idx = df[df["ID_Zgloszenia"] == sel_del_id].index[0]
                d_row = df.loc[d_idx]

                dzial_choice = st.radio(
                    "Wybierz dział docelowy:",
                    ["Dział Handlowy (Oferta / Wymiana)", "Dział Montażu (Poprawka / Przebudowa)", "Specjalista Techniczny (Konsultacja)"],
                    horizontal=True
                )

                target_email = ""
                if dzial_choice.startswith("Dział Handlowy"):
                    handlowcy_list = email_cfg.get("handlowcy", [])
                    h_opts = [f"{h['nazwa']} <{h['email']}> ({h.get('rejon', '')})" for h in handlowcy_list]
                    sel_h = st.selectbox("Wybierz Handlowca, do którego trafi oferta:", h_opts) if h_opts else ""
                    if sel_h and "<" in sel_h:
                        target_email = sel_h.split("<")[1].split(">")[0]
                elif dzial_choice.startswith("Dział Montażu"):
                    target_email = email_cfg.get("email_montaz", "")
                else:
                    target_email = email_cfg.get("email_technik", "")

                target_email_input = st.text_input("Adres(y) e-mail odbiorcy (możesz dopisać po przecinku):", value=target_email)
                notatka_dzial = st.text_area("Oficjalna notatka / wytyczne dla działu:", value=str(d_row.get("Notatka_Dla_Dzialu", "")))
                att_files = st.file_uploader("Załącz zdjęcia tabliczki / instalacji do wiadomości e-mail (opcjonalnie):", accept_multiple_files=True, key="del_att")

                if st.button("📧 WYŚLIJ E-MAIL I ODDELEGUJ ZLECENIE", type="primary"):
                    subj = f"[PROMAG SERWIS -> {dzial_choice.split(' ')[1].upper()}] Zlecenie {sel_del_id} - {d_row['Klient']} ({d_row['Miasto']})"
                    maps_link = f"https://www.google.com/maps/search/?api=1&query={d_row['Adres']}+{d_row['Kod_Pocztowy']}+{d_row['Miasto']}".replace(" ", "+")
                    body_lines = [
                        "Dzień dobry,",
                        "",
                        "Koordynator Serwisu PROMAG OZE przekazuje zgłoszenie do realizacji / przygotowania oferty:",
                        "",
                        "=========================================",
                        f"ID ZGŁOSZENIA: {sel_del_id}",
                        f"DZIAŁ DOCELOWY: {dzial_choice}",
                        f"DATA PRZEKAZANIA: {today_iso}",
                        "=========================================",
                        f"KLIENT: {d_row['Klient']}",
                        f"TELEFON: {d_row['Telefon']}",
                        f"ADRES: {d_row['Kod_Pocztowy']} {d_row['Miasto']}, {d_row['Adres']}",
                        f"MAPA GOOGLE: {maps_link}",
                        f"URZĄDZENIE: {d_row['Typ_Instalacji']} ({d_row['Typ_Serwisu']})",
                        "",
                        "OFICJALNY OPIS ZGŁOSZENIA:",
                        str(d_row['Opis_Zgloszenia']),
                        "",
                        "NOTATKA / WYTYCZNE DLA DZIAŁU:",
                        str(notatka_dzial),
                        "",
                        "NOTATKA SERWISANTA Z TERENU:",
                        str(d_row.get('Notatka_Serwisanta', 'Brak')),
                        "=========================================",
                        "UWAGA: Prosimy o przekazanie oferty lub terminu w ciągu 3 dni. Po 3 dniach system automatycznie wyśle przypomnienie."
                    ]
                    body = "\n".join(body_lines)

                    df.loc[d_idx, "Status"] = "Oddelegowane"
                    df.loc[d_idx, "Oddelegowano_Dzial"] = dzial_choice
                    df.loc[d_idx, "Oddelegowano_Email"] = target_email_input
                    df.loc[d_idx, "Data_Oddelegowania"] = today_iso
                    df.loc[d_idx, "Notatka_Dla_Dzialu"] = notatka_dzial
                    df.loc[d_idx, "Otrzymano_Oferte_Termin"] = "Nie"
                    save_tickets(df)

                    ok, msg_smtp = send_system_email(email_cfg, target_email_input, subj, body, attachments=att_files)
                    if ok:
                        st.success(f"Zlecenie oddelegowane! {msg_smtp}")
                    else:
                        st.warning(f"Zlecenie oznaczono jako Oddelegowane w bazie (Watchdog 3-dniowy aktywny). Status SMTP: {msg_smtp}")

        with del_c2:
            st.markdown("#### ⏳ Oddelegowane sprawy (Kontrola 3 dni na ofertę/termin)")
            delegated_df = df[df["Oddelegowano_Dzial"].astype(str).str.strip() != ""].copy()
            if not delegated_df.empty:
                for del_i, del_r in delegated_df.iterrows():
                    did = del_r["ID_Zgloszenia"]
                    got_offer = (del_r["Otrzymano_Oferte_Termin"] == "Tak")
                    badge = "🟢 Załatwione" if got_offer else "🔴 OCZEKUJE NA OFERTĘ / TERMIN"
                    with st.container(border=True):
                        st.markdown(f"**{did} – {del_r['Klient']}** | `{del_r['Oddelegowano_Dzial']}` ({del_r['Oddelegowano_Email']})")
                        st.caption(f"Data przekazania: {del_r['Data_Oddelegowania']} | Status: {badge} | Notatka: {del_r['Notatka_Dla_Dzialu']}")
                        if not got_offer:
                            if st.button(f"✅ Potwierdź otrzymanie oferty/terminu ({did})", key=f"conf_del_{did}"):
                                df.loc[del_i, "Otrzymano_Oferte_Termin"] = "Tak"
                                save_tickets(df)
                                st.rerun()
            else:
                st.info("Brak oddelegowanych zleceń.")

        st.divider()
        c_parts, c_leads = st.columns(2)
        with c_parts:
            st.markdown("#### 📦 Zlecenia oczekujące na zamówienie części")
            parts_df = df[df["Status"] == "Do zamówienia części"]
            if not parts_df.empty:
                for p_i, p_r in parts_df.iterrows():
                    pid = p_r["ID_Zgloszenia"]
                    with st.container(border=True):
                        st.markdown(f"**{pid} – {p_r['Klient']} ({p_r['Miasto']})** | Serwisant: {p_r['Serwisant']}")
                        st.error(f"🛠️ **Potrzebna część:** {p_r['Potrzebne_Czesci']}")
                        if st.button(f"✅ Część dotarła – przywróć do planowania ({pid})", key=f"part_ok_{pid}"):
                            df.loc[p_i, "Status"] = "Nowe"
                            save_tickets(df)
                            st.rerun()
            else:
                st.success("Brak zleceń zablokowanych przez brak części.")

        with c_leads:
            st.markdown("#### 🔥 Leady sprzedażowe przywiezione przez serwisantów")
            leads_df = df[df["Lead_Handlowy"].astype(str).str.strip() != ""]
            if not leads_df.empty:
                st.dataframe(leads_df[["ID_Zgloszenia", "Klient", "Telefon", "Miasto", "Serwisant", "Lead_Handlowy"]], use_container_width=True, hide_index=True)
            else:
                st.info("Brak nowych leadów handlowych od serwisantów.")

    # ----------------------------------------------------
    # ZAKŁADKA 4: KADRY, NARZĘDZIA, MAGAZYN BUSA I ROZLICZENIE GOTÓWKI
    # ----------------------------------------------------
    with tab_kadry:
        st.subheader("👷 Kartoteka Serwisanta, Narzędzia, Stan Materiałowy na Aucie i Gotówka")
        selected_tech = st.selectbox("Wybierz serwisanta:", list(profiles.keys()), key="kadry_tech_sel")
        t_prof = profiles[selected_tech]
        t_metrics = calculate_tech_metrics(df, selected_tech)

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Obsłużone zlecenia", t_metrics["total_closed"], f"1. wizyta: {t_metrics['ftfr_pct']}%")
        k2.metric("Obrót miesięczny", f"{t_metrics['monthly_revenue']:,.0f} zł", t_metrics["active_rate_str"])
        k3.metric("💵 Gotówka w kieszeni (do zdania)", f"{t_metrics['cash_in_hand']:,.0f} zł")
        k4.metric("🔥 Przywiezione Leady", t_metrics["leads_count"])

        if t_metrics["cash_in_hand"] > 0:
            if st.button(f"💵 Potwierdź zdanie gotówki ({t_metrics['cash_in_hand']:,.0f} zł) w biurze przez: {selected_tech}", type="primary"):
                mask_cash = (df["Serwisant"] == selected_tech) & (df["Forma_Platnosci"] == "Gotówka pobrana")
                df.loc[mask_cash, "Gotowka_Rozliczona"] = "Tak"
                save_tickets(df)
                st.success("Rozliczono gotówkę serwisanta!")
                st.rerun()

        st.divider()
        inv_c1, inv_c2 = st.columns(2)

        with inv_c1:
            st.markdown(f"#### 📦 Stan Materiałowy na Samochodzie ({selected_tech})")
            st.caption("Edytuj ilości bezpośrednio w tabeli lub wgraj plik Excel/CSV ze stanem magazynowym.")
            mat_list = t_prof.get("materialy", [])
            ed_mat_kadry = st.data_editor(
                pd.DataFrame(mat_list),
                num_rows="dynamic",
                use_container_width=True,
                key=f"kadry_mat_{selected_tech}"
            )
            if st.button(f"💾 Zapisz stan materiałowy dla: {selected_tech}", type="primary"):
                profiles[selected_tech]["materialy"] = ed_mat_kadry.to_dict(orient="records")
                save_config_and_profiles(profiles)
                st.success("Zapisano stan materiałowy na aucie!")
                st.rerun()

            up_mat_file = st.file_uploader(f"📥 Wgraj stan materiałowy z pliku CSV/XLSX dla {selected_tech}:", type=["csv", "xlsx"], key=f"up_m_{selected_tech}")
            if up_mat_file is not None:
                if st.button("Załaduj materiały z pliku"):
                    loaded_m = pd.read_csv(up_mat_file) if up_mat_file.name.endswith(".csv") else pd.read_excel(up_mat_file)
                    profiles[selected_tech]["materialy"] = loaded_m.to_dict(orient="records")
                    save_config_and_profiles(profiles)
                    st.success("Wgrano stan materiałowy!")
                    st.rerun()

        with inv_c2:
            st.markdown(f"#### 🧰 Lista Narzędzi i Wyposażenia ({selected_tech})")
            st.caption("Wykaz narzędzi przypisanych do serwisanta wraz z ilościami.")
            tools_list = t_prof.get("narzedzia", [])
            ed_tools_kadry = st.data_editor(
                pd.DataFrame(tools_list),
                num_rows="dynamic",
                use_container_width=True,
                key=f"kadry_tools_{selected_tech}"
            )
            if st.button(f"💾 Zapisz listę narzędzi dla: {selected_tech}", type="primary"):
                profiles[selected_tech]["narzedzia"] = ed_tools_kadry.to_dict(orient="records")
                save_config_and_profiles(profiles)
                st.success("Zapisano listę narzędzi!")
                st.rerun()

            up_tools_file = st.file_uploader(f"📥 Wgraj listę narzędzi z pliku CSV/XLSX dla {selected_tech}:", type=["csv", "xlsx"], key=f"up_t_{selected_tech}")
            if up_tools_file is not None:
                if st.button("Załaduj narzędzia z pliku"):
                    loaded_t = pd.read_csv(up_tools_file) if up_tools_file.name.endswith(".csv") else pd.read_excel(up_tools_file)
                    profiles[selected_tech]["narzedzia"] = loaded_t.to_dict(orient="records")
                    save_config_and_profiles(profiles)
                    st.success("Wgrano listę narzędzi!")
                    st.rerun()

        st.divider()
        with st.expander("📋 Dane osobowe, PIN, Umowa, BHP, Flota i Urlopy"):
            with st.form(f"form_kadry_full_{selected_tech}"):
                fk1, fk2, fk3 = st.columns(3)
                with fk1:
                    new_fn = st.text_input("Pełne dane:", value=t_prof.get("pelne_dane", ""))
                    new_pin = st.text_input("Kod PIN:", value=str(t_prof.get("pin", "")))
                with fk2:
                    new_sal = st.number_input("Podstawa (zł):", 0.0, 30000.0, float(t_prof.get("wynagrodzenie_bazowe", 6000.0)), 100.0)
                    new_rej = st.text_input("Nr rej. auta:", value=t_prof.get("flota", {}).get("nr_rej", ""))
                with fk3:
                    new_odo_k = st.number_input("Przebieg [km]:", 0, 1000000, int(t_prof.get("flota", {}).get("przebieg_km", 120000)), 500)
                    new_olej_k = st.number_input("Kolejny serwis olejowy przy [km]:", 0, 1000000, int(t_prof.get("flota", {}).get("limit_olej_km", 130000)), 1000)
                if st.form_submit_button("💾 Zapisz dane kadrowe i flotowe"):
                    t_prof["pelne_dane"] = new_fn
                    t_prof["pin"] = new_pin
                    t_prof["wynagrodzenie_bazowe"] = new_sal
                    t_prof["flota"]["nr_rej"] = new_rej
                    t_prof["flota"]["przebieg_km"] = new_odo_k
                    t_prof["flota"]["limit_olej_km"] = new_olej_k
                    save_config_and_profiles(profiles)
                    st.success("Zapisano!")

    # ----------------------------------------------------
    # ZAKŁADKA 5: USTAWIENIA E-MAIL (SMTP + HANDLOWCY), LOGO I IMPORT EXCEL
    # ----------------------------------------------------
    with tab_ustawienia:
        st.subheader("⚙️ Konfiguracja Konta E-mail (SMTP) i Adresów Działów / Handlowców")
        em_c1, em_c2 = st.columns(2)
        with em_c1:
            with st.form("form_smtp_settings"):
                st.markdown("#### 📧 Serwer Poczty Wychodzącej (SMTP)")
                s_host = st.text_input("Serwer SMTP (np. smtp.gmail.com):", value=email_cfg.get("smtp_server", "smtp.gmail.com"))
                s_port = st.number_input("Port SMTP (587 dla TLS lub 465 dla SSL):", value=int(email_cfg.get("smtp_port", 587)))
                s_user = st.text_input("Adres e-mail nadawcy:", value=email_cfg.get("sender_email", ""))
                s_pass = st.text_input("Hasło (lub Hasło Aplikacji Google):", value=email_cfg.get("sender_password", ""), type="password")
                st.markdown("#### 🛠️ Adresy Działu Montażu i Specjalisty Technicznego")
                e_montaz = st.text_input("E-mail Działu Montażu:", value=email_cfg.get("email_montaz", ""))
                e_technik = st.text_input("E-mail Specjalisty Technicznego:", value=email_cfg.get("email_technik", ""))

                if st.form_submit_button("💾 Zapisz konfigurację SMTP i adresy działów", type="primary"):
                    email_cfg["smtp_server"] = s_host
                    email_cfg["smtp_port"] = int(s_port)
                    email_cfg["sender_email"] = s_user
                    email_cfg["sender_password"] = s_pass
                    email_cfg["email_montaz"] = e_montaz
                    email_cfg["email_technik"] = e_technik
                    st.session_state["email_cfg"] = email_cfg
                    save_config_and_profiles(profiles, email_cfg=email_cfg)
                    st.success("Zapisano ustawienia poczty!")

            if st.button("🧪 Wyślij testowy e-mail na adres nadawcy"):
                ok_t, msg_t = send_system_email(email_cfg, email_cfg.get("sender_email", ""), "Test SMTP - PROMAG OZE", "Połączenie SMTP działa prawidłowo!")
                if ok_t:
                    st.success(msg_t)
                else:
                    st.error(msg_t)

        with em_c2:
            st.markdown("#### 💼 Lista Handlowców (Do wyboru przy oddelegowywaniu ofert)")
            st.caption("Dodaj dowolną liczbę handlowców. Przy oddelegowaniu zlecenia wybierzesz z listy, do którego handlowca ma trafić temat.")
            h_df = pd.DataFrame(email_cfg.get("handlowcy", []))
            ed_h_df = st.data_editor(h_df, num_rows="dynamic", use_container_width=True, key="editor_handlowcy")
            if st.button("💾 Zapisz listę handlowców", type="primary"):
                email_cfg["handlowcy"] = ed_h_df.to_dict(orient="records")
                st.session_state["email_cfg"] = email_cfg
                save_config_and_profiles(profiles, email_cfg=email_cfg)
                st.success("Zaktualizowano listę handlowców!")

        st.divider()
        b_c1, b_c2 = st.columns(2)
        with b_c1:
            st.subheader("🎨 Wgraj Logo Firmy")
            uploaded_logo = st.file_uploader("Wybierz plik PNG/JPG:", type=["png", "jpg", "jpeg"])
            if uploaded_logo is not None:
                img = Image.open(uploaded_logo)
                img.thumbnail((400, 120))
                buf = BytesIO()
                img.save(buf, format="PNG")
                b64_str = base64.b64encode(buf.getvalue()).decode("utf-8")
                st.session_state["logo_b64"] = b64_str
                save_config_and_profiles(profiles, logo_b64=b64_str)
                st.success("Logo zapisane!")
                st.rerun()
        with b_c2:
            st.subheader("📥 Import Bazy Zgłoszeń z pliku Excel (.xlsx)")
            uploaded_excel = st.file_uploader("Wgraj plik .xlsx:", type=["xlsx"])
            if uploaded_excel is not None:
                if st.button("⚠️ Zaimportuj dane z pliku Excel"):
                    new_df = ensure_columns(pd.read_excel(uploaded_excel))
                    save_tickets(new_df)
                    st.success(f"Zaimportowano {len(new_df)} zgłoszeń!")
                    st.rerun()

# ==========================================
# 8. PANEL ZARZĄDU
# ==========================================
elif st.session_state["logged_role"] == "zarzad":
    df = st.session_state["df_tickets"]
    profiles = st.session_state["profiles"]

    st.subheader("📊 Panel Zarządu – Raporty Operacyjne, Gotówka, Marża i Prowizje")
    month_summary = []
    for t_name, t_prof in profiles.items():
        tm = calculate_tech_metrics(df, t_name)
        base_sal = float(t_prof.get("wynagrodzenie_bazowe", 0.0))
        month_summary.append({
            "Serwisant": t_name,
            "Umowa": t_prof.get("umowa", ""),
            "Zamknięte": tm["total_closed"],
            "1. Wizyta (%)": f"{tm['ftfr_pct']}%",
            "Leady Handlowe": tm["leads_count"],
            "Gotówka nierozliczona (zł)": round(tm["cash_in_hand"], 2),
            "Obrót miesięczny (zł)": round(tm["monthly_revenue"], 2),
            "Koszt zużytych materiałów (zł)": round(tm["monthly_parts_cost"], 2),
            "Podstawa (zł)": base_sal,
            "Prowizja (zł)": round(tm["commission_pln"], 2),
            "Do wypłaty (zł)": round(base_sal + tm["commission_pln"], 2)
        })
    st.dataframe(pd.DataFrame(month_summary), use_container_width=True, hide_index=True)
