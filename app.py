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

# Import wydzielonych modułów (cennik.py, kadry.py, baza_wiedzy.py)
try:
    from cennik import (
        STAWKA_RBH_NETTO,
        KOSZT_PALIWA_LITR,
        SR_CZAS_WIZYTY_H,
        RYCZALT_DOJAZD,
        TABELA_CENNIKA,
        wylicz_cene_uslugi,
        pobierz_slownik_cennika,
        oblicz_koszyk_uslug
    )
except ImportError:
    STAWKA_RBH_NETTO = 100.0
    KOSZT_PALIWA_LITR = 8.0
    SR_CZAS_WIZYTY_H = 1.5
    RYCZALT_DOJAZD = 200.0
    TABELA_CENNIKA = [{"usluga": "Przegląd/Konserwacja PC", "cena_materialow": 50.0, "czas_pracy_h": 2.0}]
    def wylicz_cene_uslugi(w, s=100.0): return 250.0
    def pobierz_slownik_cennika(t=None, s=100.0): return {"Przegląd/Konserwacja PC": {"cena_materialow": 50.0, "czas_pracy_h": 2.0, "cena_calkowita": 250.0}}
    def oblicz_koszyk_uslug(w, t=None, s=100.0, d=True, k=200.0): return {"suma_materialow": 0, "suma_czasu_h": 0, "suma_robocizny": 0, "dojazd": 0, "razem_dla_klienta": 0, "auto_notatka": ""}

try:
    from kadry import (
        COMMISSION_TIER_1,
        COMMISSION_RATE_1,
        COMMISSION_TIER_2,
        COMMISSION_RATE_2,
        DOMYSLNE_NARZEDZIA,
        DOMYSLNE_MATERIALY,
        DEFAULT_TECHS,
        ensure_tech_profile_fields,
        get_low_stock_items,
        calculate_tech_metrics
    )
except ImportError:
    COMMISSION_TIER_1, COMMISSION_RATE_1, COMMISSION_TIER_2, COMMISSION_RATE_2 = 15000.0, 0.03, 30000.0, 0.04
    DOMYSLNE_NARZEDZIA = [{"nazwa": "Refraktometr", "ilosc": 1}]
    DOMYSLNE_MATERIALY = [{"nazwa": "Glikol 20L", "ilosc": 3, "koszt_zakupu": 180.0, "cena_klienta": 320.0, "slowo_klucz": "glikol"}]
    DEFAULT_TECHS = {"Leszek": {"pin": "4829", "narzedzia": DOMYSLNE_NARZEDZIA, "materialy": DOMYSLNE_MATERIALY, "flota": {"przebieg_km": 120000, "limit_olej_km": 130000}}}
    def ensure_tech_profile_fields(p): return p
    def get_low_stock_items(p): return []
    def calculate_tech_metrics(df, t): return {"total_closed": 0, "ftfr_pct": 100.0, "paid_count": 0, "warranty_count": 0, "leads_count": 0, "cash_in_hand": 0.0, "monthly_revenue": 0.0, "monthly_parts_cost": 0.0, "missing_tier_1": 15000.0, "missing_tier_2": 30000.0, "commission_pln": 0.0, "active_rate_str": "0%", "progress_pct": 0}

try:
    from baza_wiedzy import (
        DIAGNOSTYKA_BAZA,
        INSTRUKCJE_KONFIGURACJI_TEXT,
        POSTAL_PREFIX_COORDS,
        przeszukaj_lokalne_instrukcje,
        przeszukaj_praktyke_i_internet
    )
except ImportError:
    DIAGNOSTYKA_BAZA = {"Brak pliku baza_wiedzy.py": {"Wgraj baza_wiedzy.py": "Wgraj plik na GitHub"}}
    INSTRUKCJE_KONFIGURACJI_TEXT = ""
    POSTAL_PREFIX_COORDS = {"05": (52.08, 20.85)}
    def przeszukaj_lokalne_instrukcje(q, f="."): return []
    def przeszukaj_praktyke_i_internet(q): return [], "https://google.com", "https://google.com", "https://youtube.com"

# ==========================================
# 1. WCZESNE WCZYTANIE UPROSZCZONEGO LOGO NA PASEK PRZEGLĄDARKI (FAVICON)
# ==========================================
GOOGLE_SHEET_NAME = "Serwisowa posortowana"
LOCAL_EXCEL_FILE = "Serwisowa posortowana.xlsx"
MONTAZE_CSV_FILE = "skonczone.csv"
MONTAZE_PARQUET_CACHE = "skonczone_clean.parquet"
PROFILES_FILE = "profiles.json"
LOGO_FILE = "logo.png"

PIN_KOORDYNATOR = "1234"
PIN_ZARZAD = "9999"

def get_early_favicon():
    """Wczytuje wgrane uproszczone logo jako ikonę na pasek przeglądarki (favicon) przed st.set_page_config."""
    if os.path.exists(LOGO_FILE):
        try:
            return Image.open(LOGO_FILE)
        except Exception:
            pass
    if os.path.exists(PROFILES_FILE):
        try:
            with open(PROFILES_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
                b64 = saved.get("logo_b64", None)
                if b64:
                    return Image.open(BytesIO(base64.b64decode(b64)))
        except Exception:
            pass
    return "🌿"

st.set_page_config(
    page_title="PROMAG OZE - Serwis",
    page_icon=get_early_favicon(),
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    /* Kolorystyka inspirowana promag-gaz.pl: Głęboki granat (#0a2540 / #10375c), Ekologiczna zieleń (#16a34a / #22c55e), Czyste jasne karty */
    .stButton > button {
        width: 100%;
        min-height: 48px;
        font-size: 15px !important;
        font-weight: 700 !important;
        border-radius: 10px !important;
        margin-bottom: 6px;
        transition: all 0.2s ease-in-out;
    }
    .promag-header {
        background: linear-gradient(135deg, #0a2540 0%, #10375c 100%);
        color: #ffffff;
        padding: 16px 22px;
        border-radius: 14px;
        border-bottom: 4px solid #16a34a;
        margin-bottom: 14px;
    }
    .kpi-card {
        background: linear-gradient(135deg, #0a2540 0%, #154360 100%);
        color: #ffffff;
        padding: 16px;
        border-radius: 14px;
        border-left: 6px solid #22c55e;
        box-shadow: 0 4px 12px rgba(10, 37, 64, 0.12);
        margin-bottom: 10px;
    }
    .kpi-card-blue {
        background: linear-gradient(135deg, #10375c 0%, #1b4f72 100%);
        color: #ffffff;
        padding: 16px;
        border-radius: 14px;
        border-left: 6px solid #38bdf8;
        box-shadow: 0 4px 12px rgba(10, 37, 64, 0.12);
        margin-bottom: 10px;
    }
    .kpi-title { font-size: 12px; text-transform: uppercase; letter-spacing: 0.6px; color: #cbd5e1; }
    .kpi-value { font-size: 20px; font-weight: 800; color: #ffffff; margin-top: 2px; }
    .kpi-sub { font-size: 13px; color: #86efac; margin-top: 4px; }
    .day-card {
        background: #ffffff;
        color: #0a2540;
        border: 1px solid #cbd5e1;
        border-radius: 14px;
        padding: 16px;
        margin-bottom: 10px;
        border-left: 7px solid #0284c7;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    }
    .day-card-today {
        background: #f0fdf4;
        color: #064e3b;
        border: 1px solid #86efac;
        border-left: 7px solid #16a34a;
    }
    .install-badge {
        background-color: #f0f9ff;
        border: 1px solid #bae6fd;
        border-left: 5px solid #0284c7;
        color: #0c4a6e;
        padding: 12px;
        border-radius: 10px;
        margin-top: 8px;
        margin-bottom: 10px;
        font-size: 14px;
    }
    .client-box {
        background-color: #ffffff;
        border: 2px solid #16a34a;
        border-radius: 14px;
        padding: 18px;
        margin-top: 12px;
        margin-bottom: 12px;
        color: #0a2540;
    }
    .onboard-card {
        background: #f8fafc;
        border: 1px solid #cbd5e1;
        border-left: 6px solid #16a34a;
        border-radius: 12px;
        padding: 16px;
        margin-bottom: 12px;
        color: #0f172a;
    }
</style>
""", unsafe_allow_html=True)

DEFAULT_BURN_L_100KM = 9.0
DEFAULT_SPEED_KMH = 60.0
ROAD_TORTUOSITY = 1.3
TARGET_NET_PROFIT = 1200.0

DEFAULT_EMAIL_CONFIG = {
    "smtp_server": "smtp.gmail.com",
    "smtp_port": 587,
    "sender_email": "",
    "sender_password": "",
    "email_montaz": "montaz@promag-gaz.pl",
    "email_technik": "technik@promag-gaz.pl",
    "handlowcy": [
        {"nazwa": "Handlowiec 1 (Pompy / PV)", "email": "handlowiec1@promag-gaz.pl", "rejon": "Cała Polska"},
        {"nazwa": "Handlowiec 2 (Magazyny / Modernizacje)", "email": "handlowiec2@promag-gaz.pl", "rejon": "Mazowieckie"}
    ]
}

# ==========================================
# 2. KONWERTER DATY (TYDZIEŃ ROKU -> MIESIĄC I ROK) + PARSER BAZY MONTAŻY
# ==========================================
MIESIACE_PL = {
    1: "Styczeń", 2: "Luty", 3: "Marzec", 4: "Kwiecień",
    5: "Maj", 6: "Czerwiec", 7: "Lipiec", 8: "Sierpień",
    9: "Wrzesień", 10: "Październik", 11: "Listopad", 12: "Grudzień"
}

def convert_week_or_date_to_month_year(raw_val):
    s = str(raw_val).strip()
    if not s or s.lower() in ["nan", "none", "nat", ""]:
        return "Brak daty"

    m_iso = re.search(r'\b(20\d{2})[-/.](0?[1-9]|1[0-2])[-/.](\d{1,2})\b', s)
    if m_iso:
        yr, mn = int(m_iso.group(1)), int(m_iso.group(2))
        return f"{MIESIACE_PL.get(mn, '')} {yr} ({yr}-{mn:02d})"

    m_pl = re.search(r'\b(\d{1,2})[-/.](0?[1-9]|1[0-2])[-/.](20\d{2})\b', s)
    if m_pl:
        mn, yr = int(m_pl.group(2)), int(m_pl.group(3))
        return f"{MIESIACE_PL.get(mn, '')} {yr} ({yr}-{mn:02d})"

    m_wk1 = re.search(r'\b([1-9]|[1-4]\d|5[0-3])\s*[/._-]\s*(20[1-3]\d)\b', s)
    if m_wk1:
        wk, yr = int(m_wk1.group(1)), int(m_wk1.group(2))
        try:
            d = datetime.date.fromisocalendar(yr, min(52, max(1, wk)), 4)
            return f"{MIESIACE_PL[d.month]} {d.year} ({d.year}-{d.month:02d})"
        except Exception:
            return f"Rok {yr}"

    m_wk2 = re.search(r'\b(20[1-3]\d)\s*[/._-]?[wWtT]?\s*([1-9]|[1-4]\d|5[0-3])\b', s)
    if m_wk2:
        yr, wk = int(m_wk2.group(1)), int(m_wk2.group(2))
        try:
            d = datetime.date.fromisocalendar(yr, min(52, max(1, wk)), 4)
            return f"{MIESIACE_PL[d.month]} {d.year} ({d.year}-{d.month:02d})"
        except Exception:
            return f"Rok {yr}"

    if s.isdigit() and 1 <= int(s) <= 53:
        wk = int(s)
        d = datetime.date.fromisocalendar(2024, min(52, wk), 4)
        return f"{MIESIACE_PL[d.month]} (Tydz. {wk})"

    return s

def extract_phone_numbers(raw_phone_str, raw_client_str=""):
    combined = f"{raw_phone_str} {raw_client_str}"
    candidates = re.findall(r'(?:\+48)?\s*\d{3}[\s-]?\d{3}[\s-]?\d{3}|\d{2}[\s-]?\d{3}[\s-]?\d{2}[\s-]?\d{2}|\d{9,11}', combined)
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

def clean_and_parse_montaze_df(raw_df):
    """Przekształca tabelę skonczone.csv / skonczone_clean.parquet w bazę klientów z wyszukiwaniem po Nazwisku, Telefonie i Adresie."""
    if raw_df is None or raw_df.empty:
        return pd.DataFrame(columns=["Klient", "Telefon", "Kod_Pocztowy", "Miasto", "Adres", "Data_Montazu", "Opis_Instalacji", "Search_Label"])

    expected_ready = {"Klient", "Telefon", "Kod_Pocztowy", "Miasto", "Adres", "Data_Montazu", "Opis_Instalacji", "Search_Label"}
    if expected_ready.issubset(set(raw_df.columns)):
        raw_df = raw_df.fillna("").copy()
        # Upewnij się, że Search_Label ma na początku Nazwisko, Telefon oraz Adres
        raw_df["Search_Label"] = (
            raw_df["Klient"].astype(str) + " | Tel: " + raw_df["Telefon"].astype(str) +
            " | Adres: " + raw_df["Kod_Pocztowy"].astype(str) + " " + raw_df["Miasto"].astype(str) + " " + raw_df["Adres"].astype(str) +
            " | Montaż: " + raw_df["Data_Montazu"].astype(str)
        )
        return raw_df

    raw_df = raw_df.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)

    if any(str(c).startswith("Unnamed:") for c in raw_df.columns[:5]) and len(raw_df) > 1:
        first_row_vals = [str(v).strip() for v in raw_df.iloc[0].values]
        if any("klient" in v.lower() or "adres" in v.lower() for v in first_row_vals):
            new_cols = []
            counts = {}
            for idx, val in enumerate(raw_df.iloc[0].values):
                clean_v = re.sub(r'\s+', ' ', str(val).strip())
                if not clean_v or clean_v.lower() == 'nan':
                    clean_v = f"Kolumna_{idx}"
                if clean_v in counts:
                    counts[clean_v] += 1
                    clean_v = f"{clean_v}_{counts[clean_v]}"
                else:
                    counts[clean_v] = 1
                new_cols.append(clean_v)
            raw_df = raw_df.iloc[1:].reset_index(drop=True)
            raw_df.columns = new_cols

    cols = [re.sub(r'\s+', ' ', str(c).strip()) for c in raw_df.columns]
    raw_df.columns = cols

    if "Klient" in cols and any("Adres" in c for c in cols):
        adr_col = next((c for c in cols if "Adres" in c), cols[2] if len(cols) > 2 else "Klient")
        tydz_col = next((c for c in cols if "Tydzień realizacji" in c or "Tydzien realizacji" in c), None)
        plan_col = next((c for c in cols if "Planowany termin" in c), None)
        umowa_data_cols = [c for c in cols if any(k in c.lower() for k in ["data podpisania", "data wprowadzenia", "termin zakończenia"])]
        opis_zest_cols = [c for c in cols if any(k in c.lower() for k in ["opis zestawu", "opis z umowy", "dodatkowe podzespoły", "uruchomione instalacje", "dodatkowe opisy", "nazwa instalatora"])]

        records = []
        for _, row in raw_df.iterrows():
            raw_klient = str(row.get("Klient", "")).strip()
            if not raw_klient or raw_klient.lower() in ["nan", "none", "klient"]:
                continue

            raw_adr = str(row.get(adr_col, "")).strip()
            if raw_adr.lower() in ["nan", "none"]:
                raw_adr = ""

            # Wyciągnij telefony zarówno z kolumny Klient, jak i z kolumny Adres
            phones = extract_phone_numbers(raw_adr, raw_klient)
            tel_str = " / ".join(phones)
            klient_clean = re.sub(r'(?:\+48)?\s*\d{3}[\s-]?\d{3}[\s-]?\d{3}|\d{2}[\s-]?\d{3}[\s-]?\d{2}[\s-]?\d{2}|\d{9,11}', '', raw_klient).strip(" ,-/") or raw_klient

            zip_m = re.search(r'\b(\d{2}-\d{3})\b', raw_adr)
            kod_p = zip_m.group(1) if zip_m else ""
            adr_no_zip = raw_adr.replace(kod_p, "").strip(" ,")
            if "," in adr_no_zip:
                parts = [p.strip() for p in adr_no_zip.split(",", 1)]
                miasto, ulica = parts[0], parts[1]
            elif " ul." in adr_no_zip.lower() or " ul " in adr_no_zip.lower():
                idx_ul = adr_no_zip.lower().find(" ul")
                miasto = adr_no_zip[:idx_ul].strip(" ,")
                ulica = adr_no_zip[idx_ul:].strip(" ,")
            else:
                miasto, ulica = adr_no_zip, raw_adr

            raw_umowa_str = " ".join([str(row.get(uc, "")) for uc in umowa_data_cols])
            yr_m = re.search(r'\b(20[0-2]\d)\b', raw_umowa_str)
            yr_val = int(yr_m.group(1)) if yr_m else None
            wk_str = str(row.get(tydz_col, "") if tydz_col else "").strip()
            if not wk_str or wk_str.lower() in ["nan", "none"]:
                wk_str = str(row.get(plan_col, "") if plan_col else "").strip()

            wk_m = re.search(r'\b([1-9]|[1-4]\d|5[0-3])\b', wk_str)
            if wk_m and yr_val:
                wk_num = int(wk_m.group(1))
                try:
                    d_iso = datetime.date.fromisocalendar(yr_val, min(52, max(1, wk_num)), 4)
                    data_montazu = f"{MIESIACE_PL[d_iso.month]} {d_iso.year} ({d_iso.year}-{d_iso.month:02d})"
                except Exception:
                    data_montazu = convert_week_or_date_to_month_year(raw_umowa_str)
            else:
                data_montazu = convert_week_or_date_to_month_year(raw_umowa_str if yr_val else wk_str)

            op_list = []
            for oc in opis_zest_cols:
                v_op = str(row.get(oc, "")).strip()
                if v_op and v_op.lower() not in ["nan", "none", "-", "0", "0,00 zł"]:
                    op_list.append(f"{oc}: {v_op}")
            opis_inst = " | ".join(op_list)

            records.append({
                "Klient": klient_clean,
                "Telefon": tel_str,
                "Kod_Pocztowy": kod_p,
                "Miasto": miasto,
                "Adres": ulica,
                "Data_Montazu": data_montazu,
                "Opis_Instalacji": opis_inst,
                "Search_Label": f"{klient_clean} | Tel: {tel_str} | Adres: {kod_p} {miasto}, {ulica} | Montaż: {data_montazu}"
            })
        return pd.DataFrame(records)

    return pd.DataFrame(columns=["Klient", "Telefon", "Kod_Pocztowy", "Miasto", "Adres", "Data_Montazu", "Opis_Instalacji", "Search_Label"])

@st.cache_data(show_spinner="Wczytywanie bazy wszystkich klientów...")
def load_montaze_database():
    if os.path.exists(MONTAZE_PARQUET_CACHE):
        try:
            pq_df = pd.read_parquet(MONTAZE_PARQUET_CACHE)
            return clean_and_parse_montaze_df(pq_df)
        except Exception:
            pass
    for candidate in ["skonczone_gotowe.csv", "skonczone_lekkie.csv", MONTAZE_CSV_FILE, "Promag/skonczone_lekkie.csv", "Promag/skonczone.csv", "skończone.csv"]:
        if os.path.exists(candidate):
            try:
                try:
                    raw_df = pd.read_csv(candidate, sep=None, engine="python", on_bad_lines="skip")
                except Exception:
                    raw_df = pd.read_csv(candidate, encoding="cp1250", sep=None, engine="python", on_bad_lines="skip")
                clean_df = clean_and_parse_montaze_df(raw_df)
                try:
                    clean_df.to_parquet(MONTAZE_PARQUET_CACHE, index=False)
                except Exception:
                    pass
                return clean_df
            except Exception:
                pass
    return pd.DataFrame(columns=["Klient", "Telefon", "Kod_Pocztowy", "Miasto", "Adres", "Data_Montazu", "Opis_Instalacji", "Search_Label"])

# ==========================================
# 3. FUNKCJE GPS, KORYTARZ TRASY I EMAIL SMTP
# ==========================================
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
                             hourly_rate=STAWKA_RBH_NETTO, fuel_price=KOSZT_PALIWA_LITR,
                             burn_l_100km=DEFAULT_BURN_L_100KM, onsite_hours_per_client=SR_CZAS_WIZYTY_H):
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
        return False, "Brak skonfigurowanego konta SMTP w zakładce Ustawienia."

    to_list = [x.strip() for x in (to_emails.split(",") if isinstance(to_emails, str) else to_emails) if str(x).strip()]
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = ", ".join(to_list)
    msg.set_content(body_text)

    if attachments:
        for att in attachments:
            try:
                msg.add_attachment(att.getvalue(), maintype="application", subtype="octet-stream", filename=att.name)
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
        return False, f"Błąd SMTP: {e}"

def load_config_and_profiles():
    profiles = DEFAULT_TECHS.copy()
    logo_b64 = None
    email_cfg = DEFAULT_EMAIL_CONFIG.copy()
    cennik_tab = [dict(x) for x in TABELA_CENNIKA]
    stawka_rbh = STAWKA_RBH_NETTO
    dojazd_kw = RYCZALT_DOJAZD

    if os.path.exists(PROFILES_FILE):
        try:
            with open(PROFILES_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
                profiles = saved.get("techs", DEFAULT_TECHS)
                logo_b64 = saved.get("logo_b64", None)
                email_cfg = saved.get("email_cfg", DEFAULT_EMAIL_CONFIG)
                cennik_tab = saved.get("cennik_tab", TABELA_CENNIKA)
                stawka_rbh = float(saved.get("stawka_rbh", STAWKA_RBH_NETTO))
                dojazd_kw = float(saved.get("dojazd_kw", RYCZALT_DOJAZD))
        except Exception:
            pass
    return ensure_tech_profile_fields(profiles), logo_b64, email_cfg, cennik_tab, stawka_rbh, dojazd_kw

def save_config_and_profiles(profiles, logo_b64=None, email_cfg=None, cennik_tab=None, stawka_rbh=None, dojazd_kw=None):
    if logo_b64 is None: logo_b64 = st.session_state.get("logo_b64", None)
    if email_cfg is None: email_cfg = st.session_state.get("email_cfg", DEFAULT_EMAIL_CONFIG)
    if cennik_tab is None: cennik_tab = st.session_state.get("cennik_tab", TABELA_CENNIKA)
    if stawka_rbh is None: stawka_rbh = st.session_state.get("stawka_rbh", STAWKA_RBH_NETTO)
    if dojazd_kw is None: dojazd_kw = st.session_state.get("dojazd_kw", RYCZALT_DOJAZD)

    profiles = ensure_tech_profile_fields(profiles)
    payload = {
        "techs": profiles, "logo_b64": logo_b64, "email_cfg": email_cfg,
        "cennik_tab": cennik_tab, "stawka_rbh": float(stawka_rbh), "dojazd_kw": float(dojazd_kw)
    }
    try:
        with open(PROFILES_FILE, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def render_brand_logo(max_width=220, centered=False):
    logo_b64 = st.session_state.get("logo_b64", None)
    if logo_b64:
        align_style = "display: block; margin-left: auto; margin-right: auto;" if centered else ""
        st.markdown(
            f'<img src="data:image/png;base64,{logo_b64}" style="max-width:{max_width}px; max-height:85px; object-fit:contain; {align_style} margin-bottom:8px;" alt="PROMAG OZE">',
            unsafe_allow_html=True
        )
    elif os.path.exists(LOGO_FILE):
        st.image(LOGO_FILE, width=max_width)
    else:
        align = "text-align:center;" if centered else ""
        st.markdown(f"<div class='promag-header' style='{align}'><span style='font-size:22px; font-weight:800;'>🌿 PROMAG SERWIS OZE</span><br><span style='font-size:13px; color:#86efac;'>Fotowoltaika • Pompy Ciepła • Kolektory Słoneczne • Magazyny Energii</span></div>", unsafe_allow_html=True)

# ==========================================
# 4. PEŁNE MAPOWANIE KOLUMN ZGŁOSZEŃ SERWISOWYCH (NAPRAWIONA LISTA ZGŁOSZEŃ!)
# ==========================================
REQUIRED_COLS = [
    "ID_Zgloszenia", "Klient", "Telefon", "Kod_Pocztowy", "Miasto", "Adres",
    "Typ_Instalacji", "Opis_Zgloszenia", "Status", "Serwisant", "Data_Planowana",
    "Kwota_Przewidywana", "Zgoda_Klienta_Kwota", "Data_Montazu", "Opis_Instalacji_Montaz",
    "Priorytet_Krytyczny", "Typ_Serwisu", "Wymagany_Przeglad", "Naprawione_1_Wizyta",
    "Kwota_Koncowa", "Koszt_Materialow", "Forma_Platnosci", "Gotowka_Rozliczona",
    "Czas_Pracy_Min", "Wyniki_Przegladu", "Zdjecia_Potwierdzone", "Zalecenia_Klienta",
    "Notatka_Serwisanta", "Data_Zamkniecia", "Podpis_Zlozony",
    "Do_Oddzwonienia", "Data_Oddzwonienia", "Potwierdzony_Tel", "Trudny_Klient",
    "Potrzebne_Czesci", "Lead_Handlowy", "Oddelegowano_Dzial", "Oddelegowano_Email",
    "Data_Oddelegowania", "Notatka_Dla_Dzialu", "Otrzymano_Oferte_Termin", "Ostatnie_Przypomnienie"
]

NUMERIC_COLS = ["Kwota_Przewidywana", "Zgoda_Klienta_Kwota", "Kwota_Koncowa", "Koszt_Materialow", "Czas_Pracy_Min"]

def ensure_columns(df):
    """Gwarantuje poprawne mapowanie kolumn z pliku Serwisowa posortowana.xlsx (Klient, Telefon, Opis_Zgloszenia, Kwota)."""
    if df is None or df.empty:
        empty_df = pd.DataFrame(columns=REQUIRED_COLS)
        for nc in NUMERIC_COLS:
            empty_df[nc] = empty_df[nc].astype(float)
        return empty_df

    clean_cols, cols_to_keep = [], []
    for idx, c in enumerate(df.columns):
        c_str = str(c).strip()
        if c_str != "" and not c_str.lower().startswith("unnamed"):
            cols_to_keep.append(idx)
            clean_cols.append(c_str)

    df = df.iloc[:, cols_to_keep].copy()
    df.columns = clean_cols
    df = df.loc[:, ~df.columns.duplicated()].copy().reset_index(drop=True)

    # Mapowanie kolumn z pliku Serwisowa posortowana.xlsx (np. Opis_Od_Klienta -> Opis_Zgloszenia, Kwota -> Kwota_Przewidywana)
    col_map = {}
    used_targets = set(df.columns)
    for c in df.columns:
        cl = c.lower()
        target = None
        if cl in ["dane", "klient", "imię i nazwisko", "imie i nazwisko"]:
            target = "Klient"
        elif "opis_od_klienta" in cl or "opis od klienta" in cl or cl == "opis":
            target = "Opis_Zgloszenia"
        elif "opis_serwisowy" in cl or "informacje_dodatkowe" in cl:
            target = "Notatka_Serwisanta"
        elif cl == "kwota":
            target = "Kwota_Przewidywana"
        elif "data_serwisu" in cl or "data serwisu" in cl:
            target = "Data_Planowana"
        elif "platnosc" in cl or "płatność" in cl:
            target = "Forma_Platnosci"
        if target and target not in used_targets:
            col_map[c] = target
            used_targets.add(target)

    if col_map:
        df = df.rename(columns=col_map)

    # Jeśli w pliku były osobne kolumny Opis_Od_Klienta i Opis_Serwisowy, połącz je, żeby nic nie zginęło
    if "Opis_Zgloszenia" not in df.columns and "Opis_Od_Klienta" in df.columns:
        df["Opis_Zgloszenia"] = df["Opis_Od_Klienta"].fillna("").astype(str)
    if "Opis_Serwisowy" in df.columns:
        opis_serw = df["Opis_Serwisowy"].fillna("").astype(str)
        if "Opis_Zgloszenia" in df.columns:
            df["Opis_Zgloszenia"] = (df["Opis_Zgloszenia"].fillna("").astype(str) + " " + opis_serw).str.strip()
        else:
            df["Opis_Zgloszenia"] = opis_serw

    for rc in REQUIRED_COLS:
        if rc not in df.columns:
            if rc == "ID_Zgloszenia":
                df[rc] = [f"ZGL-{i+1001}" for i in range(len(df))]
            elif rc == "Status":
                df[rc] = "Nowe"
            elif rc in ["Priorytet_Krytyczny", "Wymagany_Przeglad", "Zdjecia_Potwierdzone", "Podpis_Zlozony", "Do_Oddzwonienia", "Potwierdzony_Tel", "Trudny_Klient", "Otrzymano_Oferte_Termin"]:
                df[rc] = "Nie"
            elif rc in ["Naprawione_1_Wizyta", "Gotowka_Rozliczona"]:
                df[rc] = "Tak"
            elif rc == "Typ_Serwisu":
                df[rc] = "Płatny"
            elif rc == "Forma_Platnosci":
                df[rc] = "Przelew"
            elif rc in NUMERIC_COLS:
                df[rc] = 0.0
            else:
                df[rc] = ""

    df = df[REQUIRED_COLS].copy()
    for col in df.columns:
        if col in NUMERIC_COLS:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0).astype(float)
        else:
            df[col] = df[col].fillna("").astype(str).replace(["None", "nan", "NaT", "<NA>"], "")

    # Automatyczne wyciąganie numeru telefonu z pola Klient oraz kodu pocztowego i miasta z pola Adres!
    for idx, r in df.iterrows():
        raw_kl = str(r["Klient"]).strip()
        raw_tel = str(r["Telefon"]).strip()
        if not raw_tel and raw_kl:
            found_p = extract_phone_numbers("", raw_kl)
            if found_p:
                df.at[idx, "Telefon"] = " / ".join(found_p)
                clean_k = re.sub(r'(?:\+48)?\s*\d{3}[\s-]?\d{3}[\s-]?\d{3}|\d{2}[\s-]?\d{3}[\s-]?\d{2}[\s-]?\d{2}|\d{9,11}', '', raw_kl).strip(" ,-/")
                if clean_k:
                    df.at[idx, "Klient"] = clean_k

        raw_adr = str(r["Adres"]).strip()
        if not str(r["Kod_Pocztowy"]).strip() and raw_adr:
            m_z = re.search(r'\b(\d{2}-\d{3})\b', raw_adr)
            if m_z:
                df.at[idx, "Kod_Pocztowy"] = m_z.group(1)
        if not str(r["Miasto"]).strip() and raw_adr:
            adr_no_z = re.sub(r'\b\d{2}-\d{3}\b', '', raw_adr).strip(" ,")
            if "," in adr_no_z:
                df.at[idx, "Miasto"] = adr_no_z.split(",", 1)[0].strip()
            elif " ul" in adr_no_z.lower():
                df.at[idx, "Miasto"] = adr_no_z[:adr_no_z.lower().find(" ul")].strip(" ,")

        st_val = str(r["Status"]).strip().lower()
        if "zako" in st_val:
            df.at[idx, "Status"] = "Zakończone"
        elif "częśc" in st_val or "czesc" in st_val:
            df.at[idx, "Status"] = "Do zamówienia części"
        elif "oddeleg" in st_val:
            df.at[idx, "Status"] = "Oddelegowane"
        elif st_val in ["", "brak", "nan"]:
            df.at[idx, "Status"] = "Nowe"

    return df.reset_index(drop=True)

def load_tickets():
    df = pd.DataFrame()
    if os.path.exists(LOCAL_EXCEL_FILE):
        try:
            df = pd.read_excel(LOCAL_EXCEL_FILE)
        except Exception:
            pass
    return ensure_columns(df)

def save_tickets(df):
    clean_df = ensure_columns(df)
    st.session_state["df_tickets"] = clean_df.copy()
    try:
        clean_df.to_excel(LOCAL_EXCEL_FILE, index=False)
    except Exception:
        pass

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
                            if (today - datetime.date.fromisoformat(ost_przyp_str)).days < 3:
                                should_send = False
                        except Exception:
                            pass
                    if should_send and email_cfg.get("sender_email") and email_cfg.get("sender_password"):
                        subj = f"🚨 PRZYPOMNIENIE (Minęło {days_diff} dni): Brak oferty/terminu - {r['ID_Zgloszenia']} ({r['Klient']})"
                        body = "\n".join([
                            "Dzień dobry,", "",
                            f"System PROMAG OZE przypomina o oddelegowanym zleceniu z dnia {data_odd_str} ({dzial}), dla którego nadal nie zaznaczono przygotowania oferty lub terminu.",
                            f"- Klient: {r['Klient']} (Tel: {r['Telefon']})",
                            f"- Adres: {r['Kod_Pocztowy']} {r['Miasto']}, {r['Adres']}",
                            f"- Opis: {r['Opis_Zgloszenia']}",
                            f"- Notatka dla działu: {r['Notatka_Dla_Dzialu']}"
                        ])
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

# ==========================================
# 5. INICJALIZACJA SESJI + LOGOWANIE TYLKO KODEM PIN + ONBOARDING
# ==========================================
if "profiles" not in st.session_state:
    p_l, l_l, e_l, c_l, s_l, d_l = load_config_and_profiles()
    st.session_state["profiles"] = p_l
    st.session_state["logo_b64"] = l_l
    st.session_state["email_cfg"] = e_l
    st.session_state["cennik_tab"] = c_l
    st.session_state["stawka_rbh"] = s_l
    st.session_state["dojazd_kw"] = d_l

if "df_tickets" not in st.session_state:
    st.session_state["df_tickets"] = load_tickets()
else:
    st.session_state["df_tickets"] = ensure_columns(st.session_state["df_tickets"])

if "df_montaze" not in st.session_state:
    st.session_state["df_montaze"] = load_montaze_database()

for k_s in ["logged_role", "logged_user", "active_day", "active_ticket_id"]:
    if k_s not in st.session_state:
        st.session_state[k_s] = None

def resolve_pin_to_user(pin_code, profiles_dict):
    """Rozpoznaje rolę i użytkownika wyłącznie na podstawie unikalnego kodu PIN."""
    p_clean = str(pin_code).strip()
    if not p_clean:
        return None, None
    if p_clean == PIN_KOORDYNATOR:
        return "koordynator", "Koordynator"
    if p_clean == PIN_ZARZAD:
        return "zarzad", "Zarząd"
    for tech_name, t_data in profiles_dict.items():
        if str(t_data.get("pin", "")).strip() == p_clean:
            return "serwisant", tech_name
    return None, None

if st.session_state["logged_role"] is None:
    qp_token = st.query_params.get("token", None)
    if qp_token:
        r_found, u_found = resolve_pin_to_user(qp_token, st.session_state["profiles"])
        if r_found:
            st.session_state["logged_role"] = r_found
            st.session_state["logged_user"] = u_found

if st.session_state["logged_role"] is None:
    render_brand_logo(max_width=240, centered=True)

    col_login, col_onboard = st.columns([1, 1.35])
    with col_login:
        with st.container(border=True):
            st.subheader("🔐 Szybkie Logowanie PIN")
            st.caption("Wpisz swój unikalny 4-cyfrowy kod PIN – system sam rozpozna Twój profil i rolę.")
            pin_input = st.text_input("Twój 4-cyfrowy kod PIN:", type="password", max_chars=8, placeholder="••••")
            remember_me = st.checkbox("🔒 Zapamiętaj mnie na tym urządzeniu", value=True)

            if st.button("🌿 Wejdź do systemu", type="primary"):
                role_res, user_res = resolve_pin_to_user(pin_input, st.session_state["profiles"])
                if role_res:
                    st.session_state["logged_role"] = role_res
                    st.session_state["logged_user"] = user_res
                    st.session_state["active_day"] = None
                    st.session_state["active_ticket_id"] = None
                    if remember_me:
                        st.query_params["token"] = str(pin_input).strip()
                    st.rerun()
                else:
                    st.error("❌ Nieznany kod PIN. Sprawdź kod i spróbuj ponownie.")

    with col_onboard:
        with st.container(border=True):
            st.subheader("🎓 Onboarding i Standardy Pracy Serwisanta PROMAG OZE")
            st.markdown("""
            <div class="onboard-card">
                <b>1️⃣ Przed wyjazdem z bazy (Rano):</b><br>
                • Otwórz swój dzień trasy i sprawdź <b>Checklistę załadunku</b> wygenerowaną z opisów usterek.<br>
                • Upewnij się, że masz na aucie glikol, anody, czujniki NTC, naczynia przeponowe oraz sprawny miernik i refraktometr.
            </div>
            <div class="onboard-card">
                <b>2️⃣ Na miejscu u klienta (3 Kroki w aplikacji):</b><br>
                • <b>Krok 1 (Diagnoza):</b> Sprawdź historię montażu w karcie klienta. W razie błędu użyj wbudowanej <b>Szukajki PDF + Praktyka Google</b>.<br>
                • <b>Krok 2 (Części i Zdjęcia):</b> Zaznacz zużyte materiały z busa (automatycznie zejdą ze stanu i doliczą się do rachunku) oraz zrób <b>zdjęcie tabliczki znamionowej i instalacji</b>.<br>
                • <b>Krok 3 (Rozliczenie i Podpis):</b> Zbierz podpis klienta palcem na ekranie i oznacz formę płatności (Gotówka / Przelew / BLIK).
            </div>
            <div class="onboard-card">
                <b>3️⃣ System Prowizyjny i Leady:</b><br>
                • Twój pasek postępu do prowizji <b>3% (od 15 000 zł)</b> oraz <b>4% (od 30 000 zł)</b> aktualizuje się na żywo po każdej zamkniętej wizycie.<br>
                • Jeśli klient pyta o rozbudowę (np. magazyn energii, wymianę PC) – zaznacz <b>🔥 Lead Handlowy</b> w Kroku 3!
            </div>
            """, unsafe_allow_html=True)
    st.stop()

# Nagłówek zalogowanego użytkownika
top_c1, top_c2, top_c3 = st.columns([2.2, 1.5, 0.8])
with top_c1:
    render_brand_logo(max_width=190, centered=False)
with top_c2:
    st.markdown(f"**Zalogowany:** `{st.session_state['logged_user']}` ({st.session_state['logged_role'].capitalize()})")
with top_c3:
    if st.button("🚪 Wyloguj"):
        st.query_params.clear()
        st.session_state["logged_role"] = None
        st.session_state["logged_user"] = None
        st.rerun()

# ==========================================
# 6. PANEL SERWISANTA (3 KROKI + MULTI-SZUKAJKA + ONBOARDING POD RĘKĄ)
# ==========================================
if st.session_state["logged_role"] == "serwisant":
    tech_name = st.session_state["logged_user"]
    df = st.session_state["df_tickets"]
    profiles = st.session_state["profiles"]
    my_prof = profiles[tech_name]
    metrics = calculate_tech_metrics(df, tech_name)

    k_c1, k_c2, k_c3 = st.columns([1.5, 1.2, 1])
    with k_c1:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-title">🏆 Twoja Premia i Postęp Miesięczny ({metrics['active_rate_str']})</div>
            <div class="kpi-value">Wypracowana premia: {metrics['commission_pln']:,.0f} zł | Obrót: {metrics['monthly_revenue']:,.0f} zł</div>
            <div class="kpi-sub">Zrealizowane wizyty: {metrics['total_closed']} (Skuteczność 1. wizyty: {metrics['ftfr_pct']}%)</div>
        </div>
        """, unsafe_allow_html=True)
        st.progress(metrics["progress_pct"] / 100.0)

    with k_c2:
        st.markdown(f"""
        <div class="kpi-card-blue">
            <div class="kpi-title">💼 Rozliczenie Gotówki i Leady</div>
            <div class="kpi-value">Gotówka do zdania: {metrics['cash_in_hand']:,.0f} zł</div>
            <div class="kpi-sub">Przekazane Leady Handlowe: {metrics['leads_count']} szt.</div>
        </div>
        """, unsafe_allow_html=True)

    with k_c3:
        if st.button("🔍 SZUKAJKA BŁĘDÓW (PDF + GOOGLE)"):
            st.session_state["show_kb"] = not st.session_state.get("show_kb", False)
        if st.button("🚐 MÓJ MAGAZYN BUSA I NARZĘDZIA"):
            st.session_state["show_van"] = not st.session_state.get("show_van", False)

    low_stock = get_low_stock_items(my_prof)
    if low_stock:
        st.warning(f"🔔 **Przypomnienie magazynowe (<= 1 szt. na aucie):** Uzupełnij w bazie: **{', '.join(low_stock)}**")

    if st.session_state.get("show_kb", False):
        with st.container(border=True):
            st.subheader("🩺 Inteligentna Szukajka Techniczna (Instrukcje PDF + Praktyka / Google)")
            kb_t1, kb_t2, kb_t3 = st.tabs([
                "📚 1. Przeszukaj Instrukcje PDF i Kody Błędów",
                "🌐 2. Praktyka z Terenu i Wyszukiwarka Google",
                "📶 3. Konfiguracja Wi-Fi i Standardy"
            ])
            with kb_t1:
                q_pdf = st.text_input("Wpisz kod błędu lub hasło (np. F08, AL007, Riso, naczynie):", key="tech_q_pdf")
                if q_pdf.strip():
                    wyniki_pdf = przeszukaj_lokalne_instrukcje(q_pdf)
                    if wyniki_pdf:
                        for w in wyniki_pdf:
                            st.markdown(f"**{w['zrodlo']} — {w['naglowek']}**")
                            st.info(w["fragment"])
                    else:
                        st.warning("Brak bezpośredniego trafienia w lokalnych plikach PDF – sprawdź zakładkę 'Praktyka z Terenu i Google' obok!")
                st.divider()
                kb_cat = st.selectbox("Lub wybierz urządzenie z listy:", list(DIAGNOSTYKA_BAZA.keys()))
                kb_err = st.selectbox("Kod błędu:", list(DIAGNOSTYKA_BAZA[kb_cat].keys()))
                st.success(DIAGNOSTYKA_BAZA[kb_cat][kb_err])

            with kb_t2:
                q_web = st.text_input("Wpisz problem własnymi słowami (np. 'Deye F08 wyłącza backup' lub 'solary zapowietrzają się'):", value=q_pdf, key="tech_q_web")
                if q_web.strip():
                    porady, l_goog, l_elek, l_yt = przeszukaj_praktyke_i_internet(q_web)
                    if porady:
                        for p in porady:
                            st.success(f"💡 **{p['tytul']}**\n\n{p['opis']}")
                    c_g1, c_g2, c_g3 = st.columns(3)
                    with c_g1:
                        st.link_button("🌐 Szukaj rozwiązania w Google", l_goog)
                    with c_g2:
                        st.link_button("🛠️ Szukaj na Forum Elektroda.pl", l_elek)
                    with c_g3:
                        st.link_button("▶️ Zobacz wideo-instrukcję na YouTube", l_yt)

            with kb_t3:
                st.markdown(INSTRUKCJE_KONFIGURACJI_TEXT)

    if st.session_state.get("show_van", False):
        with st.container(border=True):
            st.subheader(f"🚐 Twój Samochód Serwisowy ({tech_name})")
            vc1, vc2, vc3 = st.columns([1.2, 1, 1])
            with vc1:
                st.markdown("**📦 Materiały na aucie:**")
                st.dataframe(pd.DataFrame(my_prof.get("materialy", []))[["nazwa", "ilosc", "cena_klienta"]], use_container_width=True, hide_index=True)
            with vc2:
                st.markdown("**🧰 Narzędzia na aucie:**")
                st.dataframe(pd.DataFrame(my_prof.get("narzedzia", [])), use_container_width=True, hide_index=True)
            with vc3:
                flota = my_prof.get("flota", {})
                new_odo = st.number_input("Przebieg licznika [km]:", min_value=0, step=100, value=int(flota.get("przebieg_km", 120000)))
                if st.button("💾 Zaktualizuj licznik"):
                    my_prof["flota"]["przebieg_km"] = int(new_odo)
                    save_config_and_profiles(profiles)
                    st.success("Zapisano przebieg!")

    st.divider()

    my_tickets = df[(df["Serwisant"] == tech_name) & (df["Status"].isin(["Zaplanowane", "W trakcie"]))].copy()
    if my_tickets.empty:
        st.success("🌿 Wszystkie zaplanowane na dziś zlecenia zostały wykonane!")
    else:
        today_str = str(datetime.date.today())
        my_tickets["Data_Grupa"] = my_tickets["Data_Planowana"].astype(str).str.strip().replace("", today_str)

        if st.session_state["active_day"] is None:
            st.subheader("📅 Wybierz dzień trasy")
            for day_str in sorted(my_tickets["Data_Grupa"].unique()):
                day_cluster = my_tickets[my_tickets["Data_Grupa"] == day_str]
                pts_count = len(day_cluster)
                cities_list = ", ".join([m for m in day_cluster["Miasto"].astype(str).unique() if m.strip()]) or "Trasa serwisowa"
                est_rev = float(day_cluster["Kwota_Przewidywana"].sum())
                is_today = (day_str == today_str)
                card_cls = "day-card day-card-today" if is_today else "day-card"

                opisy_dnia = " ".join(day_cluster["Opis_Zgloszenia"].astype(str).tolist()).lower()
                co_spakowac = [f"{m['nazwa']} (w busie: {m['ilosc']} szt.)" for m in my_prof.get("materialy", []) if str(m.get("slowo_klucz", "")).lower() in opisy_dnia]

                with st.container(border=True):
                    c_i, c_b = st.columns([2.3, 1])
                    with c_i:
                        st.markdown(f"""
                        <div class="{card_cls}">
                            <div style="font-size:18px; font-weight:800;">🗓️ Trasa: {day_str} {'(DZISIAJ)' if is_today else ''}</div>
                            <div style="font-size:15px; margin-top:4px;">📍 <b>Miejscowości:</b> {cities_list}</div>
                            <div style="font-size:14px; margin-top:4px;">🔧 Zleceń na trasie: <b>{pts_count}</b> | 💰 Wartość zleceń: <b>{est_rev:,.0f} zł</b></div>
                        </div>
                        """, unsafe_allow_html=True)
                        if co_spakowac:
                            st.info("📦 **Weź na auto przed wyjazdem:** " + ", ".join(co_spakowac))
                    with c_b:
                        st.write("")
                        if st.button(f"🚀 OTWÓRZ TRASĘ ({day_str})", key=f"open_d_{day_str}", type="primary"):
                            st.session_state["active_day"] = day_str
                            st.session_state["active_ticket_id"] = None
                            st.rerun()

        elif st.session_state["active_ticket_id"] is None:
            act_day = st.session_state["active_day"]
            day_tickets = my_tickets[my_tickets["Data_Grupa"] == act_day].copy()
            c_h1, c_h2 = st.columns([2.2, 1])
            with c_h1:
                st.subheader(f"🚚 Klienci na trasie w dniu: {act_day} ({len(day_tickets)} pkt)")
            with c_h2:
                if st.button("⬅️ Wróć do kalendarza dni"):
                    st.session_state["active_day"] = None
                    st.rerun()

            for idx_pt, (t_idx, t_row) in enumerate(day_tickets.iterrows(), 1):
                t_id = t_row["ID_Zgloszenia"]
                with st.container(border=True):
                    if str(t_row.get("Trudny_Klient", "Nie")) == "Tak":
                        st.error("🚩 **UWAGA: Wymagana płatność z góry / trudny klient!**")
                    st.markdown(f"### Punkt #{idx_pt}: **{t_row['Klient']}** | 📞 `{t_row['Telefon']}`")
                    st.markdown(f"📍 **Adres:** `{t_row['Kod_Pocztowy']} {t_row['Miasto']}, {t_row['Adres']}`")
                    if str(t_row.get("Opis_Instalacji_Montaz", "")).strip():
                        st.markdown(f"""
                        <div class="install-badge">
                            📋 <b>Zamontowana instalacja u klienta (Montaż: {t_row.get('Data_Montazu', 'Brak daty')}):</b><br>
                            {t_row['Opis_Instalacji_Montaz']}
                        </div>
                        """, unsafe_allow_html=True)
                    st.markdown(f"🔧 **Zadanie / Opis zgłoszenia:** {t_row['Opis_Zgloszenia']}")
                    if float(t_row.get("Zgoda_Klienta_Kwota", 0.0)) > 0:
                        st.success(f"✅ **Zgoda klienta z biura na kwotę:** `{float(t_row['Zgoda_Klienta_Kwota']):,.0f} zł`")

                    b_c1, b_c2, b_c3 = st.columns(3)
                    phones = extract_phone_numbers(str(t_row["Telefon"]), str(t_row["Klient"]))
                    with b_c1:
                        if phones:
                            st.link_button(f"📞 Zadzwoń ({phones[0]})", f"tel:{phones[0]}")
                    with b_c2:
                        maps_q = f"{t_row['Adres']} {t_row['Kod_Pocztowy']} {t_row['Miasto']}".replace(" ", "+")
                        st.link_button("🗺️ Prowadź w Google Maps", f"https://www.google.com/maps/search/?api=1&query={maps_q}")
                    with b_c3:
                        if st.button(f"🛠️ OTWÓRZ KARTĘ WIZYTY ({t_id})", key=f"go_t_{t_id}", type="primary"):
                            df.loc[t_idx, "Status"] = "W trakcie"
                            save_tickets(df)
                            st.session_state["active_ticket_id"] = t_id
                            st.rerun()

        else:
            selected_id = st.session_state["active_ticket_id"]
            matching_rows = df[df["ID_Zgloszenia"] == selected_id]
            if matching_rows.empty:
                st.session_state["active_ticket_id"] = None
                st.rerun()
            t_idx = matching_rows.index[0]
            t_row = df.loc[t_idx]

            top_b1, top_b2 = st.columns([2.3, 1])
            with top_b1:
                st.subheader(f"🛠️ Wizyta u klienta: {t_row['Klient']} ({selected_id})")
            with top_b2:
                if st.button("⬅️ Wróć do trasy dnia"):
                    st.session_state["active_ticket_id"] = None
                    st.rerun()

            step1, step2, step3 = st.tabs([
                "1️⃣ KROK 1: Dane Instalacji i Diagnoza",
                "2️⃣ KROK 2: Zużyte Części z Busa i Zdjęcia",
                "3️⃣ KROK 3: Rozliczenie i Podpis Klienta"
            ])

            with step1:
                st.markdown(f"📍 **Adres:** `{t_row['Kod_Pocztowy']} {t_row['Miasto']}, {t_row['Adres']}` | 📞 **Tel:** `{t_row['Telefon']}`")
                if str(t_row.get("Opis_Instalacji_Montaz", "")).strip():
                    st.markdown(f"""
                    <div class="install-badge">
                        📋 <b>Historia z bazy montaży (Data montażu: {t_row.get('Data_Montazu', 'Brak')}):</b><br>
                        {t_row['Opis_Instalacji_Montaz']}
                    </div>
                    """, unsafe_allow_html=True)
                st.info(f"📝 **Wytyczne i wycenione pozycje od Koordynatora:**\n\n{t_row['Opis_Zgloszenia']}")
                if float(t_row.get("Zgoda_Klienta_Kwota", 0.0)) > 0:
                    st.success(f"💰 **Klient zaakceptował telefonicznie koszt do: {float(t_row['Zgoda_Klienta_Kwota']):,.0f} zł**")

                st.markdown("#### 🔍 Szybka pomoc techniczna (PDF + Praktyka / Google)")
                q_diag_inline = st.text_input("Wpisz kod błędu lub objaw usterki:", key=f"in_q_{selected_id}")
                if q_diag_inline.strip():
                    res_local = przeszukaj_lokalne_instrukcje(q_diag_inline)
                    for r_l in res_local[:3]:
                        st.warning(f"**{r_l['zrodlo']} ({r_l['naglowek']}):**\n{r_l['fragment']}")
                    porady_in, l_g, l_e, _ = przeszukaj_praktyke_i_internet(q_diag_inline)
                    for p_in in porady_in:
                        st.success(f"💡 **{p_in['tytul']}:** {p_in['opis']}")
                    st.link_button("🌐 Otwórz wyniki w Google dla tego błędu", l_g)

            with step2:
                st.markdown("#### 🚐 Wybierz materiały zużyte z Twojego samochodu")
                van_materials = my_prof.get("materialy", [])
                mat_df = pd.DataFrame(van_materials)
                if not mat_df.empty:
                    mat_df["Zużyto_szt"] = 0
                    edited_mat_df = st.data_editor(
                        mat_df[["nazwa", "ilosc", "cena_klienta", "koszt_zakupu", "Zużyto_szt"]],
                        column_config={
                            "nazwa": st.column_config.TextColumn("Materiał na aucie", disabled=True),
                            "ilosc": st.column_config.NumberColumn("Stan w busie", disabled=True),
                            "cena_klienta": st.column_config.NumberColumn("Cena klienta (zł)", disabled=True),
                            "koszt_zakupu": st.column_config.NumberColumn("Koszt (zł)", disabled=True),
                            "Zużyto_szt": st.column_config.NumberColumn("Zużyto (szt.)", min_value=0, max_value=20, step=1)
                        },
                        use_container_width=True, hide_index=True, key=f"van_ed_{selected_id}"
                    )
                    used_materials_price = float((edited_mat_df["Zużyto_szt"] * edited_mat_df["cena_klienta"]).sum())
                    used_materials_cost = float((edited_mat_df["Zużyto_szt"] * edited_mat_df["koszt_zakupu"]).sum())
                    used_items_desc = [f"{r['nazwa']} x{int(r['Zużyto_szt'])}" for _, r in edited_mat_df.iterrows() if int(r["Zużyto_szt"]) > 0]
                else:
                    used_materials_price, used_materials_cost, used_items_desc = 0.0, 0.0, []

                st.divider()
                st.markdown("#### 📸 Zdjęcia z realizacji (Tabliczka znamionowa + Instalacja)")
                uploaded_photos = st.file_uploader("Dodaj zdjęcia z aparatu:", type=["png", "jpg", "jpeg"], accept_multiple_files=True, key=f"ph_{selected_id}")
                confirm_photos = st.checkbox("✅ Zdjęcia tabliczki oraz instalacji zostały załączone", key=f"cph_{selected_id}")

            with step3:
                cennik_slownik = pobierz_slownik_cennika(st.session_state["cennik_tab"], st.session_state["stawka_rbh"])
                wybrane_uslugi_tech = st.multiselect("Dodatkowe czynności z cennika (jeśli doszły na miejscu):", list(cennik_slownik.keys()), key=f"srv_t_{selected_id}")
                extra_srv_sum = sum(cennik_slownik[u]["cena_calkowita"] for u in wybrane_uslugi_tech)

                rc1, rc2 = st.columns(2)
                with rc1:
                    forma_plat = st.selectbox("Forma płatności:", ["Gotówka pobrana", "Przelew (Faktura)", "BLIK / Terminal", "Gwarancja (0 zł)"])
                    typ_serwisu_real = st.selectbox("Kwalifikacja:", ["Płatny", "Gwarancyjny"], index=0 if t_row["Typ_Serwisu"] != "Gwarancyjny" else 1)
                with rc2:
                    wynik_wizyty = st.radio("Wynik wizyty:", ["✅ Naprawione / Zakończone", "📦 Wymagane zamówienie części"])
                    work_minutes = st.number_input("Czas pracy na miejscu (min):", 15, 600, 90, 15)

                potrzebne_czesci_txt = ""
                if wynik_wizyty.startswith("📦"):
                    potrzebne_czesci_txt = st.text_input("Wpisz dokładny model części do zamówienia:")

                czy_lead = st.checkbox("🔥 Klient prosi o ofertę handlową (przekaż Lead do Handlowca)")
                lead_opis = st.text_input("Czego dotyczy oferta? (np. Magazyn energii, wymiana PC):") if czy_lead else ""

                base_expected = float(t_row["Kwota_Przewidywana"])
                suggested_total = 0.0 if (typ_serwisu_real == "Gwarancyjny" or forma_plat == "Gwarancja (0 zł)") else (base_expected + extra_srv_sum + used_materials_price)
                final_price = st.number_input("💰 Kwota końcowa do zapłaty przez klienta (zł):", 0.0, 50000.0, float(suggested_total), 50.0)
                internal_note = st.text_area("Twoja notatka po wizycie:", value=str(t_row["Notatka_Serwisanta"]))

                st.markdown('<div class="client-box">', unsafe_allow_html=True)
                st.markdown(f"#### 📄 Podpis Klienta — Kwota: **{final_price:,.2f} zł** ({forma_plat})")
                client_recommendations = st.text_input("Zalecenia dla klienta:", value=str(t_row["Zalecenia_Klienta"]))
                canvas_result = st_canvas(
                    fill_color="rgba(255, 255, 255, 0)", stroke_width=3, stroke_color="#0a2540",
                    background_color="#ffffff", height=150, width=580, drawing_mode="freedraw", key=f"canv_{selected_id}"
                )
                st.markdown('</div>', unsafe_allow_html=True)

                ch1, ch2 = st.columns(2)
                ch1.markdown("✅ **Zdjęcia potwierdzone**" if (confirm_photos and uploaded_photos) else "⬜ *Dodaj zdjęcie i zaznacz checkbox w Kroku 2*")
                ch2.markdown(f"✅ **Kwota rozliczenia: {final_price:,.0f} zł**")

                if st.button("🌿 ZAKOŃCZ WIZYTĘ I DOPISZ DO PREMII", type="primary"):
                    if not confirm_photos or not uploaded_photos:
                        st.error("Wróć do Kroku 2 i załącz min. 1 zdjęcie tabliczki/instalacji!")
                    elif wynik_wizyty.startswith("📦") and not potrzebne_czesci_txt.strip():
                        st.error("Wpisz nazwę części do zamówienia!")
                    else:
                        if not mat_df.empty:
                            for i_m, r_m in edited_mat_df.iterrows():
                                u_q = int(r_m["Zużyto_szt"])
                                if u_q > 0:
                                    my_prof["materialy"][i_m]["ilosc"] = max(0, int(my_prof["materialy"][i_m]["ilosc"]) - u_q)
                            save_config_and_profiles(profiles)

                        df.loc[t_idx, "Status"] = "Do zamówienia części" if wynik_wizyty.startswith("📦") else "Zakończone"
                        df.loc[t_idx, "Naprawione_1_Wizyta"] = "Nie" if wynik_wizyty.startswith("📦") else "Tak"
                        df.loc[t_idx, "Potrzebne_Czesci"] = potrzebne_czesci_txt.strip()
                        df.loc[t_idx, "Typ_Serwisu"] = typ_serwisu_real
                        df.loc[t_idx, "Forma_Platnosci"] = forma_plat
                        df.loc[t_idx, "Gotowka_Rozliczona"] = "Nie" if (forma_plat == "Gotówka pobrana" and final_price > 0) else "Tak"
                        df.loc[t_idx, "Kwota_Koncowa"] = float(final_price)
                        df.loc[t_idx, "Koszt_Materialow"] = float(used_materials_cost)
                        df.loc[t_idx, "Czas_Pracy_Min"] = float(work_minutes)
                        df.loc[t_idx, "Wyniki_Przegladu"] = "; ".join(used_items_desc + wybrane_uslugi_tech)
                        df.loc[t_idx, "Zdjecia_Potwierdzone"] = f"Tak ({len(uploaded_photos)} zdj.)"
                        df.loc[t_idx, "Zalecenia_Klienta"] = str(client_recommendations)
                        df.loc[t_idx, "Notatka_Serwisanta"] = str(internal_note)
                        if czy_lead and lead_opis.strip():
                            df.loc[t_idx, "Lead_Handlowy"] = lead_opis.strip()
                        df.loc[t_idx, "Data_Zamkniecia"] = str(datetime.date.today())
                        df.loc[t_idx, "Podpis_Zlozony"] = "Tak"
                        save_tickets(df)
                        st.session_state["active_ticket_id"] = None
                        st.success("Zlecenie zamknięte, magazyn busa zaktualizowany!")
                        st.rerun()

# ==========================================
# 7. PANEL KOORDYNATORA (PODPOWIEDZI NA ŻYWO PO NAZWISKU/TELEFONIE/ADRESIE + PEŁNA TABELA)
# ==========================================
elif st.session_state["logged_role"] == "koordynator":
    df = st.session_state["df_tickets"]
    df_montaze = st.session_state["df_montaze"]
    profiles = st.session_state["profiles"]
    email_cfg = st.session_state["email_cfg"]
    cennik_tab = st.session_state["cennik_tab"]
    stawka_rbh = st.session_state["stawka_rbh"]
    dojazd_kw = st.session_state["dojazd_kw"]

    overdue_delegated, sent_cnt = check_and_send_3day_reminders(df, email_cfg)
    if overdue_delegated:
        st.error(f"🚨 **PRZETERMINOWANE ODDELEGOWANIA (>3 DNI BEZ OFERTY/TERMINU):** `{', '.join(overdue_delegated)}`")

    for t_n, t_p in profiles.items():
        low_s = get_low_stock_items(t_p)
        if low_s:
            st.warning(f"🚐 **Uzupełnij auto serwisanta {t_n} (<=1 szt.):** {', '.join(low_s)}")

    now_time = datetime.datetime.now().time()
    today_iso = str(datetime.date.today())
    callbacks_df = df[df["Do_Oddzwonienia"] == "Tak"].copy()
    is_after_1530 = (now_time >= datetime.time(15, 30))

    if not callbacks_df.empty:
        header_cb = "🔔 KLIENCI DO ODDZWONIENIA (PO 15:30 - OBOWIĄZKOWA LISTA DO SKUTKU!)" if is_after_1530 else f"📞 Klienci do oddzwonienia ({len(callbacks_df)} poz. - główny blok o 15:30)"
        with st.expander(header_cb, expanded=(is_after_1530 or len(callbacks_df) > 0)):
            for cb_idx, cb_row in callbacks_df.iterrows():
                cb_id = cb_row["ID_Zgloszenia"]
                cb_phones = extract_phone_numbers(str(cb_row["Telefon"]), str(cb_row["Klient"]))
                cb_phone_main = cb_phones[0] if cb_phones else ""
                ost_wpis = cb_row["Notatka_Serwisanta"] or cb_row["Opis_Zgloszenia"]
                with st.container(border=True):
                    cc1, cc2, cc3 = st.columns([1.6, 1.2, 1.2])
                    with cc1:
                        st.markdown(f"🔴 **{cb_row['Klient']}** (`{cb_id}`) – 📞 `{cb_row['Telefon']}` – 📅 Data kontaktu: **{cb_row['Data_Oddzwonienia'] or today_iso}**")
                        st.markdown(f"📍 `{cb_row['Kod_Pocztowy']} {cb_row['Miasto']}, {cb_row['Adres']}`")
                        st.caption(f"💬 Ostatni wpis: {ost_wpis}")
                    with cc2:
                        if cb_phone_main:
                            st.link_button(f"📞 Zadzwoń: {cb_phone_main}", f"tel:{cb_phone_main}")
                            sms_m = "Dzien dobry, kontaktowalismy sie z Dzialu Serwisu PROMAG OZE. Prosimy o kontakt zwrotny."
                            st.link_button("📲 Wyślij SMS (Nie odbiera)", f"sms:{cb_phone_main}?body={sms_m.replace(' ', '%20')}")
                        nowa_notatka_cb = st.text_input("Ustalenia z rozmowy:", key=f"cb_n_{cb_id}")
                    with cc3:
                        if st.button("✅ Dodzwoniono (Zdejmij)", key=f"cb_ok_{cb_id}", type="primary"):
                            df.loc[cb_idx, "Do_Oddzwonienia"] = "Nie"
                            df.loc[cb_idx, "Potwierdzony_Tel"] = "Tak"
                            if nowa_notatka_cb.strip():
                                df.loc[cb_idx, "Notatka_Serwisanta"] = f"[{today_iso} Tel OK]: {nowa_notatka_cb.strip()} | " + str(df.loc[cb_idx, "Notatka_Serwisanta"])
                            save_tickets(df)
                            st.rerun()
                        if st.button("⏳ Nie odbiera (Zostaw do skutku)", key=f"cb_w_{cb_id}"):
                            stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
                            df.loc[cb_idx, "Notatka_Serwisanta"] = f"[Próba {stamp} - nie odbiera] " + str(df.loc[cb_idx, "Notatka_Serwisanta"])
                            save_tickets(df)
                            st.rerun()

    tab_klienci, tab_plan, tab_oddeleguj, tab_cennik_kadry, tab_ustawienia = st.tabs([
        "📋 1. Przyjęcie Zgłoszenia na Żywo + Lista Zgłoszeń (Start)",
        "🗺️ 2. Planer Korytarzowy i Czysty Zysk (4-5 zleceń / 1200 zł)",
        "📤 3. Oddelegowania (E-mail 3 dni), Części i Leady",
        "💰 4. Cennik (100 zł/h), Kadry, Narzędzia i Magazyn Busa",
        "⚙️ 5. Uproszczone Logo (Favicon), Baza Montaży i SMTP"
    ])

    # ----------------------------------------------------
    # ZAKŁADKA 1: JEDNOPOLOWA WYSZUKIWARKA NA ŻYWO (PO NAZWISKU, TELEFONIE I ADRESIE) + PEŁNA TABELA ZGŁOSZEŃ
    # ----------------------------------------------------
    with tab_klienci:
        st.subheader("➕ Szybkie Przyjęcie Zgłoszenia (Wyszukiwarka na żywo + Zintegrowany Kalkulator)")

        with st.container(border=True):
            st.markdown("#### Krok 1: Wyszukaj klienta na żywo (po Nazwisku, Numerze Telefonu lub Adresie)")
            st.caption("Kliknij w poniższe pole i zacznij pisać nazwisko, cyfry telefonu lub ulicę/miasto – lista podpowiedzi rozwija się i filtruje natychmiast w trakcie pisania!")

            # Budujemy zunifikowaną listę podpowiedzi na żywo (Nazwisko + Telefon + Adres) z obu baz
            @st.cache_data(show_spinner=False)
            def build_unified_search_options(montaze_df, tickets_df):
                opts = ["➕ Nowy klient (wpisz dane ręcznie poniżej)"]
                lookup = {}

                if tickets_df is not None and not tickets_df.empty:
                    for _, sr in tickets_df.iterrows():
                        kl = str(sr.get("Klient", "")).strip()
                        tel = str(sr.get("Telefon", "")).strip()
                        adr = f"{sr.get('Kod_Pocztowy', '')} {sr.get('Miasto', '')}, {sr.get('Adres', '')}".strip(" ,")
                        if kl or tel or adr:
                            label = f"🛠️ {kl} | 📞 {tel} | 📍 {adr} [{sr.get('ID_Zgloszenia', '')}]"
                            if label not in lookup:
                                opts.append(label)
                                lookup[label] = {
                                    "Klient": kl, "Telefon": tel,
                                    "Kod_Pocztowy": str(sr.get("Kod_Pocztowy", "")),
                                    "Miasto": str(sr.get("Miasto", "")),
                                    "Adres": str(sr.get("Adres", "")),
                                    "Data_Montazu": str(sr.get("Data_Montazu", "")),
                                    "Opis_Instalacji": str(sr.get("Opis_Instalacji_Montaz", "") or sr.get("Typ_Instalacji", ""))
                                }

                if montaze_df is not None and not montaze_df.empty:
                    for _, mr in montaze_df.iterrows():
                        kl = str(mr.get("Klient", "")).strip()
                        tel = str(mr.get("Telefon", "")).strip()
                        adr = f"{mr.get('Kod_Pocztowy', '')} {mr.get('Miasto', '')}, {mr.get('Adres', '')}".strip(" ,")
                        dm = str(mr.get("Data_Montazu", "")).strip()
                        label = f"🏠 {kl} | 📞 {tel} | 📍 {adr} | Montaż: {dm}"
                        if label not in lookup:
                            opts.append(label)
                            lookup[label] = {
                                "Klient": kl, "Telefon": tel,
                                "Kod_Pocztowy": str(mr.get("Kod_Pocztowy", "")),
                                "Miasto": str(mr.get("Miasto", "")),
                                "Adres": str(mr.get("Adres", "")),
                                "Data_Montazu": dm,
                                "Opis_Instalacji": str(mr.get("Opis_Instalacji", ""))
                            }
                return opts, lookup

            all_options, options_lookup = build_unified_search_options(df_montaze, df)

            chosen_client_option = st.selectbox(
                "🔍 Wpisz w tym polu Nazwisko, Telefon lub Adres klienta (podpowiedzi filtrują się na żywo poniżej):",
                options=all_options,
                index=0,
                key="live_autocomplete_client"
            )

            selected_prefill = options_lookup.get(chosen_client_option, {
                "Klient": "", "Telefon": "", "Kod_Pocztowy": "", "Miasto": "",
                "Adres": "", "Data_Montazu": "", "Opis_Instalacji": ""
            })

            if selected_prefill["Klient"] or selected_prefill["Telefon"]:
                st.markdown(f"""
                <div class="install-badge">
                    👤 <b>Wybrany klient:</b> <b>{selected_prefill['Klient']}</b> | 📞 <b>Telefon:</b> <code>{selected_prefill['Telefon'] or 'Brak'}</code> | 📍 <b>Adres:</b> <code>{selected_prefill['Kod_Pocztowy']} {selected_prefill['Miasto']}, {selected_prefill['Adres']}</code><br>
                    📅 <b>Data zakończenia montażu (Miesiąc i Rok):</b> <code>{selected_prefill['Data_Montazu'] or 'Brak w bazie'}</code><br>
                    📋 <b>Opis instalacji z tabeli montaży:</b> {selected_prefill['Opis_Instalacji'] or 'Brak szczegółowego opisu w bazie montaży'}
                </div>
                """, unsafe_allow_html=True)

            st.divider()
            st.markdown("#### Krok 2: Wybierz czynności z cennika (Kalkulator na żywo nalicza klienta i sam buduje notatkę dla serwisanta)")

            cennik_slownik = pobierz_slownik_cennika(cennik_tab, stawka_rbh)
            opcje_cennika = list(cennik_slownik.keys())

            k2_c1, k2_c2 = st.columns([2.2, 1])
            with k2_c1:
                wybrane_czynnosci = st.multiselect(
                    "🛠️ Wybierz pozycje do wyceny (dodanie pozycji zwiększa cenę i wpisuje ją do notatki; wykreślenie od razu odejmuje kwotę):",
                    options=opcje_cennika,
                    key="live_services_multiselect"
                )
            with k2_c2:
                czy_doliczyc_dojazd = st.checkbox(f"🚗 Dolicz dojazd z cennika (+{dojazd_kw:.0f} zł)", value=True)
                uwagi_dodatkowe = st.text_input("Dodatkowe uwagi od klienta (dopiszą się za czynnościami):", placeholder="np. cieknie przy grupie pompowej...")

            koszyk = oblicz_koszyk_uslug(
                wybrane_czynnosci, tabela=cennik_tab, stawka_rbh=stawka_rbh,
                dolicz_dojazd=czy_doliczyc_dojazd, kwota_dojazdu=dojazd_kw
            )

            auto_opis_pelny = koszyk["auto_notatka"]
            if uwagi_dodatkowe.strip():
                auto_opis_pelny = (auto_opis_pelny + " | UWAGI KLIENTA: " + uwagi_dodatkowe.strip()).strip(" |")

            m_k1, m_k2, m_k3, m_k4 = st.columns(4)
            m_k1.metric("Materiały z cennika", f"{koszyk['suma_materialow']:,.0f} zł")
            m_k2.metric(f"Robocizna ({koszyk['suma_czasu_h']}h x {stawka_rbh:.0f} zł)", f"{koszyk['suma_robocizny']:,.0f} zł")
            m_k3.metric("Opłata za dojazd", f"{koszyk['dojazd']:,.0f} zł")
            m_k4.metric("💰 RAZEM DLA KLIENTA", f"{koszyk['razem_dla_klienta']:,.0f} zł")

            st.divider()
            st.markdown("#### Krok 3: Dane zgłoszenia i zapis")

            f_c1, f_c2, f_c3 = st.columns(3)
            with f_c1:
                in_klient = st.text_input("Imię i nazwisko klienta:", value=selected_prefill["Klient"])
                in_tel = st.text_input("Telefon:", value=selected_prefill["Telefon"])
                in_kod = st.text_input("Kod pocztowy (XX-XXX):", value=selected_prefill["Kod_Pocztowy"])
            with f_c2:
                in_miasto = st.text_input("Miejscowość:", value=selected_prefill["Miasto"])
                in_adres = st.text_input("Ulica i nr domu:", value=selected_prefill["Adres"])
                in_data_montazu = st.text_input("Miesiąc i Rok montażu:", value=selected_prefill["Data_Montazu"])
            with f_c3:
                in_typ_inst = st.selectbox("Kategoria instalacji:", [
                    "Kolektory Słoneczne (Solar)", "Pompa Ciepła ARGOM", "Falownik Deye (PV / Hybryda)",
                    "Falownik GoodWe (PV / Hybryda)", "Magazyn Energii", "Pompa Ciepła + PV", "Inne"
                ])
                in_kryt = st.selectbox("Priorytet krytyczny (Awaria)?", ["Nie", "Tak"])
                in_trudny = st.selectbox("🚩 Trudny klient / Przedpłata?", ["Nie", "Tak"])

            in_opis_instalacji = st.text_input("📋 Opis instalacji z bazy montaży (widoczny w karcie klienta i u serwisanta):", value=selected_prefill["Opis_Instalacji"])
            in_notatka_konserwatora = st.text_area("📝 Notatka / Zakres prac przekazywany serwisantowi (uzupełnia się sama z kalkulatora):", value=auto_opis_pelny, height=85)

            b_k1, b_k2, b_k3, b_k4 = st.columns(4)
            with b_k1:
                in_kwota = st.number_input("Wycena zlecenia (zł):", 0.0, 50000.0, float(koszyk["razem_dla_klienta"]), 50.0)
            with b_k2:
                in_zgoda = st.number_input("✅ Zgoda klienta na kwotę (zł):", 0.0, 50000.0, float(koszyk["razem_dla_klienta"]), 50.0)
            with b_k3:
                in_oddzwonic = st.checkbox("📞 Do oddzwonienia (blok 15:30)?", value=False)
            with b_k4:
                in_data_oddzw = st.date_input("Data oddzwonienia:", value=datetime.date.today())

            if st.button("🌿 ZAPISZ ZGŁOSZENIE W SYSTEMIE", type="primary"):
                if not in_klient.strip():
                    st.error("Podaj imię i nazwisko klienta!")
                else:
                    new_id = f"ZGL-{len(df) + 1001}"
                    new_row = {
                        "ID_Zgloszenia": new_id, "Klient": in_klient.strip(), "Telefon": in_tel.strip(),
                        "Kod_Pocztowy": in_kod.strip(), "Miasto": in_miasto.strip(), "Adres": in_adres.strip(),
                        "Typ_Instalacji": in_typ_inst, "Data_Montazu": in_data_montazu.strip(),
                        "Opis_Instalacji_Montaz": in_opis_instalacji.strip(),
                        "Wymagany_Przeglad": "Tak" if "Przegląd" in in_notatka_konserwatora else "Nie",
                        "Opis_Zgloszenia": in_notatka_konserwatora.strip(),
                        "Priorytet_Krytyczny": in_kryt,
                        "Typ_Serwisu": "Gwarancyjny" if float(in_kwota) == 0 else "Płatny",
                        "Serwisant": "", "Data_Planowana": "", "Status": "Nowe", "Naprawione_1_Wizyta": "Tak",
                        "Kwota_Przewidywana": float(in_kwota), "Zgoda_Klienta_Kwota": float(in_zgoda),
                        "Kwota_Koncowa": 0.0, "Koszt_Materialow": float(koszyk["suma_materialow"]),
                        "Forma_Platnosci": "Przelew", "Gotowka_Rozliczona": "Tak", "Czas_Pracy_Min": float(koszyk["suma_czasu_h"] * 60),
                        "Wyniki_Przegladu": "", "Zdjecia_Potwierdzone": "Nie", "Zalecenia_Klienta": "",
                        "Notatka_Serwisanta": "", "Data_Zamkniecia": "", "Podpis_Zlozony": "Nie",
                        "Do_Oddzwonienia": "Tak" if in_oddzwonic else "Nie",
                        "Data_Oddzwonienia": str(in_data_oddzw) if in_oddzwonic else "",
                        "Potwierdzony_Tel": "Nie" if in_oddzwonic else "Tak",
                        "Trudny_Klient": in_trudny, "Potrzebne_Czesci": "", "Lead_Handlowy": "",
                        "Oddelegowano_Dzial": "", "Oddelegowano_Email": "", "Data_Oddelegowania": "",
                        "Notatka_Dla_Dzialu": "", "Otrzymano_Oferte_Termin": "Nie", "Ostatnie_Przypomnienie": ""
                    }
                    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                    save_tickets(df)
                    st.success(f"✅ Dodano zgłoszenie {new_id} dla klienta {in_klient} (Wycena: {in_kwota:.0f} zł)!")
                    st.rerun()

        st.divider()
        st.subheader("📋 Lista Aktywnych Zgłoszeń Serwisowych (Pełne Dane: Klient, Telefon, Adres, Treść Zgłoszenia)")

        flt1, flt2, flt3 = st.columns([1.4, 1, 1])
        with flt1:
            q_table_filter = st.text_input("🔍 Filtruj tabelę na żywo (wpisz nazwisko, telefon, adres lub treść usterki):", placeholder="Szukaj w zgłoszeniach...")
        with flt2:
            stat_filter = st.multiselect("Pokaż statusy:", ["Nowe", "Zaplanowane", "W trakcie", "Do zamówienia części", "Oddelegowane", "Zakończone"], default=["Nowe", "Zaplanowane", "W trakcie", "Do zamówienia części", "Oddelegowane"])
        with flt3:
            tech_filter = st.multiselect("Serwisant:", [""] + list(profiles.keys()))

        view_tickets = df.copy()
        if stat_filter:
            view_tickets = view_tickets[view_tickets["Status"].isin(stat_filter)]
        if tech_filter:
            view_tickets = view_tickets[view_tickets["Serwisant"].isin(tech_filter)]
        if q_table_filter.strip():
            qf = q_table_filter.strip()
            m_f = (
                view_tickets["Klient"].astype(str).str.contains(qf, case=False, na=False) |
                view_tickets["Telefon"].astype(str).str.contains(qf, case=False, na=False) |
                view_tickets["Adres"].astype(str).str.contains(qf, case=False, na=False) |
                view_tickets["Miasto"].astype(str).str.contains(qf, case=False, na=False) |
                view_tickets["Opis_Zgloszenia"].astype(str).str.contains(qf, case=False, na=False) |
                view_tickets["ID_Zgloszenia"].astype(str).str.contains(qf, case=False, na=False)
            )
            view_tickets = view_tickets[m_f]

        # Kluczowe kolumny ustawione na samym przodzie tabeli!
        priority_display_cols = [
            "ID_Zgloszenia", "Klient", "Telefon", "Kod_Pocztowy", "Miasto", "Adres",
            "Typ_Instalacji", "Opis_Zgloszenia", "Kwota_Przewidywana", "Zgoda_Klienta_Kwota",
            "Status", "Serwisant", "Data_Planowana", "Data_Montazu", "Opis_Instalacji_Montaz"
        ]
        avail_cols = [c for c in priority_display_cols if c in view_tickets.columns]
        st.dataframe(
            view_tickets[avail_cols],
            column_config={
                "ID_Zgloszenia": st.column_config.TextColumn("Nr Zgł."),
                "Klient": st.column_config.TextColumn("Nazwisko i Imię Klienta"),
                "Telefon": st.column_config.TextColumn("Telefon"),
                "Kod_Pocztowy": st.column_config.TextColumn("Kod"),
                "Miasto": st.column_config.TextColumn("Miasto"),
                "Adres": st.column_config.TextColumn("Ulica i nr"),
                "Typ_Instalacji": st.column_config.TextColumn("Instalacja"),
                "Opis_Zgloszenia": st.column_config.TextColumn("Treść Zgłoszenia / Opis Usterki", width="large"),
                "Kwota_Przewidywana": st.column_config.NumberColumn("Wycena (zł)"),
                "Zgoda_Klienta_Kwota": st.column_config.NumberColumn("Zgoda (zł)"),
                "Status": st.column_config.TextColumn("Status"),
                "Serwisant": st.column_config.TextColumn("Serwisant"),
                "Data_Planowana": st.column_config.TextColumn("Data Serwisu")
            },
            use_container_width=True,
            hide_index=True
        )

    # ----------------------------------------------------
    # ZAKŁADKA 2: PLANER KORYTARZOWY (CEL: 4-5 ZLECEŃ DZIENNIE / 1200 ZŁ CZYSTEGO ZYSKU)
    # ----------------------------------------------------
    with tab_plan:
        st.subheader("🗺️ Planer Korytarzowy („Po Drodze”) – Model 4–5 Zleceń Dziennie & Czysty Zysk")
        pc1, pc2, pc3, pc4, pc5 = st.columns(5)
        with pc1:
            sel_tech_plan = st.selectbox("Serwisant:", list(profiles.keys()), key="cor_tech")
        with pc2:
            start_postal = st.text_input("Baza Startowa:", value="05-552")
        with pc3:
            target_postal = st.text_input("Kod Docelowy (Cel):", value="05-800")
        with pc4:
            corridor_width_km = st.slider("Szerokość korytarza [km]:", 10, 150, 40, 5)
        with pc5:
            plan_date = st.date_input("Data trasy:", value=datetime.date.today() + datetime.timedelta(days=1))

        start_lat, start_lon = get_coords_for_postal(start_postal)
        dest_lat, dest_lon = get_coords_for_postal(target_postal)

        pending_df = df[df["Status"].isin(["Nowe", "Do zaplanowania", ""])].copy()
        if not pending_df.empty:
            dist_cor, dist_base, coords_l = [], [], []
            for _, r in pending_df.iterrows():
                rlat, rlon = extract_coords_from_row(r)
                coords_l.append((rlat, rlon))
                dist_base.append(round(haversine_km(start_lat, start_lon, rlat, rlon) * ROAD_TORTUOSITY, 1))
                dist_cor.append(round(distance_from_segment_km(rlat, rlon, start_lat, start_lon, dest_lat, dest_lon), 1))
            pending_df["Km_od_Osi_Trasy"] = dist_cor
            pending_df["Km_od_Bazy"] = dist_base
            pending_df["_coords"] = coords_l

            filtered_plan = pending_df[pending_df["Km_od_Osi_Trasy"] <= corridor_width_km].sort_values(by=["Priorytet_Krytyczny", "Km_od_Bazy"], ascending=[False, True]).copy()
            if not filtered_plan.empty:
                disp_cols = ["ID_Zgloszenia", "Km_od_Osi_Trasy", "Km_od_Bazy", "Klient", "Telefon", "Kod_Pocztowy", "Miasto", "Adres", "Kwota_Przewidywana", "Opis_Zgloszenia"]
                ed_df = filtered_plan[disp_cols].copy()
                ed_df.insert(0, "Wybierz", False)
                edited_route = st.data_editor(ed_df, use_container_width=True, hide_index=True, key="cor_ed")

                chosen_ids = edited_route[edited_route["Wybierz"] == True]["ID_Zgloszenia"].tolist()
                chosen_rows = filtered_plan[filtered_plan["ID_Zgloszenia"].isin(chosen_ids)]
                gross_rev = float(chosen_rows["Kwota_Przewidywana"].sum())
                parts_c = float(chosen_rows["Koszt_Materialow"].sum())

                econ = estimate_route_economics(
                    chosen_rows["_coords"].tolist(), (start_lat, start_lon), gross_rev, parts_cost=parts_c,
                    hourly_rate=stawka_rbh, fuel_price=KOSZT_PALIWA_LITR,
                    burn_l_100km=DEFAULT_BURN_L_100KM, onsite_hours_per_client=SR_CZAS_WIZYTY_H
                )

                ek1, ek2, ek3, ek4, ek5 = st.columns(5)
                ek1.metric("Liczba punktów (Cel: 4-5)", f"{len(chosen_ids)} pkt", "Optymalnie!" if len(chosen_ids) >= 4 else "Dobierz do 4-5 pkt")
                ek2.metric("Przychód z trasy (Do prowizji)", f"{gross_rev:,.0f} zł", f"Przy 20 dniach: {gross_rev*20:,.0f} zł/mies.")
                ek3.metric("Dystans i Paliwo (8 zł/l)", f"{econ['road_km']} km (-{econ['fuel_cost']:,.0f} zł)")
                ek4.metric(f"Koszt czasu ({stawka_rbh:.0f} zł/h)", f"-{(econ['labor_drive_cost'] + econ['labor_onsite_cost']):,.0f} zł")
                ek5.metric("💎 CZYSTY ZYSK NETTO", f"{econ['net_profit']:,.0f} zł", delta=f"{econ['net_profit'] - TARGET_NET_PROFIT:,.0f} zł vs 1200 zł")

                if len(chosen_ids) > 0:
                    if st.button(f"🚀 PRZYPISZ TRASĘ ({len(chosen_ids)} zleceń) DO: {sel_tech_plan} NA DZIEŃ {plan_date}", type="primary"):
                        for cid in chosen_ids:
                            idx_m = df[df["ID_Zgloszenia"] == cid].index[0]
                            df.loc[idx_m, "Serwisant"] = sel_tech_plan
                            df.loc[idx_m, "Data_Planowana"] = str(plan_date)
                            df.loc[idx_m, "Status"] = "Zaplanowane"
                        save_tickets(df)
                        st.success("Trasa zaplanowana!")
                        st.rerun()

    # ----------------------------------------------------
    # ZAKŁADKA 3: ODDELEGOWANIA (EMAIL + 3 DNI WATCHDOG), CZĘŚCI I LEADY
    # ----------------------------------------------------
    with tab_oddeleguj:
        st.subheader("📤 Oddelegowanie Zlecenia do Działu Handlowego / Montażu / Specjalisty")
        del_c1, del_c2 = st.columns([1.2, 1])
        with del_c1:
            if not df.empty:
                del_opts = [f"{r['ID_Zgloszenia']} | {r['Klient']} | Tel: {r['Telefon']} ({r['Miasto']})" for _, r in df.iterrows()]
                sel_del_str = st.selectbox("Wybierz zlecenie:", del_opts)
                sel_del_id = sel_del_str.split(" | ")[0]
                d_idx = df[df["ID_Zgloszenia"] == sel_del_id].index[0]
                d_row = df.loc[d_idx]

                dzial_choice = st.radio("Dział docelowy:", ["Dział Handlowy", "Dział Montażu", "Specjalista Techniczny"], horizontal=True)
                target_email = ""
                if dzial_choice == "Dział Handlowy":
                    h_list = email_cfg.get("handlowcy", [])
                    h_opts = [f"{h['nazwa']} <{h['email']}>" for h in h_list]
                    sel_h = st.selectbox("Wybierz Handlowca z listy:", h_opts) if h_opts else ""
                    if "<" in sel_h:
                        target_email = sel_h.split("<")[1].split(">")[0]
                elif dzial_choice == "Dział Montażu":
                    target_email = email_cfg.get("email_montaz", "")
                else:
                    target_email = email_cfg.get("email_technik", "")

                target_email_in = st.text_input("Adres e-mail odbiorcy:", value=target_email)
                notatka_dzial = st.text_area("Notatka / szczegóły dla działu:", value=str(d_row.get("Notatka_Dla_Dzialu", "")))

                if st.button("📧 WYŚLIJ E-MAIL I ODDELEGUJ (WATCHDOG 3 DNI)", type="primary"):
                    subj = f"[PROMAG SERWIS -> {dzial_choice.upper()}] {sel_del_id} - {d_row['Klient']}"
                    body = "\n".join([
                        f"Przekazano zgłoszenie {sel_del_id} do działu: {dzial_choice}",
                        f"Klient: {d_row['Klient']} | Tel: {d_row['Telefon']}",
                        f"Adres: {d_row['Kod_Pocztowy']} {d_row['Miasto']}, {d_row['Adres']}",
                        f"Montaż: {d_row.get('Data_Montazu', '')} | Instalacja: {d_row.get('Opis_Instalacji_Montaz', '')}",
                        f"Opis zgłoszenia: {d_row['Opis_Zgloszenia']}",
                        f"Notatka dla działu: {notatka_dzial}"
                    ])
                    df.loc[d_idx, "Status"] = "Oddelegowane"
                    df.loc[d_idx, "Oddelegowano_Dzial"] = dzial_choice
                    df.loc[d_idx, "Oddelegowano_Email"] = target_email_in
                    df.loc[d_idx, "Data_Oddelegowania"] = today_iso
                    df.loc[d_idx, "Notatka_Dla_Dzialu"] = notatka_dzial
                    df.loc[d_idx, "Otrzymano_Oferte_Termin"] = "Nie"
                    save_tickets(df)
                    ok, msg_s = send_system_email(email_cfg, target_email_in, subj, body)
                    st.info(f"Zapisano oddelegowanie (Watchdog 3-dniowy aktywny). Status SMTP: {msg_s}")

        with del_c2:
            st.markdown("#### ⏳ Oddelegowane zlecenia (Oczekujące na ofertę/termin)")
            del_df = df[df["Oddelegowano_Dzial"].astype(str).str.strip() != ""]
            for di, dr in del_df.iterrows():
                did = dr["ID_Zgloszenia"]
                got_o = (dr["Otrzymano_Oferte_Termin"] == "Tak")
                with st.container(border=True):
                    st.markdown(f"**{did} – {dr['Klient']}** (`{dr['Oddelegowano_Dzial']}` -> `{dr['Oddelegowano_Email']}`)")
                    st.caption(f"Przekazano: {dr['Data_Oddelegowania']} | Status: {'🟢 Otrzymano ofertę/termin' if got_o else '🔴 Czeka na reakcję działu'}")
                    if not got_o:
                        if st.button(f"✅ Zaznacz: Otrzymano ofertę / termin ({did})", key=f"ok_del_{did}"):
                            df.loc[di, "Otrzymano_Oferte_Termin"] = "Tak"
                            save_tickets(df)
                            st.rerun()

    # ----------------------------------------------------
    # ZAKŁADKA 4: CENNIK (100 ZŁ/H), KADRY, NARZĘDZIA I MAGAZYN BUSA
    # ----------------------------------------------------
    with tab_cennik_kadry:
        st.subheader("💰 Zestandaryzowany Cennik Usług (Nazwa usługi | Cena materiałów | Czas pracy x Stawka godzinowa)")
        ck1, ck2 = st.columns(2)
        with ck1:
            new_stawka_rbh = st.number_input("Stawka godzinowa pracy serwisanta (zł netto / h):", 50.0, 500.0, float(stawka_rbh), 10.0)
        with ck2:
            new_dojazd_kw = st.number_input("Stała kwota za dojazd doliczana do zlecenia (zł):", 0.0, 1000.0, float(dojazd_kw), 20.0)

        cennik_df_view = pd.DataFrame(cennik_tab)
        cennik_df_view["Wyliczona_Cena_Uslugi_ZL"] = cennik_df_view.apply(lambda r: wylicz_cene_uslugi(r, new_stawka_rbh), axis=1)
        ed_cennik_df = st.data_editor(
            cennik_df_view,
            column_config={
                "usluga": st.column_config.TextColumn("Nazwa usługi / czynności"),
                "cena_materialow": st.column_config.NumberColumn("Cena materiałów (zł netto)", min_value=0.0, step=10.0),
                "czas_pracy_h": st.column_config.NumberColumn("Czas pracy (h)", min_value=0.0, step=0.5),
                "Wyliczona_Cena_Uslugi_ZL": st.column_config.NumberColumn("Cena pozycji (Materiały + h * Stawka)", disabled=True)
            },
            num_rows="dynamic", use_container_width=True, key="ed_cennik_table"
        )
        if st.button("💾 Zapisz Cennik i Stawki", type="primary"):
            clean_c_list = ed_cennik_df[["usluga", "cena_materialow", "czas_pracy_h"]].to_dict(orient="records")
            st.session_state["cennik_tab"] = clean_c_list
            st.session_state["stawka_rbh"] = float(new_stawka_rbh)
            st.session_state["dojazd_kw"] = float(new_dojazd_kw)
            save_config_and_profiles(profiles, cennik_tab=clean_c_list, stawka_rbh=new_stawka_rbh, dojazd_kw=new_dojazd_kw)
            st.success("Zaktualizowano cennik!")
            st.rerun()

        st.divider()
        st.subheader("👷 Kadry, Unikalne Kody PIN, Narzędzia i Stan Materiałowy na Samochodach")
        sel_t_k = st.selectbox("Wybierz serwisanta:", list(profiles.keys()))
        t_prof = profiles[sel_t_k]
        t_m = calculate_tech_metrics(df, sel_t_k)

        pin_c1, pin_c2 = st.columns([1, 2])
        with pin_c1:
            new_tech_pin = st.text_input(f"🔑 Unikalny kod PIN dla {sel_t_k}:", value=str(t_prof.get("pin", "")))
            if st.button(f"💾 Zapisz PIN ({sel_t_k})"):
                profiles[sel_t_k]["pin"] = new_tech_pin.strip()
                save_config_and_profiles(profiles)
                st.success("Zapisano nowy PIN!")

        if t_m["cash_in_hand"] > 0:
            if st.button(f"💵 Rozlicz gotówkę zdaną w biurze przez {sel_t_k} ({t_m['cash_in_hand']:,.0f} zł)"):
                df.loc[(df["Serwisant"] == sel_t_k) & (df["Forma_Platnosci"] == "Gotówka pobrana"), "Gotowka_Rozliczona"] = "Tak"
                save_tickets(df)
                st.rerun()

        ic1, ic2 = st.columns(2)
        with ic1:
            st.markdown(f"**📦 Stan materiałowy na aucie ({sel_t_k}):**")
            ed_m_k = st.data_editor(pd.DataFrame(t_prof.get("materialy", [])), num_rows="dynamic", use_container_width=True, key=f"km_{sel_t_k}")
            if st.button(f"💾 Zapisz materiały ({sel_t_k})"):
                profiles[sel_t_k]["materialy"] = ed_m_k.to_dict(orient="records")
                save_config_and_profiles(profiles)
                st.success("Zapisano!")
        with ic2:
            st.markdown(f"**🧰 Narzędzia przypisane ({sel_t_k}):**")
            ed_t_k = st.data_editor(pd.DataFrame(t_prof.get("narzedzia", [])), num_rows="dynamic", use_container_width=True, key=f"kt_{sel_t_k}")
            if st.button(f"💾 Zapisz narzędzia ({sel_t_k})"):
                profiles[sel_t_k]["narzedzia"] = ed_t_k.to_dict(orient="records")
                save_config_and_profiles(profiles)
                st.success("Zapisano!")

    # ----------------------------------------------------
    # ZAKŁADKA 5: UPROSZCZONE LOGO (FAVICON + APLIKACJA), BAZA MONTAŻY I SMTP
    # ----------------------------------------------------
    with tab_ustawienia:
        st.subheader("🎨 Uproszczone Logo Aplikacji i Ikona Paska Przeglądarki (Favicon)")
        st.caption("Wgraj tutaj uproszczone logo (PNG/JPG). System automatycznie ustawi je zarówno jako główne logo w nagłówku aplikacji, jak i ikonę na pasku karty przeglądarki!")
        up_logo = st.file_uploader("Wgraj uproszczone logo (PNG / JPG):", type=["png", "jpg", "jpeg"])
        if up_logo is not None:
            if st.button("💾 Zapisz uproszczone logo na pasku przeglądarki i w aplikacji", type="primary"):
                img_l = Image.open(up_logo)
                img_l.thumbnail((360, 120))
                img_l.save(LOGO_FILE, format="PNG")
                buf = BytesIO()
                img_l.save(buf, format="PNG")
                b64_str = base64.b64encode(buf.getvalue()).decode("utf-8")
                st.session_state["logo_b64"] = b64_str
                save_config_and_profiles(profiles, logo_b64=b64_str)
                st.success("Zapisano! Logo jest już widoczne na pasku przeglądarki oraz w nagłówku.")
                st.rerun()

        st.divider()
        st.subheader(f"🏠 Pełna Baza Montaży (`skonczone_clean.parquet` / `skonczone.csv`) — Załadowano rekordów: {len(df_montaze)}")
        up_montaze = st.file_uploader("📥 Wgraj plik `skonczone_clean.parquet`, `.csv` lub `.xlsx`:", type=["parquet", "csv", "xlsx"])
        if up_montaze is not None:
            if st.button("⚡ Przetwórz i połącz bazę wszystkich klientów", type="primary"):
                if up_montaze.name.endswith(".parquet"):
                    r_df = pd.read_parquet(up_montaze)
                elif up_montaze.name.endswith(".csv"):
                    try:
                        r_df = pd.read_csv(up_montaze, sep=None, engine="python", on_bad_lines="skip")
                    except Exception:
                        up_montaze.seek(0)
                        r_df = pd.read_csv(up_montaze, encoding="cp1250", sep=None, engine="python", on_bad_lines="skip")
                else:
                    r_df = pd.read_excel(up_montaze)
                c_df = clean_and_parse_montaze_df(r_df)
                st.session_state["df_montaze"] = c_df
                try:
                    c_df.to_parquet(MONTAZE_PARQUET_CACHE, index=False)
                except Exception:
                    pass
                st.cache_data.clear()
                st.success(f"Załadowano i oczyszczono {len(c_df)} klientów!")
                st.rerun()

        st.divider()
        st.subheader("📧 Konfiguracja E-mail (SMTP) i Listy Handlowców")
        ec1, ec2 = st.columns(2)
        with ec1:
            s_host = st.text_input("Serwer SMTP:", value=email_cfg.get("smtp_server", "smtp.gmail.com"))
            s_port = st.number_input("Port SMTP:", value=int(email_cfg.get("smtp_port", 587)))
            s_user = st.text_input("E-mail nadawcy:", value=email_cfg.get("sender_email", ""))
            s_pass = st.text_input("Hasło aplikacyjne:", value=email_cfg.get("sender_password", ""), type="password")
            e_mont = st.text_input("E-mail Działu Montażu:", value=email_cfg.get("email_montaz", ""))
            e_tech = st.text_input("E-mail Specjalisty Technicznego:", value=email_cfg.get("email_technik", ""))
            if st.button("💾 Zapisz SMTP"):
                email_cfg.update({"smtp_server": s_host, "smtp_port": int(s_port), "sender_email": s_user, "sender_password": s_pass, "email_montaz": e_mont, "email_technik": e_tech})
                st.session_state["email_cfg"] = email_cfg
                save_config_and_profiles(profiles, email_cfg=email_cfg)
                st.success("Zapisano SMTP!")
        with ec2:
            st.markdown("**💼 Lista Handlowców:**")
            ed_h = st.data_editor(pd.DataFrame(email_cfg.get("handlowcy", [])), num_rows="dynamic", use_container_width=True)
            if st.button("💾 Zapisz Handlowców"):
                email_cfg["handlowcy"] = ed_h.to_dict(orient="records")
                st.session_state["email_cfg"] = email_cfg
                save_config_and_profiles(profiles, email_cfg=email_cfg)
                st.success("Zapisano listę handlowców!")

# ==========================================
# 8. PANEL ZARZĄDU
# ==========================================
elif st.session_state["logged_role"] == "zarzad":
    df = st.session_state["df_tickets"]
    profiles = st.session_state["profiles"]
    st.subheader("📊 Panel Zarządu – Wyniki Finansowe, Prowizje i Kasa")
    rows_z = []
    for t_name, t_prof in profiles.items():
        tm = calculate_tech_metrics(df, t_name)
        base_s = float(t_prof.get("wynagrodzenie_bazowe", 6000.0))
        rows_z.append({
            "Serwisant": t_name, "Zamknięte wizyty": tm["total_closed"], "Skuteczność 1. wizyty": f"{tm['ftfr_pct']}%",
            "Leady Handlowe": tm["leads_count"], "Gotówka do zdania (zł)": tm["cash_in_hand"],
            "Obrót miesięczny (zł)": tm["monthly_revenue"], "Prowizja (zł)": tm["commission_pln"],
            "Łącznie do wypłaty (zł)": base_s + tm["commission_pln"]
        })
    st.dataframe(pd.DataFrame(rows_z), use_container_width=True, hide_index=True)
