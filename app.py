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

try:
    from cennik import (
        STAWKA_RBH_NETTO, KOSZT_PALIWA_LITR, SR_CZAS_WIZYTY_H,
        RYCZALT_DOJAZD, TABELA_CENNIKA, wylicz_cene_uslugi,
        pobierz_slownik_cennika, oblicz_koszyk_uslug
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
        COMMISSION_TIER_1, COMMISSION_RATE_1, COMMISSION_TIER_2, COMMISSION_RATE_2,
        DOMYSLNE_NARZEDZIA, DOMYSLNE_MATERIALY, DEFAULT_TECHS,
        ensure_tech_profile_fields, get_low_stock_items, calculate_tech_metrics
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
        DIAGNOSTYKA_BAZA, INSTRUKCJE_KONFIGURACJI_TEXT, POSTAL_PREFIX_COORDS,
        przeszukaj_lokalne_instrukcje, przeszukaj_praktyke_i_internet
    )
except ImportError:
    DIAGNOSTYKA_BAZA = {"Brak pliku baza_wiedzy.py": {"Wgraj baza_wiedzy.py": "Wgraj plik na GitHub"}}
    INSTRUKCJE_KONFIGURACJI_TEXT = ""
    POSTAL_PREFIX_COORDS = {"05": (52.08, 20.85)}
    def przeszukaj_lokalne_instrukcje(q, f="."): return []
    def przeszukaj_praktyke_i_internet(q): return [], "https://google.com", "https://google.com", "https://youtube.com"

GOOGLE_SHEET_NAME = "Serwisowa posortowana"
LOCAL_EXCEL_FILE = "Serwisowa posortowana.xlsx"
MONTAZE_CSV_FILE = "skonczone.csv"
MONTAZE_PARQUET_CACHE = "skonczone_clean.parquet"
PROFILES_FILE = "profiles.json"
LOGO_FILE = "logo.png"

PIN_KOORDYNATOR = "1234"
PIN_ZARZAD = "9999"

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

MIESIACE_PL = {
    1: "Styczeń", 2: "Luty", 3: "Marzec", 4: "Kwiecień",
    5: "Maj", 6: "Czerwiec", 7: "Lipiec", 8: "Sierpień",
    9: "Wrzesień", 10: "Październik", 11: "Listopad", 12: "Grudzień"
}

def get_early_favicon():
    if os.path.exists(LOGO_FILE):
        try: return Image.open(LOGO_FILE)
        except Exception: pass
    if os.path.exists(PROFILES_FILE):
        try:
            with open(PROFILES_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
                b64 = saved.get("logo_b64", None)
                if b64: return Image.open(BytesIO(base64.b64decode(b64)))
        except Exception: pass
    return "🌿"

st.set_page_config(
    page_title="PROMAG OZE - Serwis",
    page_icon=get_early_favicon(),
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    .stButton > button { width: 100%; min-height: 48px; font-size: 15px !important; font-weight: 700 !important; border-radius: 10px !important; margin-bottom: 6px; }
    .promag-header { background: linear-gradient(135deg, #0a2540 0%, #10375c 100%); color: #ffffff; padding: 16px 22px; border-radius: 14px; border-bottom: 4px solid #16a34a; margin-bottom: 14px; }
    .kpi-card { background: linear-gradient(135deg, #0a2540 0%, #154360 100%); color: #ffffff; padding: 16px; border-radius: 14px; border-left: 6px solid #22c55e; margin-bottom: 10px; }
    .kpi-card-blue { background: linear-gradient(135deg, #10375c 0%, #1b4f72 100%); color: #ffffff; padding: 16px; border-radius: 14px; border-left: 6px solid #38bdf8; margin-bottom: 10px; }
    .kpi-title { font-size: 12px; text-transform: uppercase; letter-spacing: 0.6px; color: #cbd5e1; }
    .kpi-value { font-size: 20px; font-weight: 800; color: #ffffff; margin-top: 2px; }
    .kpi-sub { font-size: 13px; color: #86efac; margin-top: 4px; }
    .day-card { background: #ffffff; color: #0a2540; border: 1px solid #cbd5e1; border-radius: 14px; padding: 16px; margin-bottom: 10px; border-left: 7px solid #0284c7; }
    .day-card-today { background: #f0fdf4; color: #064e3b; border: 1px solid #86efac; border-left: 7px solid #16a34a; }
    .install-badge { background-color: #f0f9ff; border: 1px solid #bae6fd; border-left: 5px solid #0284c7; color: #0c4a6e; padding: 12px; border-radius: 10px; margin-top: 8px; margin-bottom: 10px; font-size: 14px; }
    .client-box { background-color: #ffffff; border: 2px solid #16a34a; border-radius: 14px; padding: 18px; margin-top: 12px; margin-bottom: 12px; color: #0a2540; }
    .onboard-card { background: #f8fafc; border: 1px solid #cbd5e1; border-left: 6px solid #16a34a; border-radius: 12px; padding: 16px; margin-bottom: 12px; color: #0f172a; }
</style>
""", unsafe_allow_html=True)

def convert_week_or_date_to_month_year(raw_val):
    s = str(raw_val).strip()
    if not s or s.lower() in ["nan", "none", "nat", ""]: return "Brak daty"
    m_iso = re.search(r'\b(20\d{2})[-/.](0?[1-9]|1[0-2])[-/.](\d{1,2})\b', s)
    if m_iso: return f"{MIESIACE_PL.get(int(m_iso.group(2)), '')} {m_iso.group(1)} ({m_iso.group(1)}-{int(m_iso.group(2)):02d})"
    m_pl = re.search(r'\b(\d{1,2})[-/.](0?[1-9]|1[0-2])[-/.](20\d{2})\b', s)
    if m_pl: return f"{MIESIACE_PL.get(int(m_pl.group(2)), '')} {m_pl.group(3)} ({m_pl.group(3)}-{int(m_pl.group(2)):02d})"
    m_wk1 = re.search(r'\b([1-9]|[1-4]\d|5[0-3])\s*[/._-]\s*(20[1-3]\d)\b', s)
    if m_wk1:
        try:
            d = datetime.date.fromisocalendar(int(m_wk1.group(2)), min(52, max(1, int(m_wk1.group(1)))), 4)
            return f"{MIESIACE_PL[d.month]} {d.year} ({d.year}-{d.month:02d})"
        except Exception: return f"Rok {m_wk1.group(2)}"
    if s.isdigit() and 1 <= int(s) <= 53:
        d = datetime.date.fromisocalendar(2024, min(52, int(s)), 4)
        return f"{MIESIACE_PL[d.month]} (Tydz. {s})"
    return s

def extract_phone_numbers(raw_phone_str, raw_client_str=""):
    combined = f"{raw_phone_str} {raw_client_str}"
    candidates = re.findall(r'(?:\+48)?\s*\d{3}[\s-]?\d{3}[\s-]?\d{3}|\d{2}[\s-]?\d{3}[\s-]?\d{2}[\s-]?\d{2}|\d{9,11}', combined)
    phones = []
    for c in candidates:
        digits = re.sub(r'\D', '', c)
        if len(digits) >= 9:
            sn = digits[-9:]
            if sn not in phones: phones.append(sn)
    return phones

def clean_and_parse_montaze_df(raw_df):
    if raw_df is None or raw_df.empty:
        return pd.DataFrame(columns=["Klient", "Telefon", "Kod_Pocztowy", "Miasto", "Adres", "Data_Montazu", "Opis_Instalacji", "Search_Label"])
    expected = {"Klient", "Telefon", "Kod_Pocztowy", "Miasto", "Adres", "Data_Montazu", "Opis_Instalacji", "Search_Label"}
    if expected.issubset(set(raw_df.columns)):
        raw_df = raw_df.fillna("").copy()
        raw_df["Search_Label"] = raw_df["Klient"].astype(str) + " | Tel: " + raw_df["Telefon"].astype(str) + " | Adres: " + raw_df["Kod_Pocztowy"].astype(str) + " " + raw_df["Miasto"].astype(str) + " " + raw_df["Adres"].astype(str) + " | Montaż: " + raw_df["Data_Montazu"].astype(str)
        return raw_df
    raw_df = raw_df.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)
    if any(str(c).startswith("Unnamed:") for c in raw_df.columns[:5]) and len(raw_df) > 1:
        first_vals = [str(v).strip() for v in raw_df.iloc[0].values]
        if any("klient" in v.lower() or "adres" in v.lower() for v in first_vals):
            new_cols, counts = [], {}
            for idx, val in enumerate(raw_df.iloc[0].values):
                cv = re.sub(r'\s+', ' ', str(val).strip())
                if not cv or cv.lower() == 'nan': cv = f"Kolumna_{idx}"
                counts[cv] = counts.get(cv, 0) + 1
                new_cols.append(cv if counts[cv] == 1 else f"{cv}_{counts[cv]}")
            raw_df = raw_df.iloc[1:].reset_index(drop=True)
            raw_df.columns = new_cols
    cols = [re.sub(r'\s+', ' ', str(c).strip()) for c in raw_df.columns]
    raw_df.columns = cols
    if "Klient" in cols and any("Adres" in c for c in cols):
        adr_col = next((c for c in cols if "Adres" in c), "Klient")
        tydz_col = next((c for c in cols if "Tydzień realizacji" in c or "Tydzien realizacji" in c), None)
        plan_col = next((c for c in cols if "Planowany termin" in c), None)
        umowa_cols = [c for c in cols if any(k in c.lower() for k in ["data podpisania", "data wprowadzenia", "termin zakończenia"])]
        opis_cols = [c for c in cols if any(k in c.lower() for k in ["opis zestawu", "opis z umowy", "dodatkowe podzespoły", "uruchomione instalacje", "dodatkowe opisy", "nazwa instalatora"])]
        records = []
        for _, row in raw_df.iterrows():
            raw_kl = str(row.get("Klient", "")).strip()
            if not raw_kl or raw_kl.lower() in ["nan", "none", "klient"]: continue
            raw_adr = str(row.get(adr_col, "")).strip()
            if raw_adr.lower() in ["nan", "none"]: raw_adr = ""
            phones = extract_phone_numbers(raw_adr, raw_kl)
            tel_str = " / ".join(phones)
            klient_clean = re.sub(r'(?:\+48)?\s*\d{3}[\s-]?\d{3}[\s-]?\d{3}|\d{2}[\s-]?\d{3}[\s-]?\d{2}[\s-]?\d{2}|\d{9,11}', '', raw_kl).strip(" ,-/") or raw_kl
            zip_m = re.search(r'\b(\d{2}-\d{3})\b', raw_adr)
            kod_p = zip_m.group(1) if zip_m else ""
            adr_no_zip = raw_adr.replace(kod_p, "").strip(" ,")
            if "," in adr_no_zip:
                miasto, ulica = [p.strip() for p in adr_no_zip.split(",", 1)]
            elif " ul" in adr_no_zip.lower():
                idx_u = adr_no_zip.lower().find(" ul")
                miasto, ulica = adr_no_zip[:idx_u].strip(" ,"), adr_no_zip[idx_u:].strip(" ,")
            else:
                miasto, ulica = adr_no_zip, raw_adr
            umowa_str = " ".join([str(row.get(uc, "")) for uc in umowa_cols])
            yr_m = re.search(r'\b(20[0-2]\d)\b', umowa_str)
            yr_val = int(yr_m.group(1)) if yr_m else None
            wk_str = str(row.get(tydz_col, "") if tydz_col else "").strip()
            if not wk_str or wk_str.lower() in ["nan", "none"]: wk_str = str(row.get(plan_col, "") if plan_col else "").strip()
            wk_m = re.search(r'\b([1-9]|[1-4]\d|5[0-3])\b', wk_str)
            if wk_m and yr_val:
                d_iso = datetime.date.fromisocalendar(yr_val, min(52, max(1, int(wk_m.group(1)))), 4)
                data_m = f"{MIESIACE_PL[d_iso.month]} {d_iso.year} ({d_iso.year}-{d_iso.month:02d})"
            else:
                data_m = convert_week_or_date_to_month_year(umowa_str if yr_val else wk_str)
            op_list = [f"{oc}: {str(row.get(oc, '')).strip()}" for oc in opis_cols if str(row.get(oc, '')).strip() and str(row.get(oc, '')).strip().lower() not in ["nan", "none", "-", "0", "0,00 zł"]]
            records.append({
                "Klient": klient_clean, "Telefon": tel_str, "Kod_Pocztowy": kod_p,
                "Miasto": miasto, "Adres": ulica, "Data_Montazu": data_m,
                "Opis_Instalacji": " | ".join(op_list),
                "Search_Label": f"{klient_clean} | Tel: {tel_str} | Adres: {kod_p} {miasto}, {ulica} | Montaż: {data_m}"
            })
        return pd.DataFrame(records)
    return pd.DataFrame(columns=["Klient", "Telefon", "Kod_Pocztowy", "Miasto", "Adres", "Data_Montazu", "Opis_Instalacji", "Search_Label"])

@st.cache_data(show_spinner="Wczytywanie bazy montaży...")
def load_montaze_database():
    if os.path.exists(MONTAZE_PARQUET_CACHE):
        try:
            return clean_and_parse_montaze_df(pd.read_parquet(MONTAZE_PARQUET_CACHE))
        except Exception: pass
    for cand in ["skonczone_clean.parquet", MONTAZE_CSV_FILE, "Promag/skonczone.csv", "skończone.xlsx"]:
        if os.path.exists(cand):
            try:
                if cand.endswith(".parquet"): r_df = pd.read_parquet(cand)
                elif cand.endswith(".csv"):
                    try: r_df = pd.read_csv(cand, sep=None, engine="python", on_bad_lines="skip")
                    except Exception: r_df = pd.read_csv(cand, encoding="cp1250", sep=None, engine="python", on_bad_lines="skip")
                else: r_df = pd.read_excel(cand)
                cdf = clean_and_parse_montaze_df(r_df)
                try: cdf.to_parquet(MONTAZE_PARQUET_CACHE, index=False)
                except Exception: pass
                return cdf
            except Exception: pass
    return pd.DataFrame(columns=["Klient", "Telefon", "Kod_Pocztowy", "Miasto", "Adres", "Data_Montazu", "Opis_Instalacji", "Search_Label"])

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
    if df is None or df.empty:
        edf = pd.DataFrame(columns=REQUIRED_COLS)
        for nc in NUMERIC_COLS: edf[nc] = edf[nc].astype(float)
        return edf
    cc, ctk = [], []
    for idx, c in enumerate(df.columns):
        cs = str(c).strip()
        if cs != "" and not cs.lower().startswith("unnamed"):
            ctk.append(idx); cc.append(cs)
    df = df.iloc[:, ctk].copy()
    df.columns = cc
    df = df.loc[:, ~df.columns.duplicated()].copy().reset_index(drop=True)
    
    cmap = {}
    used_t = set(df.columns)
    for c in df.columns:
        cl = c.lower()
        t = None
        if cl in ["dane", "klient", "imię i nazwisko", "imie i nazwisko"]: t = "Klient"
        elif "opis_od_klienta" in cl or "opis od klienta" in cl or cl == "opis": t = "Opis_Zgloszenia"
        elif "opis_serwisowy" in cl or "informacje_dodatkowe" in cl: t = "Notatka_Serwisanta"
        elif cl == "kwota": t = "Kwota_Przewidywana"
        elif "data_serwisu" in cl or "data serwisu" in cl: t = "Data_Planowana"
        elif "platnosc" in cl or "płatność" in cl: t = "Forma_Platnosci"
        if t and t not in used_t: cmap[c] = t; used_t.add(t)
    if cmap: df = df.rename(columns=cmap)

    for rc in REQUIRED_COLS:
        if rc not in df.columns:
            if rc == "ID_Zgloszenia": df[rc] = [f"ZGL-{i+1001}" for i in range(len(df))]
            elif rc == "Status": df[rc] = "Nowe"
            elif rc in ["Priorytet_Krytyczny", "Wymagany_Przeglad", "Zdjecia_Potwierdzone", "Podpis_Zlozony", "Do_Oddzwonienia", "Potwierdzony_Tel", "Trudny_Klient", "Otrzymano_Oferte_Termin"]: df[rc] = "Nie"
            elif rc in ["Naprawione_1_Wizyta", "Gotowka_Rozliczona"]: df[rc] = "Tak"
            elif rc == "Typ_Serwisu": df[rc] = "Płatny"
            elif rc == "Forma_Platnosci": df[rc] = "Przelew"
            elif rc in NUMERIC_COLS: df[rc] = 0.0
            else: df[rc] = ""
    df = df[REQUIRED_COLS].copy()
    for col in df.columns:
        if col in NUMERIC_COLS: df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0).astype(float)
        else: df[col] = df[col].fillna("").astype(str).replace(["None", "nan", "NaT", "<NA>"], "")
    return df.reset_index(drop=True)

def load_tickets():
    if not os.path.exists(LOCAL_EXCEL_FILE):
        empty_df = pd.DataFrame(columns=REQUIRED_COLS)
        try: empty_df.to_excel(LOCAL_EXCEL_FILE, index=False)
        except Exception: pass
        return empty_df
    try:
        return ensure_columns(pd.read_excel(LOCAL_EXCEL_FILE))
    except Exception:
        return ensure_columns(pd.DataFrame(columns=REQUIRED_COLS))

def save_tickets(df):
    clean_df = ensure_columns(df)
    st.session_state["df_tickets"] = clean_df.copy()
    try:
        clean_df.to_excel(LOCAL_EXCEL_FILE, index=False)
    except Exception:
        pass

if "profiles" not in st.session_state:
    p_l, l_l, e_l, c_l, s_l, d_l = load_config_and_profiles() if 'load_config_and_profiles' in globals() else (DEFAULT_TECHS, None, DEFAULT_EMAIL_CONFIG, TABELA_CENNIKA, STAWKA_RBH_NETTO, RYCZALT_DOJAZD)
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
    p_clean = str(pin_code).strip()
    if not p_clean: return None, None
    if p_clean == PIN_KOORDYNATOR: return "koordynator", "Koordynator"
    if p_clean == PIN_ZARZAD: return "zarzad", "Zarząd"
    for tech_name, t_data in profiles_dict.items():
        if str(t_data.get("pin", "")).strip() == p_clean:
            return "serwisant", tech_name
    return None, None

if st.session_state["logged_role"] is None:
    qp_token = st.query_params.get("token", None)
    if qp_token:
        r_f, u_f = resolve_pin_to_user(qp_token, st.session_state["profiles"])
        if r_f: st.session_state["logged_role"] = r_f; st.session_state["logged_user"] = u_f

if st.session_state["logged_role"] is None:
    render_brand_logo(max_width=240, centered=True) if 'render_brand_logo' in globals() else st.title("PROMAG OZE")
    col_login, col_onboard = st.columns([1, 1.35])
    with col_login:
        with st.container(border=True):
            st.subheader("🔐 Szybkie Logowanie PIN")
            pin_input = st.text_input("Kod PIN:", type="password", max_chars=8, placeholder="••••")
            remember_me = st.checkbox("🔒 Zapamiętaj mnie", value=True)
            if st.button("🌿 Wejdź do systemu", type="primary"):
                r_res, u_res = resolve_pin_to_user(pin_input, st.session_state["profiles"])
                if r_res:
                    st.session_state["logged_role"] = r_res
                    st.session_state["logged_user"] = u_res
                    if remember_me: st.query_params["token"] = str(pin_input).strip()
                    st.rerun()
                else: st.error("❌ Nieznany kod PIN.")
    with col_onboard:
        with st.container(border=True):
            st.subheader("🎓 Onboarding Serwisanta PROMAG OZE")
            st.markdown("<div class='onboard-card'><b>1. Rano:</b> Sprawdź checklistę załadunku na busie.</div><div class='onboard-card'><b>2. U klienta:</b> Diagnoza, części i zdjęcia, podpis palcem.</div>", unsafe_allow_html=True)
    st.stop()

top_c1, top_c2, top_c3 = st.columns([2.2, 1.5, 0.8])
with top_c1: render_brand_logo(max_width=190, centered=False) if 'render_brand_logo' in globals() else st.write("")
with top_c2: st.markdown(f"**Zalogowany:** `{st.session_state['logged_user']}` ({st.session_state['logged_role'].capitalize()})")
with top_c3:
    if st.button("🚪 Wyloguj"):
        st.query_params.clear()
        st.session_state["logged_role"] = None
        st.session_state["logged_user"] = None
        st.rerun()

if st.session_state["logged_role"] == "koordynator":
    df = st.session_state["df_tickets"]
    df_montaze = st.session_state["df_montaze"]
    profiles = st.session_state["profiles"]
    email_cfg = st.session_state["email_cfg"]
    cennik_tab = st.session_state["cennik_tab"]
    stawka_rbh = st.session_state["stawka_rbh"]
    dojazd_kw = st.session_state["dojazd_kw"]

    tab_klienci, tab_plan, tab_oddeleguj, tab_cennik_kadry, tab_ustawienia = st.tabs([
        "📋 1. Przyjęcie Zgłoszenia na Żywo + Lista Zgłoszeń (Start)",
        "🗺️ 2. Planer Korytarzowy i Zysk (1200 zł)",
        "📤 3. Oddelegowania",
        "💰 4. Cennik i Kadry",
        "⚙️ 5. Logo i Baza"
    ])

    with tab_klienci:
        st.subheader("➕ Szybkie Przyjęcie Zgłoszenia (Wyszukiwarka na żywo + Kalkulator)")
        with st.container(border=True):
            st.markdown("#### Krok 1: Wyszukaj klienta na żywo (Nazwisko, Telefon lub Adres)")
            
            @st.cache_data(show_spinner=False)
            def build_live_options(m_df, t_df):
                opts = ["➕ Nowy klient (wpisz dane ręcznie poniżej)"]
                lookup = {}
                if t_df is not None and not t_df.empty:
                    for _, sr in t_df.iterrows():
                        kl = str(sr.get("Klient", "")).strip()
                        tel = str(sr.get("Telefon", "")).strip()
                        adr = f"{sr.get('Kod_Pocztowy', '')} {sr.get('Miasto', '')}, {sr.get('Adres', '')}".strip(" ,")
                        if kl or tel or adr:
                            lbl = f"🛠️ {kl} | Tel: {tel} | Adres: {adr} [{sr.get('ID_Zgloszenia', '')}]"
                            if lbl not in lookup:
                                opts.append(lbl)
                                lookup[lbl] = {"Klient": kl, "Telefon": tel, "Kod_Pocztowy": str(sr.get("Kod_Pocztowy", "")), "Miasto": str(sr.get("Miasto", "")), "Adres": str(sr.get("Adres", "")), "Data_Montazu": str(sr.get("Data_Montazu", "")), "Opis_Instalacji": str(sr.get("Opis_Instalacji_Montaz", "") or sr.get("Typ_Instalacji", ""))}
                if m_df is not None and not m_df.empty:
                    for _, mr in m_df.iterrows():
                        kl = str(mr.get("Klient", "")).strip()
                        tel = str(mr.get("Telefon", "")).strip()
                        adr = f"{mr.get('Kod_Pocztowy', '')} {mr.get('Miasto', '')}, {mr.get('Adres', '')}".strip(" ,")
                        dm = str(mr.get("Data_Montazu", "")).strip()
                        lbl = f"🏠 {kl} | Tel: {tel} | Adres: {adr} | Montaż: {dm}"
                        if lbl not in lookup:
                            opts.append(lbl)
                            lookup[lbl] = {"Klient": kl, "Telefon": tel, "Kod_Pocztowy": str(mr.get("Kod_Pocztowy", "")), "Miasto": str(mr.get("Miasto", "")), "Adres": str(mr.get("Adres", "")), "Data_Montazu": dm, "Opis_Instalacji": str(mr.get("Opis_Instalacji", ""))}
                return opts, lookup

            all_opts, opts_map = build_live_options(df_montaze, df)
            chosen_client = st.selectbox("🔍 Wybierz klienta z bazy (wpisuj nazwisko, telefon lub miasto):", options=all_opts, key="coord_live_sel")
            selected_prefill = opts_map.get(chosen_client, {"Klient": "", "Telefon": "", "Kod_Pocztowy": "", "Miasto": "", "Adres": "", "Data_Montazu": "", "Opis_Instalacji": ""})

            if selected_prefill["Klient"] or selected_prefill["Telefon"]:
                st.markdown(f"""
                <div class="install-badge">
                    👤 <b>Wybrany klient:</b> <b>{selected_prefill['Klient']}</b> | 📞 <b>Telefon:</b> <code>{selected_prefill['Telefon'] or 'Brak'}</code> | 📍 <b>Adres:</b> <code>{selected_prefill['Kod_Pocztowy']} {selected_prefill['Miasto']}, {selected_prefill['Adres']}</code><br>
                    📅 <b>Data montażu:</b> <code>{selected_prefill['Data_Montazu'] or 'Brak'}</code> | 📋 <b>Opis:</b> {selected_prefill['Opis_Instalacji'] or 'Brak'}
                </div>
                """, unsafe_allow_html=True)

            st.divider()
            cennik_slownik = pobierz_slownik_cennika(cennik_tab, stawka_rbh) if 'pobierz_slownik_cennika' in globals() else {}
            wybrane_c = st.multiselect("🛠️ Wybierz czynności z cennika:", options=list(cennik_slownik.keys()), key="coord_srv_sel")
            koszyk = oblicz_koszyk_uslug(wybrane_c, tabela=cennik_tab, stawka_rbh=stawka_rbh, dolicz_dojazd=True, kwota_dojazdu=dojazd_kw) if 'oblicz_koszyk_uslug' in globals() else {"suma_materialow": 0, "suma_czasu_h": 0, "suma_robocizny": 0, "dojazd": 0, "razem_dla_klienta": 0, "auto_notatka": ""}

            st.markdown(f"💰 **Szacowana kwota dla klienta:** `{koszyk['razem_dla_klienta']:,.0f} zł`")

            f1, f2, f3 = st.columns(3)
            with f1:
                in_k = st.text_input("Imię i nazwisko klienta:", value=selected_prefill["Klient"])
                in_t = st.text_input("Telefon:", value=selected_prefill["Telefon"])
            with f2:
                in_kod = st.text_input("Kod pocztowy:", value=selected_prefill["Kod_Pocztowy"])
                in_mia = st.text_input("Miejscowość:", value=selected_prefill["Miasto"])
            with f3:
                in_adr = st.text_input("Ulica i nr:", value=selected_prefill["Adres"])
                in_kw = st.number_input("Wycena zlecenia (zł):", 0.0, 50000.0, float(koszyk["razem_dla_klienta"]), 50.0)

            in_opis = st.text_area("📝 Opis usterki / notatka dla serwisanta:", value=koszyk["auto_notatka"], height=80)

            if st.button("🌿 ZAPISZ ZGŁOSZENIE", type="primary"):
                if not in_k.strip(): st.error("Podaj imię i nazwisko klienta!")
                else:
                    new_id = f"ZGL-{len(df) + 1001}"
                    new_row = {
                        "ID_Zgloszenia": new_id, "Klient": in_k.strip(), "Telefon": in_t.strip(),
                        "Kod_Pocztowy": in_kod.strip(), "Miasto": in_mia.strip(), "Adres": in_adr.strip(),
                        "Typ_Instalacji": "Pompa Ciepła / PV", "Data_Montazu": selected_prefill["Data_Montazu"],
                        "Opis_Instalacji_Montaz": selected_prefill["Opis_Instalacji"], "Opis_Zgloszenia": in_opis.strip(),
                        "Status": "Nowe", "Kwota_Przewidywana": float(in_kw), "Zgoda_Klienta_Kwota": float(in_kw),
                        "Typ_Serwisu": "Płatny", "Forma_Platnosci": "Przelew", "Gotowka_Rozliczona": "Tak"
                    }
                    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                    save_tickets(df)
                    st.success(f"✅ Zapisano zgłoszenie {new_id}!")
                    st.rerun()

        st.divider()
        st.subheader("📋 Lista Aktywnych Zgłoszeń Serwisowych")
        st.dataframe(df, use_container_width=True, hide_index=True)

    with tab_plan: st.subheader("Planer Korytarzowy")
    with tab_oddeleguj: st.subheader("Oddelegowania")
    with tab_cennik_kadry: st.subheader("Cennik i Kadry")
    with tab_ustawienia: st.subheader("Ustawienia i Logo")

elif st.session_state["logged_role"] == "serwisant":
    st.subheader(f"Panel Serwisanta - {st.session_state['logged_user']}")
    st.info("Pracujesz w trybie mobilnym serwisanta.")
