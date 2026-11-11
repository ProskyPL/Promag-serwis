import streamlit as st
import pandas as pd
import numpy as np
import datetime
import os
import json
import base64
import math
import re
from io import BytesIO
from PIL import Image
from streamlit_drawable_canvas import st_canvas
import gspread
from google.oauth2.service_account import Credentials

# ==========================================
# 1. KONFIGURACJA STRONY I RESPONSYWNY CSS (TELEFON + TABLET)
# ==========================================
st.set_page_config(
    page_title="PROMAG OZE - System Serwisowy",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    /* Optymalizacja pod ekrany dotykowe (Telefony i Tablety) */
    .stButton > button {
        width: 100%;
        min-height: 52px;
        font-size: 16px !important;
        font-weight: 600 !important;
        border-radius: 10px !important;
        margin-bottom: 6px;
        transition: all 0.2s;
    }
    .stSelectbox, .stTextInput, .stNumberInput, .stDateInput {
        font-size: 16px !important;
    }
    /* Kafelki motywacyjne i KPI */
    .kpi-card {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        color: #ffffff;
        padding: 16px;
        border-radius: 12px;
        border-left: 6px solid #10b981;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.15);
        margin-bottom: 12px;
    }
    .kpi-card-warn {
        border-left: 6px solid #f59e0b;
    }
    .kpi-title {
        font-size: 13px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        color: #94a3b8;
        margin-bottom: 4px;
    }
    .kpi-value {
        font-size: 24px;
        font-weight: 800;
        color: #f8fafc;
    }
    .kpi-sub {
        font-size: 13px;
        color: #38bdf8;
        margin-top: 4px;
    }
    /* Nagłówek protokołu klienta z brandingiem */
    .client-box {
        background-color: #f8fafc;
        border: 2px solid #cbd5e1;
        border-radius: 12px;
        padding: 18px;
        margin-top: 12px;
        margin-bottom: 12px;
        color: #0f172a;
    }
    @media (max-width: 768px) {
        .kpi-value { font-size: 20px; }
        .block-container { padding-top: 1.5rem !important; padding-left: 0.8rem !important; padding-right: 0.8rem !important; }
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. STAŁE, CENNIK I MAPA GEOGRAFICZNA POLSKI (DO KLASTRÓW <= 50 KM)
# ==========================================
GOOGLE_SHEET_NAME = "Serwisowa posortowana"
LOCAL_EXCEL_FILE = "Serwisowa posortowana.xlsx"
PROFILES_FILE = "profiles.json"
LOGO_FILE = "logo.png"

PIN_KOORDYNATOR = "1234"
PIN_ZARZAD = "9999"

DEFAULT_TECHS = {
    "Leszek": {
        "pin": "4829",
        "pelne_dane": "Leszek - Serwisant OZE",
        "telefon": "",
        "umowa": "Umowa o pracę",
        "wynagrodzenie_bazowe": 6500.0,
        "bhp_data": "2026-01-15",
        "bhp_protokol": "BHP/01/2026 - Ważne",
        "medycyna_data": "2026-01-10",
        "medycyna_zakres": "Praca powyżej 3m, pomiary elektryczne do 1kV, prawo jazdy kat. B",
        "urlopy_l4": [],
        "notatki_poufne": "Doświadczony w pompach ciepła i falownikach hybrydowych."
    },
    "Mikołaj": {
        "pin": "7315",
        "pelne_dane": "Mikołaj - Serwisant OZE",
        "telefon": "",
        "umowa": "B2B",
        "wynagrodzenie_bazowe": 7000.0,
        "bhp_data": "2026-02-01",
        "bhp_protokol": "BHP/02/2026 - Ważne",
        "medycyna_data": "2026-01-25",
        "medycyna_zakres": "Praca na wysokości pow. 3m, SEP E+D, F-gazy",
        "urlopy_l4": [],
        "notatki_poufne": ""
    },
    "Wojtek": {
        "pin": "5942",
        "pelne_dane": "Wojtek - Serwisant OZE",
        "telefon": "",
        "umowa": "Umowa o pracę",
        "wynagrodzenie_bazowe": 6000.0,
        "bhp_data": "2026-03-01",
        "bhp_protokol": "BHP/03/2026 - Ważne",
        "medycyna_data": "2026-02-20",
        "medycyna_zakres": "Praca na wysokości, SEP E, kat. B",
        "urlopy_l4": [],
        "notatki_poufne": ""
    },
    "Serwisant 4": {
        "pin": "8164",
        "pelne_dane": "Wakat (Serwisant 4) - zmień imię w panelu Kadry",
        "telefon": "",
        "umowa": "Umowa zlecenie",
        "wynagrodzenie_bazowe": 6000.0,
        "bhp_data": "2026-01-01",
        "bhp_protokol": "Brak",
        "medycyna_data": "2026-01-01",
        "medycyna_zakres": "Do uzupełnienia",
        "urlopy_l4": [],
        "notatki_poufne": "Konto przygotowane dla 4. serwisanta."
    }
}

# Progi prowizyjne (15 000 zł -> 3%, 30 000 zł -> 4%)
COMMISSION_TIER_1 = 15000.0
COMMISSION_RATE_1 = 0.03
COMMISSION_TIER_2 = 30000.0
COMMISSION_RATE_2 = 0.04

CENNIK_BAZOWY = {
    "Przegląd/Konserwacja PC (Powietrze)": 650.0,
    "Przegląd/Konserwacja PC (Grunt)": 850.0,
    "Przegląd instalacji PV (<10 kWp)": 550.0,
    "Przegląd instalacji PV (>10 kWp)": 800.0,
    "Diagnostyka / Awaria Pompy Ciepła": 750.0,
    "Diagnostyka / Awaria Falownika PV": 600.0,
    "Czyszczenie filtrów i układu CO": 400.0,
    "Pomiary elektryczne + Protokół": 650.0,
    "Modernizacja / Rozbudowa rozdzielnicy": 950.0,
    "Serwis Gwarancyjny (Rozliczenie wewnętrzne)": 0.0
}

# Przybliżone współrzędne geograficzne polskich stref kodowych (pierwsze 2 cyfry kodu pocztowego XX-xxx)
POSTAL_PREFIX_COORDS = {
    "00": (52.23, 21.01), "01": (52.25, 20.95), "02": (52.18, 21.00), "03": (52.27, 21.08),
    "04": (52.22, 21.15), "05": (52.17, 20.80), "06": (52.88, 20.61), "07": (52.60, 21.80),
    "08": (52.16, 22.28), "09": (52.55, 19.70), "10": (53.78, 20.48), "11": (54.03, 21.00),
    "12": (53.60, 21.50), "13": (53.35, 19.90), "14": (53.85, 19.80), "15": (53.13, 23.16),
    "16": (53.65, 23.10), "17": (52.75, 23.40), "18": (53.18, 22.06), "19": (53.83, 22.35),
    "20": (51.25, 22.57), "21": (51.65, 22.80), "22": (51.05, 23.47), "23": (50.80, 22.30),
    "24": (51.40, 21.97), "25": (50.87, 20.63), "26": (51.40, 21.15), "27": (50.85, 21.45),
    "28": (50.45, 20.70), "29": (50.80, 19.95), "30": (50.06, 19.94), "31": (50.08, 20.00),
    "32": (50.00, 19.65), "33": (50.01, 20.99), "34": (49.65, 19.95), "35": (50.04, 22.00),
    "36": (50.25, 21.95), "37": (50.00, 22.68), "38": (49.68, 21.76), "39": (50.29, 21.42),
    "40": (50.26, 19.02), "41": (50.30, 18.95), "42": (50.70, 19.15), "43": (49.95, 18.95),
    "44": (50.20, 18.60), "45": (50.67, 17.92), "46": (50.75, 18.10), "47": (50.35, 18.20),
    "48": (50.47, 17.33), "49": (50.75, 17.50), "50": (51.11, 17.04), "51": (51.15, 17.10),
    "52": (51.07, 17.05), "53": (51.08, 16.98), "54": (51.14, 16.92), "55": (51.05, 17.25),
    "56": (51.40, 17.35), "57": (50.44, 16.65), "58": (50.84, 16.30), "59": (51.20, 16.16),
    "60": (52.41, 16.93), "61": (52.39, 16.95), "62": (52.35, 17.50), "63": (51.76, 17.80),
    "64": (52.75, 16.60), "65": (51.94, 15.51), "66": (52.25, 15.25), "67": (51.66, 16.08),
    "68": (51.63, 15.14), "69": (52.35, 14.65), "70": (53.43, 14.55), "71": (53.46, 14.50),
    "72": (53.75, 14.90), "73": (53.34, 15.05), "74": (53.00, 14.60), "75": (54.19, 16.18),
    "76": (54.46, 17.03), "77": (53.85, 17.25), "78": (53.70, 16.20), "80": (54.35, 18.65),
    "81": (54.52, 18.53), "82": (54.10, 19.05), "83": (54.10, 18.45), "84": (54.60, 18.20),
    "85": (53.12, 18.01), "86": (53.45, 18.55), "87": (53.01, 18.60), "88": (52.80, 18.25),
    "89": (53.58, 17.85), "90": (51.76, 19.46), "91": (51.80, 19.42), "92": (51.75, 19.55),
    "93": (51.71, 19.47), "94": (51.74, 19.38), "95": (51.88, 19.35), "96": (51.96, 20.15),
    "97": (51.40, 19.70), "98": (51.60, 18.95), "99": (52.10, 19.25)
}

# ==========================================
# 3. OBSŁUGA GOOGLE SHEETS, PROFILI KADROWYCH I BRANDINGU
# ==========================================
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

def load_config_and_profiles():
    profiles = DEFAULT_TECHS.copy()
    logo_b64 = None

    if os.path.exists(PROFILES_FILE):
        try:
            with open(PROFILES_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
                profiles = saved.get("techs", DEFAULT_TECHS)
                logo_b64 = saved.get("logo_b64", None)
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
                if raw_json:
                    profiles = json.loads(raw_json)
                if raw_logo:
                    logo_b64 = raw_logo
            except gspread.WorksheetNotFound:
                ws = sh.add_worksheet(title="Config_System", rows="10", cols="5")
                ws.update_acell("A1", json.dumps(profiles, ensure_ascii=False))
        except Exception:
            pass

    return profiles, logo_b64

def save_config_and_profiles(profiles, logo_b64=None):
    if logo_b64 is None and "logo_b64" in st.session_state:
        logo_b64 = st.session_state["logo_b64"]

    payload = {"techs": profiles, "logo_b64": logo_b64}
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

# ==========================================
# 4. WCZYTYWANIE I ZAPIS BAZY ZGŁOSZEŃ
# ==========================================
REQUIRED_COLS = [
    "ID_Zgloszenia", "Klient", "Adres", "Kod_Pocztowy", "Miasto", "Telefon",
    "Typ_Instalacji", "Opis_Zgloszenia", "Priorytet_Krytyczny", "Typ_Serwisu",
    "Serwisant", "Data_Planowana", "Status", "Naprawione_1_Wizyta",
    "Kwota_Przewidywana", "Kwota_Koncowa", "Czas_Pracy_Min",
    "Zalecenia_Klienta", "Notatka_Serwisanta", "Data_Zamkniecia", "Podpis_Zlozony"
]

def ensure_columns(df):
    if df.empty:
        return pd.DataFrame(columns=REQUIRED_COLS)

    col_map = {}
    for c in df.columns:
        cl = str(c).lower().strip()
        if "klient" in cl or "nazwisko" in cl or "imie" in cl or "imię" in cl:
            if "Klient" not in df.columns: col_map[c] = "Klient"
        elif "adres" in cl or "ulica" in cl:
            if "Adres" not in df.columns: col_map[c] = "Adres"
        elif "kod" in cl and "poczt" in cl:
            if "Kod_Pocztowy" not in df.columns: col_map[c] = "Kod_Pocztowy"
        elif "miasto" in cl or "miejscowo" in cl:
            if "Miasto" not in df.columns: col_map[c] = "Miasto"
        elif "tel" in cl:
            if "Telefon" not in df.columns: col_map[c] = "Telefon"
        elif "opis" in cl or "usterk" in cl or "uwag" in cl:
            if "Opis_Zgloszenia" not in df.columns: col_map[c] = "Opis_Zgloszenia"

    if col_map:
        df = df.rename(columns=col_map)

    for rc in REQUIRED_COLS:
        if rc not in df.columns:
            if rc == "ID_Zgloszenia":
                df[rc] = [f"ZGL-{i+1001}" for i in range(len(df))]
            elif rc == "Status":
                df[rc] = "Nowe"
            elif rc == "Priorytet_Krytyczny":
                df[rc] = "Nie"
            elif rc == "Typ_Serwisu":
                df[rc] = "Płatny"
            elif rc == "Naprawione_1_Wizyta":
                df[rc] = "Tak"
            elif rc in ["Kwota_Przewidywana", "Kwota_Koncowa", "Czas_Pracy_Min"]:
                df[rc] = 650.0 if rc == "Kwota_Przewidywana" else 0.0
            else:
                df[rc] = ""

    df["Kwota_Przewidywana"] = pd.to_numeric(df["Kwota_Przewidywana"], errors="coerce").fillna(650.0)
    df["Kwota_Koncowa"] = pd.to_numeric(df["Kwota_Koncowa"], errors="coerce").fillna(0.0)
    df["Czas_Pracy_Min"] = pd.to_numeric(df["Czas_Pracy_Min"], errors="coerce").fillna(0.0)
    df = df.fillna("")
    return df

def load_tickets():
    client = get_gsheet_client()
    if client:
        try:
            sh = client.open(GOOGLE_SHEET_NAME)
            ws = sh.get_worksheet(0)
            data = ws.get_all_records()
            if data:
                return ensure_columns(pd.DataFrame(data))
        except Exception:
            pass

    if os.path.exists(LOCAL_EXCEL_FILE):
        try:
            df = pd.read_excel(LOCAL_EXCEL_FILE)
            return ensure_columns(df)
        except Exception:
            pass

    sample = pd.DataFrame([
        {
            "ID_Zgloszenia": "ZGL-1001", "Klient": "Jan Kowalski", "Adres": "ul. Słoneczna 12",
            "Kod_Pocztowy": "05-800", "Miasto": "Pruszków", "Telefon": "500100200",
            "Typ_Instalacji": "Pompa Ciepła + PV", "Opis_Zgloszenia": "Brak grzania CO - błąd sprężarki (Awaria)",
            "Priorytet_Krytyczny": "Tak", "Typ_Serwisu": "Płatny", "Serwisant": "Leszek",
            "Data_Planowana": str(datetime.date.today()), "Status": "Zaplanowane",
            "Naprawione_1_Wizyta": "Tak", "Kwota_Przewidywana": 850.0, "Kwota_Koncowa": 0.0,
            "Czas_Pracy_Min": 0, "Zalecenia_Klienta": "", "Notatka_Serwisanta": "",
            "Data_Zamkniecia": "", "Podpis_Zlozony": "Nie"
        },
        {
            "ID_Zgloszenia": "ZGL-1002", "Klient": "Anna Nowak", "Adres": "ul. Lipowa 4",
            "Kod_Pocztowy": "05-840", "Miasto": "Brwinów", "Telefon": "600200300",
            "Typ_Instalacji": "Pompa Ciepła", "Opis_Zgloszenia": "Przegląd roczny przed sezonem + czyszczenie filtrów",
            "Priorytet_Krytyczny": "Nie", "Typ_Serwisu": "Płatny", "Serwisant": "Leszek",
            "Data_Planowana": str(datetime.date.today()), "Status": "Zaplanowane",
            "Naprawione_1_Wizyta": "Tak", "Kwota_Przewidywana": 650.0, "Kwota_Koncowa": 0.0,
            "Czas_Pracy_Min": 0, "Zalecenia_Klienta": "", "Notatka_Serwisanta": "",
            "Data_Zamkniecia": "", "Podpis_Zlozony": "Nie"
        }
    ])
    return ensure_columns(sample)

def save_tickets(df):
    st.session_state["df_tickets"] = df.copy()
    try:
        df.to_excel(LOCAL_EXCEL_FILE, index=False)
    except Exception:
        pass

    client = get_gsheet_client()
    if client:
        try:
            sh = client.open(GOOGLE_SHEET_NAME)
            ws = sh.get_worksheet(0)
            ws.clear()
            df_str = df.astype(str)
            ws.update([df_str.columns.values.tolist()] + df_str.values.tolist())
        except Exception as e:
            st.warning(f"Zapisano lokalnie. Synchronizacja z Google Sheets napotkała problem: {e}")

# ==========================================
# 5. ALGORYTM PLANOWANIA TRAS I KLASTRÓW
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
    return 52.23, 21.01

def haversine_km(lat1, lon1, lat2, lon2):
    R = 6371.0
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlambda = math.radians(lon2 - lon1)
    a = math.sin(dphi / 2.0)**2 + math.cos(phi1) * math.cos(phi2) * math.sin(dlambda / 2.0)**2
    return 2 * R * math.atan2(math.sqrt(a), math.sqrt(1 - a))

def is_tech_available(tech_name, target_date_str, profiles):
    tech_data = profiles.get(tech_name, {})
    absences = tech_data.get("urlopy_l4", [])
    for ab in absences:
        start_d = ab.get("od", "")
        end_d = ab.get("do", "")
        if start_d <= target_date_str <= end_d:
            return False
    return True

def generate_smart_routes(df, profiles, start_date, days_ahead=5):
    pending = df[df["Status"].isin(["Nowe", "Do zaplanowania", ""])].copy()
    if pending.empty:
        return [], "Brak nowych, nieprzypisanych zgłoszeń w bazie."

    coords = pending.apply(extract_coords_from_row, axis=1)
    pending["lat"] = [c[0] for c in coords]
    pending["lon"] = [c[1] for c in coords]
    pending["is_crit"] = pending.apply(
        lambda r: 1 if str(r.get("Priorytet_Krytyczny", "")).lower() == "tak" or "awaria" in str(r.get("Opis_Zgloszenia", "")).lower() else 0,
        axis=1
    )

    pending = pending.sort_values(by=["is_crit", "Kwota_Przewidywana"], ascending=[False, False])
    used_indices = set()
    valid_clusters = []

    for idx, seed in pending.iterrows():
        if idx in used_indices:
            continue
        cluster_indices = [idx]
        for cand_idx, cand in pending.iterrows():
            if cand_idx == idx or cand_idx in used_indices:
                continue
            if len(cluster_indices) >= 4:
                break
            dist = haversine_km(seed["lat"], seed["lon"], cand["lat"], cand["lon"])
            if dist <= 50.0:
                cluster_indices.append(cand_idx)

        cluster_df = pending.loc[cluster_indices]
        total_rev = cluster_df["Kwota_Przewidywana"].sum()
        crit_count = cluster_df["is_crit"].sum()

        if total_rev >= 1200.0 or crit_count >= 2:
            for c_i in cluster_indices:
                used_indices.add(c_i)
            valid_clusters.append({
                "indices": cluster_indices,
                "total_rev": total_rev,
                "crit_count": int(crit_count),
                "miasta": ", ".join(cluster_df["Miasto"].astype(str).unique()),
                "count": len(cluster_indices)
            })

    if not valid_clusters:
        return [], "Żaden z klastrów nie spełnił warunku minimum 1200 zł lub min. 2 zgłoszeń krytycznych w promieniu 50 km."

    tech_names = list(profiles.keys())
    proposals = []
    cluster_ptr = 0

    for d_offset in range(days_ahead):
        curr_date = start_date + datetime.timedelta(days=d_offset)
        curr_date_str = str(curr_date)
        for t_name in tech_names:
            if cluster_ptr >= len(valid_clusters):
                break
            if is_tech_available(t_name, curr_date_str, profiles):
                cl = valid_clusters[cluster_ptr]
                cl["serwisant"] = t_name
                cl["data"] = curr_date_str
                proposals.append(cl)
                cluster_ptr += 1

    return proposals, f"Wygenerowano {len(proposals)} opłacalnych tras (klastrów <=50 km)."

# ==========================================
# 6. OBLICZANIE WYNIKÓW I PROWIZJI SERWISANTA
# ==========================================
def calculate_tech_metrics(df, tech_name):
    tech_df = df[df["Serwisant"] == tech_name]
    closed_df = tech_df[tech_df["Status"] == "Zakończone"].copy()

    total_closed = len(closed_df)
    if total_closed > 0:
        ftfr_count = len(closed_df[closed_df["Naprawione_1_Wizyta"] == "Tak"])
        ftfr_pct = round((ftfr_count / total_closed) * 100.0, 1)
    else:
        ftfr_pct = 100.0

    paid_count = len(closed_df[closed_df["Typ_Serwisu"] == "Płatny"])
    warranty_count = len(closed_df[closed_df["Typ_Serwisu"] == "Gwarancyjny"])

    current_ym = datetime.date.today().strftime("%Y-%m")
    closed_df["ym"] = closed_df["Data_Zamkniecia"].astype(str).str.slice(0, 7)
    monthly_df = closed_df[(closed_df["ym"] == current_ym) | (closed_df["ym"] == "")]
    monthly_revenue = monthly_df["Kwota_Koncowa"].sum()

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
        "monthly_revenue": monthly_revenue,
        "missing_tier_1": missing_tier_1,
        "missing_tier_2": missing_tier_2,
        "commission_pln": commission_pln,
        "active_rate_str": active_rate_str
    }

# ==========================================
# 7. INICJALIZACJA STANU SESJI
# ==========================================
if "profiles" not in st.session_state or "logo_b64" not in st.session_state:
    p_loaded, l_loaded = load_config_and_profiles()
    st.session_state["profiles"] = p_loaded
    st.session_state["logo_b64"] = l_loaded

if "df_tickets" not in st.session_state:
    st.session_state["df_tickets"] = load_tickets()

if "logged_role" not in st.session_state:
    st.session_state["logged_role"] = None
if "logged_user" not in st.session_state:
    st.session_state["logged_user"] = None

# ==========================================
# 8. EKRAN GŁÓWNY LOGOWANIA + SZKOLENIE WSTĘPNE
# ==========================================
if st.session_state["logged_role"] is None:
    render_brand_logo(max_width=280, centered=True)
    st.markdown("<p style='text-align:center; color:#64748b; margin-top:-5px;'>Mobilny System Zarządzania Serwisem OZE (Telefon / Tablet)</p>", unsafe_allow_html=True)
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

        if st.button("🚀 Zaloguj się", type="primary"):
            if user_choice == "--- Wybierz z listy ---":
                st.error("Wybierz swój profil z listy powyżej.")
            elif user_choice == "Koordynator Serwisu":
                if pin_input == PIN_KOORDYNATOR:
                    st.session_state["logged_role"] = "koordynator"
                    st.session_state["logged_user"] = "Koordynator"
                    st.rerun()
                else:
                    st.error("Nieprawidłowy kod PIN Koordynatora.")
            elif user_choice == "Zarząd (Raporty)":
                if pin_input == PIN_ZARZAD:
                    st.session_state["logged_role"] = "zarzad"
                    st.session_state["logged_user"] = "Zarząd"
                    st.rerun()
                else:
                    st.error("Nieprawidłowy kod PIN Zarządu.")
            elif user_choice in st.session_state["profiles"]:
                expected_pin = str(st.session_state["profiles"][user_choice].get("pin", ""))
                if pin_input == expected_pin:
                    st.session_state["logged_role"] = "serwisant"
                    st.session_state["logged_user"] = user_choice
                    st.rerun()
                else:
                    st.error("Nieprawidłowy kod PIN serwisanta.")

    with col_onboard:
        st.subheader("🎓 Dla nowych pracowników")
        st.info("Rozpoczynasz pracę w naszym zespole serwisowym? Zapoznaj się z procedurami bezpieczeństwa i standardem obsługi klienta.")
        if st.button("📘 Otwórz Szkolenie Wstępne (Onboarding)"):
            st.session_state["show_onboarding"] = not st.session_state.get("show_onboarding", False)

        if st.session_state.get("show_onboarding", False):
            with st.container(border=True):
                st.markdown("### Moduł Szkolenia Wstępnego")
                st.markdown("""
                1. **Standard wizyty serwisowej** – weryfikacja układu, dokumentacja zdjęciowa przed i po, kultura pracy.
                2. **Procedury BHP przy OZE** – bezpieczne odłączanie strony DC/AC w instalacjach PV oraz praca z czynnikiem chłodniczym w pompach ciepła.
                3. **Obsługa aplikacji mobilnej** – każde zlecenie kończymy czytelnymi zaleceniami oraz podpisem klienta na ekranie tabletu/telefonu.
                
                *(Szczegółowe materiały wideo oraz testy wdrożeniowe zostaną opublikowane w tym module w kolejnym etapie).*
                """)

    st.stop()

# Pasek górny po zalogowaniu
top_c1, top_c2, top_c3 = st.columns([2, 2, 1])
with top_c1:
    render_brand_logo(max_width=200, centered=False)
with top_c2:
    st.markdown(f"**Zalogowany jako:** `{st.session_state['logged_user']}`")
with top_c3:
    if st.button("🚪 Wyloguj"):
        st.session_state["logged_role"] = None
        st.session_state["logged_user"] = None
        st.rerun()

st.divider()

# ==========================================
# 9. PANEL SERWISANTA
# ==========================================
if st.session_state["logged_role"] == "serwisant":
    tech_name = st.session_state["logged_user"]
    df = st.session_state["df_tickets"]
    metrics = calculate_tech_metrics(df, tech_name)

    m_col1, m_col2, m_col3 = st.columns([1.4, 1.4, 1])
    with m_col1:
        if metrics["missing_tier_1"] > 0:
            headline = f"Brakuje: {metrics['missing_tier_1']:,.0f} zł do prowizji 3%"
            subline = f"Próg I (15 000 zł) | Do progu 4% (30 000 zł) brakuje: {metrics['missing_tier_2']:,.0f} zł"
        elif metrics["missing_tier_2"] > 0:
            headline = f"🔥 Masz już 3%! Brakuje: {metrics['missing_tier_2']:,.0f} zł do 4%"
            subline = f"Wypracowany obrót: {metrics['monthly_revenue']:,.0f} zł (Premia: {metrics['commission_pln']:,.0f} zł)"
        else:
            headline = f"🏆 Maksymalny próg 4% zdobyty!"
            subline = f"Wypracowany obrót: {metrics['monthly_revenue']:,.0f} zł (Premia: {metrics['commission_pln']:,.0f} zł)"

        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-title">🎯 Twój Cel Prowizyjny (Bieżący Miesiąc)</div>
            <div class="kpi-value">{headline}</div>
            <div class="kpi-sub">{subline}</div>
        </div>
        """, unsafe_allow_html=True)

    with m_col2:
        st.markdown(f"""
        <div class="kpi-card kpi-card-warn">
            <div class="kpi-title">⚡ Skuteczność 1. Wizyty (First-Time Fix)</div>
            <div class="kpi-value">{metrics['ftfr_pct']}% naprawionych od ręki</div>
            <div class="kpi-sub">Obsłużone: {metrics['total_closed']} (Płatne: {metrics['paid_count']} | Gwarancyjne: {metrics['warranty_count']})</div>
        </div>
        """, unsafe_allow_html=True)

    with m_col3:
        st.write("")
        if st.button("📚 BAZA WIEDZY (Schematy / Kody błędów)"):
            st.session_state["show_kb"] = not st.session_state.get("show_kb", False)

    progress_val = min(1.0, float(metrics["monthly_revenue"] / COMMISSION_TIER_2))
    st.progress(progress_val, text=f"Postęp obrotu w tym miesiącu: {metrics['monthly_revenue']:,.0f} zł / 30 000 zł ({metrics['active_rate_str']})")

    if st.session_state.get("show_kb", False):
        with st.expander("📚 Moduł Bazy Wiedzy Technicznej OZE", expanded=True):
            st.info("W tym miejscu pojawią się kody błędów falowników i pomp ciepła, instrukcje serwisowe oraz listy kontrolne przeglądów (moduł w przygotowaniu).")

    st.divider()

    my_tickets = df[(df["Serwisant"] == tech_name) & (df["Status"] != "Zakończone")].copy()

    st.subheader(f"📋 Twoje aktywne zlecenia ({len(my_tickets)})")
    if my_tickets.empty:
        st.success("Brak otwartych zleceń przypisanych do Twojego konta.")
    else:
        ticket_options = [
            f"{row['ID_Zgloszenia']} | {row['Miasto']} ({row['Adres']}) - {row['Klient']} [{'🔥 KRYTYCZNE' if row['Priorytet_Krytyczny']=='Tak' else row['Typ_Serwisu']}]"
            for _, row in my_tickets.iterrows()
        ]
        selected_option = st.selectbox("Wybierz zlecenie do obsługi:", ticket_options)
        selected_id = selected_option.split(" | ")[0]
        t_idx = df[df["ID_Zgloszenia"] == selected_id].index[0]
        t_row = df.loc[t_idx]

        c_info1, c_info2 = st.columns([1.2, 1])
        with c_info1:
            st.markdown(f"### 👤 Klient: **{t_row['Klient']}**")
            st.markdown(f"📍 **Adres:** `{t_row['Kod_Pocztowy']} {t_row['Miasto']}, {t_row['Adres']}`")
            st.markdown(f"📞 **Telefon:** [{t_row['Telefon']}](tel:{t_row['Telefon']})")
            st.markdown(f"🔧 **Instalacja:** {t_row['Typ_Instalacji']} | **Typ:** `{t_row['Typ_Serwisu']}`")
            st.markdown(f"📝 **Opis zgłoszenia:** {t_row['Opis_Zgloszenia']}")

        with c_info2:
            maps_query = f"{t_row['Adres']} {t_row['Kod_Pocztowy']} {t_row['Miasto']}".replace(" ", "+")
            st.link_button("🗺️ Nawiguj w Google Maps", f"https://www.google.com/maps/search/?api=1&query={maps_query}")
            sms_body = f"Dzien dobry, tu serwis OZE. Przypominamy o zaplanowanej wizycie serwisowej pod adresem {t_row['Adres']}."
            st.link_button("💬 Wyślij SMS do klienta", f"sms:{t_row['Telefon']}?body={sms_body.replace(' ', '%20')}")

        st.divider()

        st.subheader("🛠️ Realizacja wizyty i rozliczenie")
        calc_c1, calc_c2 = st.columns(2)
        with calc_c1:
            selected_services = st.multiselect(
                "Wykonane czynności z cennika:",
                options=list(CENNIK_BAZOWY.keys()),
                default=[]
            )
            extra_parts = st.number_input("Koszt użytych części / materiałów (zł):", min_value=0.0, step=50.0, value=0.0)
            km_dojazd = st.number_input("Dojazd (km w obie strony powyżej ryczałtu):", min_value=0, step=10, value=0)

        with calc_c2:
            work_minutes = st.number_input("Czas pracy na miejscu (minuty):", min_value=15, step=15, value=int(t_row["Czas_Pracy_Min"]) if t_row["Czas_Pracy_Min"] > 0 else 60)
            typ_serwisu_real = st.selectbox("Kwalifikacja zlecenia:", ["Płatny", "Gwarancyjny"], index=0 if t_row["Typ_Serwisu"] != "Gwarancyjny" else 1)
            fixed_first_visit = st.radio(
                "Czy usterka / zgłoszenie zostało w pełni rozwiązane podczas tej wizyty?",
                ["Tak", "Nie (Wymagana kolejna wizyta / zamówienie części)"],
                horizontal=True
            )

        base_sum = sum(CENNIK_BAZOWY[s] for s in selected_services)
        if base_sum == 0 and typ_serwisu_real == "Płatny":
            base_sum = float(t_row["Kwota_Przewidywana"])
        auto_total = 0.0 if typ_serwisu_real == "Gwarancyjny" else (base_sum + extra_parts + (km_dojazd * 2.5))

        final_price = st.number_input("💰 Kwota końcowa do rozliczenia (zł brutto):", min_value=0.0, step=50.0, value=float(auto_total))
        internal_note = st.text_area("Notatka wewnętrzna dla Koordynatora (niewidoczna dla klienta):", value=str(t_row["Notatka_Serwisanta"]))

        st.markdown('<div class="client-box">', unsafe_allow_html=True)
        render_brand_logo(max_width=220, centered=False)
        st.markdown("### 📄 Protokół Odbioru Serwisu i Zalecenia dla Klienta")
        st.markdown(f"**Klient:** {t_row['Klient']} | **Adres:** {t_row['Kod_Pocztowy']} {t_row['Miasto']}, {t_row['Adres']} | **Data:** {datetime.date.today()}")
        client_recommendations = st.text_area(
            "Zalecenia po-serwisowe dla Klienta (widoczne na protokole):",
            value=str(t_row["Zalecenia_Klienta"]),
            placeholder="Wpisz zalecenia dotyczące eksploatacji instalacji, ciśnienia w układzie lub terminu kolejnego przeglądu..."
        )
        st.markdown(f"#### Do zapłaty / Rozliczenie: **{final_price:,.2f} zł** ({typ_serwisu_real})")
        st.markdown("**Podpis Klienta (potwierdzenie wykonania prac i zapoznania się z zaleceniami):**")

        canvas_result = st_canvas(
            fill_color="rgba(255, 255, 255, 0)",
            stroke_width=3,
            stroke_color="#0f172a",
            background_color="#ffffff",
            height=180,
            width=620,
            drawing_mode="freedraw",
            key=f"canvas_{selected_id}"
        )
        st.markdown('</div>', unsafe_allow_html=True)

        if st.button("✅ ZAKOŃCZ ZLECENIE I ZAPISZ PROTOKÓŁ", type="primary"):
            df.loc[t_idx, "Status"] = "Zakończone"
            df.loc[t_idx, "Typ_Serwisu"] = typ_serwisu_real
            df.loc[t_idx, "Naprawione_1_Wizyta"] = "Tak" if fixed_first_visit.startswith("Tak") else "Nie"
            df.loc[t_idx, "Kwota_Koncowa"] = final_price
            df.loc[t_idx, "Czas_Pracy_Min"] = work_minutes
            df.loc[t_idx, "Zalecenia_Klienta"] = client_recommendations
            df.loc[t_idx, "Notatka_Serwisanta"] = internal_note
            df.loc[t_idx, "Data_Zamkniecia"] = str(datetime.date.today())
            df.loc[t_idx, "Podpis_Zlozony"] = "Tak" if (canvas_result.image_data is not None) else "Nie"

            save_tickets(df)
            st.success("Zlecenie zostało zamknięte, a Twoje statystyki prowizyjne zaktualizowane!")
            st.rerun()

# ==========================================
# 10. PANEL KOORDYNATORA
# ==========================================
elif st.session_state["logged_role"] == "koordynator":
    df = st.session_state["df_tickets"]
    profiles = st.session_state["profiles"]

    tab_plan, tab_kadry, tab_zlecenia, tab_brand = st.tabs([
        "🗺️ Harmonogram i Auto-Planer Tras",
        "👷 Profilowanie Serwisantów (Kadry i Wyniki)",
        "📋 Baza Zgłoszeń i Dodawanie",
        "🎨 Branding (Logo) i Plik Excel"
    ])

    with tab_plan:
        st.subheader("🗺️ Inteligentne Planowanie Wyjazdów (Cała Polska)")
        st.markdown("""
        **Zasady algorytmu:**
        * Brak sztywnej rejonizacji – serwisanci wyruszają z domu lub z hotelu z poprzedniego dnia.
        * Grupowanie do **maks. 4 punktów** oddalonych od siebie o **maks. 50 km**.
        * Trasa jest zatwierdzana, jeśli jej przewidywany przychód wynosi **min. 1200 zł** LUB zawiera **min. 2 zgłoszenia krytyczne** (np. awaria pompy ciepła).
        * Automatyczne wykluczanie serwisantów mających wpisany urlop lub L4 w danym dniu.
        """)

        p_col1, p_col2 = st.columns(2)
        with p_col1:
            plan_start_date = st.date_input("Planuj wyjazdy począwszy od dnia:", value=datetime.date.today() + datetime.timedelta(days=1))
        with p_col2:
            plan_days = st.slider("Liczba dni do przodu:", min_value=1, max_value=10, value=5)

        if st.button("🚀 Wygeneruj propozycje tras na kolejne dni", type="primary"):
            proposals, msg = generate_smart_routes(df, profiles, plan_start_date, plan_days)
            st.session_state["route_proposals"] = proposals
            st.session_state["route_msg"] = msg

        if "route_msg" in st.session_state:
            st.info(st.session_state["route_msg"])

        proposals = st.session_state.get("route_proposals", [])
        if proposals:
            for i, prop in enumerate(proposals):
                with st.container(border=True):
                    st.markdown(
                        f"#### 📅 Dzień: `{prop['data']}` | 👷 Serwisant: **{prop['serwisant']}** | "
                        f"📍 Punkty: **{prop['count']}/4** ({prop['miasta']}) | "
                        f"💰 Przewidywany przychód: **{prop['total_rev']:,.0f} zł** | 🔥 Krytyczne: **{prop['crit_count']}**"
                    )
                    sub_df = df.loc[prop["indices"], ["ID_Zgloszenia", "Klient", "Kod_Pocztowy", "Miasto", "Adres", "Opis_Zgloszenia", "Priorytet_Krytyczny", "Kwota_Przewidywana"]]
                    st.dataframe(sub_df, use_container_width=True, hide_index=True)

            if st.button("✅ ZATWIERDŹ I PRZYPISZ WSZYSTKIE POWYŻSZE TRASY DO SERWISANTÓW", type="primary"):
                for prop in proposals:
                    for idx in prop["indices"]:
                        df.loc[idx, "Serwisant"] = prop["serwisant"]
                        df.loc[idx, "Data_Planowana"] = prop["data"]
                        df.loc[idx, "Status"] = "Zaplanowane"
                save_tickets(df)
                st.session_state["route_proposals"] = []
                st.success("Trasy zostały przypisane do harmonogramu serwisantów!")
                st.rerun()

    with tab_kadry:
        st.subheader("👷 Kartoteka Kadrowa, Uprawnienia i Wyniki Serwisantów")
        selected_tech = st.selectbox("Wybierz serwisanta do edycji profilu:", list(profiles.keys()))
        t_prof = profiles[selected_tech]
        t_metrics = calculate_tech_metrics(df, selected_tech)

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Obsłużone zgłoszenia", t_metrics["total_closed"], f"Płatne: {t_metrics['paid_count']} | Gwar.: {t_metrics['warranty_count']}")
        k2.metric("Naprawione przy 1. wizycie", f"{t_metrics['ftfr_pct']}%")
        k3.metric("Obrót w tym miesiącu", f"{t_metrics['monthly_revenue']:,.0f} zł", t_metrics["active_rate_str"])
        k4.metric("Brakuje do progów (3% / 4%)", f"I: {t_metrics['missing_tier_1']:,.0f} zł", f"II: {t_metrics['missing_tier_2']:,.0f} zł")

        st.divider()

        with st.form(f"form_profile_{selected_tech}"):
            st.markdown("#### 1. Dane osobowe, hasło (PIN) i warunki zatrudnienia")
            fc1, fc2, fc3 = st.columns(3)
            with fc1:
                new_full_name = st.text_input("Pełne imię, nazwisko i stanowisko:", value=t_prof.get("pelne_dane", ""))
                new_phone = st.text_input("Telefon służbowy:", value=t_prof.get("telefon", ""))
            with fc2:
                new_pin = st.text_input("Stałe 4-cyfrowe hasło (PIN):", value=str(t_prof.get("pin", "")))
                umowa_opts = ["Umowa o pracę", "B2B", "Umowa zlecenie", "Inna"]
                curr_umowa = t_prof.get("umowa", "Umowa o pracę")
                new_umowa = st.selectbox("Rodzaj umowy:", umowa_opts, index=umowa_opts.index(curr_umowa) if curr_umowa in umowa_opts else 0)
            with fc3:
                new_base_sal = st.number_input("Wynagrodzenie zasadnicze / podstawa (zł):", min_value=0.0, step=100.0, value=float(t_prof.get("wynagrodzenie_bazowe", 6000.0)))

            st.markdown("#### 2. Szkolenia BHP oraz Badania Medycyny Pracy")
            bh1, bh2 = st.columns(2)
            with bh1:
                new_bhp_date = st.text_input("Data szkolenia BHP (RRRR-MM-DD):", value=t_prof.get("bhp_data", ""))
                new_bhp_prot = st.text_input("Numer i status protokołu BHP:", value=t_prof.get("bhp_protokol", ""))
            with bh2:
                new_med_date = st.text_input("Data badań medycyny pracy (RRRR-MM-DD):", value=t_prof.get("medycyna_data", ""))
                new_med_scope = st.text_area("Zakres badań medycyny pracy (np. wysokość >3m, SEP, kat. B):", value=t_prof.get("medycyna_zakres", ""))

            st.markdown("#### 3. Poufne notatki Koordynatora / Zarządu")
            new_secret_notes = st.text_area("Poufne notatki (widoczne wyłącznie dla Koordynatora i Zarządu):", value=t_prof.get("notatki_poufne", ""))

            if st.form_submit_button("💾 Zapisz kartotekę serwisanta", type="primary"):
                profiles[selected_tech]["pelne_dane"] = new_full_name
                profiles[selected_tech]["telefon"] = new_phone
                profiles[selected_tech]["pin"] = new_pin
                profiles[selected_tech]["umowa"] = new_umowa
                profiles[selected_tech]["wynagrodzenie_bazowe"] = new_base_sal
                profiles[selected_tech]["bhp_data"] = new_bhp_date
                profiles[selected_tech]["bhp_protokol"] = new_bhp_prot
                profiles[selected_tech]["medycyna_data"] = new_med_date
                profiles[selected_tech]["medycyna_zakres"] = new_med_scope
                profiles[selected_tech]["notatki_poufne"] = new_secret_notes
                save_config_and_profiles(profiles)
                st.success(f"Zapisano profil serwisanta {selected_tech}!")

        st.markdown("#### 🏖️ Urlopy i Zwolnienia Lekarskie (L4)")
        u_col1, u_col2, u_col3, u_col4 = st.columns([1, 1, 1, 1])
        with u_col1:
            ab_type = st.selectbox("Typ nieobecności:", ["Urlop wypoczynkowy", "Zwolnienie lekarskie (L4)", "Szkolenie / Inne"])
        with u_col2:
            ab_from = st.date_input("Od dnia:", value=datetime.date.today(), key="ab_from")
        with u_col3:
            ab_to = st.date_input("Do dnia:", value=datetime.date.today() + datetime.timedelta(days=3), key="ab_to")
        with u_col4:
            st.write("")
            if st.button("➕ Dodaj nieobecność"):
                profiles[selected_tech].setdefault("urlopy_l4", []).append({
                    "typ": ab_type,
                    "od": str(ab_from),
                    "do": str(ab_to)
                })
                save_config_and_profiles(profiles)
                st.success("Dodano wpis o nieobecności (Auto-Planer automatycznie pominie te dni).")
                st.rerun()

        if t_prof.get("urlopy_l4"):
            st.table(pd.DataFrame(t_prof["urlopy_l4"]))
            if st.button("🗑️ Wyczyść historię urlopów/L4 tego serwisanta"):
                profiles[selected_tech]["urlopy_l4"] = []
                save_config_and_profiles(profiles)
                st.rerun()

    with tab_zlecenia:
        st.subheader("➕ Dodaj nowe zgłoszenie serwisowe")
        with st.expander("Kliknij, aby rozwinąć formularz nowego zgłoszenia", expanded=False):
            with st.form("new_ticket_form"):
                nc1, nc2, nc3 = st.columns(3)
                with nc1:
                    n_klient = st.text_input("Imię i nazwisko / Firma klienta:")
                    n_tel = st.text_input("Telefon:")
                    n_typ_inst = st.selectbox("Typ instalacji:", ["Pompa Ciepła", "Fotowoltaika (PV)", "Pompa Ciepła + PV", "Magazyn Energii", "Inne"])
                with nc2:
                    n_kod = st.text_input("Kod pocztowy (XX-XXX):", placeholder="np. 05-800")
                    n_miasto = st.text_input("Miejscowość:")
                    n_adres = st.text_input("Ulica i numer:")
                with nc3:
                    n_kryt = st.selectbox("Czy zgłoszenie krytyczne (np. awaria PC zimą)?", ["Nie", "Tak"])
                    n_typ_serw = st.selectbox("Rodzaj serwisu:", ["Płatny", "Gwarancyjny"])
                    n_kwota = st.number_input("Przewidywany przychód (zł):", min_value=0.0, step=50.0, value=650.0)
                    n_tech = st.selectbox("Przypisz od razu serwisanta (lub zostaw dla Auto-Planera):", ["Auto-Planer"] + list(profiles.keys()))

                n_opis = st.text_area("Opis usterki / zakres prac:")

                if st.form_submit_button("➕ Zapisz zgłoszenie w bazie", type="primary"):
                    new_id = f"ZGL-{len(df) + 1001}"
                    new_row = {
                        "ID_Zgloszenia": new_id,
                        "Klient": n_klient,
                        "Adres": n_adres,
                        "Kod_Pocztowy": n_kod,
                        "Miasto": n_miasto,
                        "Telefon": n_tel,
                        "Typ_Instalacji": n_typ_inst,
                        "Opis_Zgloszenia": n_opis,
                        "Priorytet_Krytyczny": n_kryt,
                        "Typ_Serwisu": n_typ_serw,
                        "Serwisant": "" if n_tech == "Auto-Planer" else n_tech,
                        "Data_Planowana": "" if n_tech == "Auto-Planer" else str(datetime.date.today()),
                        "Status": "Nowe" if n_tech == "Auto-Planer" else "Zaplanowane",
                        "Naprawione_1_Wizyta": "Tak",
                        "Kwota_Przewidywana": n_kwota,
                        "Kwota_Koncowa": 0.0,
                        "Czas_Pracy_Min": 0,
                        "Zalecenia_Klienta": "",
                        "Notatka_Serwisanta": "",
                        "Data_Zamkniecia": "",
                        "Podpis_Zlozony": "Nie"
                    }
                    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                    save_tickets(df)
                    st.success(f"Dodano zgłoszenie {new_id}!")
                    st.rerun()

                st.subheader("📋 Wszystkie zgłoszenia (Edycja bezpośrednia)")
        changes = st.data_editor(df, use_container_width=True, num_rows="dynamic", key="tickets_editor")
        if st.button("💾 Zapisz zmiany w tabeli zgłoszeń"):
            save_tickets(changes)
            st.success("Baza zgłoszeń została zaktualizowana!")
            st.rerun()


    with tab_brand:
        st.subheader("🎨 Branding Aplikacji – Wgraj Logo Firmy")
        st.info("Rekomendowane wymiary logo: **400 × 120 px** (proporcje ok. 3:1), format **PNG (z przezroczystym tłem)** lub JPG. Logo wyświetla się na ekranie głównym oraz na protokołach podpisywanych przez klienta.")

        uploaded_logo = st.file_uploader("Wybierz plik graficzny z logo (PNG / JPG):", type=["png", "jpg", "jpeg"])
        if uploaded_logo is not None:
            img = Image.open(uploaded_logo)
            img.thumbnail((400, 120))
            buf = BytesIO()
            img.save(buf, format="PNG")
            b64_str = base64.b64encode(buf.getvalue()).decode("utf-8")
            st.session_state["logo_b64"] = b64_str
            with open(LOGO_FILE, "wb") as f:
                f.write(buf.getvalue())
            save_config_and_profiles(profiles, b64_str)
            st.success("Logo zostało zapisane w chmurze i zaktualizowane w całej aplikacji!")
            st.rerun()

        st.divider()
        st.subheader("📥 Import / Podmiana bazy z pliku Excel (.xlsx)")
        uploaded_excel = st.file_uploader("Wgraj plik Excel z bazą zgłoszeń:", type=["xlsx"])
        if uploaded_excel is not None:
            if st.button("⚠️ Wczytaj dane z tego pliku Excel i zapisz w Google Sheets"):
                new_df = ensure_columns(pd.read_excel(uploaded_excel))
                save_tickets(new_df)
                st.success("Baza zgłoszeń została zaimportowana!")
                st.rerun()

# ==========================================
# 11. PANEL ZARZĄDU
# ==========================================
elif st.session_state["logged_role"] == "zarzad":
    df = st.session_state["df_tickets"]
    profiles = st.session_state["profiles"]

    st.subheader("📊 Panel Zarządu – Raporty Operacyjne i Finansowe")

    rep_tab_day, rep_tab_month = st.tabs(["📅 Raport Dzienny z Pracy Serwisantów", "📈 Raport Miesięczny i Prowizje"])

    with rep_tab_day:
        sel_date = st.date_input("Wybierz dzień raportu:", value=datetime.date.today())
        sel_date_str = str(sel_date)

        day_df = df[(df["Data_Zamkniecia"] == sel_date_str) | (df["Data_Planowana"] == sel_date_str)].copy()
        closed_day = day_df[day_df["Status"] == "Zakończone"]

        dc1, dc2, dc3, dc4 = st.columns(4)
        dc1.metric("Przychód w wybranym dniu", f"{closed_day['Kwota_Koncowa'].sum():,.2f} zł")
        dc2.metric("Zlecenia zamknięte dzisiaj", f"{len(closed_day)} / {len(day_df)}")
        ftfr_day = round((len(closed_day[closed_day["Naprawione_1_Wizyta"] == "Tak"]) / len(closed_day) * 100.0), 1) if len(closed_day) > 0 else 100.0
        dc3.metric("Skuteczność 1. wizyty (Dzień)", f"{ftfr_day}%")
        dc4.metric("Łączny czas pracy w terenie", f"{int(closed_day['Czas_Pracy_Min'].sum())} min")

        st.markdown("#### Zestawienie dzienne według serwisantów")
        day_rows = []
        for t_name in profiles.keys():
            t_day = closed_day[closed_day["Serwisant"] == t_name]
            day_rows.append({
                "Serwisant": t_name,
                "Zamknięte zlecenia": len(t_day),
                "Serwisy Płatne": len(t_day[t_day["Typ_Serwisu"] == "Płatny"]),
                "Serwisy Gwarancyjne": len(t_day[t_day["Typ_Serwisu"] == "Gwarancyjny"]),
                "Naprawione od ręki (1. wizyta)": len(t_day[t_day["Naprawione_1_Wizyta"] == "Tak"]),
                "Czas pracy (min)": int(t_day["Czas_Pracy_Min"].sum()),
                "Przychód (zł)": t_day["Kwota_Koncowa"].sum()
            })
        st.dataframe(pd.DataFrame(day_rows), use_container_width=True, hide_index=True)

        st.markdown("#### Szczegółowy wykaz zleceń z wybranego dnia")
        st.dataframe(day_df, use_container_width=True, hide_index=True)

    with rep_tab_month:
        months_avail = sorted(list(set(
            [datetime.date.today().strftime("%Y-%m")] +
            [str(x)[:7] for x in df["Data_Zamkniecia"].unique() if len(str(x)) >= 7]
        )), reverse=True)
        sel_month = st.selectbox("Wybierz miesiąc rozliczeniowy (RRRR-MM):", months_avail)

        m_df = df[(df["Status"] == "Zakończone") & (df["Data_Zamkniecia"].astype(str).str.startswith(sel_month))].copy()

        mc1, mc2, mc3, mc4 = st.columns(4)
        mc1.metric("Przychód miesięczny ogółem", f"{m_df['Kwota_Koncowa'].sum():,.2f} zł")
        mc2.metric("Obsłużone zgłoszenia", len(m_df))
        mc3.metric("Płatne vs Gwarancyjne", f"{len(m_df[m_df['Typ_Serwisu']=='Płatny'])} / {len(m_df[m_df['Typ_Serwisu']=='Gwarancyjny'])}")
        ftfr_m = round((len(m_df[m_df["Naprawione_1_Wizyta"] == "Tak"]) / len(m_df) * 100.0), 1) if len(m_df) > 0 else 100.0
        mc4.metric("Skuteczność 1. wizyty (Miesiąc)", f"{ftfr_m}%")

        st.markdown("#### Podsumowanie wyników, wynagrodzeń i prowizji serwisantów")
        month_summary = []
        for t_name, t_prof in profiles.items():
            tm = calculate_tech_metrics(df, t_name)
            base_sal = float(t_prof.get("wynagrodzenie_bazowe", 0.0))
            month_summary.append({
                "Serwisant": t_name,
                "Umowa": t_prof.get("umowa", ""),
                "Zlecenia ogółem": tm["total_closed"],
                "Płatne": tm["paid_count"],
                "Gwarancyjne": tm["warranty_count"],
                "Skuteczność 1. wizyty (%)": f"{tm['ftfr_pct']}%",
                "Obrót miesięczny (zł)": round(tm["monthly_revenue"], 2),
                "Brakuje do 3% (15 tys.)": round(tm["missing_tier_1"], 2),
                "Brakuje do 4% (30 tys.)": round(tm["missing_tier_2"], 2),
                "Podstawa (zł)": base_sal,
                "Należna prowizja (zł)": round(tm["commission_pln"], 2),
                "Łącznie do wypłaty (zł)": round(base_sal + tm["commission_pln"], 2)
            })
        st.dataframe(pd.DataFrame(month_summary), use_container_width=True, hide_index=True)
