import streamlit as st
import pandas as pd
import datetime
import os
import shutil
import urllib.parse

# Próba załadowania bibliotek Google Drive / Sheets
try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSPREAD_AVAILABLE = True
except ImportError:
    GSPREAD_AVAILABLE = False

# Próba załadowania biblioteki do podpisu palcem na ekranie
try:
    from streamlit_drawable_canvas import st_canvas
    from PIL import Image
    CANVAS_AVAILABLE = True
except ImportError:
    CANVAS_AVAILABLE = False

st.set_page_config(
    page_title="PROMAG SERWIS - System Obsługi",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Wsparcie dla PWA + Blokada pływania ekranu na boki + Ergonomia jednej dłoni
st.markdown("""
<link rel="manifest" href="manifest.json">
<meta name="theme-color" content="#2b6cb0">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="PROMAG">
<style>
/* Całkowita blokada poziomego przesuwania ekranu na telefonach */
html, body, [data-testid="stAppViewContainer"], .main {
    overflow-x: hidden !important;
    max-width: 100vw !important;
}
.block-container {
    padding-top: 1.2rem !important;
    padding-left: 0.8rem !important;
    padding-right: 0.8rem !important;
    padding-bottom: 3rem !important;
    max-width: 100% !important;
}
/* Ukrycie zbędnych elementów interfejsu Streamlita na mobile */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

/* Duże, wygodne przyciski pod kciuk (min. 48px wysokości) */
.stButton > button {
    min-height: 48px !important;
    border-radius: 10px !important;
    font-weight: 700 !important;
    font-size: 1rem !important;
}

/* Karty zleceń */
.mobile-card {
    background-color: #f8f9fa;
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    padding: 14px;
    margin-bottom: 12px;
    box-shadow: 0 2px 4px rgba(0,0,0,0.04);
    word-wrap: break-word;
}
.mobile-card-route {
    background-color: #f0fff4;
    border: 2px solid #48bb78;
    border-radius: 12px;
    padding: 14px;
    margin-bottom: 12px;
    box-shadow: 0 3px 6px rgba(0,0,0,0.08);
    word-wrap: break-word;
}
.badge {
    display: inline-block;
    padding: 4px 10px;
    border-radius: 12px;
    font-size: 0.75rem;
    font-weight: 700;
    text-transform: uppercase;
    margin-right: 4px;
    margin-bottom: 4px;
}
.badge-route { background-color: #2f855a; color: #ffffff; }
.badge-solar { background-color: #feebc8; color: #c05621; }
.badge-foto { background-color: #fefcbf; color: #975a16; }
.badge-pompa { background-color: #bee3f8; color: #2b6cb0; }
.badge-magazyn { background-color: #e9d8fd; color: #6b46c1; }
.wizard-step-box {
    background-color: #ebf8ff;
    border-left: 5px solid #2b6cb0;
    padding: 12px;
    border-radius: 8px;
    margin-bottom: 14px;
    font-weight: 600;
    color: #2c5282;
}
</style>
""", unsafe_allow_html=True)

# Pliki i katalogi
ATTACHMENTS_DIR = "zalaczniki"
NOWE_FILE = "nowe_zgloszenia.csv"
EXCEL_FILE = "Serwisowa posortowana.xlsx"
CENNIK_FILE = "cennik.csv"
CENNIK_HIST_FILE = "historia_cen.csv"
GSHEET_NAME = "PROMAG_BAZA_SERWISOWA"
os.makedirs(ATTACHMENTS_DIR, exist_ok=True)

# SZTYWNA DATA STARTU NOWEGO SYSTEMU
CUTOFF_DATE = "2026-10-08"

# Domyślny cennik startowy (jeśli plik cennik.csv jeszcze nie istnieje)
DEFAULT_PRICELIST = [
    {"Kod": "USL-01", "Kategoria": "Ogólne", "Nazwa": "Diagnoza / Dojazd serwisowy", "Cena": 300.0},
    {"Kod": "SOL-01", "Kategoria": "Solar", "Nazwa": "Przegląd instalacji solarnej + uzupełnienie glikolu", "Cena": 450.0},
    {"Kod": "SOL-02", "Kategoria": "Solar", "Nazwa": "Wymiana pompy obiegowej / grupy pompowej", "Cena": 650.0},
    {"Kod": "SOL-03", "Kategoria": "Solar", "Nazwa": "Wymiana naczynia przeponowego 18L/24L", "Cena": 480.0},
    {"Kod": "PC-01", "Kategoria": "Pompa ciepła", "Nazwa": "Przegląd okresowy pompy ciepła + czyszczenie filtrów", "Cena": 550.0},
    {"Kod": "PC-02", "Kategoria": "Pompa ciepła", "Nazwa": "Wymiana czujnika przepływu / zaworu 3D", "Cena": 600.0},
    {"Kod": "PV-01", "Kategoria": "Fotowoltaika", "Nazwa": "Pomiary i przegląd instalacji PV + aktualizacja falownika", "Cena": 450.0},
    {"Kod": "PV-02", "Kategoria": "Fotowoltaika", "Nazwa": "Konfiguracja połączenia Wi-Fi / Aplikacji monitoringowej", "Cena": 250.0},
    {"Kod": "MAG-01", "Kategoria": "Magazyn energii", "Nazwa": "Diagnostyka i kalibracja BMS magazynu energii", "Cena": 500.0},
]

REGION_DEFINITIONS = [
    {"kod": "0", "nazwa": "Rejon 0 (Warszawa / Mazowieckie)", "krotka": "Rejon 0 - Mazowieckie"},
    {"kod": "2", "nazwa": "Rejon 2 (Lubelskie / Świętokrzyskie)", "krotka": "Rejon 2 - Lubelskie/Święt."},
    {"kod": "6", "nazwa": "Rejon 6 (Wielkopolskie / Lubuskie)", "krotka": "Rejon 6 - Wielkopolskie"},
    {"kod": "8", "nazwa": "Rejon 8 (Pomorskie / Kujawsko-Pomorskie)", "krotka": "Rejon 8 - Pomorskie/Kujawy"},
    {"kod": "9", "nazwa": "Rejon 9 (Łódzkie)", "krotka": "Rejon 9 - Łódzkie"},
    {"kod": "3", "nazwa": "Rejon 3 (Małopolskie / Podkarpackie)", "krotka": "Rejon 3 - Małopolskie/Podk."},
    {"kod": "1", "nazwa": "Rejon 1 (Podlaskie / Warmińsko-Mazurskie)", "krotka": "Rejon 1 - Podlaskie/Warm."},
    {"kod": "4", "nazwa": "Rejon 4 (Śląskie / Opolskie)", "krotka": "Rejon 4 - Śląskie/Opole"},
    {"kod": "5", "nazwa": "Rejon 5 (Dolnośląskie)", "krotka": "Rejon 5 - Dolnośląskie"},
    {"kod": "7", "nazwa": "Rejon 7 (Zachodniopomorskie)", "krotka": "Rejon 7 - Zachodniopomorskie"},
]

PIN_CONFIG = {
    "1110": {"rejon_kod": "0", "nazwa_rejonu": "Rejon 0 (Warszawa / Mazowieckie)", "domyslny_serwisant": "Leszek"},
    "1111": {"rejon_kod": "2", "nazwa_rejonu": "Rejon 2 (Lubelskie / Świętokrzyskie)", "domyslny_serwisant": "Stanisław"},
    "1112": {"rejon_kod": "6", "nazwa_rejonu": "Rejon 6 (Wielkopolskie / Lubuskie)", "domyslny_serwisant": "Wojtek"},
    "1113": {"rejon_kod": "8", "nazwa_rejonu": "Rejon 8 (Pomorskie / Kujawsko-Pomorskie)", "domyslny_serwisant": "Mikołaj"},
    "1114": {"rejon_kod": "9", "nazwa_rejonu": "Rejon 9 (Łódzkie)", "domyslny_serwisant": "Artur"},
    "1115": {"rejon_kod": "3", "nazwa_rejonu": "Rejon 3 (Małopolskie / Podkarpackie)", "domyslny_serwisant": "Piotr"},
    "1116": {"rejon_kod": "1", "nazwa_rejonu": "Rejon 1 (Podlaskie / Warmińsko-Mazurskie)", "domyslny_serwisant": "Alex"},
    "1117": {"rejon_kod": "4", "nazwa_rejonu": "Rejon 4 (Śląskie / Opolskie)", "domyslny_serwisant": "Dominik"},
    "1118": {"rejon_kod": "5", "nazwa_rejonu": "Rejon 5 (Dolnośląskie)", "domyslny_serwisant": "Hubert"},
    "1119": {"rejon_kod": "7", "nazwa_rejonu": "Rejon 7 (Zachodniopomorskie)", "domyslny_serwisant": "Piasta"},
    "9999": {"rejon_kod": "ALL", "nazwa_rejonu": "Dostęp ogólny / Wszystkie", "domyslny_serwisant": "Koordynator"}
}

QUICK_ACTIONS = {
    "solar": [
        "Odpowietrzono układ solarny", "Uzupełniono i sprawdzono stężenie glikolu",
        "Wymieniono pompę obiegową", "Wymieniono naczynie przeponowe",
        "Skonfigurowano sterownik solarny", "Uszczelniono połączenia śrubunkowe"
    ],
    "pompa": [
        "Wyczyszczono filtr siatkowy / magnetyczny", "Skorygowano krzywą grzewczą",
        "Skasowano błąd i zresetowano sterownik", "Sprawdzono ciśnienie układu CO",
        "Odpowietrzono układ pompy ciepła", "Wymieniono czujnik temperatury"
    ],
    "foto": [
        "Zresetowano i zaktualizowano falownik", "Sprawdzono napięcia stringów DC",
        "Dokręcono złącza w rozdzielnicy AC/DC", "Ponownie skonfigurowano Wi-Fi / monitoring",
        "Wymieniono konektory MC4", "Przetestowano zabezpieczenia nadprądowe"
    ],
    "magazyn": [
        "Zaktualizowano oprogramowanie BMS", "Wykonano kalibrację SOC baterii",
        "Sprawdzono komunikację CAN/RS485 z falownikiem", "Zresetowano moduł bateryjny",
        "Zweryfikowano poprawność ładowania/rozładowania"
    ]
}

QUICK_PARTS = {
    "solar": [
        "Glikol solarny 5L", "Glikol solarny 10L", "Pompa obiegowa Wilo 25/6",
        "Naczynie przeponowe 18L", "Naczynie przeponowe 24L", "Odpowietrznik automatyczny",
        "Zawór zwrotny 3/4 cala", "Czujnik temperatury PT1000"
    ],
    "pompa": [
        "Zawór trójdrogowy 1 cal", "Czujnik przepływu", "Filtr magnetyczny 1 cal",
        "Odpowietrznik automatyczny", "Czujnik temperatury NTC", "Grupa bezpieczeństwa CO"
    ],
    "foto": [
        "Komplet złączy MC4", "Ogranicznik przepięć AC/DC", "Wyłącznik nadprądowy B16/C16",
        "Moduł komunikacyjny Wi-Fi / Logger", "Przewód solarny 6mm2 (5m)"
    ],
    "magazyn": [
        "Przewód komunikacyjny CAN/RS485", "Bezpiecznik DC 125A", "Moduł BMS",
        "Zestaw przewodów mocy DC"
    ]
}

QUICK_RECOMMENDATIONS = [
    "Kontrola ciśnienia na manometrze za 48h",
    "Nie wyłączać zasilania głównego sterownika",
    "Obserwować pracę instalacji przez najbliższe 7 dni",
    "Zgłosić do biura w razie ponownego pojawienia się błędu",
    "Zalecany przegląd okresowy przed kolejnym sezonem",
    "Instalacja sprawna, poinstruowano klienta o obsłudze"
]

# =========================================================================
# OBSŁUGA GOOGLE SHEETS ORAZ PLIKÓW LOKALNYCH
# =========================================================================
def get_gsheet_connection():
    if not GSPREAD_AVAILABLE:
        return None
    try:
        if "gcp_service_account" not in st.secrets:
            return None
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        creds = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"], scopes=scopes
        )
        client = gspread.authorize(creds)
        return client.open(GSHEET_NAME)
    except Exception as e:
        print(f"Błąd połączenia z Google Sheets: {e}")
        return None

def get_gsheet_worksheet(sheet_title="Wszystkie posortowane", create_if_missing=False):
    sh = get_gsheet_connection()
    if sh is None:
        return None
    try:
        return sh.worksheet(sheet_title)
    except Exception:
        if create_if_missing:
            try:
                return sh.add_worksheet(title=sheet_title, rows="500", cols="20")
            except Exception:
                return None
        if sheet_title == "Wszystkie posortowane":
            try:
                return sh.get_worksheet(0)
            except Exception:
                return None
        return None

def sync_to_gsheet(df_to_save):
    ws = get_gsheet_worksheet("Wszystkie posortowane", create_if_missing=False)
    if ws is None or df_to_save.empty:
        return False
    try:
        clean_df = df_to_save.drop(columns=['Klient_ID'], errors='ignore').fillna('')
        clean_df = clean_df.astype(str)
        data_to_upload = [clean_df.columns.values.tolist()] + clean_df.values.tolist()
        ws.clear()
        ws.update(data_to_upload)
        return True
    except Exception as e:
        print(f"Błąd zapisu do Google Sheets: {e}")
        return False

# =========================================================================
# MODUŁ CENNIKA I HISTORII ZMIAN CEN
# =========================================================================
def load_pricelist():
    ws_cennik = get_gsheet_worksheet("Cennik", create_if_missing=True)
    if ws_cennik is not None:
        try:
            records = ws_cennik.get_all_records()
            if records:
                df_c = pd.DataFrame(records)
                df_c['Cena'] = pd.to_numeric(df_c['Cena'], errors='coerce').fillna(0.0)
                return df_c
        except Exception:
            pass

    if os.path.exists(CENNIK_FILE):
        try:
            df_c = pd.read_csv(CENNIK_FILE)
            df_c['Cena'] = pd.to_numeric(df_c['Cena'], errors='coerce').fillna(0.0)
            return df_c
        except Exception:
            pass

    df_default = pd.DataFrame(DEFAULT_PRICELIST)
    save_pricelist(df_default)
    return df_default

def save_pricelist(df_c):
    try:
        df_c.to_csv(CENNIK_FILE, index=False, encoding='utf-8-sig')
    except Exception:
        pass
    ws_cennik = get_gsheet_worksheet("Cennik", create_if_missing=True)
    if ws_cennik is not None:
        try:
            clean_c = df_c.fillna('').astype(str)
            ws_cennik.clear()
            ws_cennik.update([clean_c.columns.values.tolist()] + clean_c.values.tolist())
        except Exception:
            pass

def load_price_history():
    ws_hist = get_gsheet_worksheet("Historia_Cen", create_if_missing=True)
    if ws_hist is not None:
        try:
            records = ws_hist.get_all_records()
            if records:
                return pd.DataFrame(records)
        except Exception:
            pass

    if os.path.exists(CENNIK_HIST_FILE):
        try:
            return pd.read_csv(CENNIK_HIST_FILE)
        except Exception:
            pass

    return pd.DataFrame(columns=["Data zmiany", "Kod", "Nazwa pozycji", "Stara cena (zł)", "Nowa cena (zł)", "Zmienił"])

def append_price_history(history_entries):
    if not history_entries:
        return
    df_hist = load_price_history()
    new_df = pd.DataFrame(history_entries)
    updated_hist = pd.concat([new_df, df_hist], ignore_index=True)
    try:
        updated_hist.to_csv(CENNIK_HIST_FILE, index=False, encoding='utf-8-sig')
    except Exception:
        pass
    ws_hist = get_gsheet_worksheet("Historia_Cen", create_if_missing=True)
    if ws_hist is not None:
        try:
            clean_h = updated_hist.fillna('').astype(str)
            ws_hist.clear()
            ws_hist.update([clean_h.columns.values.tolist()] + clean_h.values.tolist())
        except Exception:
            pass
    st.session_state.price_history_df = updated_hist

# =========================================================================
# FUNKCJE POMOCNICZE I ŁADOWANIE BAZY ZGŁOSZEŃ
# =========================================================================
def get_clean_client_id(row):
    tel = str(row.get('Numer telefonu', '')).strip()
    tel_digits = "".join([c for c in tel if c.isdigit()])
    imie = str(row.get('Imię', '')).strip().lower()
    nazwisko = str(row.get('Nazwisko', '')).strip().lower()
    if tel_digits:
        return f"tel_{tel_digits}"
    return f"osoba_{imie}_{nazwisko}"

def get_region_name_from_code(kod_pocztowy):
    clean = str(kod_pocztowy).strip()
    if clean and clean[0].isdigit():
        cyfra = clean[0]
        for r in REGION_DEFINITIONS:
            if r['kod'] == cyfra:
                return r['nazwa']
    return "Rejon 0 (Warszawa / Mazowieckie)"

def build_google_maps_multi_route(tickets_df):
    addresses = []
    for _, r in tickets_df.iterrows():
        addr = f"{r.get('Adres', '')}, {r.get('Kod pocztowy', '')} {r.get('Miejscowość', '')}".strip(', ')
        if addr and addr != "":
            addresses.append(urllib.parse.quote(addr))
    if not addresses:
        return ""
    return "https://www.google.com/maps/dir/" + "/".join(addresses)

def wyswietl_multimedia(ticket_nr):
    ticket_folder = os.path.join(ATTACHMENTS_DIR, str(ticket_nr))
    if os.path.exists(ticket_folder):
        files = os.listdir(ticket_folder)
        if files:
            st.markdown("📷 **Załączone materiały (zdjęcia / podpis / wideo):**")
            cols = st.columns(3)
            for idx, filename in enumerate(files):
                file_path = os.path.join(ticket_folder, filename)
                ext = filename.split('.')[-1].lower()
                with cols[idx % 3]:
                    if ext in ['jpg', 'jpeg', 'png', 'webp']:
                        st.image(file_path, caption=filename, use_container_width=True)
                    elif ext in ['mp4', 'mov', 'avi']:
                        st.video(file_path)
                        st.caption(f"🎥 {filename}")

def load_all_data():
    df = pd.DataFrame()
    ws = get_gsheet_worksheet("Wszystkie posortowane", create_if_missing=False)
    loaded_from_cloud = False

    if ws is not None:
        try:
            records = ws.get_all_records()
            if records:
                df = pd.DataFrame(records)
                loaded_from_cloud = True
                st.session_state.cloud_active = True
        except Exception as e:
            print(f"Błąd odczytu z Google Sheets: {e}")

    if not loaded_from_cloud:
        st.session_state.cloud_active = False
        base_df = pd.DataFrame()
        if os.path.exists(EXCEL_FILE):
            try:
                base_df = pd.read_excel(EXCEL_FILE, sheet_name="Wszystkie posortowane")
            except Exception:
                base_df = pd.read_excel(EXCEL_FILE, sheet_name=0)

        if os.path.exists(NOWE_FILE):
            try:
                nowe_df = pd.read_csv(NOWE_FILE, dtype=str)
                df = pd.concat([nowe_df, base_df], ignore_index=True)
                df = df.drop_duplicates(subset=['nr.'], keep='first')
            except Exception:
                df = base_df
        else:
            df = base_df

        if ws is not None and not df.empty:
            if sync_to_gsheet(df):
                st.session_state.cloud_active = True

    cols_to_str = ['nr.', 'INSTALACJA', 'PRZYJMUJĄCY', 'Imię', 'Nazwisko', 'Numer telefonu', 
                   'Kod pocztowy', 'Miejscowość', 'Adres', 'Rejon Polski', 'STATUS ZGŁOSZENIA', 
                   'SERWISANT', 'OPIS OD KLIENTA', 'OPIS', 'INFORMACJE DODATKOWE', 'Kwota', 'Płatność ', 'E-mail']
    for c in cols_to_str:
        if c in df.columns:
            df[c] = df[c].fillna('').astype(str)

    if 'DATA ZGŁOSZENIA' in df.columns:
        df['DATA ZGŁOSZENIA'] = df['DATA ZGŁOSZENIA'].fillna('').astype(str)

    for opt_col in ['Zużyte materiały', 'Czynności dodatkowe', 'Ilość przyjazdów', 'Kwota przyjęta', 
                    'Komentarz serwisanta', 'Email handlowca', 'Etap oferty', 'Data oferty', 
                    'Komentarz koordynatora', 'Zalecenia dla klienta', 'Przyczyna nieukończenia', 
                    'Brakujące materiały', 'Szczegóły oferty', 'Tabliczki_OK',
                    'Kolejnosc_trasy', 'Data_trasy', 'Trasa_aktywna', 'Odbior_klienta',
                    'Zdalne', 'Termin_koncowy', 'Pozycje_cennika', 'Cena_bazowa', 'Rabat_proc']:
        if opt_col not in df.columns:
            df[opt_col] = ''
        else:
            df[opt_col] = df[opt_col].fillna('').astype(str)

    if 'Wymaga oferty' not in df.columns:
        df['Wymaga oferty'] = False
    else:
        df['Wymaga oferty'] = df['Wymaga oferty'].astype(str).str.lower().isin(['true', '1', 'tak'])

    df['Klient_ID'] = df.apply(get_clean_client_id, axis=1)
    return df

def sync_excel_file(df_to_save):
    if df_to_save.empty:
        return
    if sync_to_gsheet(df_to_save):
        st.session_state.cloud_active = True

    try:
        if os.path.exists(EXCEL_FILE):
            shutil.copyfile(EXCEL_FILE, f"{EXCEL_FILE}.bak")
        clean_save_df = df_to_save.drop(columns=['Klient_ID'], errors='ignore')
        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
            clean_save_df.to_excel(writer, sheet_name="Wszystkie posortowane", index=False)
    except Exception as e:
        print(f"Błąd zapisu lokalnego Excel: {e}")

def save_ticket_persistently(ticket_dict, current_full_df=None):
    df_single = pd.DataFrame([ticket_dict])
    if os.path.exists(NOWE_FILE):
        try:
            existing_nowe = pd.read_csv(NOWE_FILE, dtype=str)
            if str(ticket_dict.get('nr.', '')) in existing_nowe['nr.'].astype(str).values:
                existing_nowe = existing_nowe[existing_nowe['nr.'].astype(str) != str(ticket_dict.get('nr.', ''))]
            updated = pd.concat([df_single, existing_nowe], ignore_index=True)
            updated.to_csv(NOWE_FILE, index=False, encoding='utf-8-sig')
        except Exception:
            df_single.to_csv(NOWE_FILE, index=False, encoding='utf-8-sig')
    else:
        df_single.to_csv(NOWE_FILE, index=False, encoding='utf-8-sig')
    
    if current_full_df is not None:
        sync_excel_file(current_full_df)

def save_multiple_tickets_persistently(updated_df_rows, current_full_df=None):
    if updated_df_rows.empty:
        return
    if os.path.exists(NOWE_FILE):
        try:
            existing_nowe = pd.read_csv(NOWE_FILE, dtype=str)
            updated_nrs = updated_df_rows['nr.'].astype(str).values
            existing_nowe = existing_nowe[~existing_nowe['nr.'].astype(str).isin(updated_nrs)]
            merged = pd.concat([updated_df_rows, existing_nowe], ignore_index=True)
            merged.to_csv(NOWE_FILE, index=False, encoding='utf-8-sig')
        except Exception:
            updated_df_rows.to_csv(NOWE_FILE, index=False, encoding='utf-8-sig')
    else:
        updated_df_rows.to_csv(NOWE_FILE, index=False, encoding='utf-8-sig')

    if current_full_df is not None:
        sync_excel_file(current_full_df)

# Inicjalizacja stanu aplikacji
if 'cloud_active' not in st.session_state:
    st.session_state.cloud_active = False
if 'df' not in st.session_state:
    st.session_state.df = load_all_data()
if 'pricelist_df' not in st.session_state:
    st.session_state.pricelist_df = load_pricelist()
if 'price_history_df' not in st.session_state:
    st.session_state.price_history_df = load_price_history()
if 'user_role' not in st.session_state:
    st.session_state.user_role = None
if 'tech_pin' not in st.session_state:
    st.session_state.tech_pin = None
if 'selected_client_id' not in st.session_state:
    st.session_state.selected_client_id = None
if 'show_new_ticket_modal' not in st.session_state:
    st.session_state.show_new_ticket_modal = False
if 'show_new_client_modal' not in st.session_state:
    st.session_state.show_new_client_modal = False
if 'show_edit_client_modal' not in st.session_state:
    st.session_state.show_edit_client_modal = False
if 'active_ticket_nr' not in st.session_state:
    st.session_state.active_ticket_nr = None
if 'day_started' not in st.session_state:
    st.session_state.day_started = False
if 'last_closed_summary' not in st.session_state:
    st.session_state.last_closed_summary = None
if 'coord_sms_notification' not in st.session_state:
    st.session_state.coord_sms_notification = None

# Stan dla ekranowego kreatora (Wizard) w panelu serwisanta
if 'tech_main_screen' not in st.session_state:
    st.session_state.tech_main_screen = "otwarte"
if 'wizard_step' not in st.session_state:
    st.session_state.wizard_step = 1
if 'wizard_data' not in st.session_state:
    st.session_state.wizard_data = {}

def get_next_ticket_id():
    existing_nrs = st.session_state.df['nr.'].astype(str).tolist()
    max_id = 0
    for n in existing_nrs:
        clean = n.replace('N', '').replace('#', '').replace('K', '').strip()
        if clean.isdigit():
            val = int(clean)
            if val > max_id:
                max_id = val
    return max_id + 1

# --- EKRAN STARTOWY: WYBÓR ROLI ---
if not st.session_state.user_role:
    st.markdown("<h2 style='text-align: center; margin-top: 1.5rem;'>⚡ PROMAG SERWIS</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #4a5568;'>Wybierz moduł pracy:</p>", unsafe_allow_html=True)
    st.write("")
    
    col_r1, col_r2 = st.columns(2)
    with col_r1:
        st.info("### 🖥️ Panel Koordynatora")
        st.write("Baza klientów, zlecenia zdalne, dynamiczny cennik z historią i rabatami, powiadomienia SMS dla klientów i planowanie tras.")
        if st.button("Wejdź jako Koordynator", use_container_width=True, type="primary"):
            st.session_state.user_role = "koordynator"
            st.rerun()

    with col_r2:
        st.success("### 📱 Panel Serwisanta")
        st.write("Wygodny interfejs jednoekranowy pod obsługę jedną dłonią, trasa na dziś, szybkie kafelki i podpis klienta na ekranie.")
        if st.button("Wejdź jako Serwisant (PIN)", use_container_width=True):
            st.session_state.user_role = "serwisant"
            st.rerun()
    st.stop()

# =========================================================================
# 1. MODUŁ KOORDYNATORA
# =========================================================================
if st.session_state.user_role == "koordynator":
    df = st.session_state.df
    pricelist_df = st.session_state.pricelist_df

    st.sidebar.title("⚡ Koordynator")
    if st.session_state.cloud_active:
        st.sidebar.success("☁️ Baza Google Drive: AKTYWNA")
    else:
        st.sidebar.info("💾 Tryb lokalny (Excel / CSV)")

    if st.sidebar.button("🔄 Odśwież dane z chmury", use_container_width=True):
        st.session_state.df = load_all_data()
        st.session_state.pricelist_df = load_pricelist()
        st.session_state.price_history_df = load_price_history()
        st.rerun()

    if st.sidebar.button("🚪 Zmień moduł (Wyloguj)", use_container_width=True):
        st.session_state.user_role = None
        st.session_state.selected_client_id = None
        st.session_state.show_edit_client_modal = False
        st.session_state.show_new_ticket_modal = False
        st.session_state.active_ticket_nr = None
        st.rerun()

    st.sidebar.markdown("---")
    st.sidebar.subheader("📍 Rejony Pocztowe")
    st.sidebar.caption(f"Sortowanie wg zgłoszeń od **{CUTOFF_DATE}**:")

    is_active_since_cutoff = df['DATA ZGŁOSZENIA'].str[:10] >= CUTOFF_DATE

    region_stats = []
    for r_def in REGION_DEFINITIONS:
        r_kod = r_def['kod']
        mask_r = (
            df['Rejon Polski'].str.contains(f"Rejon {r_kod}", case=False, na=False) |
            df['Kod pocztowy'].str.startswith(r_kod)
        )
        total_in_r = mask_r.sum()
        active_in_r = (mask_r & is_active_since_cutoff).sum()
        pending_in_r = (mask_r & is_active_since_cutoff & ~df['STATUS ZGŁOSZENIA'].str.lower().str.contains('zakończ|anulow', na=False)).sum()
        
        region_stats.append({
            'kod': r_kod,
            'krotka': r_def['krotka'],
            'pelna': r_def['nazwa'],
            'active': active_in_r,
            'pending': pending_in_r,
            'total': total_in_r
        })

    region_stats_sorted = sorted(region_stats, key=lambda x: x['active'], reverse=True)
    all_region_keys = ["ALL"] + [r['kod'] for r in region_stats_sorted]
    total_active_all = is_active_since_cutoff.sum()

    region_labels = {
        "ALL": f"🌐 Wszystkie rejony ({total_active_all} zgłoszeń od {CUTOFF_DATE})"
    }
    for r in region_stats_sorted:
        lbl = f"📍 {r['krotka']} — {r['active']} zgłoszeń"
        if r['pending'] > 0:
            lbl += f" ({r['pending']} w toku)"
        region_labels[r['kod']] = lbl

    selected_region_code = st.sidebar.radio(
        "Wybierz rejon do podglądu:",
        options=all_region_keys,
        format_func=lambda x: region_labels.get(x, x)
    )

    st.title("🖥️ Pulpit Koordynatora")

    # Baner z gotowym SMS-em dla klienta po utworzeniu zlecenia
    if st.session_state.coord_sms_notification:
        sms_info = st.session_state.coord_sms_notification
        st.success(f"✅ Zlecenie **#{sms_info['nr']}** dla klienta **{sms_info['klient']}** zostało zapisane! Wyślij teraz powiadomienie SMS:")
        
        col_sms_btn1, col_sms_btn2, col_sms_close = st.columns([2, 2, 1])
        with col_sms_btn1:
            if sms_info['tel_clean']:
                encoded_body = urllib.parse.quote(sms_info['tekst'])
                st.markdown(
                    f'<a href="sms:{sms_info["tel_clean"]}?body={encoded_body}" target="_blank" style="text-decoration:none;">'
                    f'<button style="width:100%; height:48px; border-radius:10px; background:#2b6cb0; color:white; border:none; font-weight:bold; font-size:0.95rem; cursor:pointer;">'
                    f'📲 1. Wyślij automatycznie (Aplikacja SMS / Telefon)</button></a>',
                    unsafe_allow_html=True
                )
        with col_sms_btn2:
            st.markdown(
                f'<a href="https://messages.google.com/web/conversations/new" target="_blank" style="text-decoration:none;">'
                f'<button style="width:100%; height:48px; border-radius:10px; background:#38a169; color:white; border:none; font-weight:bold; font-size:0.95rem; cursor:pointer;">'
                f'💬 2. Otwórz Wiadomości Google (Web) na komputerze</button></a>',
                unsafe_allow_html=True
            )
        with col_sms_close:
            if st.button("✖ Ukryj", use_container_width=True):
                st.session_state.coord_sms_notification = None
                st.rerun()

        st.caption(f"Numer telefonu klienta: **{sms_info['tel_clean']}** — Kliknij ikonę kopiowania w prawym górnym rogu poniższej ramki, aby skopiować treść do Wiadomości Google:")
        st.code(sms_info['tekst'], language=None)
        st.markdown("---")

    # Monitor pilnych ofert handlowych z terenu
    is_wymaga_oferty = df['Wymaga oferty'].astype(str).str.lower().isin(['true', '1', 'tak'])
    is_pilne_etap = df['Etap oferty'].astype(str).str.contains('PILNE', case=False, na=False)
    pilne_oferty_df = df[is_wymaga_oferty & is_pilne_etap]

    if not pilne_oferty_df.empty:
        st.warning(f"🚨 **UWAGA: Masz {len(pilne_oferty_df)} zgłoszeń z terenu wymagających PILNEJ oferty handlowej!** Sprawdź zakładkę *Lejek Ofert Handlowych*.")

    # Licznik zgłoszeń zdalnych do rozwiązania przez koordynatora
    is_zdalne_open = (
        df['Zdalne'].astype(str).str.lower().isin(['true', '1', 'tak']) &
        ~df['STATUS ZGŁOSZENIA'].str.lower().str.contains('zakończ|anulow', na=False)
    )
    zdalne_open_df = df[is_zdalne_open]

    tab_k1, tab_k_zdalne, tab_k2, tab_k3, tab_k_cennik = st.tabs([
        "🔍 Baza Klientów & Zgłoszenia",
        f"📞 Zgłoszenia Zdalne ({len(zdalne_open_df)})",
        "📑 Lejek Ofert Handlowych",
        "🗺️ Planowanie Tras",
        "💰 Cennik i Historia Cen"
    ])

    # --- ZAKŁADKA: ZGŁOSZENIA ZDALNE (DO ROZWIĄZANIA PRZEZ KOORDYNATORA) ---
    with tab_k_zdalne:
        st.subheader("📞 Zgłoszenia do rozwiązania zdalnie (Helpdesk / Wsparcie telefoniczne)")
        st.caption("Tutaj trafiają sprawy zaznaczone przy rejestracji jako zdalne. Nie obciążają one trasy serwisantów – możesz je rozwiązać i zamknąć bezpośrednio tutaj.")

        if zdalne_open_df.empty:
            st.success("🎉 Wszystkie zgłoszenia zdalne zostały rozwiązane!")
        else:
            for idx_z, z_row in zdalne_open_df.iterrows():
                z_nr = z_row.get('nr.', '')
                z_klient = f"{z_row.get('Imię', '')} {z_row.get('Nazwisko', '')}"
                z_tel = z_row.get('Numer telefonu', '')
                z_miasto = f"{z_row.get('Kod pocztowy', '')} {z_row.get('Miejscowość', '')}"
                z_inst = str(z_row.get('INSTALACJA', '')).upper()

                with st.expander(f"📞 Zdalne #{z_nr} | {z_klient} (📞 {z_tel}) | {z_inst} | Data: {str(z_row.get('DATA ZGŁOSZENIA', ''))[:16]}", expanded=True):
                    c_z1, c_z2 = st.columns(2)
                    with c_z1:
                        st.markdown(f"**Klient:** {z_klient} (**📞 {z_tel}**)")
                        st.markdown(f"**Adres:** {z_miasto}, {z_row.get('Adres', '')}")
                        st.markdown(f"**Opis problemu:** {z_row.get('OPIS OD KLIENTA', '')}")
                    with c_z2:
                        st.markdown(f"**Wycena:** {z_row.get('Kwota', '0')} zł (Rabat: {z_row.get('Rabat_proc', '0')}%)")
                        st.markdown(f"**Termin końcowy:** {z_row.get('Termin_koncowy', 'Brak')}")
                        if z_row.get('Komentarz koordynatora', ''):
                            st.info(f"🔒 **Notatka:** {z_row.get('Komentarz koordynatora')}")

                    roz_notatka = st.text_input(
                        "Jak rozwiązano problem zdalnie (notatka do zamknięcia):",
                        placeholder="np. Pomoc telefoniczna – zrestartowano falownik i skonfigurowano Wi-Fi...",
                        key=f"zdalne_note_{z_nr}"
                    )

                    col_zb1, col_zb2 = st.columns(2)
                    with col_zb1:
                        if st.button("✅ ZAMKNIJ (Rozwiązano zdalnie)", key=f"btn_z_close_{z_nr}", type="primary", use_container_width=True):
                            df.at[idx_z, 'STATUS ZGŁOSZENIA'] = 'zakończone (zdalnie)'
                            df.at[idx_z, 'DATA SERWISU'] = datetime.date.today().strftime("%Y-%m-%d")
                            df.at[idx_z, 'SERWISANT'] = 'Zdalnie (Koordynator)'
                            df.at[idx_z, 'Komentarz serwisanta'] = roz_notatka.strip() or "Rozwiązano zdalnie przez telefon."
                            st.session_state.df = df
                            save_ticket_persistently(df.loc[idx_z].to_dict(), current_full_df=st.session_state.df)
                            st.success(f"Zgłoszenie #{z_nr} zostało zamknięte zdalnie!")
                            st.rerun()
                    with col_zb2:
                        if st.button("🚗 Przekaż jednak do wyjazdu dla serwisanta", key=f"btn_z_to_field_{z_nr}", use_container_width=True):
                            df.at[idx_z, 'Zdalne'] = 'False'
                            df.at[idx_z, 'STATUS ZGŁOSZENIA'] = 'w toku'
                            st.session_state.df = df
                            save_ticket_persistently(df.loc[idx_z].to_dict(), current_full_df=st.session_state.df)
                            st.info(f"Zgłoszenie #{z_nr} przekazano do kolejki terenowej serwisanta.")
                            st.rerun()

    # --- ZAKŁADKA: CENNIK I HISTORIA CEN ---
    with tab_k_cennik:
        st.subheader("💰 Zarządzanie Cennikiem Usług i Części")
        st.caption("Każda zmiana ceny w poniższej tabeli jest zapisywana na stałe w bazie oraz automatycznie odnotowywana w tabeli Historii Cen poniżej.")

        edited_pricelist = st.data_editor(
            pricelist_df,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            key="pricelist_editor_table"
        )

        if st.button("💾 ZAPISZ ZMIANY W CENNIKU (Z ZACHOWANIEM HISTORII)", type="primary"):
            now_dt = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
            old_map = {str(r['Nazwa']).strip(): float(r['Cena']) for _, r in pricelist_df.iterrows() if str(r.get('Nazwa', '')).strip()}
            history_logs = []

            for _, new_r in edited_pricelist.iterrows():
                nazwa_p = str(new_r.get('Nazwa', '')).strip()
                kod_p = str(new_r.get('Kod', '')).strip() or "NOWY"
                if not nazwa_p:
                    continue
                try:
                    nowa_cena = float(new_r.get('Cena', 0.0))
                except Exception:
                    nowa_cena = 0.0

                if nazwa_p in old_map:
                    stara_cena = old_map[nazwa_p]
                    if round(stara_cena, 2) != round(nowa_cena, 2):
                        history_logs.append({
                            "Data zmiany": now_dt,
                            "Kod": kod_p,
                            "Nazwa pozycji": nazwa_p,
                            "Stara cena (zł)": f"{stara_cena:.2f}",
                            "Nowa cena (zł)": f"{nowa_cena:.2f}",
                            "Zmienił": "Koordynator"
                        })
                else:
                    history_logs.append({
                        "Data zmiany": now_dt,
                        "Kod": kod_p,
                        "Nazwa pozycji": nazwa_p,
                        "Stara cena (zł)": "Nowa pozycja",
                        "Nowa cena (zł)": f"{nowa_cena:.2f}",
                        "Zmienił": "Koordynator"
                    })

            clean_pricelist = edited_pricelist[edited_pricelist['Nazwa'].astype(str).str.strip() != ''].copy()
            clean_pricelist['Cena'] = pd.to_numeric(clean_pricelist['Cena'], errors='coerce').fillna(0.0)
            save_pricelist(clean_pricelist)
            st.session_state.pricelist_df = clean_pricelist

            if history_logs:
                append_price_history(history_logs)
                st.success(f"✅ Zapisano nowy cennik oraz dodano {len(history_logs)} wpis(ów) do historii zmian cen!")
            else:
                st.info("Zapisano cennik (brak zmian w kwotach).")
            st.rerun()

        st.markdown("---")
        st.subheader("📜 Historia zmian cen")
        hist_df = st.session_state.price_history_df
        if hist_df.empty:
            st.info("Brak zarejestrowanych zmian cen w historii.")
        else:
            st.dataframe(hist_df, use_container_width=True, hide_index=True)

    # --- ZAKŁADKA: PLANOWANIE TRAS ---
    with tab_k3:
        st.subheader("🗺️ Zaplanuj i zatwierdź trasę dla serwisanta")
        st.caption("Wybierz rejon, ułóż kolejność wizyt na jutro (lub dziś) i zatwierdź. Zlecenia zdalne są automatycznie odfiltrowane.")

        col_tr1, col_tr2 = st.columns(2)
        with col_tr1:
            plan_reg_options = [r['kod'] for r in REGION_DEFINITIONS]
            plan_reg_names = {r['kod']: r['nazwa'] for r in REGION_DEFINITIONS}
            default_plan_reg = selected_region_code if selected_region_code != "ALL" else "0"
            wybrany_rejon_trasy = st.selectbox(
                "1. Wybierz rejon do zaplanowania trasy:",
                options=plan_reg_options,
                index=plan_reg_options.index(default_plan_reg) if default_plan_reg in plan_reg_options else 0,
                format_func=lambda x: plan_reg_names.get(x, x)
            )
        with col_tr2:
            jutro_default = datetime.date.today() + datetime.timedelta(days=1)
            data_planowanej_trasy = st.date_input("2. Data realizacji trasy:", value=jutro_default)

        mask_plan_reg = (
            df['Rejon Polski'].str.contains(f"Rejon {wybrany_rejon_trasy}", case=False, na=False) |
            df['Kod pocztowy'].str.startswith(wybrany_rejon_trasy)
        )
        mask_open_plan = ~df['STATUS ZGŁOSZENIA'].str.lower().str.contains('zakończ|anulow|nowy klient', na=False)
        mask_not_remote = ~df['Zdalne'].astype(str).str.lower().isin(['true', '1', 'tak'])
        mask_cutoff_plan = df['DATA ZGŁOSZENIA'].str[:10] >= CUTOFF_DATE
        mask_ready_offer = (
            (df['Etap oferty'] != 'Czeka na ofertę handlowca') &
            (df['Etap oferty'] != 'Oczekiwanie na akceptację klienta')
        )

        kandydaci_df = df[mask_plan_reg & mask_open_plan & mask_not_remote & mask_cutoff_plan & mask_ready_offer].copy()

        if kandydaci_df.empty:
            st.info("Brak otwartych zleceń terenowych w tym rejonie do zaplanowania trasy.")
        else:
            st.markdown("##### 3. Wybierz zlecenia do trasy i ustaw ich kolejność (1, 2, 3...):")
            route_editor_rows = []
            for _, r_k in kandydaci_df.iterrows():
                curr_order = str(r_k.get('Kolejnosc_trasy', '')).strip()
                is_in_route = str(r_k.get('Trasa_aktywna', '')).lower() in ['true', '1', 'tak']
                order_num = int(curr_order) if curr_order.isdigit() else 99
                route_editor_rows.append({
                    "Włącz do trasy": is_in_route,
                    "Kolejność (1,2,3...)": order_num if is_in_route and order_num < 99 else 1,
                    "Nr": str(r_k.get('nr.', '')),
                    "Termin do": str(r_k.get('Termin_koncowy', '')),
                    "Miejscowość": r_k.get('Miejscowość', ''),
                    "Adres": r_k.get('Adres', ''),
                    "Klient": f"{r_k.get('Imię', '')} {r_k.get('Nazwisko', '')}",
                    "Instalacja": str(r_k.get('INSTALACJA', '')).upper(),
                    "Opis usterki": r_k.get('OPIS OD KLIENTA', '') or r_k.get('OPIS', '')
                })

            route_edit_df = pd.DataFrame(route_editor_rows)
            edited_route = st.data_editor(
                route_edit_df,
                use_container_width=True,
                hide_index=True,
                disabled=["Nr", "Termin do", "Miejscowość", "Adres", "Klient", "Instalacja", "Opis usterki"],
                key=f"route_editor_{wybrany_rejon_trasy}"
            )

            selected_for_route = edited_route[edited_route["Włącz do trasy"] == True].sort_values(by="Kolejność (1,2,3...)")

            if not selected_for_route.empty:
                st.markdown("##### 📍 Podgląd ułożonej trasy:")
                st.markdown(" ➔ ".join([f"**[{i}] {r['Miejscowość']}**" for i, (_, r) in enumerate(selected_for_route.iterrows(), 1)]))

                selected_nrs = selected_for_route['Nr'].astype(str).tolist()
                ordered_df_for_map = pd.concat([df[df['nr.'].astype(str) == n] for n in selected_nrs], ignore_index=True)
                gmaps_url = build_google_maps_multi_route(ordered_df_for_map)

                col_act1, col_act2, col_act3 = st.columns([2, 2, 1])
                with col_act1:
                    if st.button("✅ ZATWIERDŹ I OPUBLIKUJ TRASĘ DLA SERWISANTA", type="primary", use_container_width=True):
                        df.loc[mask_plan_reg, 'Trasa_aktywna'] = 'False'
                        df.loc[mask_plan_reg, 'Kolejnosc_trasy'] = ''

                        for final_pos, (_, s_row) in enumerate(selected_for_route.iterrows(), 1):
                            nr_val = str(s_row['Nr'])
                            m_idx = df['nr.'].astype(str) == nr_val
                            df.loc[m_idx, 'Trasa_aktywna'] = 'True'
                            df.loc[m_idx, 'Kolejnosc_trasy'] = str(final_pos)
                            df.loc[m_idx, 'Data_trasy'] = str(data_planowanej_trasy)

                        st.session_state.df = df
                        save_multiple_tickets_persistently(df[mask_plan_reg], current_full_df=st.session_state.df)
                        st.success(f"🚀 Trasa na dzień {data_planowanej_trasy} ({len(selected_for_route)} pkt) wysłana do serwisanta!")
                        st.rerun()

                with col_act2:
                    if gmaps_url:
                        st.markdown(f'<a href="{gmaps_url}" target="_blank" style="text-decoration:none;"><button style="width:100%; height:48px; border-radius:10px; background:#38a169; color:white; border:none; font-weight:bold; cursor:pointer;">🗺️ Sprawdź pętlę w Google Maps</button></a>', unsafe_allow_html=True)

                with col_act3:
                    if st.button("🗑️ Wyczyść trasę", use_container_width=True):
                        df.loc[mask_plan_reg, 'Trasa_aktywna'] = 'False'
                        df.loc[mask_plan_reg, 'Kolejnosc_trasy'] = ''
                        st.session_state.df = df
                        save_multiple_tickets_persistently(df[mask_plan_reg], current_full_df=st.session_state.df)
                        st.info("Wyczyszczono zaplanowaną trasę.")
                        st.rerun()

    # --- ZAKŁADKA: LEJEK OFERTOWY ---
    with tab_k2:
        st.subheader("📑 Zgłoszenia wymagające oferty handlowej")
        oferty_df = df[df['Wymaga oferty'].astype(str).str.lower().isin(['true', '1', 'tak'])]

        if oferty_df.empty:
            st.info("Brak zgłoszeń w trakcie ofertowania.")
        else:
            for idx, o_row in oferty_df.iterrows():
                o_nr = o_row['nr.']
                o_etap = o_row.get('Etap oferty', 'Czeka na ofertę handlowca')
                o_email = o_row.get('Email handlowca', 'Nie podano')
                o_klient = f"{o_row.get('Imię', '')} {o_row.get('Nazwisko', '')}"
                o_tel = o_row.get('Numer telefonu', '')
                o_miasto = f"{o_row.get('Kod pocztowy', '')} {o_row.get('Miejscowość', '')}"
                is_urgent_service = "PILNE" in str(o_etap)

                title_prefix = "🚨 [PILNE Z TERENU] " if is_urgent_service else ""
                with st.expander(f"{title_prefix}Zgłoszenie #{o_nr} | Klient: {o_klient} ({o_miasto}) | Status: {o_etap}", expanded=is_urgent_service):
                    if is_urgent_service:
                        st.error(f"🚨 **SERWISANT ({o_row.get('SERWISANT', 'Serwisant')}) ZGŁOSIŁ POTRZEBĘ OFERTY Z OBIEKTU!**\n\n**Co wycenić:** {o_row.get('Szczegóły oferty', 'Brak opisu')}\n\n**Zalecenia:** {o_row.get('Zalecenia dla klienta', 'Brak')}")

                    c_o1, c_o2 = st.columns(2)
                    with c_o1:
                        st.markdown(f"**Klient:** {o_klient} (📞 {o_tel})")
                        st.markdown(f"**Lokalizacja:** {o_miasto}, {o_row.get('Adres', '')}")
                        st.markdown(f"**Usterka bazowa:** {o_row.get('OPIS OD KLIENTA', '')}")
                    with c_o2:
                        st.markdown(f"**Przypisany handlowiec:** `{o_email}`")
                        st.markdown(f"**Aktualny etap:** **{o_etap}**")

                    wyswietl_multimedia(o_nr)

                    col_b1, col_b2 = st.columns(2)
                    if "Czeka" in str(o_etap) or "PILNE" in str(o_etap):
                        with col_b1:
                            if st.button("✅ Oferta wysłana do klienta", key=f"btn_off_sent_{o_nr}", type="primary"):
                                df.at[idx, 'Etap oferty'] = 'Oczekiwanie na akceptację klienta'
                                df.at[idx, 'STATUS ZGŁOSZENIA'] = 'Oczekiwanie na akceptację'
                                st.session_state.df = df
                                save_ticket_persistently(df.loc[idx].to_dict(), current_full_df=st.session_state.df)
                                st.rerun()
                    elif o_etap == "Oczekiwanie na akceptację klienta":
                        with col_b1:
                            if st.button("🚀 Klient zaakceptował → Przekaż do serwisu", key=f"btn_off_accept_{o_nr}", type="primary"):
                                df.at[idx, 'Etap oferty'] = 'Zaakceptowano (w serwisie)'
                                df.at[idx, 'STATUS ZGŁOSZENIA'] = 'w toku'
                                st.session_state.df = df
                                save_ticket_persistently(df.loc[idx].to_dict(), current_full_df=st.session_state.df)
                                st.rerun()
                        with col_b2:
                            if st.button("❌ Klient odrzucił ofertę", key=f"btn_off_rej_{o_nr}"):
                                df.at[idx, 'Etap oferty'] = 'Odrzucona przez klienta'
                                df.at[idx, 'STATUS ZGŁOSZENIA'] = 'anulowane'
                                save_ticket_persistently(df.loc[idx].to_dict(), current_full_df=st.session_state.df)
                                st.rerun()

    # --- ZAKŁADKA 1: BAZA KLIENTÓW & DODAWANIE ZGŁOSZEŃ ---
    with tab_k1:
        if selected_region_code != "ALL":
            region_mask = (
                df['Rejon Polski'].str.contains(f"Rejon {selected_region_code}", case=False, na=False) |
                df['Kod pocztowy'].str.startswith(selected_region_code)
            )
            region_df = df[region_mask]
            r_info = next((item for item in region_stats if item["kod"] == selected_region_code), None)
            r_title = r_info['pelna'] if r_info else f"Rejon {selected_region_code}"
            st.markdown(f"### 📍 Widok: {r_title}")
            
            active_in_curr_region = region_df[region_df['DATA ZGŁOSZENIA'].str[:10] >= CUTOFF_DATE]
            with st.expander(f"📋 Zgłoszenia w tym rejonie od {CUTOFF_DATE} ({len(active_in_curr_region)} spraw) ⬇", expanded=True):
                if active_in_curr_region.empty:
                    st.info(f"Brak nowych zgłoszeń w tym rejonie od {CUTOFF_DATE}.")
                else:
                    t_rows = []
                    for _, p_row in active_in_curr_region.iterrows():
                        t_rows.append({
                            "Nr": p_row.get('nr.', ''),
                            "Tryb": "📞 Zdalne" if str(p_row.get('Zdalne', '')).lower() == 'true' else ("🚗 Trasa #" + str(p_row.get('Kolejnosc_trasy')) if str(p_row.get('Trasa_aktywna')).lower() == 'true' else "Teren"),
                            "Termin do": p_row.get('Termin_koncowy', ''),
                            "Klient": f"{p_row.get('Imię', '')} {p_row.get('Nazwisko', '')}",
                            "Telefon": p_row.get('Numer telefonu', ''),
                            "Miejscowość": p_row.get('Miejscowość', ''),
                            "Instalacja": str(p_row.get('INSTALACJA', '')).upper(),
                            "Wycena": f"{p_row.get('Kwota', '')} zł" if p_row.get('Kwota', '') else "—",
                            "Status": p_row.get('STATUS ZGŁOSZENIA', 'w toku'),
                            "Opis usterki": p_row.get('OPIS OD KLIENTA', '') or p_row.get('OPIS', ''),
                            "ID": p_row['Klient_ID']
                        })
                    t_df = pd.DataFrame(t_rows)
                    st.dataframe(t_df.drop(columns=['ID']), use_container_width=True, hide_index=True)
                    
                    ticket_select_options = [f"#{row['Nr']} | {row['Klient']} ({row['Miejscowość']})" for _, row in t_df.iterrows()]
                    quick_pick = st.selectbox("Szybki wybór klienta z powyższej listy:", ["-- Wybierz z listy --"] + ticket_select_options)
                    if quick_pick != "-- Wybierz z listy --":
                        picked_idx = ticket_select_options.index(quick_pick)
                        st.session_state.selected_client_id = t_df.iloc[picked_idx]['ID']
                        st.rerun()
            filtered_df = region_df.copy()
        else:
            st.markdown(f"### 🌐 Widok ogólnopolski (Wszystkie zgłoszenia od {CUTOFF_DATE})")
            active_all = df[df['DATA ZGŁOSZENIA'].str[:10] >= CUTOFF_DATE]
            with st.expander(f"📋 Wszystkie bieżące zgłoszenia od {CUTOFF_DATE} ({len(active_all)} spraw) ⬇", expanded=False):
                if not active_all.empty:
                    t_all_rows = []
                    for _, p_row in active_all.iterrows():
                        t_all_rows.append({
                            "Nr": p_row.get('nr.', ''),
                            "Tryb": "📞 Zdalne" if str(p_row.get('Zdalne', '')).lower() == 'true' else "Teren",
                            "Termin do": p_row.get('Termin_koncowy', ''),
                            "Rejon": p_row.get('Rejon Polski', ''),
                            "Klient": f"{p_row.get('Imię', '')} {p_row.get('Nazwisko', '')}",
                            "Telefon": p_row.get('Numer telefonu', ''),
                            "Miejscowość": p_row.get('Miejscowość', ''),
                            "Instalacja": str(p_row.get('INSTALACJA', '')).upper(),
                            "Wycena": f"{p_row.get('Kwota', '')} zł" if p_row.get('Kwota', '') else "—",
                            "Status": p_row.get('STATUS ZGŁOSZENIA', 'w toku'),
                            "Opis usterki": p_row.get('OPIS OD KLIENTA', '') or p_row.get('OPIS', '')
                        })
                    st.dataframe(pd.DataFrame(t_all_rows), use_container_width=True, hide_index=True)
            filtered_df = df.copy()

        st.markdown("---")
        c_title, c_add_btn = st.columns([3, 1])
        with c_title:
            st.subheader("1. Baza Klientów (Wyszukaj lub Dodaj)")
        with c_add_btn:
            if st.button("➕ DODAJ NOWEGO KLIENTA", type="primary", use_container_width=True):
                st.session_state.show_new_client_modal = True

        if st.session_state.show_new_client_modal:
            st.info("👤 Wprowadź dane nowego klienta do bazy:")
            with st.form("form_dodaj_klienta"):
                nc_c1, nc_c2 = st.columns(2)
                with nc_c1:
                    nc_imie = st.text_input("Imię:*", placeholder="np. Jan")
                    nc_nazwisko = st.text_input("Nazwisko:*", placeholder="np. Kowalski")
                    nc_tel = st.text_input("Numer telefonu:*", placeholder="np. 600 123 456")
                    nc_email = st.text_input("E-mail:", placeholder="np. klient@gmail.com")
                with nc_c2:
                    nc_kod = st.text_input("Kod pocztowy:*", placeholder="np. 05-800")
                    nc_miasto = st.text_input("Miejscowość:*", placeholder="np. Pruszków")
                    nc_adres = st.text_input("Adres (ulica i nr):*", placeholder="np. Spacerowa 12")
                    auto_r = get_region_name_from_code(nc_kod)
                    all_region_names = [r['nazwa'] for r in REGION_DEFINITIONS]
                    idx_r = all_region_names.index(auto_r) if auto_r in all_region_names else 0
                    nc_rejon = st.selectbox("Rejon Polski (auto z kodu):", all_region_names, index=idx_r)

                nc_save, nc_cancel = st.columns([1, 4])
                with nc_save:
                    nc_submitted = st.form_submit_button("💾 ZAPISZ KLIENTA W BAZIE", type="primary")
                with nc_cancel:
                    nc_cancelled = st.form_submit_button("Anuluj")

                if nc_submitted:
                    if not nc_nazwisko or not nc_tel:
                        st.error("❌ Nazwisko i numer telefonu są wymagane!")
                    else:
                        new_client_row = {
                            "nr.": f"K{get_next_ticket_id()}",
                            "INSTALACJA": "Brak (nowy klient)",
                            "DATA ZGŁOSZENIA": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
                            "PRZYJMUJĄCY": "Koordynator",
                            "Imię": nc_imie.strip(),
                            "Nazwisko": nc_nazwisko.strip(),
                            "Numer telefonu": nc_tel.strip(),
                            "E-mail": nc_email.strip(),
                            "Kod pocztowy": nc_kod.strip(),
                            "Miejscowość": nc_miasto.strip(),
                            "Adres": nc_adres.strip(),
                            "Rejon Polski": nc_rejon,
                            "STATUS ZGŁOSZENIA": "Nowy klient",
                            "SERWISANT": "",
                            "OPIS OD KLIENTA": "Nowy klient dodany ręcznie przez koordynatora.",
                            "OPIS": "",
                            "INFORMACJE DODATKOWE": "",
                            "Kwota": "",
                            "Płatność ": "",
                            "Zużyte materiały": "",
                            "Czynności dodatkowe": "",
                            "Ilość przyjazdów": "0",
                            "Kwota przyjęta": "",
                            "Komentarz serwisanta": "",
                            "Komentarz koordynatora": "",
                            "Zalecenia dla klienta": "",
                            "Przyczyna nieukończenia": "",
                            "Brakujące materiały": "",
                            "Szczegóły oferty": "",
                            "Tabliczki_OK": False,
                            "Wymaga oferty": False,
                            "Email handlowca": "",
                            "Etap oferty": "",
                            "Data oferty": "",
                            "Kolejnosc_trasy": "",
                            "Data_trasy": "",
                            "Trasa_aktywna": "False",
                            "Odbior_klienta": "",
                            "Zdalne": "False",
                            "Termin_koncowy": "",
                            "Pozycje_cennika": "",
                            "Cena_bazowa": "",
                            "Rabat_proc": "0",
                            "Klient_ID": get_clean_client_id({'Numer telefonu': nc_tel, 'Imię': nc_imie, 'Nazwisko': nc_nazwisko})
                        }
                        st.session_state.df = pd.concat([pd.DataFrame([new_client_row]), st.session_state.df], ignore_index=True)
                        save_ticket_persistently(new_client_row, current_full_df=st.session_state.df)
                        st.session_state.selected_client_id = new_client_row['Klient_ID']
                        st.session_state.show_new_client_modal = False
                        st.success(f"✅ Klient {nc_imie} {nc_nazwisko} dodany do bazy!")
                        st.rerun()

                if nc_cancelled:
                    st.session_state.show_new_client_modal = False
                    st.rerun()

        col_s1, col_s2 = st.columns([3, 1])
        with col_s1:
            search_query = st.text_input(
                "Wyszukaj w bazie (nazwisko, imię, numer telefonu, miejscowość lub adres):",
                placeholder="np. Kowalski, 602 123 456, Radom, Topolowa...",
            ).strip().lower()
        with col_s2:
            if st.button("🔄 Resetuj wybór", use_container_width=True):
                st.session_state.selected_client_id = None
                st.session_state.show_edit_client_modal = False
                st.rerun()

        if search_query:
            query_mask = (
                filtered_df['Imię'].str.lower().str.contains(search_query, na=False) |
                filtered_df['Nazwisko'].str.lower().str.contains(search_query, na=False) |
                filtered_df['Numer telefonu'].str.lower().str.contains(search_query, na=False) |
                filtered_df['Miejscowość'].str.lower().str.contains(search_query, na=False) |
                filtered_df['Adres'].str.lower().str.contains(search_query, na=False)
            )
            filtered_df = filtered_df[query_mask]

        clients_unique = filtered_df.drop_duplicates(subset=['Klient_ID'])
        st.write(f"Znaleziono pasujących klientów: **{len(clients_unique)}**")

        if len(clients_unique) > 0:
            client_choices = []
            for _, r in clients_unique.iterrows():
                name = f"{r.get('Imię', '')} {r.get('Nazwisko', '')}".strip() or "Brak nazwiska"
                tel = r.get('Numer telefonu', '')
                miasto = r.get('Miejscowość', '')
                ulica = r.get('Adres', '')
                cid = r['Klient_ID']
                lbl = f"{name} | 📞 {tel} | 📍 {miasto}, {ulica}"
                client_choices.append({'ID': cid, 'Etykieta': lbl})

            current_choice_idx = 0
            if st.session_state.selected_client_id:
                for i, c in enumerate(client_choices):
                    if c['ID'] == st.session_state.selected_client_id:
                        current_choice_idx = i
                        break

            choice_labels = [c['Etykieta'] for c in client_choices]
            selected_label = st.selectbox("Wybierz klienta z listy:", options=choice_labels, index=current_choice_idx)
            for c in client_choices:
                if c['Etykieta'] == selected_label:
                    st.session_state.selected_client_id = c['ID']
                    break

        # --- KARTA WYBRANEGO KLIENTA ---
        if st.session_state.selected_client_id:
            client_records = st.session_state.df[st.session_state.df['Klient_ID'] == st.session_state.selected_client_id]
            
            if not client_records.empty:
                first_record = client_records.iloc[0]

                st.divider()
                st.subheader(f"👤 Karta Klienta: {first_record.get('Imię', '')} {first_record.get('Nazwisko', '')}")
                
                col_k1, col_k2, col_k3 = st.columns(3)
                with col_k1:
                    st.markdown(f"**📞 Telefon:** {first_record.get('Numer telefonu', 'Brak')}")
                    st.markdown(f"**✉️ E-mail:** {first_record.get('E-mail', 'Brak')}")
                with col_k2:
                    st.markdown(f"**📍 Adres:** {first_record.get('Adres', '')}")
                    st.markdown(f"**📮 Kod i miasto:** {first_record.get('Kod pocztowy', '')} {first_record.get('Miejscowość', '')}")
                with col_k3:
                    st.markdown(f"**🌐 Rejon:** {first_record.get('Rejon Polski', '')}")
                    st.markdown(f"**Liczba zgłoszeń:** **{len(client_records)}**")

                btn_c1, btn_c2, _ = st.columns([1.5, 1.5, 3])
                with btn_c1:
                    if st.button("➕ NOWE ZGŁOSZENIE", type="primary", use_container_width=True):
                        st.session_state.show_new_ticket_modal = True
                        st.session_state.show_edit_client_modal = False
                with btn_c2:
                    if st.button("✏️ EDYTUJ DANE KLIENTA", use_container_width=True):
                        st.session_state.show_edit_client_modal = not st.session_state.show_edit_client_modal
                        st.session_state.show_new_ticket_modal = False

                # FORMULARZ NOWEGO ZGŁOSZENIA Z CENNIKIEM, RABATEM, ZDALNYM I SMS-EM
                if st.session_state.show_new_ticket_modal:
                    st.info("📝 Rejestracja nowego zgłoszenia (z automatyczną wyceną i generatorem SMS dla klienta):")
                    
                    next_id = get_next_ticket_id()
                    new_ticket_nr_preview = f"N{next_id}"

                    # Kalkulator wyceny i rabatu na żywo przed formularzem
                    st.markdown("##### 💰 1. Kalkulator wyceny z cennika i rabat")
                    price_options = [f"{r['Nazwa']} ({float(r['Cena']):.0f} zł)" for _, r in pricelist_df.iterrows()]
                    price_lookup = {f"{r['Nazwa']} ({float(r['Cena']):.0f} zł)": float(r['Cena']) for _, r in pricelist_df.iterrows()}

                    col_p1, col_p2, col_p3 = st.columns([2.5, 1, 1])
                    with col_p1:
                        wybrane_pozycje_cennika = st.multiselect(
                            "Wybierz usługi / części z cennika (automatycznie sumuje kwotę):",
                            options=price_options,
                            key="new_ticket_pricelist_select"
                        )
                    suma_bazowa = sum([price_lookup.get(item, 0.0) for item in wybrane_pozycje_cennika])
                    with col_p2:
                        reczna_korekta = st.number_input("Lub wpisz kwotę bazową ręcznie (zł):", min_value=0.0, value=float(suma_bazowa), step=50.0)
                    with col_p3:
                        rabat_procent = st.number_input("Rabat dla klienta (%):", min_value=0, max_value=100, value=0, step=5)

                    kwota_po_rabacie = round(reczna_korekta * (1.0 - rabat_procent / 100.0), 2)
                    st.success(f"💵 **Podsumowanie wyceny:** Cena bazowa: **{reczna_korekta:.0f} zł** | Rabat: **{rabat_procent}%** ➔ **Kwota końcowa dla klienta: {kwota_po_rabacie:.0f} zł**")

                    st.markdown("##### 📋 2. Szczegóły zlecenia i termin końcowy")
                    f_col1, f_col2 = st.columns(2)
                    with f_col1:
                        st.text_input("Numer zgłoszenia (auto):", value=f"#{new_ticket_nr_preview}", disabled=True)
                        forma = st.radio("Forma zgłoszenia:", ["telefon", "SMS", "WhatsApp", "e-mail"], horizontal=True)
                        priorytet = st.selectbox("Priorytet:", ["Standardowy", "Pilny", "Awaria", "Niski"], index=0)
                        czy_zdalne = st.checkbox("📞 Zgłoszenie do rozwiązania ZDALNIE (wisi w panelu koordynatora)")
                    with f_col2:
                        usterka_dotyczy = st.selectbox("Usterka dotyczy:", ["solar", "foto", "pompa", "magazyn"], index=0)
                        producent = st.text_input("Producent instalacji:", placeholder="np. Hewalex, Deye, Fronius, Sofar...")
                        default_deadline = datetime.date.today() + datetime.timedelta(days=7)
                        termin_koncowy = st.date_input("📅 Ostateczny termin realizacji (wpisywany do SMS):", value=default_deadline)

                    domyslny_zakres = ", ".join([p.split(" (")[0] for p in wybrane_pozycje_cennika])
                    opis_zgloszenia = st.text_area(
                        "Zakres prac / Opis dla serwisanta (trafia również do SMS dla klienta):*",
                        value=domyslny_zakres,
                        placeholder="Wpisz zakres prac lub opis usterki..."
                    )

                    komentarz_koordynatora = st.text_area(
                        "🔒 Wewnętrzny komentarz koordynatora (POUFNE — niewidoczne dla serwisanta i klienta):",
                        placeholder="Uwagi wewnętrzne biura..."
                    )

                    uploaded_media = st.file_uploader(
                        "📷 Załącz zdjęcia lub filmiki od klienta:",
                        type=['png', 'jpg', 'jpeg', 'webp', 'mp4', 'mov', 'avi'],
                        accept_multiple_files=True,
                        key="uploader_new_ticket"
                    )

                    wymaga_oferty = st.checkbox("💼 Zgłoszenie wymaga oferty handlowej (przekaż do handlowca)")
                    email_handlowca = ""
                    if wymaga_oferty:
                        email_handlowca = st.text_input("Wpisz adres e-mail handlowca:", placeholder="np. handlowiec@twojafirma.pl")

                    c_save, c_cancel = st.columns([2, 1])
                    with c_save:
                        if st.button("💾 ZAPISZ ZLECENIE I PRZYGOTUJ SMS DLA KLIENTA", type="primary", use_container_width=True):
                            now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
                            new_ticket_nr = f"N{next_id}"
                            etap_startowy = "Czeka na ofertę handlowca" if wymaga_oferty else ""
                            if czy_zdalne:
                                status_startowy = "Zdalne - do realizacji"
                            elif wymaga_oferty:
                                status_startowy = "Czeka na ofertę handlowca"
                            else:
                                status_startowy = "w toku"

                            if uploaded_media:
                                target_dir = os.path.join(ATTACHMENTS_DIR, str(new_ticket_nr))
                                os.makedirs(target_dir, exist_ok=True)
                                for file in uploaded_media:
                                    with open(os.path.join(target_dir, file.name), "wb") as f_out:
                                        f_out.write(file.getbuffer())

                            kwota_str = f"{kwota_po_rabacie:.0f}" if kwota_po_rabacie > 0 else "0"
                            termin_str = termin_koncowy.strftime("%Y-%m-%d")

                            new_row_dict = {
                                "nr.": new_ticket_nr,
                                "INSTALACJA": usterka_dotyczy,
                                "DATA ZGŁOSZENIA": now_str,
                                "PRZYJMUJĄCY": "Koordynator",
                                "Imię": first_record.get('Imię', ''),
                                "Nazwisko": first_record.get('Nazwisko', ''),
                                "Numer telefonu": first_record.get('Numer telefonu', ''),
                                "E-mail": first_record.get('E-mail', ''),
                                "Kod pocztowy": first_record.get('Kod pocztowy', ''),
                                "Miejscowość": first_record.get('Miejscowość', ''),
                                "Adres": first_record.get('Adres', ''),
                                "Rejon Polski": first_record.get('Rejon Polski', ''),
                                "STATUS ZGŁOSZENIA": status_startowy,
                                "SERWISANT": "Zdalnie (Koordynator)" if czy_zdalne else "",
                                "OPIS OD KLIENTA": opis_zgloszenia.strip(),
                                "OPIS": "",
                                "INFORMACJE DODATKOWE": f"Forma: {forma} | Priorytet: {priorytet} | Producent: {producent}",
                                "Kwota": kwota_str,
                                "Płatność ": "",
                                "Zużyte materiały": "",
                                "Czynności dodatkowe": "",
                                "Ilość przyjazdów": "1",
                                "Kwota przyjęta": "",
                                "Komentarz serwisanta": "",
                                "Komentarz koordynatora": komentarz_koordynatora.strip(),
                                "Zalecenia dla klienta": "",
                                "Przyczyna nieukończenia": "",
                                "Brakujące materiały": "",
                                "Szczegóły oferty": "",
                                "Tabliczki_OK": first_record.get('Tabliczki_OK', False),
                                "Wymaga oferty": bool(wymaga_oferty),
                                "Email handlowca": email_handlowca,
                                "Etap oferty": etap_startowy,
                                "Data oferty": now_str[:10],
                                "Kolejnosc_trasy": "",
                                "Data_trasy": "",
                                "Trasa_aktywna": "False",
                                "Odbior_klienta": "",
                                "Zdalne": "True" if czy_zdalne else "False",
                                "Termin_koncowy": termin_str,
                                "Pozycje_cennika": "; ".join(wybrane_pozycje_cennika),
                                "Cena_bazowa": f"{reczna_korekta:.0f}",
                                "Rabat_proc": str(rabat_procent),
                                "Klient_ID": first_record['Klient_ID']
                            }

                            st.session_state.df = pd.concat([pd.DataFrame([new_row_dict]), st.session_state.df], ignore_index=True)
                            save_ticket_persistently(new_row_dict, current_full_df=st.session_state.df)

                            # Przygotowanie powiadomienia SMS dla klienta
                            raw_tel = str(first_record.get('Numer telefonu', '')).strip()
                            clean_tel = "".join([c for c in raw_tel.split(',')[0].split('/')[0] if c.isdigit()])
                            rabat_txt = f" (po rabacie {rabat_procent}%)" if rabat_procent > 0 else ""
                            sms_text = (
                                f"PROMAG SERWIS: Przyjelismy zlecenie nr #{new_ticket_nr}. "
                                f"Zakres prac: {opis_zgloszenia.strip() or usterka_dotyczy}. "
                                f"Szacowana wycena: {kwota_str} zl{rabat_txt}. "
                                f"Przewidywany termin realizacji do: {termin_str}. "
                                f"Informujemy, ze jest to termin koncowy, a serwisant bedzie kontaktowal sie z Panstwem rano w dniu wizyty."
                            )
                            st.session_state.coord_sms_notification = {
                                "nr": new_ticket_nr,
                                "klient": f"{first_record.get('Imię', '')} {first_record.get('Nazwisko', '')}",
                                "tel_clean": clean_tel,
                                "tekst": sms_text
                            }

                            st.session_state.show_new_ticket_modal = False
                            st.rerun()

                    with c_cancel:
                        if st.button("Anuluj", use_container_width=True):
                            st.session_state.show_new_ticket_modal = False
                            st.rerun()

                st.markdown("### 📋 Historia i status zgłoszeń tego klienta:")
                for i, (_, row) in enumerate(client_records.iterrows(), 1):
                    nr_t = str(row.get('nr.', 'Brak nr'))
                    tryb_lbl = "📞 ZDALNE" if str(row.get('Zdalne', '')).lower() == 'true' else "🚗 TEREN"
                    with st.expander(f"Zgłoszenie #{nr_t} [{tryb_lbl}] | Typ: {row.get('INSTALACJA', '')} | Termin do: {row.get('Termin_koncowy', '—')} | Status: {row.get('STATUS ZGŁOSZENIA', '')}", expanded=(i==1)):
                        c1, c2 = st.columns(2)
                        with c1:
                            st.markdown(f"**Data zgłoszenia:** {row.get('DATA ZGŁOSZENIA', '')}")
                            st.markdown(f"**Termin końcowy:** {row.get('Termin_koncowy', 'Brak')}")
                            st.markdown(f"**Serwisant:** {row.get('SERWISANT', 'Nieprzypisany')}")
                        with c2:
                            st.markdown(f"**Wycena końcowa:** **{row.get('Kwota', '0')} zł** (Bazowa: {row.get('Cena_bazowa', '—')} zł, Rabat: {row.get('Rabat_proc', '0')}%)")
                            st.markdown(f"**Pozycje cennika:** {row.get('Pozycje_cennika', 'Brak')}")
                            st.markdown(f"**Status:** {row.get('STATUS ZGŁOSZENIA', '')}")

                        st.markdown(f"**Zakres prac / Opis dla serwisanta:** {row.get('OPIS OD KLIENTA', 'Brak')}")
                        if row.get('Komentarz koordynatora', ''):
                            st.info(f"🔒 **Poufna notatka koordynatora:** {row.get('Komentarz koordynatora')}")
                        if row.get('Komentarz serwisanta', ''):
                            st.write(f"🛠️ **Raport serwisanta:** {row.get('Komentarz serwisanta')}")
                        if row.get('Zużyte materiały', ''):
                            st.write(f"📦 **Zużyte materiały:** {row.get('Zużyte materiały')}")
                        if row.get('Zalecenia dla klienta', ''):
                            st.write(f"💡 **Zalecenia dla klienta:** {row.get('Zalecenia dla klienta')}")

                        wyswietl_multimedia(nr_t)

# =========================================================================
# 2. MODUŁ SERWISANTA (POJEDYNCZE EKRANY - WIZARD UX POD JEDNĄ DŁOŃ)
# =========================================================================
elif st.session_state.user_role == "serwisant":
    df = st.session_state.df

    if not st.session_state.tech_pin:
        if st.button("⬅️ Wróć do wyboru modułu", use_container_width=True):
            st.session_state.user_role = None
            st.rerun()

        st.markdown("<h3 style='text-align: center;'>🔒 Logowanie Serwisanta</h3>", unsafe_allow_html=True)
        with st.form("form_pin"):
            pin_input = st.text_input("Wpisz 4-cyfrowy kod PIN:", type="password", max_chars=4, placeholder="np. 1110")
            submit_pin = st.form_submit_button("ZALOGUJ SIĘ", type="primary", use_container_width=True)

            if submit_pin:
                if pin_input in PIN_CONFIG:
                    st.session_state.tech_pin = pin_input
                    st.session_state.tech_main_screen = "otwarte"
                    st.rerun()
                else:
                    st.error("❌ Błędny kod PIN.")

        with st.expander("ℹ️ Ściągawka kodów PIN"):
            st.write("• **1110** - Rejon 0 (Leszek) | **1111** - Rejon 2 (Stanisław) | **1112** - Rejon 6 (Wojtek)")
            st.write("• **1113** - Rejon 8 (Mikołaj) | **1114** - Rejon 9 (Artur) | **9999** - Wszystkie rejony")
        st.stop()

    cfg = PIN_CONFIG[st.session_state.tech_pin]
    rejon_kod = cfg["rejon_kod"]
    rejon_nazwa = cfg["nazwa_rejonu"]
    domyslny_tech = cfg["domyslny_serwisant"]

    # Zwięzły pasek nagłówka na telefonie
    col_h1, col_h2 = st.columns([3, 1])
    with col_h1:
        st.markdown(f"**👷 {domyslny_tech}** | *{rejon_nazwa}*")
    with col_h2:
        if st.button("Wyloguj", use_container_width=True):
            st.session_state.tech_pin = None
            st.session_state.active_ticket_nr = None
            st.session_state.day_started = False
            st.rerun()

    # -------------------------------------------------------------------------
    # EKRANOWY KREATOR ROZLICZENIA ZLECENIA (KROK 1 -> KROK 2 -> KROK 3)
    # -------------------------------------------------------------------------
    if st.session_state.active_ticket_nr:
        curr_rows = df[df['nr.'].astype(str) == str(st.session_state.active_ticket_nr)]
        if curr_rows.empty:
            st.session_state.active_ticket_nr = None
            st.rerun()

        ticket = curr_rows.iloc[0]
        nr_zlec = str(ticket['nr.'])
        typ_inst = str(ticket.get('INSTALACJA', 'solar')).lower().strip()
        if typ_inst not in QUICK_ACTIONS:
            typ_inst = "solar"

        step = st.session_state.wizard_step
        w_data = st.session_state.wizard_data

        st.markdown(f"""<div class="wizard-step-box">🛠️ Zlecenie #{nr_zlec} ({ticket.get('Imię','')} {ticket.get('Nazwisko','')}) — KROK {step} z 3</div>""", unsafe_allow_html=True)
        st.progress(step / 3.0)

        # ---------------- KROK 1: WYNIK PRAC, CZYNNOŚCI I CZĘŚCI ----------------
        if step == 1:
            st.subheader("Krok 1: Co zostało zrobione?")
            outcome = st.radio(
                "Wynik wizyty serwisowej:",
                ["TAK — Usterka usunięta", "NIE — Wymaga ponownego przyjazdu / części"],
                index=0 if w_data.get("is_success", True) else 1
            )
            is_succ = "TAK" in outcome

            if is_succ:
                q_act = st.multiselect("⚡ Szybki wybór wykonanych czynności:", options=QUICK_ACTIONS.get(typ_inst, QUICK_ACTIONS["solar"]), default=w_data.get("q_act", []))
                q_prt = st.multiselect("🔩 Szybki wybór zużytych części:", options=QUICK_PARTS.get(typ_inst, QUICK_PARTS["solar"]), default=w_data.get("q_prt", []))
                default_opis = ", ".join(q_act) if q_act else w_data.get("opis_prac", ticket.get('Komentarz serwisanta', ''))
                default_czesci = ", ".join(q_prt) if q_prt else w_data.get("zuzyte_czesci", ticket.get('Zużyte materiały', ''))
                opis_prac = st.text_area("Opis wykonanych prac:*", value=default_opis)
                zuzyte_czesci = st.text_area("Zużyte części / materiały:", value=default_czesci)
                przyczyna_nieuk = ""
                brakujace_mat = ""
            else:
                q_act = []
                q_prt = st.multiselect("🔩 Szybki wybór brakujących części:", options=QUICK_PARTS.get(typ_inst, QUICK_PARTS["solar"]), default=w_data.get("q_prt", []))
                default_braki = ", ".join(q_prt) if q_prt else w_data.get("brakujace_mat", ticket.get('Brakujące materiały', ''))
                przyczyna_nieuk = st.text_area("Dlaczego nie udało się ukończyć zlecenia:*", value=w_data.get("przyczyna_nieuk", ticket.get('Przyczyna nieukończenia', '')))
                brakujace_mat = st.text_area("Jakich części brakuje do zamówienia:*", value=default_braki)
                opis_prac = ""
                zuzyte_czesci = ""

            q_rec = st.multiselect("💡 Szybki wybór zaleceń dla klienta:", options=QUICK_RECOMMENDATIONS, default=w_data.get("q_rec", []))
            default_zal = ". ".join(q_rec) if q_rec else w_data.get("zalecenia", ticket.get('Zalecenia dla klienta', ''))
            zalecenia = st.text_area("Zalecenia dla klienta (OBOWIĄZKOWE):*", value=default_zal)

            col_nav1, col_nav2 = st.columns(2)
            with col_nav1:
                if st.button("❌ Anuluj i wróć do listy", use_container_width=True):
                    st.session_state.active_ticket_nr = None
                    st.session_state.wizard_step = 1
                    st.session_state.wizard_data = {}
                    st.rerun()
            with col_nav2:
                if st.button("DALEJ: Zdjęcia i Oferta ➡️", type="primary", use_container_width=True):
                    if not zalecenia.strip():
                        st.error("❌ Wpisz lub wybierz zalecenia dla klienta!")
                    elif is_succ and not opis_prac.strip():
                        st.error("❌ Opisz krótko, co zostało zrobione!")
                    elif not is_succ and (not przyczyna_nieuk.strip() or not brakujace_mat.strip()):
                        st.error("❌ Opisz przyczynę oraz brakujące części!")
                    else:
                        w_data.update({
                            "is_success": is_succ,
                            "q_act": q_act,
                            "q_prt": q_prt,
                            "q_rec": q_rec,
                            "opis_prac": opis_prac.strip(),
                            "zuzyte_czesci": zuzyte_czesci.strip(),
                            "przyczyna_nieuk": przyczyna_nieuk.strip(),
                            "brakujace_mat": brakujace_mat.strip(),
                            "zalecenia": zalecenia.strip()
                        })
                        st.session_state.wizard_data = w_data
                        st.session_state.wizard_step = 2
                        st.rerun()

        # ---------------- KROK 2: ZDJĘCIA, TABLICZKI I OFERTA HANDLOWA ----------------
        elif step == 2:
            st.subheader("Krok 2: Zdjęcia z obiektu i Oferta")
            klient_id_curr = ticket.get('Klient_ID', '')
            klient_ma_juz_tabliczki = (
                (df['Klient_ID'] == klient_id_curr) & 
                (df['Tabliczki_OK'].astype(str).str.lower().isin(['true', '1', 'tak']))
            ).any()

            if klient_ma_juz_tabliczki:
                st.success("✅ Tabliczki znamionowe tego klienta są już w bazie.")
                potwierdzenie_tabliczek = True
            else:
                st.warning("📌 Brak tabliczek znamionowych w bazie! Zrób im zdjęcie poniżej.")
                potwierdzenie_tabliczek = st.checkbox(
                    "✅ POTWIERDZAM: Zrobiłem zdjęcia tabliczek znamionowych",
                    value=w_data.get("potwierdzenie_tabliczek", False)
                )

            tech_media = st.file_uploader("Wgraj zdjęcia z galerii telefonu:", type=['png', 'jpg', 'jpeg', 'webp', 'mp4'], accept_multiple_files=True)
            camera_photo = st.camera_input("Lub zrób zdjęcie aparatem teraz:")

            # Zapis zdjęć od razu do folderu zlecenia
            target_dir = os.path.join(ATTACHMENTS_DIR, nr_zlec)
            os.makedirs(target_dir, exist_ok=True)
            if tech_media:
                for file in tech_media:
                    with open(os.path.join(target_dir, file.name), "wb") as f_out:
                        f_out.write(file.getbuffer())
                st.toast("📸 Zapisano zdjęcia z galerii!")
            if camera_photo:
                c_name = f"aparat_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.jpg"
                with open(os.path.join(target_dir, c_name), "wb") as f_c:
                    f_c.write(camera_photo.getbuffer())
                st.toast("📸 Zapisano zdjęcie z aparatu!")

            st.markdown("---")
            potrzebna_oferta = st.checkbox(
                "💼 Klient prosi o OFERTĘ HANDLOWĄ (rozbudowa / modyfikacja instalacji)",
                value=w_data.get("potrzebna_oferta", False)
            )
            szczegoly_oferty = ""
            if potrzebna_oferta:
                szczegoly_oferty = st.text_area(
                    "Co dokładnie biuro ma wycenić dla klienta:*",
                    value=w_data.get("szczegoly_oferty", ""),
                    placeholder="np. Dobór i montaż dodatkowego bufora 200L..."
                )

            col_nav1, col_nav2 = st.columns(2)
            with col_nav1:
                if st.button("⬅️ WSTECZ (Krok 1)", use_container_width=True):
                    st.session_state.wizard_step = 1
                    st.rerun()
            with col_nav2:
                if st.button("DALEJ: Rozliczenie i Podpis ➡️", type="primary", use_container_width=True):
                    if not potwierdzenie_tabliczek:
                        st.error("❌ Potwierdź wykonanie zdjęć tabliczek znamionowych!")
                    elif potrzebna_oferta and not szczegoly_oferty.strip():
                        st.error("❌ Opisz, co biuro ma wycenić dla klienta!")
                    else:
                        w_data.update({
                            "potwierdzenie_tabliczek": potwierdzenie_tabliczek,
                            "potrzebna_oferta": potrzebna_oferta,
                            "szczegoly_oferty": szczegoly_oferty.strip()
                        })
                        st.session_state.wizard_data = w_data
                        st.session_state.wizard_step = 3
                        st.rerun()

        # ---------------- KROK 3: ROZLICZENIE FINANSOWE I PODPIS KLIENTA ----------------
        elif step == 3:
            st.subheader("Krok 3: Kwota i Podpis Klienta")
            wycena_biura = ticket.get('Kwota', '0') or '0'
            rabat_biura = ticket.get('Rabat_proc', '0') or '0'
            st.info(f"💰 Wycena ustalona przez koordynatora: **{wycena_biura} zł** (uwzględniony rabat: **{rabat_biura}%**)")

            kwota_pobrana = st.text_input("Kwota pobrana od klienta na miejscu (zł):", value=str(wycena_biura if w_data.get("is_success", True) else "0"))
            forma_plat = st.selectbox("Forma płatności:", ["Gotówka", "Przelew", "Gwarancja (0 zł)", "Czeka na przelew"])
            przyjazdy = st.selectbox("Ilość przyjazdów:", ["1", "2", "3", "Więcej"])
            
            status_opcje = ["zakończone", "do ponownego przyjazdu (brak części)", "w toku (wymaga części)", "anulowane"]
            dom_stat_idx = 0 if w_data.get("is_success", True) else 1
            status_nowy = st.selectbox("Końcowy status zlecenia:", status_opcje, index=dom_stat_idx)

            st.markdown("---")
            st.markdown("##### ✍️ Podpis klienta na ekranie telefonu (potwierdzenie odbioru)")
            if CANVAS_AVAILABLE:
                st.caption("Przekaż telefon klientowi, aby podpisał się palcem w poniższej ramce:")
                canvas_result = st_canvas(
                    fill_color="rgba(255, 255, 255, 1)",
                    stroke_width=3,
                    stroke_color="#1a202c",
                    background_color="#ffffff",
                    height=180,
                    drawing_mode="freedraw",
                    key=f"sig_canvas_{nr_zlec}"
                )
            else:
                canvas_result = None
                st.caption("Moduł rysowania niedostępny – użyj potwierdzenia poniżej:")

            odbior_klienta = st.checkbox("✅ Klient obecny na miejscu — potwierdził odbiór prac i zalecenia", value=True)

            col_nav1, col_nav2 = st.columns(2)
            with col_nav1:
                if st.button("⬅️ WSTECZ (Krok 2)", use_container_width=True):
                    st.session_state.wizard_step = 2
                    st.rerun()
            with col_nav2:
                if st.button("💾 ZAKOŃCZ I ZAPISZ ZLECENIE", type="primary", use_container_width=True):
                    idx = curr_rows.index[0]
                    is_succ = w_data.get("is_success", True)

                    if is_succ:
                        df.at[idx, 'Komentarz serwisanta'] = w_data.get("opis_prac", "")
                        df.at[idx, 'Zużyte materiały'] = w_data.get("zuzyte_czesci", "")
                        df.at[idx, 'Przyczyna nieukończenia'] = ""
                        df.at[idx, 'Brakujące materiały'] = ""
                    else:
                        df.at[idx, 'Komentarz serwisanta'] = w_data.get("przyczyna_nieuk", "")
                        df.at[idx, 'Przyczyna nieukończenia'] = w_data.get("przyczyna_nieuk", "")
                        df.at[idx, 'Brakujące materiały'] = w_data.get("brakujace_mat", "")
                        df.at[idx, 'Zużyte materiały'] = ""

                    df.at[idx, 'Zalecenia dla klienta'] = w_data.get("zalecenia", "")

                    klient_id_curr = ticket.get('Klient_ID', '')
                    if w_data.get("potwierdzenie_tabliczek", False):
                        df.loc[df['Klient_ID'] == klient_id_curr, 'Tabliczki_OK'] = 'True'
                        df.at[idx, 'Tabliczki_OK'] = 'True'

                    # Zapis podpisu z ekranu jako obraz PNG w teczce zlecenia
                    if CANVAS_AVAILABLE and canvas_result is not None and canvas_result.image_data is not None:
                        try:
                            target_dir = os.path.join(ATTACHMENTS_DIR, nr_zlec)
                            os.makedirs(target_dir, exist_ok=True)
                            sig_img = Image.fromarray(canvas_result.image_data.astype('uint8'), 'RGBA')
                            sig_path = os.path.join(target_dir, f"podpis_klienta_{nr_zlec}.png")
                            sig_img.save(sig_path)
                        except Exception as e:
                            print(f"Błąd zapisu podpisu: {e}")

                    if w_data.get("potrzebna_oferta", False):
                        df.at[idx, 'Wymaga oferty'] = True
                        df.at[idx, 'Etap oferty'] = '🚨 PILNE: Zgłoszenie z serwisu (wymaga wyceny modyfikacji)'
                        df.at[idx, 'Szczegóły oferty'] = w_data.get("szczegoly_oferty", "")
                        df.at[idx, 'Data oferty'] = datetime.date.today().strftime("%Y-%m-%d")
                    else:
                        df.at[idx, 'Wymaga oferty'] = False

                    df.at[idx, 'Ilość przyjazdów'] = przyjazdy
                    df.at[idx, 'Kwota przyjęta'] = kwota_pobrana
                    df.at[idx, 'Płatność '] = forma_plat
                    df.at[idx, 'STATUS ZGŁOSZENIA'] = status_nowy
                    df.at[idx, 'DATA SERWISU'] = datetime.date.today().strftime("%Y-%m-%d")
                    df.at[idx, 'SERWISANT'] = domyslny_tech
                    df.at[idx, 'Odbior_klienta'] = "Podpisano na ekranie" if (CANVAS_AVAILABLE and odbior_klienta) else ("Potwierdzony" if odbior_klienta else "Brak")

                    st.session_state.df = df
                    save_ticket_persistently(df.loc[idx].to_dict(), current_full_df=st.session_state.df)

                    tel_raw = str(ticket.get('Numer telefonu', '')).strip()
                    tel_digits_only = "".join([c for c in tel_raw.split(',')[0].split('/')[0] if c.isdigit()])
                    st.session_state.last_closed_summary = {
                        "nr": nr_zlec,
                        "tel": tel_digits_only,
                        "zalecenia": w_data.get("zalecenia", ""),
                        "kwota": kwota_pobrana or "0"
                    }

                    st.session_state.active_ticket_nr = None
                    st.session_state.wizard_step = 1
                    st.session_state.wizard_data = {}
                    st.rerun()

        st.stop()

    # Powiadomienie po zamknięciu zlecenia z szybkim SMS-em podsumowującym
    if st.session_state.last_closed_summary:
        summ = st.session_state.last_closed_summary
        st.success(f"✅ Zlecenie #{summ['nr']} zostało rozliczone i zapisane!")
        if summ.get('tel'):
            sms_podsumowanie = urllib.parse.quote(
                f"PROMAG SERWIS: Potwierdzamy wykonanie serwisu (Zlecenie #{summ['nr']}). "
                f"Zalecenia: {summ['zalecenia']}. Kwota: {summ['kwota']} zl."
            )
            st.markdown(
                f'<a href="sms:{summ["tel"]}?body={sms_podsumowanie}" target="_blank" style="text-decoration:none;">'
                f'<button style="width:100%; height:48px; border-radius:10px; background:#2b6cb0; color:white; border:none; font-weight:bold; margin-bottom:8px; cursor:pointer;">'
                f'📲 Wyślij SMS z potwierdzeniem do klienta</button></a>',
                unsafe_allow_html=True
            )
        if st.button("OK, zamknij powiadomienie", use_container_width=True):
            st.session_state.last_closed_summary = None
            st.rerun()

    # Filtrowanie zleceń dla rejonu serwisanta (bez zleceń zdalnych!)
    if rejon_kod == "ALL":
        tech_df = df.copy()
    else:
        mask_rejon = (
            df['Rejon Polski'].str.contains(f"Rejon {rejon_kod}", case=False, na=False) |
            df['Kod pocztowy'].str.startswith(f"{rejon_kod}")
        )
        tech_df = df[mask_rejon]

    not_remote_mask = ~tech_df['Zdalne'].astype(str).str.lower().isin(['true', '1', 'tak'])
    ready_for_service_mask = (
        (tech_df['Etap oferty'] != 'Czeka na ofertę handlowca') &
        (tech_df['Etap oferty'] != 'Oczekiwanie na akceptację klienta')
    )
    tech_df = tech_df[not_remote_mask & ready_for_service_mask]

    tech_recent_mask = tech_df['DATA ZGŁOSZENIA'].str[:10] >= CUTOFF_DATE
    tech_operational_df = tech_df[tech_recent_mask]

    is_done = tech_operational_df['STATUS ZGŁOSZENIA'].str.lower().str.contains('zakończ|anulow|nowy klient', na=False)
    otwarte_df = tech_operational_df[~is_done].copy()
    zakonczone_df = tech_operational_df[is_done].copy()

    planned_route_df = otwarte_df[otwarte_df['Trasa_aktywna'].astype(str).str.lower().isin(['true', '1', 'tak'])].copy()
    if not planned_route_df.empty:
        planned_route_df['sort_order'] = pd.to_numeric(planned_route_df['Kolejnosc_trasy'], errors='coerce').fillna(99)
        planned_route_df = planned_route_df.sort_values(by='sort_order')

    # Nawigacja ekranowa (zamiast szerokich zakładek, które rozpychają ekran na boki)
    nav_c1, nav_c2, nav_c3, nav_c4 = st.columns(4)
    with nav_c1:
        if st.button(f"⏳ Zlecenia ({len(otwarte_df)})", type="primary" if st.session_state.tech_main_screen == "otwarte" else "secondary", use_container_width=True):
            st.session_state.tech_main_screen = "otwarte"
            st.rerun()
    with nav_c2:
        if st.button(f"✅ Zrobione ({len(zakonczone_df)})", type="primary" if st.session_state.tech_main_screen == "zakonczone" else "secondary", use_container_width=True):
            st.session_state.tech_main_screen = "zakonczone"
            st.rerun()
    with nav_c3:
        if st.button("🔍 Szukaj", type="primary" if st.session_state.tech_main_screen == "szukaj" else "secondary", use_container_width=True):
            st.session_state.tech_main_screen = "szukaj"
            st.rerun()
    with nav_c4:
        if st.button("📘 Błędy", type="primary" if st.session_state.tech_main_screen == "wiedza" else "secondary", use_container_width=True):
            st.session_state.tech_main_screen = "wiedza"
            st.rerun()

    st.markdown("---")

    # NAPRAWIONA FUNKCJA RENDERUJĄCA KARTĘ (BRAK WCIĘĆ W HTML = ZERO BŁĘDÓW WYŚWIETLANIA KODU)
    def render_ticket_card(row, is_open=True):
        nr = str(row.get('nr.', '')).strip()
        typ = str(row.get('INSTALACJA', '')).lower().strip()
        imie_nazwisko = f"{row.get('Imię', '')} {row.get('Nazwisko', '')}".strip() or "Brak danych"
        tel = str(row.get('Numer telefonu', '')).strip()
        tel_clean = "".join([c for c in tel.split(',')[0].split('/')[0] if c.isdigit()])
        adres_full = f"{row.get('Kod pocztowy', '')} {row.get('Miejscowość', '')}, {row.get('Adres', '')}".strip(', ')
        map_query = urllib.parse.quote(f"{row.get('Adres', '')}, {row.get('Miejscowość', '')}")
        curr_status = str(row.get('STATUS ZGŁOSZENIA', 'w toku')).strip()
        kwota_val = str(row.get('Kwota', '')).strip() or "0"
        termin_val = str(row.get('Termin_koncowy', '')).strip()
        data_zgl = str(row.get('DATA ZGŁOSZENIA', ''))[:10]
        opis_klienta = str(row.get('OPIS OD KLIENTA', '')).strip() or "Brak opisu"

        is_route_item = str(row.get('Trasa_aktywna', '')).lower() in ['true', '1', 'tak']
        route_pos = str(row.get('Kolejnosc_trasy', '')).strip()

        badge_type = "badge-solar" if "solar" in typ else ("badge-foto" if "foto" in typ else ("badge-pompa" if "pompa" in typ else "badge-magazyn"))
        card_class = "mobile-card-route" if (is_route_item and is_open) else "mobile-card"
        route_badge_html = f'<span class="badge badge-route">🎯 TRASA #{route_pos}</span>' if (is_route_item and is_open and route_pos) else ""
        termin_html = f'<span>⏳ Termin do: <b>{termin_val}</b></span>' if termin_val else f'<span>📅 {data_zgl}</span>'

        card_html = (
            f'<div class="{card_class}">'
            f'<div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; margin-bottom:8px;">'
            f'<span style="font-weight:bold; font-size:1.1rem; color:#1a202c;">#{nr}</span>'
            f'<div>{route_badge_html}<span class="badge {badge_type}">{typ.upper() or "SERWIS"}</span>'
            f'<span class="badge" style="background:#e2e8f0; color:#2d3748;">{curr_status.upper()}</span></div>'
            f'</div>'
            f'<div style="font-size:1.05rem; font-weight:700; color:#2d3748;">👤 {imie_nazwisko}</div>'
            f'<div style="font-size:0.95rem; color:#4a5568; margin-top:3px;">📍 {adres_full}</div>'
            f'<div style="margin-top:8px; padding:8px; background:#edf2f7; border-radius:8px; font-size:0.9rem; color:#2d3748;">'
            f'<b>Zakres / Usterka:</b> {opis_klienta}</div>'
            f'<div style="display:flex; justify-content:space-between; margin-top:8px; font-size:0.85rem; color:#718096;">'
            f'<span>💰 Wycena: <b>{kwota_val} zł</b></span>{termin_html}</div>'
            f'</div>'
        )
        st.markdown(card_html, unsafe_allow_html=True)

        cid = row.get('Klient_ID', '')
        prev_visits = df[(df['Klient_ID'] == cid) & (df['nr.'].astype(str) != nr)]
        if not prev_visits.empty:
            with st.expander(f"📂 Historia tego klienta ({len(prev_visits)} wizyt)"):
                for _, pv in prev_visits.iterrows():
                    st.markdown(
                        f"• **#{pv.get('nr.')}** ({str(pv.get('DATA ZGŁOSZENIA', ''))[:10]}) | "
                        f"*{pv.get('SERWISANT', 'Brak')}*: {pv.get('Komentarz serwisanta', '') or pv.get('OPIS OD KLIENTA', '')}"
                    )

        wyswietl_multimedia(nr)

        col_a, col_sms, col_b = st.columns(3)
        with col_a:
            if tel_clean:
                st.markdown(f'<a href="tel:{tel_clean}" target="_blank" style="text-decoration:none;"><button style="width:100%; height:48px; border-radius:10px; background:#2b6cb0; color:white; border:none; font-weight:bold; cursor:pointer;">📞 Zadzwoń</button></a>', unsafe_allow_html=True)
        with col_sms:
            if tel_clean:
                sms_msg = urllib.parse.quote("Dzień dobry, z tej strony serwis PROMAG. Będę u Państwa dzisiaj zgodnie z planem. Zadzwonię bezpośrednio przed dojazdem.")
                st.markdown(f'<a href="sms:{tel_clean}?body={sms_msg}" target="_blank" style="text-decoration:none;"><button style="width:100%; height:48px; border-radius:10px; background:#805ad5; color:white; border:none; font-weight:bold; cursor:pointer;">💬 SMS</button></a>', unsafe_allow_html=True)
        with col_b:
            if row.get('Miejscowość', ''):
                st.markdown(f'<a href="https://www.google.com/maps/search/?api=1&query={map_query}" target="_blank" style="text-decoration:none;"><button style="width:100%; height:48px; border-radius:10px; background:#38a169; color:white; border:none; font-weight:bold; cursor:pointer;">🗺️ Nawiguj</button></a>', unsafe_allow_html=True)

        if is_open:
            btn_col1, btn_col2 = st.columns(2)
            with btn_col1:
                if curr_status.lower() in ['w trakcie realizacji', 'rozpoczęte']:
                    st.button("⚙️ W TRAKCIE PRAC", key=f"btn_started_{nr}", disabled=True, use_container_width=True)
                else:
                    if st.button("▶️ Rozpocznij pracę", key=f"btn_start_{nr}", use_container_width=True):
                        match_idx = df[df['nr.'].astype(str) == nr].index
                        if not match_idx.empty:
                            df.at[match_idx[0], 'STATUS ZGŁOSZENIA'] = 'w trakcie realizacji'
                            df.at[match_idx[0], 'SERWISANT'] = domyslny_tech
                            st.session_state.df = df
                            save_ticket_persistently(df.loc[match_idx[0]].to_dict(), current_full_df=st.session_state.df)
                            st.rerun()
            with btn_col2:
                if st.button("✅ Rozlicz (3 kroki)", key=f"btn_close_{nr}", type="primary", use_container_width=True):
                    st.session_state.active_ticket_nr = nr
                    st.session_state.wizard_step = 1
                    st.session_state.wizard_data = {}
                    st.rerun()
        st.write("")

    # EKRAN 1: OTWARTE ZLECENIA I TRASA
    if st.session_state.tech_main_screen == "otwarte":
        if not planned_route_df.empty:
            route_date_str = planned_route_df.iloc[0].get('Data_trasy', 'Dziś')
            if not st.session_state.day_started:
                st.info(f"📅 **Masz zaplanowaną trasę na dzień: {route_date_str}** ({len(planned_route_df)} pkt).")
                if st.button("🚀 ROZPOCZNIJ DZIEŃ (Pokaż moją trasę)", type="primary", use_container_width=True):
                    st.session_state.day_started = True
                    st.rerun()
            else:
                st.success(f"🚗 **TRASA AKTYWNA ({route_date_str})**")
                route_map_url = build_google_maps_multi_route(planned_route_df)
                if route_map_url:
                    st.markdown(
                        f'<a href="{route_map_url}" target="_blank" style="text-decoration:none;">'
                        f'<button style="width:100%; height:48px; border-radius:10px; background:#2f855a; color:white; border:none; font-weight:bold; font-size:1rem; margin-bottom:8px; cursor:pointer;">'
                        f'🗺️ OTWÓRZ CAŁĄ TRASĘ W GOOGLE MAPS ({len(planned_route_df)} pkt)</button></a>',
                        unsafe_allow_html=True
                    )
                if st.button("Pokaż wszystkie zlecenia w rejonie", use_container_width=True):
                    st.session_state.day_started = False
                    st.rerun()
                st.markdown("---")

        if otwarte_df.empty:
            st.info("🎉 Brak otwartych zleceń w Twoim rejonie!")
        else:
            if st.session_state.day_started and not planned_route_df.empty:
                for _, row in planned_route_df.iterrows():
                    render_ticket_card(row, is_open=True)
            else:
                q = st.text_input("🔍 Szukaj na liście (miejscowość / nazwisko):", key="search_open").strip().lower()
                view_df = otwarte_df.copy()
                if q:
                    view_df = view_df[
                        view_df['Nazwisko'].str.lower().str.contains(q, na=False) |
                        view_df['Miejscowość'].str.lower().str.contains(q, na=False) |
                        view_df['Adres'].str.lower().str.contains(q, na=False)
                    ]
                view_df['is_r'] = view_df['Trasa_aktywna'].astype(str).str.lower().isin(['true', '1', 'tak'])
                view_df['r_ord'] = pd.to_numeric(view_df['Kolejnosc_trasy'], errors='coerce').fillna(999)
                view_df = view_df.sort_values(by=['is_r', 'r_ord'], ascending=[False, True])

                for _, row in view_df.head(30).iterrows():
                    render_ticket_card(row, is_open=True)

    # EKRAN 2: ZAKOŃCZONE ZLECENIA
    elif st.session_state.tech_main_screen == "zakonczone":
        st.subheader("✅ Zakończone zlecenia")
        if zakonczone_df.empty:
            st.info("Brak zakończonych zleceń od 08.10 w Twoim rejonie.")
        else:
            for _, row in zakonczone_df.head(20).iterrows():
                render_ticket_card(row, is_open=False)

    # EKRAN 3: WYSZUKIWARKA HISTORII KLIENTA
    elif st.session_state.tech_main_screen == "szukaj":
        st.subheader("🔍 Sprawdź pełną historię klienta")
        tech_search = st.text_input("Wpisz nazwisko, telefon lub miejscowość:").strip().lower()
        if tech_search:
            t_match = df[
                df['Nazwisko'].str.lower().str.contains(tech_search, na=False) |
                df['Numer telefonu'].str.lower().str.contains(tech_search, na=False) |
                df['Miejscowość'].str.lower().str.contains(tech_search, na=False)
            ]
            t_clients = t_match.drop_duplicates(subset=['Klient_ID'])
            st.write(f"Znaleziono klientów: **{len(t_clients)}**")
            for _, c_row in t_clients.head(10).iterrows():
                c_name = f"{c_row.get('Imię', '')} {c_row.get('Nazwisko', '')}"
                c_id = c_row['Klient_ID']
                c_history = df[df['Klient_ID'] == c_id]
                with st.expander(f"👤 {c_name} | 📞 {c_row.get('Numer telefonu', '')} | 📍 {c_row.get('Miejscowość', '')} ({len(c_history)})"):
                    for _, h_row in c_history.iterrows():
                        hist_html = (
                            f'<div style="background:#edf2f7; padding:10px; border-radius:8px; margin-bottom:8px; font-size:0.9rem;">'
                            f'<b>#{h_row.get("nr.", "")}</b> | Typ: <b>{str(h_row.get("INSTALACJA", "")).upper()}</b> | '
                            f'Data: {str(h_row.get("DATA ZGŁOSZENIA", ""))[:10]} | Status: <b>{h_row.get("STATUS ZGŁOSZENIA", "")}</b><br>'
                            f'<b>Usterka:</b> {h_row.get("OPIS OD KLIENTA", "") or h_row.get("OPIS", "Brak")}<br>'
                            f'<b>Raport:</b> {h_row.get("Komentarz serwisanta", "Brak")} | <b>Części:</b> {h_row.get("Zużyte materiały", "Brak")}'
                            f'</div>'
                        )
                        st.markdown(hist_html, unsafe_allow_html=True)

    # EKRAN 4: BAZA WIEDZY TECHNICZNEJ
    elif st.session_state.tech_main_screen == "wiedza":
        st.subheader("📘 Ściągawka Techniczna & Kody Błędów")
        with st.expander("☀️ INSTALACJE SOLARNE (Hewalex / Geco)", expanded=True):
            st.markdown("""
* **Ciśnienie spoczynkowe (zimny układ):** **1.5 – 2.0 bar**.
* **Ciśnienie w naczyniu przeponowym:** o **0.2 – 0.3 bar niżej** niż robocze.
* **Czujnik PT1000:** w temp. 20°C rezystancja wynosi ok. **1078 Ω** (w 0°C = 1000 Ω).
""")
        with st.expander("❄️ POMPY CIEPŁA — Kody błędów"):
            st.markdown("""
* **E03 / E04 (Brak przepływu):** wyczyścić filtr siatkowy i magnetyczny, odpowietrzyć układ.
* **HP / E01 (Wysokie ciśnienie):** otworzyć pętle grzewcze, odpowietrzyć wymiennik.
* **LP / E02 (Niskie ciśnienie):** sprawdzić czystość parownika i ilość czynnika.
""")
        with st.expander("⚡ FOTOWOLTAIKA & MAGAZYNY (Deye / Sofar / Fronius)"):
            st.markdown("""
* **Deye F18 (AC Overcurrent):** sprawdzić zwarcie na AC/Backup, restart: AC ➔ DC ➔ Bateria.
* **Deye F56 / Sofar ID05 (Niskie napięcie baterii):** wymusić doładowanie z sieci (Grid Charge), sprawdzić kabel CAN (piny 4 i 5).
""")
