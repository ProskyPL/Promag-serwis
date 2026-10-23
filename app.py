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

# Zwarty blok CSS bez pustych linii - eliminuje błąd wyświetlania stylów jako tekstu w Streamlit
st.markdown("""<link rel="manifest" href="manifest.json">
<meta name="theme-color" content="#2b6cb0">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
<meta name="apple-mobile-web-app-capable" content="yes">
<style>
html, body, [data-testid="stAppViewContainer"], .main { overflow-x: hidden !important; max-width: 100vw !important; }
.block-container { padding-top: 1.2rem !important; padding-left: 0.8rem !important; padding-right: 0.8rem !important; padding-bottom: 3rem !important; max-width: 100% !important; }
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
.stButton > button { min-height: 48px !important; border-radius: 10px !important; font-weight: 700 !important; font-size: 1rem !important; }
.mobile-card { background-color: #f8f9fa; border: 1px solid #e2e8f0; border-radius: 12px; padding: 14px; margin-bottom: 12px; box-shadow: 0 2px 4px rgba(0,0,0,0.04); word-wrap: break-word; }
.mobile-card-route { background-color: #f0fff4; border: 2px solid #48bb78; border-radius: 12px; padding: 14px; margin-bottom: 12px; box-shadow: 0 3px 6px rgba(0,0,0,0.08); word-wrap: break-word; }
.badge { display: inline-block; padding: 4px 10px; border-radius: 12px; font-size: 0.75rem; font-weight: 700; text-transform: uppercase; margin-right: 4px; margin-bottom: 4px; }
.badge-route { background-color: #2f855a; color: #ffffff; }
.badge-solar { background-color: #feebc8; color: #c05621; }
.badge-foto { background-color: #fefcbf; color: #975a16; }
.badge-pompa { background-color: #bee3f8; color: #2b6cb0; }
.badge-magazyn { background-color: #e9d8fd; color: #6b46c1; }
.wizard-step-box { background-color: #ebf8ff; border-left: 5px solid #2b6cb0; padding: 12px; border-radius: 8px; margin-bottom: 14px; font-weight: 600; color: #2c5282; }
</style>""", unsafe_allow_html=True)

ATTACHMENTS_DIR = "zalaczniki"
NOWE_FILE = "nowe_zgloszenia.csv"
EXCEL_FILE = "Serwisowa posortowana.xlsx"
CENNIK_FILE = "cennik.csv"
CENNIK_HIST_FILE = "historia_cen.csv"
GSHEET_NAME = "PROMAG_BAZA_SERWISOWA"
os.makedirs(ATTACHMENTS_DIR, exist_ok=True)

CUTOFF_DATE = "2026-10-08"

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
    "solar": ["Odpowietrzono układ solarny", "Uzupełniono i sprawdzono stężenie glikolu", "Wymieniono pompę obiegową", "Wymieniono naczynie przeponowe", "Skonfigurowano sterownik solarny", "Uszczelniono połączenia śrubunkowe"],
    "pompa": ["Wyczyszczono filtr siatkowy / magnetyczny", "Skorygowano krzywą grzewczą", "Skasowano błąd i zresetowano sterownik", "Sprawdzono ciśnienie układu CO", "Odpowietrzono układ pompy ciepła", "Wymieniono czujnik temperatury"],
    "foto": ["Zresetowano i zaktualizowano falownik", "Sprawdzono napięcia stringów DC", "Dokręcono złącza w rozdzielnicy AC/DC", "Ponownie skonfigurowano Wi-Fi / monitoring", "Wymieniono konektory MC4", "Przetestowano zabezpieczenia nadprądowe"],
    "magazyn": ["Zaktualizowano oprogramowanie BMS", "Wykonano kalibrację SOC baterii", "Sprawdzono komunikację CAN/RS485 z falownikiem", "Zresetowano moduł bateryjny", "Zweryfikowano poprawność ładowania/rozładowania"]
}

QUICK_PARTS = {
    "solar": ["Glikol solarny 5L", "Glikol solarny 10L", "Pompa obiegowa Wilo 25/6", "Naczynie przeponowe 18L", "Naczynie przeponowe 24L", "Odpowietrznik automatyczny", "Zawór zwrotny 3/4 cala", "Czujnik temperatury PT1000"],
    "pompa": ["Zawór trójdrogowy 1 cal", "Czujnik przepływu", "Filtr magnetyczny 1 cal", "Odpowietrznik automatyczny", "Czujnik temperatury NTC", "Grupa bezpieczeństwa CO"],
    "foto": ["Komplet złączy MC4", "Ogranicznik przepięć AC/DC", "Wyłącznik nadprądowy B16/C16", "Moduł komunikacyjny Wi-Fi / Logger", "Przewód solarny 6mm2 (5m)"],
    "magazyn": ["Przewód komunikacyjny CAN/RS485", "Bezpiecznik DC 125A", "Moduł BMS", "Zestaw przewodów mocy DC"]
}

QUICK_RECOMMENDATIONS = [
    "Kontrola ciśnienia na manometrze za 48h",
    "Nie wyłączać zasilania głównego sterownika",
    "Obserwować pracę instalacji przez najbliższe 7 dni",
    "Zgłosić do biura w razie ponownego pojawienia się błędu",
    "Zalecany przegląd okresowy przed kolejnym sezonem",
    "Instalacja sprawna, poinstruowano klienta o obsłudze"
]

if 'cloud_error' not in st.session_state:
    st.session_state.cloud_error = ""

# =========================================================================
# OBSŁUGA GOOGLE SHEETS Z JAWNĄ DIAGNOSTYKĄ BŁĘDÓW
# =========================================================================
def get_gsheet_connection():
    if not GSPREAD_AVAILABLE:
        st.session_state.cloud_error = "Brak biblioteki 'gspread' lub 'google-auth' w pliku requirements.txt na GitHubie!"
        return None
    try:
        if "gcp_service_account" not in st.secrets:
            st.session_state.cloud_error = "Brak sekcji [gcp_service_account] w ustawieniach Secrets w Streamlit Cloud!"
            return None
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        creds = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"], scopes=scopes
        )
        client = gspread.authorize(creds)
        sh = client.open(GSHEET_NAME)
        st.session_state.cloud_error = ""
        return sh
    except gspread.exceptions.SpreadsheetNotFound:
        bot_email = st.secrets["gcp_service_account"].get("client_email", "konto-bota")
        st.session_state.cloud_error = f"Nie znaleziono Arkusza Google o nazwie '{GSHEET_NAME}' lub nie udostępniono go dla: {bot_email} (ważne: na Google Drive musi to być zapisany 'Arkusz Google', a nie zwykły plik .xlsx!)."
        return None
    except Exception as e:
        st.session_state.cloud_error = f"Błąd połączenia z Google Drive: {str(e)}"
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
                return sh.add_worksheet(title=sheet_title, rows="1000", cols="35")
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
        st.session_state.cloud_active = True
        st.session_state.cloud_error = ""
        return True
    except Exception as e:
        st.session_state.cloud_error = f"Błąd podczas zapisu do Arkusza Google: {str(e)}"
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

def ensure_required_columns(df):
    required_cols = [
        'nr.', 'INSTALACJA', 'DATA ZGŁOSZENIA', 'PRZYJMUJĄCY', 'Imię', 'Nazwisko',
        'Numer telefonu', 'E-mail', 'Kod pocztowy', 'Miejscowość', 'Adres',
        'Rejon Polski', 'STATUS ZGŁOSZENIA', 'SERWISANT', 'DATA SERWISU',
        'OPIS OD KLIENTA', 'OPIS', 'INFORMACJE DODATKOWE', 'Kwota', 'Płatność ',
        'Zużyte materiały', 'Czynności dodatkowe', 'Ilość przyjazdów', 'Kwota przyjęta',
        'Komentarz serwisanta', 'Email handlowca', 'Etap oferty', 'Data oferty',
        'Komentarz koordynatora', 'Zalecenia dla klienta', 'Przyczyna nieukończenia',
        'Brakujące materiały', 'Szczegóły oferty', 'Tabliczki_OK',
        'Kolejnosc_trasy', 'Data_trasy', 'Trasa_aktywna', 'Odbior_klienta',
        'Zdalne', 'Poczekalnia', 'Powod_poczekalni', 'Termin_koncowy',
        'Pozycje_cennika', 'Cena_bazowa', 'Rabat_proc',
        'Czas_start', 'Czas_koniec', 'Czas_pracy_min', 'Czas_trasy_min'
    ]
    for col in required_cols:
        if col not in df.columns:
            df[col] = ''
        else:
            df[col] = df[col].fillna('').astype(str)

    if 'Wymaga oferty' not in df.columns:
        df['Wymaga oferty'] = False
    else:
        df['Wymaga oferty'] = df['Wymaga oferty'].astype(str).str.lower().isin(['true', '1', 'tak'])

    if not df.empty:
        df['Klient_ID'] = df.apply(get_clean_client_id, axis=1)
    else:
        df['Klient_ID'] = pd.Series(dtype='str')
    return df

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
            st.session_state.cloud_error = f"Błąd odczytu rekordów z Arkusza Google: {str(e)}"

    if not loaded_from_cloud:
        st.session_state.cloud_active = False
        base_df = pd.DataFrame()
        if os.path.exists(EXCEL_FILE):
            try:
                base_df = pd.read_excel(EXCEL_FILE, sheet_name="Wszystkie posortowane")
            except Exception:
                try:
                    base_df = pd.read_excel(EXCEL_FILE, sheet_name=0)
                except Exception:
                    base_df = pd.DataFrame()

        if os.path.exists(NOWE_FILE):
            try:
                nowe_df = pd.read_csv(NOWE_FILE, dtype=str)
                df = pd.concat([nowe_df, base_df], ignore_index=True)
                if 'nr.' in df.columns:
                    df = df.drop_duplicates(subset=['nr.'], keep='first')
            except Exception:
                df = base_df
        else:
            df = base_df

        df = ensure_required_columns(df)
        if ws is not None and not df.empty:
            if sync_to_gsheet(df):
                st.session_state.cloud_active = True
        return df

    return ensure_required_columns(df)

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
if 'board_authenticated' not in st.session_state:
    st.session_state.board_authenticated = False
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
if 'tech_main_screen' not in st.session_state:
    st.session_state.tech_main_screen = "otwarte"
if 'wizard_step' not in st.session_state:
    st.session_state.wizard_step = 1
if 'wizard_data' not in st.session_state:
    st.session_state.wizard_data = {}

def get_next_ticket_id():
    if st.session_state.df.empty:
        return 1
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
    st.markdown("<h2 style='text-align: center; margin-top: 1rem;'>⚡ PROMAG SERWIS</h2>", unsafe_allow_html=True)
    if st.session_state.cloud_active:
        st.markdown("<p style='text-align: center; color: #2f855a; font-weight: bold;'>☁️ Połączono z bazą Google Drive (Zapis trwały aktywny)</p>", unsafe_allow_html=True)
    elif st.session_state.cloud_error:
        st.warning(f"⚠️ **Uwaga dot. zapisu na Google Drive:** {st.session_state.cloud_error}")

    col_r1, col_r2, col_r3 = st.columns(3)
    with col_r1:
        st.info("### 🖥️ Panel Koordynatora")
        st.write("Baza klientów, poczekalnia, zgłoszenia zdalne, cennik z historią i rabatami, SMS oraz trasy.")
        if st.button("Wejdź jako Koordynator", use_container_width=True, type="primary"):
            st.session_state.user_role = "koordynator"
            st.rerun()

    with col_r2:
        st.success("### 📱 Panel Serwisanta")
        st.write("Wygodny interfejs jednoekranowy pod jedną dłoń, trasa na dziś, szybkie kafelki i podpis klienta.")
        if st.button("Wejdź jako Serwisant (PIN)", use_container_width=True):
            st.session_state.user_role = "serwisant"
            st.rerun()

    with col_r3:
        st.warning("### 📊 Panel Zarządu")
        st.write("Podgląd zleceń i archiwum, statystyki serwisantów: zyski, czas pracy na instalacjach i czas w trasie.")
        if st.button("Wejdź jako Zarząd (PIN)", use_container_width=True):
            st.session_state.user_role = "zarzad"
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
        st.sidebar.error("💾 Tryb lokalny (Brak połączenia z chmurą)")
        if st.session_state.cloud_error:
            st.sidebar.caption(f"Powód: {st.session_state.cloud_error}")

    if st.sidebar.button("🔄 Odśwież / Połącz z Google Drive", use_container_width=True):
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
            'kod': r_kod, 'krotka': r_def['krotka'], 'pelna': r_def['nazwa'],
            'active': active_in_r, 'pending': pending_in_r, 'total': total_in_r
        })

    region_stats_sorted = sorted(region_stats, key=lambda x: x['active'], reverse=True)
    all_region_keys = ["ALL"] + [r['kod'] for r in region_stats_sorted]
    total_active_all = is_active_since_cutoff.sum()

    region_labels = {"ALL": f"🌐 Wszystkie rejony ({total_active_all} zgłoszeń od {CUTOFF_DATE})"}
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

    # Ratunkowe wgrywanie bazy Excel, jeśli chmura lub kontener były puste
    if df.empty:
        st.error("⚠️ Obecnie baza zgłoszeń jest pusta! Możesz poniżej wgrać swój plik `Serwisowa posortowana.xlsx` – zostanie on od razu wysłany i zapisany na stałe w Twoim Arkuszu Google.")
        rescue_file = st.file_uploader("Wgraj plik Excel z bazą (Serwisowa posortowana.xlsx):", type=['xlsx'])
        if rescue_file is not None:
            try:
                loaded_rescue = pd.read_excel(rescue_file, sheet_name=0)
                loaded_rescue = ensure_required_columns(loaded_rescue)
                st.session_state.df = loaded_rescue
                sync_excel_file(loaded_rescue)
                st.success(f"✅ Wgrano {len(loaded_rescue)} rekordów i zsynchronizowano z bazą!")
                st.rerun()
            except Exception as e:
                st.error(f"Błąd wczytywania pliku: {e}")

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
        st.caption(f"Numer telefonu klienta: **{sms_info['tel_clean']}** — Kliknij ikonę kopiowania w prawym górnym rogu poniższej ramki:")
        st.code(sms_info['tekst'], language=None)
        st.markdown("---")

    # Liczniki do zakładek specjalnych
    is_poczekalnia_open = (
        df['Poczekalnia'].astype(str).str.lower().isin(['true', '1', 'tak']) &
        ~df['STATUS ZGŁOSZENIA'].str.lower().str.contains('zakończ|anulow', na=False)
    )
    poczekalnia_df = df[is_poczekalnia_open]

    is_zdalne_open = (
        df['Zdalne'].astype(str).str.lower().isin(['true', '1', 'tak']) &
        ~df['STATUS ZGŁOSZENIA'].str.lower().str.contains('zakończ|anulow', na=False)
    )
    zdalne_open_df = df[is_zdalne_open]

    tab_k1, tab_k_pocz, tab_k_zdalne, tab_k2, tab_k3, tab_k_cennik = st.tabs([
        "🔍 Baza Klientów & Zgłoszenia",
        f"⏳ Poczekalnia ({len(poczekalnia_df)})",
        f"📞 Zgłoszenia Zdalne ({len(zdalne_open_df)})",
        "📑 Lejek Ofert Handlowych",
        "🗺️ Planowanie Tras",
        "💰 Cennik i Historia Cen"
    ])

    # --- ZAKŁADKA: POCZEKALNIA (UKRYTE PRZED SERWISANTEM) ---
    with tab_k_pocz:
        st.subheader("⏳ Poczekalnia (Zgłoszenia wstrzymane — niewidoczne dla serwisantów)")
        st.caption("Tutaj znajdują się zgłoszenia, w których czekasz np. na dosłanie zdjęć z obiektu, faktury lub decyzję klienta. Po uzupełnieniu danych kliknij 'Zwolnij do serwisu'.")

        if poczekalnia_df.empty:
            st.success("🎉 Poczekalnia jest pusta! Wszystkie zgłoszenia są przekazane do realizacji.")
        else:
            for idx_p, p_row in poczekalnia_df.iterrows():
                p_nr = p_row.get('nr.', '')
                p_klient = f"{p_row.get('Imię', '')} {p_row.get('Nazwisko', '')}"
                p_tel = p_row.get('Numer telefonu', '')
                p_miasto = f"{p_row.get('Kod pocztowy', '')} {p_row.get('Miejscowość', '')}"
                p_powod = p_row.get('Powod_poczekalni', 'Czekamy na zdjęcia / dokumenty od klienta')

                with st.expander(f"⏳ Wstrzymane #{p_nr} | {p_klient} (📞 {p_tel}) | Powód: {p_powod}", expanded=True):
                    c_p1, c_p2 = st.columns(2)
                    with c_p1:
                        st.markdown(f"**Klient:** {p_klient} (**📞 {p_tel}**)")
                        st.markdown(f"**Lokalizacja:** {p_miasto}, {p_row.get('Adres', '')} ({p_row.get('Rejon Polski', '')})")
                        st.markdown(f"**Zakres / Opis:** {p_row.get('OPIS OD KLIENTA', '')}")
                    with c_p2:
                        st.warning(f"📌 **Powód wstrzymania:** {p_powod}")
                        st.markdown(f"**Wycena:** {p_row.get('Kwota', '0')} zł | **Termin końcowy:** {p_row.get('Termin_koncowy', '—')}")

                    wyswietl_multimedia(p_nr)
                    nowe_pliki_pocz = st.file_uploader(f"📎 Dosłano zdjęcia/dokumenty? Wgraj je do zlecenia #{p_nr}:", type=['png', 'jpg', 'jpeg', 'webp', 'mp4'], accept_multiple_files=True, key=f"up_pocz_{p_nr}")
                    if nowe_pliki_pocz:
                        t_dir = os.path.join(ATTACHMENTS_DIR, str(p_nr))
                        os.makedirs(t_dir, exist_ok=True)
                        for f in nowe_pliki_pocz:
                            with open(os.path.join(t_dir, f.name), "wb") as f_out:
                                f_out.write(f.getbuffer())
                        st.success("Załączono nowe pliki!")

                    col_pb1, col_pb2 = st.columns(2)
                    with col_pb1:
                        if st.button("🚀 ZWOLNIJ Z POCZEKALNI (Przekaż serwisantowi do realizacji)", key=f"btn_rel_{p_nr}", type="primary", use_container_width=True):
                            df.at[idx_p, 'Poczekalnia'] = 'False'
                            df.at[idx_p, 'STATUS ZGŁOSZENIA'] = 'w toku'
                            st.session_state.df = df
                            save_ticket_persistently(df.loc[idx_p].to_dict(), current_full_df=st.session_state.df)
                            st.success(f"Zlecenie #{p_nr} zostało przekazane serwisantowi!")
                            st.rerun()
                    with col_pb2:
                        if st.button("❌ Anuluj zgłoszenie", key=f"btn_canc_pocz_{p_nr}", use_container_width=True):
                            df.at[idx_p, 'Poczekalnia'] = 'False'
                            df.at[idx_p, 'STATUS ZGŁOSZENIA'] = 'anulowane'
                            st.session_state.df = df
                            save_ticket_persistently(df.loc[idx_p].to_dict(), current_full_df=st.session_state.df)
                            st.rerun()

    # --- ZAKŁADKA: ZGŁOSZENIA ZDALNE ---
    with tab_k_zdalne:
        st.subheader("📞 Zgłoszenia do rozwiązania zdalnie (Helpdesk)")
        if zdalne_open_df.empty:
            st.success("🎉 Wszystkie zgłoszenia zdalne zostały rozwiązane!")
        else:
            for idx_z, z_row in zdalne_open_df.iterrows():
                z_nr = z_row.get('nr.', '')
                z_klient = f"{z_row.get('Imię', '')} {z_row.get('Nazwisko', '')}"
                z_tel = z_row.get('Numer telefonu', '')
                with st.expander(f"📞 Zdalne #{z_nr} | {z_klient} (📞 {z_tel}) | Data: {str(z_row.get('DATA ZGŁOSZENIA', ''))[:16]}", expanded=True):
                    st.markdown(f"**Opis problemu:** {z_row.get('OPIS OD KLIENTA', '')}")
                    roz_notatka = st.text_input("Jak rozwiązano problem zdalnie:", key=f"zdalne_note_{z_nr}")
                    col_zb1, col_zb2 = st.columns(2)
                    with col_zb1:
                        if st.button("✅ ZAMKNIJ (Rozwiązano zdalnie)", key=f"btn_z_close_{z_nr}", type="primary", use_container_width=True):
                            df.at[idx_z, 'STATUS ZGŁOSZENIA'] = 'zakończone (zdalnie)'
                            df.at[idx_z, 'DATA SERWISU'] = datetime.date.today().strftime("%Y-%m-%d")
                            df.at[idx_z, 'SERWISANT'] = 'Zdalnie (Koordynator)'
                            df.at[idx_z, 'Komentarz serwisanta'] = roz_notatka.strip() or "Rozwiązano zdalnie przez telefon."
                            st.session_state.df = df
                            save_ticket_persistently(df.loc[idx_z].to_dict(), current_full_df=st.session_state.df)
                            st.rerun()
                    with col_zb2:
                        if st.button("🚗 Przekaż do wyjazdu dla serwisanta", key=f"btn_z_to_field_{z_nr}", use_container_width=True):
                            df.at[idx_z, 'Zdalne'] = 'False'
                            df.at[idx_z, 'STATUS ZGŁOSZENIA'] = 'w toku'
                            st.session_state.df = df
                            save_ticket_persistently(df.loc[idx_z].to_dict(), current_full_df=st.session_state.df)
                            st.rerun()

    # --- ZAKŁADKA: CENNIK I HISTORIA CEN ---
    with tab_k_cennik:
        st.subheader("💰 Zarządzanie Cennikiem Usług i Części")
        edited_pricelist = st.data_editor(pricelist_df, num_rows="dynamic", use_container_width=True, hide_index=True, key="pricelist_editor_table")
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
                        history_logs.append({"Data zmiany": now_dt, "Kod": kod_p, "Nazwa pozycji": nazwa_p, "Stara cena (zł)": f"{stara_cena:.2f}", "Nowa cena (zł)": f"{nowa_cena:.2f}", "Zmienił": "Koordynator"})
                else:
                    history_logs.append({"Data zmiany": now_dt, "Kod": kod_p, "Nazwa pozycji": nazwa_p, "Stara cena (zł)": "Nowa pozycja", "Nowa cena (zł)": f"{nowa_cena:.2f}", "Zmienił": "Koordynator"})
            clean_pricelist = edited_pricelist[edited_pricelist['Nazwa'].astype(str).str.strip() != ''].copy()
            clean_pricelist['Cena'] = pd.to_numeric(clean_pricelist['Cena'], errors='coerce').fillna(0.0)
            save_pricelist(clean_pricelist)
            st.session_state.pricelist_df = clean_pricelist
            if history_logs:
                append_price_history(history_logs)
                st.success(f"✅ Zapisano nowy cennik oraz dodano {len(history_logs)} wpis(ów) do historii!")
            st.rerun()
        st.markdown("---")
        st.subheader("📜 Historia zmian cen")
        st.dataframe(st.session_state.price_history_df, use_container_width=True, hide_index=True)

    # --- ZAKŁADKA: PLANOWANIE TRAS ---
    with tab_k3:
        st.subheader("🗺️ Zaplanuj i zatwierdź trasę dla serwisanta")
        col_tr1, col_tr2 = st.columns(2)
        with col_tr1:
            plan_reg_options = [r['kod'] for r in REGION_DEFINITIONS]
            plan_reg_names = {r['kod']: r['nazwa'] for r in REGION_DEFINITIONS}
            default_plan_reg = selected_region_code if selected_region_code != "ALL" else "0"
            wybrany_rejon_trasy = st.selectbox("1. Wybierz rejon:", options=plan_reg_options, index=plan_reg_options.index(default_plan_reg) if default_plan_reg in plan_reg_options else 0, format_func=lambda x: plan_reg_names.get(x, x))
        with col_tr2:
            data_planowanej_trasy = st.date_input("2. Data realizacji trasy:", value=datetime.date.today() + datetime.timedelta(days=1))

        mask_plan_reg = df['Rejon Polski'].str.contains(f"Rejon {wybrany_rejon_trasy}", case=False, na=False) | df['Kod pocztowy'].str.startswith(wybrany_rejon_trasy)
        mask_open_plan = ~df['STATUS ZGŁOSZENIA'].str.lower().str.contains('zakończ|anulow|nowy klient', na=False)
        mask_not_remote = ~df['Zdalne'].astype(str).str.lower().isin(['true', '1', 'tak'])
        mask_not_waiting = ~df['Poczekalnia'].astype(str).str.lower().isin(['true', '1', 'tak'])
        mask_cutoff_plan = df['DATA ZGŁOSZENIA'].str[:10] >= CUTOFF_DATE

        kandydaci_df = df[mask_plan_reg & mask_open_plan & mask_not_remote & mask_not_waiting & mask_cutoff_plan].copy()
        if kandydaci_df.empty:
            st.info("Brak gotowych zleceń terenowych w tym rejonie do zaplanowania trasy.")
        else:
            route_editor_rows = []
            for _, r_k in kandydaci_df.iterrows():
                curr_order = str(r_k.get('Kolejnosc_trasy', '')).strip()
                is_in_route = str(r_k.get('Trasa_aktywna', '')).lower() in ['true', '1', 'tak']
                order_num = int(curr_order) if curr_order.isdigit() else 1
                route_editor_rows.append({
                    "Włącz do trasy": is_in_route, "Kolejność (1,2,3...)": order_num,
                    "Nr": str(r_k.get('nr.', '')), "Termin do": str(r_k.get('Termin_koncowy', '')),
                    "Miejscowość": r_k.get('Miejscowość', ''), "Adres": r_k.get('Adres', ''),
                    "Klient": f"{r_k.get('Imię', '')} {r_k.get('Nazwisko', '')}",
                    "Instalacja": str(r_k.get('INSTALACJA', '')).upper(),
                    "Opis usterki": r_k.get('OPIS OD KLIENTA', '') or r_k.get('OPIS', '')
                })
            edited_route = st.data_editor(pd.DataFrame(route_editor_rows), use_container_width=True, hide_index=True, disabled=["Nr", "Termin do", "Miejscowość", "Adres", "Klient", "Instalacja", "Opis usterki"], key=f"route_editor_{wybrany_rejon_trasy}")
            selected_for_route = edited_route[edited_route["Włącz do trasy"] == True].sort_values(by="Kolejność (1,2,3...)")
            if not selected_for_route.empty:
                if st.button("✅ ZATWIERDŹ I OPUBLIKUJ TRASĘ DLA SERWISANTA", type="primary", use_container_width=True):
                    df.loc[mask_plan_reg, 'Trasa_aktywna'] = 'False'
                    df.loc[mask_plan_reg, 'Kolejnosc_trasy'] = ''
                    for final_pos, (_, s_row) in enumerate(selected_for_route.iterrows(), 1):
                        m_idx = df['nr.'].astype(str) == str(s_row['Nr'])
                        df.loc[m_idx, 'Trasa_aktywna'] = 'True'
                        df.loc[m_idx, 'Kolejnosc_trasy'] = str(final_pos)
                        df.loc[m_idx, 'Data_trasy'] = str(data_planowanej_trasy)
                    st.session_state.df = df
                    save_multiple_tickets_persistently(df[mask_plan_reg], current_full_df=st.session_state.df)
                    st.success("🚀 Trasa wysłana do serwisanta!")
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
                o_klient = f"{o_row.get('Imię', '')} {o_row.get('Nazwisko', '')}"
                with st.expander(f"Zgłoszenie #{o_nr} | Klient: {o_klient} | Status: {o_etap}", expanded=True):
                    st.markdown(f"**Szczegóły oferty:** {o_row.get('Szczegóły oferty', '') or o_row.get('OPIS OD KLIENTA', '')}")
                    wyswietl_multimedia(o_nr)

    # --- ZAKŁADKA 1: BAZA KLIENTÓW & DODAWANIE ZGŁOSZEŃ ---
    with tab_k1:
        filtered_df = df.copy()
        if selected_region_code != "ALL":
            region_mask = df['Rejon Polski'].str.contains(f"Rejon {selected_region_code}", case=False, na=False) | df['Kod pocztowy'].str.startswith(selected_region_code)
            filtered_df = df[region_mask].copy()

        c_title, c_add_btn = st.columns([3, 1])
        with c_title:
            st.subheader("1. Baza Klientów (Wyszukaj lub Dodaj)")
        with c_add_btn:
            if st.button("➕ DODAJ NOWEGO KLIENTA", type="primary", use_container_width=True):
                st.session_state.show_new_client_modal = True

        if st.session_state.show_new_client_modal:
            with st.form("form_dodaj_klienta"):
                nc_c1, nc_c2 = st.columns(2)
                with nc_c1:
                    nc_imie = st.text_input("Imię:*")
                    nc_nazwisko = st.text_input("Nazwisko:*")
                    nc_tel = st.text_input("Numer telefonu:*")
                    nc_email = st.text_input("E-mail:")
                with nc_c2:
                    nc_kod = st.text_input("Kod pocztowy:*")
                    nc_miasto = st.text_input("Miejscowość:*")
                    nc_adres = st.text_input("Adres (ulica i nr):*")
                    all_region_names = [r['nazwa'] for r in REGION_DEFINITIONS]
                    nc_rejon = st.selectbox("Rejon Polski:", all_region_names)
                if st.form_submit_button("💾 ZAPISZ KLIENTA W BAZIE", type="primary"):
                    if nc_nazwisko and nc_tel:
                        new_client_row = {
                            "nr.": f"K{get_next_ticket_id()}", "INSTALACJA": "Brak (nowy klient)",
                            "DATA ZGŁOSZENIA": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
                            "PRZYJMUJĄCY": "Koordynator", "Imię": nc_imie.strip(), "Nazwisko": nc_nazwisko.strip(),
                            "Numer telefonu": nc_tel.strip(), "E-mail": nc_email.strip(), "Kod pocztowy": nc_kod.strip(),
                            "Miejscowość": nc_miasto.strip(), "Adres": nc_adres.strip(), "Rejon Polski": nc_rejon,
                            "STATUS ZGŁOSZENIA": "Nowy klient", "Zdalne": "False", "Poczekalnia": "False",
                            "Klient_ID": get_clean_client_id({'Numer telefonu': nc_tel, 'Imię': nc_imie, 'Nazwisko': nc_nazwisko})
                        }
                        st.session_state.df = ensure_required_columns(pd.concat([pd.DataFrame([new_client_row]), st.session_state.df], ignore_index=True))
                        save_ticket_persistently(new_client_row, current_full_df=st.session_state.df)
                        st.session_state.selected_client_id = new_client_row['Klient_ID']
                        st.session_state.show_new_client_modal = False
                        st.rerun()

        search_query = st.text_input("Wyszukaj klienta (nazwisko, telefon, miejscowość lub adres):").strip().lower()
        if search_query:
            query_mask = (
                filtered_df['Imię'].str.lower().str.contains(search_query, na=False) |
                filtered_df['Nazwisko'].str.lower().str.contains(search_query, na=False) |
                filtered_df['Numer telefonu'].str.lower().str.contains(search_query, na=False) |
                filtered_df['Miejscowość'].str.lower().str.contains(search_query, na=False) |
                filtered_df['Adres'].str.lower().str.contains(search_query, na=False)
            )
            filtered_df = filtered_df[query_mask]

        clients_unique = filtered_df.drop_duplicates(subset=['Klient_ID']) if not filtered_df.empty else pd.DataFrame()
        st.write(f"Znaleziono pasujących klientów: **{len(clients_unique)}**")

        if len(clients_unique) > 0:
            client_choices = [{'ID': r['Klient_ID'], 'Etykieta': f"{r.get('Imię', '')} {r.get('Nazwisko', '')} | 📞 {r.get('Numer telefonu', '')} | 📍 {r.get('Miejscowość', '')}, {r.get('Adres', '')}"} for _, r in clients_unique.iterrows()]
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

        if st.session_state.selected_client_id:
            client_records = st.session_state.df[st.session_state.df['Klient_ID'] == st.session_state.selected_client_id]
            if not client_records.empty:
                first_record = client_records.iloc[0]
                st.divider()
                st.subheader(f"👤 Karta Klienta: {first_record.get('Imię', '')} {first_record.get('Nazwisko', '')} (📞 {first_record.get('Numer telefonu', '')})")
                if st.button("➕ NOWE ZGŁOSZENIE DLA TEGO KLIENTA", type="primary"):
                    st.session_state.show_new_ticket_modal = not st.session_state.show_new_ticket_modal

                if st.session_state.show_new_ticket_modal:
                    st.info("📝 Rejestracja nowego zgłoszenia:")
                    next_id = get_next_ticket_id()
                    new_ticket_nr_preview = f"N{next_id}"

                    price_options = [f"{r['Nazwa']} ({float(r['Cena']):.0f} zł)" for _, r in pricelist_df.iterrows()]
                    price_lookup = {f"{r['Nazwa']} ({float(r['Cena']):.0f} zł)": float(r['Cena']) for _, r in pricelist_df.iterrows()}

                    col_p1, col_p2, col_p3 = st.columns([2.5, 1, 1])
                    with col_p1:
                        wybrane_pozycje_cennika = st.multiselect("Wybierz pozycje z cennika:", options=price_options)
                    suma_bazowa = sum([price_lookup.get(item, 0.0) for item in wybrane_pozycje_cennika])
                    with col_p2:
                        reczna_korekta = st.number_input("Kwota bazowa (zł):", min_value=0.0, value=float(suma_bazowa), step=50.0)
                    with col_p3:
                        rabat_procent = st.number_input("Rabat (%):", min_value=0, max_value=100, value=0, step=5)

                    kwota_po_rabacie = round(reczna_korekta * (1.0 - rabat_procent / 100.0), 2)
                    st.success(f"💵 **Kwota końcowa dla klienta: {kwota_po_rabacie:.0f} zł** (Rabat: {rabat_procent}%)")

                    f_col1, f_col2 = st.columns(2)
                    with f_col1:
                        forma = st.radio("Forma zgłoszenia:", ["telefon", "SMS", "WhatsApp", "e-mail"], horizontal=True)
                        czy_zdalne = st.checkbox("📞 Zgłoszenie do rozwiązania ZDALNIE (wisi u koordynatora)")
                        czy_poczekalnia = st.checkbox("⏳ Wstrzymaj w POCZEKALNI (serwisant go nie widzi - np. czekam na zdjęcia)")
                        powod_pocz = ""
                        if czy_poczekalnia:
                            powod_pocz = st.text_input("Powód wstrzymania w poczekalni:", value="Czekam na zdjęcia z obiektu od klienta")
                    with f_col2:
                        usterka_dotyczy = st.selectbox("Usterka dotyczy:", ["solar", "foto", "pompa", "magazyn"], index=0)
                        producent = st.text_input("Producent instalacji:", placeholder="np. Hewalex, Deye, Fronius...")
                        termin_koncowy = st.date_input("📅 Ostateczny termin realizacji (do SMS):", value=datetime.date.today() + datetime.timedelta(days=7))

                    domyslny_zakres = ", ".join([p.split(" (")[0] for p in wybrane_pozycje_cennika])
                    opis_zgloszenia = st.text_area("Zakres prac / Opis dla serwisanta (trafia też do SMS):*", value=domyslny_zakres)
                    komentarz_koordynatora = st.text_area("🔒 Poufna notatka koordynatora (niewidoczna dla serwisanta):")
                    uploaded_media = st.file_uploader("📷 Załącz zdjęcia/wideo:", type=['png', 'jpg', 'jpeg', 'webp', 'mp4'], accept_multiple_files=True)

                    if st.button("💾 ZAPISZ ZLECENIE I WYGENERUJ SMS", type="primary", use_container_width=True):
                        now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
                        new_ticket_nr = f"N{next_id}"
                        if czy_poczekalnia:
                            status_startowy = "W poczekalni (oczekuje)"
                        elif czy_zdalne:
                            status_startowy = "Zdalne - do realizacji"
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
                            "nr.": new_ticket_nr, "INSTALACJA": usterka_dotyczy, "DATA ZGŁOSZENIA": now_str,
                            "PRZYJMUJĄCY": "Koordynator", "Imię": first_record.get('Imię', ''),
                            "Nazwisko": first_record.get('Nazwisko', ''), "Numer telefonu": first_record.get('Numer telefonu', ''),
                            "E-mail": first_record.get('E-mail', ''), "Kod pocztowy": first_record.get('Kod pocztowy', ''),
                            "Miejscowość": first_record.get('Miejscowość', ''), "Adres": first_record.get('Adres', ''),
                            "Rejon Polski": first_record.get('Rejon Polski', ''), "STATUS ZGŁOSZENIA": status_startowy,
                            "SERWISANT": "Zdalnie (Koordynator)" if czy_zdalne else "",
                            "OPIS OD KLIENTA": opis_zgloszenia.strip(), "OPIS": "",
                            "INFORMACJE DODATKOWE": f"Forma: {forma} | Producent: {producent}",
                            "Kwota": kwota_str, "Płatność ": "", "Ilość przyjazdów": "1",
                            "Komentarz koordynatora": komentarz_koordynatora.strip(),
                            "Zdalne": "True" if czy_zdalne else "False",
                            "Poczekalnia": "True" if czy_poczekalnia else "False",
                            "Powod_poczekalni": powod_pocz,
                            "Termin_koncowy": termin_str,
                            "Pozycje_cennika": "; ".join(wybrane_pozycje_cennika),
                            "Cena_bazowa": f"{reczna_korekta:.0f}", "Rabat_proc": str(rabat_procent),
                            "Klient_ID": first_record['Klient_ID']
                        }

                        st.session_state.df = ensure_required_columns(pd.concat([pd.DataFrame([new_row_dict]), st.session_state.df], ignore_index=True))
                        save_ticket_persistently(new_row_dict, current_full_df=st.session_state.df)

                        raw_tel = str(first_record.get('Numer telefonu', '')).strip()
                        clean_tel = "".join([c for c in raw_tel.split(',')[0].split('/')[0] if c.isdigit()])
                        rabat_txt = f" (po rabacie {rabat_procent}%)" if rabat_procent > 0 else ""
                        sms_text = (
                            f"PROMAG SERWIS: Przyjelismy zlecenie nr #{new_ticket_nr}. "
                            f"Zakres prac: {opis_zgloszenia.strip() or usterka_dotyczy}. "
                            f"Wycena: {kwota_str} zl{rabat_txt}. "
                            f"Przewidywany termin realizacji do: {termin_str}. "
                            f"Jest to termin koncowy, a serwisant bedzie kontaktowal sie z Panstwem rano w dniu wizyty."
                        )
                        st.session_state.coord_sms_notification = {
                            "nr": new_ticket_nr, "klient": f"{first_record.get('Imię', '')} {first_record.get('Nazwisko', '')}",
                            "tel_clean": clean_tel, "tekst": sms_text
                        }
                        st.session_state.show_new_ticket_modal = False
                        st.rerun()

                st.markdown("### 📋 Historia zleceń tego klienta:")
                for i, (_, row) in enumerate(client_records.iterrows(), 1):
                    nr_t = str(row.get('nr.', ''))
                    with st.expander(f"Zgłoszenie #{nr_t} | Typ: {row.get('INSTALACJA', '')} | Status: {row.get('STATUS ZGŁOSZENIA', '')}", expanded=(i==1)):
                        st.write(f"**Opis:** {row.get('OPIS OD KLIENTA', '')} | **Kwota:** {row.get('Kwota', '0')} zł | **Termin:** {row.get('Termin_koncowy', '—')}")
                        wyswietl_multimedia(nr_t)

# =========================================================================
# 2. MODUŁ SERWISANTA (WIZARD 3-KROKOWY + POMIAR CZASU PRACY I TRASY)
# =========================================================================
elif st.session_state.user_role == "serwisant":
    df = st.session_state.df

    if not st.session_state.tech_pin:
        if st.button("⬅️ Wróć do wyboru modułu", use_container_width=True):
            st.session_state.user_role = None
            st.rerun()
        st.markdown("<h3 style='text-align: center;'>🔒 Logowanie Serwisanta</h3>", unsafe_allow_html=True)
        with st.form("form_pin"):
            pin_input = st.text_input("Wpisz 4-cyfrowy kod PIN:", type="password", max_chars=4)
            if st.form_submit_button("ZALOGUJ SIĘ", type="primary", use_container_width=True):
                if pin_input in PIN_CONFIG:
                    st.session_state.tech_pin = pin_input
                    st.session_state.tech_main_screen = "otwarte"
                    st.rerun()
                else:
                    st.error("❌ Błędny kod PIN.")
        st.stop()

    cfg = PIN_CONFIG[st.session_state.tech_pin]
    rejon_kod = cfg["rejon_kod"]
    rejon_nazwa = cfg["nazwa_rejonu"]
    domyslny_tech = cfg["domyslny_serwisant"]

    col_h1, col_h2 = st.columns([3, 1])
    with col_h1:
        st.markdown(f"**👷 {domyslny_tech}** | *{rejon_nazwa}*")
    with col_h2:
        if st.button("Wyloguj", use_container_width=True):
            st.session_state.tech_pin = None
            st.session_state.active_ticket_nr = None
            st.rerun()

    # EKRANOWY KREATOR ROZLICZENIA (3 KROKI)
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

        st.markdown(f'<div class="wizard-step-box">🛠️ Zlecenie #{nr_zlec} ({ticket.get("Imię","")} {ticket.get("Nazwisko","")}) — KROK {step} z 3</div>', unsafe_allow_html=True)
        st.progress(step / 3.0)

        if step == 1:
            st.subheader("Krok 1: Co zostało zrobione?")
            outcome = st.radio("Wynik wizyty:", ["TAK — Usterka usunięta", "NIE — Wymaga ponownego przyjazdu / części"], index=0 if w_data.get("is_success", True) else 1)
            is_succ = "TAK" in outcome
            if is_succ:
                q_act = st.multiselect("⚡ Szybki wybór wykonanych czynności:", options=QUICK_ACTIONS.get(typ_inst, QUICK_ACTIONS["solar"]), default=w_data.get("q_act", []))
                q_prt = st.multiselect("🔩 Szybki wybór zużytych części:", options=QUICK_PARTS.get(typ_inst, QUICK_PARTS["solar"]), default=w_data.get("q_prt", []))
                opis_prac = st.text_area("Opis wykonanych prac:*", value=", ".join(q_act) if q_act else w_data.get("opis_prac", ""))
                zuzyte_czesci = st.text_area("Zużyte części / materiały:", value=", ".join(q_prt) if q_prt else w_data.get("zuzyte_czesci", ""))
                przyczyna_nieuk, brakujace_mat = "", ""
            else:
                q_act = []
                q_prt = st.multiselect("🔩 Szybki wybór brakujących części:", options=QUICK_PARTS.get(typ_inst, QUICK_PARTS["solar"]), default=w_data.get("q_prt", []))
                przyczyna_nieuk = st.text_area("Dlaczego nie udało się ukończyć:*", value=w_data.get("przyczyna_nieuk", ""))
                brakujace_mat = st.text_area("Jakich części brakuje:*", value=", ".join(q_prt) if q_prt else w_data.get("brakujace_mat", ""))
                opis_prac, zuzyte_czesci = "", ""

            q_rec = st.multiselect("💡 Szybki wybór zaleceń dla klienta:", options=QUICK_RECOMMENDATIONS, default=w_data.get("q_rec", []))
            zalecenia = st.text_area("Zalecenia dla klienta (OBOWIĄZKOWE):*", value=". ".join(q_rec) if q_rec else w_data.get("zalecenia", ""))

            col_nav1, col_nav2 = st.columns(2)
            with col_nav1:
                if st.button("❌ Wróć do listy", use_container_width=True):
                    st.session_state.active_ticket_nr = None
                    st.session_state.wizard_step = 1
                    st.session_state.wizard_data = {}
                    st.rerun()
            with col_nav2:
                if st.button("DALEJ: Zdjęcia ➡️", type="primary", use_container_width=True):
                    if not zalecenia.strip():
                        st.error("❌ Wpisz zalecenia dla klienta!")
                    else:
                        w_data.update({"is_success": is_succ, "q_act": q_act, "q_prt": q_prt, "q_rec": q_rec, "opis_prac": opis_prac.strip(), "zuzyte_czesci": zuzyte_czesci.strip(), "przyczyna_nieuk": przyczyna_nieuk.strip(), "brakujace_mat": brakujace_mat.strip(), "zalecenia": zalecenia.strip()})
                        st.session_state.wizard_data = w_data
                        st.session_state.wizard_step = 2
                        st.rerun()

        elif step == 2:
            st.subheader("Krok 2: Zdjęcia i Oferta")
            potwierdzenie_tabliczek = st.checkbox("✅ POTWIERDZAM: Zrobiłem zdjęcia tabliczek znamionowych", value=w_data.get("potwierdzenie_tabliczek", True))
            tech_media = st.file_uploader("Wgraj zdjęcia z galerii:", type=['png', 'jpg', 'jpeg', 'webp', 'mp4'], accept_multiple_files=True)
            camera_photo = st.camera_input("Lub zrób zdjęcie aparatem:")

            target_dir = os.path.join(ATTACHMENTS_DIR, nr_zlec)
            os.makedirs(target_dir, exist_ok=True)
            if tech_media:
                for file in tech_media:
                    with open(os.path.join(target_dir, file.name), "wb") as f_out:
                        f_out.write(file.getbuffer())
            if camera_photo:
                c_name = f"aparat_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.jpg"
                with open(os.path.join(target_dir, c_name), "wb") as f_c:
                    f_c.write(camera_photo.getbuffer())

            potrzebna_oferta = st.checkbox("💼 Klient prosi o OFERTĘ HANDLOWĄ (rozbudowa/modyfikacja)", value=w_data.get("potrzebna_oferta", False))
            szczegoly_oferty = st.text_area("Co wycenić dla klienta:*", value=w_data.get("szczegoly_oferty", "")) if potrzebna_oferta else ""

            col_nav1, col_nav2 = st.columns(2)
            with col_nav1:
                if st.button("⬅️ WSTECZ", use_container_width=True):
                    st.session_state.wizard_step = 1
                    st.rerun()
            with col_nav2:
                if st.button("DALEJ: Podpis i Czas ➡️", type="primary", use_container_width=True):
                    w_data.update({"potwierdzenie_tabliczek": potwierdzenie_tabliczek, "potrzebna_oferta": potrzebna_oferta, "szczegoly_oferty": szczegoly_oferty.strip()})
                    st.session_state.wizard_data = w_data
                    st.session_state.wizard_step = 3
                    st.rerun()

        elif step == 3:
            st.subheader("Krok 3: Kwota, Czas i Podpis Klienta")
            wycena_biura = ticket.get('Kwota', '0') or '0'
            st.info(f"💰 Wycena z biura: **{wycena_biura} zł**")

            # Automatyczne wyliczenie czasu pracy na instalacji (od kliknięcia "Rozpocznij pracę")
            auto_work_min = 60
            czas_start_str = str(ticket.get('Czas_start', '')).strip()
            if czas_start_str:
                try:
                    t_start = datetime.datetime.strptime(czas_start_str, "%Y-%m-%d %H:%M")
                    diff_min = int((datetime.datetime.now() - t_start).total_seconds() // 60)
                    if 5 <= diff_min <= 720:
                        auto_work_min = diff_min
                except Exception:
                    pass

            c_f1, c_f2 = st.columns(2)
            with c_f1:
                kwota_pobrana = st.text_input("Kwota pobrana (zł):", value=str(wycena_biura if w_data.get("is_success", True) else "0"))
                czas_pracy_min = st.number_input("⏱️ Czas pracy na instalacji (min):", min_value=5, max_value=720, value=auto_work_min, step=15)
            with c_f2:
                forma_plat = st.selectbox("Forma płatności:", ["Gotówka", "Przelew", "Gwarancja (0 zł)", "Czeka na przelew"])
                czas_trasy_min = st.number_input("🚗 Czas dojazdu / w trasie (min):", min_value=0, max_value=480, value=30, step=15)

            status_opcje = ["zakończone", "do ponownego przyjazdu (brak części)", "w toku (wymaga części)", "anulowane"]
            status_nowy = st.selectbox("Końcowy status zlecenia:", status_opcje, index=0 if w_data.get("is_success", True) else 1)

            st.markdown("##### ✍️ Podpis klienta na ekranie telefonu")
            if CANVAS_AVAILABLE:
                canvas_result = st_canvas(fill_color="rgba(255, 255, 255, 1)", stroke_width=3, stroke_color="#1a202c", background_color="#ffffff", height=160, drawing_mode="freedraw", key=f"sig_canvas_{nr_zlec}")
            else:
                canvas_result = None

            odbior_klienta = st.checkbox("✅ Klient potwierdził odbiór prac i zalecenia", value=True)

            col_nav1, col_nav2 = st.columns(2)
            with col_nav1:
                if st.button("⬅️ WSTECZ", use_container_width=True):
                    st.session_state.wizard_step = 2
                    st.rerun()
            with col_nav2:
                if st.button("💾 ZAKOŃCZ I ZAPISZ", type="primary", use_container_width=True):
                    idx = curr_rows.index[0]
                    is_succ = w_data.get("is_success", True)
                    df.at[idx, 'Komentarz serwisanta'] = w_data.get("opis_prac", "") if is_succ else w_data.get("przyczyna_nieuk", "")
                    df.at[idx, 'Zużyte materiały'] = w_data.get("zuzyte_czesci", "") if is_succ else ""
                    df.at[idx, 'Przyczyna nieukończenia'] = "" if is_succ else w_data.get("przyczyna_nieuk", "")
                    df.at[idx, 'Brakujące materiały'] = "" if is_succ else w_data.get("brakujace_mat", "")
                    df.at[idx, 'Zalecenia dla klienta'] = w_data.get("zalecenia", "")
                    df.at[idx, 'Kwota przyjęta'] = kwota_pobrana
                    df.at[idx, 'Płatność '] = forma_plat
                    df.at[idx, 'STATUS ZGŁOSZENIA'] = status_nowy
                    df.at[idx, 'DATA SERWISU'] = datetime.date.today().strftime("%Y-%m-%d")
                    df.at[idx, 'Czas_koniec'] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
                    df.at[idx, 'Czas_pracy_min'] = str(czas_pracy_min)
                    df.at[idx, 'Czas_trasy_min'] = str(czas_trasy_min)
                    df.at[idx, 'SERWISANT'] = domyslny_tech
                    df.at[idx, 'Odbior_klienta'] = "Podpisano" if odbior_klienta else "Brak"

                    if CANVAS_AVAILABLE and canvas_result is not None and canvas_result.image_data is not None:
                        try:
                            target_dir = os.path.join(ATTACHMENTS_DIR, nr_zlec)
                            os.makedirs(target_dir, exist_ok=True)
                            Image.fromarray(canvas_result.image_data.astype('uint8'), 'RGBA').save(os.path.join(target_dir, f"podpis_{nr_zlec}.png"))
                        except Exception:
                            pass

                    st.session_state.df = df
                    save_ticket_persistently(df.loc[idx].to_dict(), current_full_df=st.session_state.df)
                    st.session_state.active_ticket_nr = None
                    st.session_state.wizard_step = 1
                    st.session_state.wizard_data = {}
                    st.rerun()
        st.stop()

    # Filtrowanie zleceń dla serwisanta (Pomija Zdalne oraz Poczekalnię!)
    tech_df = df.copy() if rejon_kod == "ALL" else df[df['Rejon Polski'].str.contains(f"Rejon {rejon_kod}", case=False, na=False) | df['Kod pocztowy'].str.startswith(f"{rejon_kod}")]
    not_remote_mask = ~tech_df['Zdalne'].astype(str).str.lower().isin(['true', '1', 'tak'])
    not_waiting_mask = ~tech_df['Poczekalnia'].astype(str).str.lower().isin(['true', '1', 'tak'])
    tech_operational_df = tech_df[not_remote_mask & not_waiting_mask & (tech_df['DATA ZGŁOSZENIA'].str[:10] >= CUTOFF_DATE)]

    is_done = tech_operational_df['STATUS ZGŁOSZENIA'].str.lower().str.contains('zakończ|anulow|nowy klient', na=False)
    otwarte_df = tech_operational_df[~is_done].copy()
    zakonczone_df = tech_operational_df[is_done].copy()

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
        opis_klienta = str(row.get('OPIS OD KLIENTA', '')).strip() or "Brak opisu"
        badge_type = "badge-solar" if "solar" in typ else ("badge-foto" if "foto" in typ else ("badge-pompa" if "pompa" in typ else "badge-magazyn"))

        card_html = f'<div class="mobile-card"><div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;"><span style="font-weight:bold; font-size:1.1rem; color:#1a202c;">#{nr}</span><div><span class="badge {badge_type}">{typ.upper() or "SERWIS"}</span><span class="badge" style="background:#e2e8f0; color:#2d3748;">{curr_status.upper()}</span></div></div><div style="font-size:1.05rem; font-weight:700; color:#2d3748;">👤 {imie_nazwisko}</div><div style="font-size:0.95rem; color:#4a5568; margin-top:3px;">📍 {adres_full}</div><div style="margin-top:8px; padding:8px; background:#edf2f7; border-radius:8px; font-size:0.9rem; color:#2d3748;"><b>Zakres / Usterka:</b> {opis_klienta}</div><div style="display:flex; justify-content:space-between; margin-top:8px; font-size:0.85rem; color:#718096;"><span>💰 Wycena: <b>{kwota_val} zł</b></span><span>⏳ Do: <b>{row.get("Termin_koncowy", "—")}</b></span></div></div>'
        st.markdown(card_html, unsafe_allow_html=True)

        col_a, col_sms, col_b = st.columns(3)
        with col_a:
            if tel_clean:
                st.markdown(f'<a href="tel:{tel_clean}" target="_blank" style="text-decoration:none;"><button style="width:100%; height:48px; border-radius:10px; background:#2b6cb0; color:white; border:none; font-weight:bold;">📞 Zadzwoń</button></a>', unsafe_allow_html=True)
        with col_sms:
            if tel_clean:
                sms_msg = urllib.parse.quote("Dzień dobry, z tej strony serwis PROMAG. Będę u Państwa dzisiaj zgodnie z planem.")
                st.markdown(f'<a href="sms:{tel_clean}?body={sms_msg}" target="_blank" style="text-decoration:none;"><button style="width:100%; height:48px; border-radius:10px; background:#805ad5; color:white; border:none; font-weight:bold;">💬 SMS</button></a>', unsafe_allow_html=True)
        with col_b:
            if row.get('Miejscowość', ''):
                st.markdown(f'<a href="https://www.google.com/maps/search/?api=1&query={map_query}" target="_blank" style="text-decoration:none;"><button style="width:100%; height:48px; border-radius:10px; background:#38a169; color:white; border:none; font-weight:bold;">🗺️ Nawiguj</button></a>', unsafe_allow_html=True)

        if is_open:
            btn_col1, btn_col2 = st.columns(2)
            with btn_col1:
                if curr_status.lower() in ['w trakcie realizacji', 'rozpoczęte']:
                    st.button(f"⚙️ W TRAKCIE (od {str(row.get('Czas_start',''))[-5:]})", key=f"btn_started_{nr}", disabled=True, use_container_width=True)
                else:
                    if st.button("▶️ Rozpocznij pracę", key=f"btn_start_{nr}", use_container_width=True):
                        match_idx = df[df['nr.'].astype(str) == nr].index
                        if not match_idx.empty:
                            df.at[match_idx[0], 'STATUS ZGŁOSZENIA'] = 'w trakcie realizacji'
                            df.at[match_idx[0], 'SERWISANT'] = domyslny_tech
                            df.at[match_idx[0], 'Czas_start'] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
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

    if st.session_state.tech_main_screen == "otwarte":
        if otwarte_df.empty:
            st.info("🎉 Brak otwartych zleceń w Twoim rejonie!")
        else:
            for _, row in otwarte_df.head(30).iterrows():
                render_ticket_card(row, is_open=True)
    elif st.session_state.tech_main_screen == "zakonczone":
        for _, row in zakonczone_df.head(20).iterrows():
            render_ticket_card(row, is_open=False)
    elif st.session_state.tech_main_screen == "szukaj":
        tech_search = st.text_input("Wpisz nazwisko, telefon lub miejscowość:").strip().lower()
        if tech_search:
            t_match = df[df['Nazwisko'].str.lower().str.contains(tech_search, na=False) | df['Numer telefonu'].str.lower().str.contains(tech_search, na=False) | df['Miejscowość'].str.lower().str.contains(tech_search, na=False)]
            for _, h_row in t_match.head(15).iterrows():
                render_ticket_card(h_row, is_open=False)
    elif st.session_state.tech_main_screen == "wiedza":
        st.subheader("📘 Ściągawka Techniczna & Kody Błędów")
        st.info("☀️ **SOLARY:** Ciśnienie zimne 1.5–2.0 bar | PT1000 w 20°C = 1078 Ω.\n\n❄️ **POMPY CIEPŁA:** E03/E04 = wyczyść filtr siatkowy/magnetyczny.\n\n⚡ **FOTOWOLTAIKA:** Deye F18 = sprawdź zwarcie AC; F56 = doładuj baterię z sieci.")

# =========================================================================
# 3. PANEL ZARZĄDU (ANALITYKA SERWISANTÓW, BIEŻĄCE ZLECENIA I ARCHIWUM)
# =========================================================================
elif st.session_state.user_role == "zarzad":
    df = st.session_state.df

    if not st.session_state.board_authenticated:
        if st.button("⬅️ Wróć do wyboru modułu"):
            st.session_state.user_role = None
            st.rerun()
        st.markdown("<h3 style='text-align: center;'>📊 Autoryzacja Panelu Zarządu</h3>", unsafe_allow_html=True)
        with st.form("board_auth_form"):
            b_pin = st.text_input("Wprowadź kod PIN Zarządu (domyślnie: 0000):", type="password")
            if st.form_submit_button("WEJDŹ DO PANELU ZARZĄDU", type="primary", use_container_width=True):
                if b_pin in ["0000", "9999"]:
                    st.session_state.board_authenticated = True
                    st.rerun()
                else:
                    st.error("❌ Nieprawidłowy kod PIN.")
        st.stop()

    col_b_top1, col_b_top2 = st.columns([4, 1])
    with col_b_top1:
        st.title("📊 Panel Zarządu — Analityka i Nadzór Serwisu")
    with col_b_top2:
        if st.button("🚪 Wyloguj", use_container_width=True):
            st.session_state.board_authenticated = False
            st.session_state.user_role = None
            st.rerun()

    tab_z1, tab_z2, tab_z3 = st.tabs([
        "👷 Wyniki Serwisantów (Zyski i Czas)",
        f"⚡ Bieżące Zlecenia (od {CUTOFF_DATE})",
        f"🗄️ Pełne Archiwum (przed {CUTOFF_DATE})"
    ])

    active_df = df[df['DATA ZGŁOSZENIA'].str[:10] >= CUTOFF_DATE].copy()
    archive_df = df[df['DATA ZGŁOSZENIA'].str[:10] < CUTOFF_DATE].copy()

    # --- ZAKŁADKA 1: ROZLICZENIE SERWISANTÓW ---
    with tab_z1:
        st.subheader("📈 Zestawienie efektywności, czasu pracy i zysków na serwisanta")
        lista_serwisantow = ["Leszek", "Stanisław", "Wojtek", "Mikołaj", "Artur", "Piotr", "Alex", "Dominik", "Hubert", "Piasta", "Zdalnie (Koordynator)"]

        stat_rows = []
        for tech_name in lista_serwisantow:
            # Przypisane bezpośrednio lub wynikające z domyślnego rejonu serwisanta
            tech_pins = [cfg['rejon_kod'] for _, cfg in PIN_CONFIG.items() if cfg['domyslny_serwisant'].lower() == tech_name.lower()]
            if tech_pins and tech_pins[0] != "ALL":
                r_code = tech_pins[0]
                mask_t = (active_df['SERWISANT'].str.lower() == tech_name.lower()) | (
                    (active_df['SERWISANT'] == '') & (active_df['Rejon Polski'].str.contains(f"Rejon {r_code}", case=False, na=False) | active_df['Kod pocztowy'].str.startswith(r_code))
                )
            else:
                mask_t = active_df['SERWISANT'].str.lower() == tech_name.lower()

            t_records = active_df[mask_t & ~active_df['STATUS ZGŁOSZENIA'].str.lower().str.contains('nowy klient|anulow', na=False)]
            done_mask = t_records['STATUS ZGŁOSZENIA'].str.lower().str.contains('zakończ', na=False)
            w_trakcie_cnt = (~done_mask).sum()
            zrealizowane_cnt = done_mask.sum()

            # Liczenie zysków z zamkniętych zleceń
            done_records = t_records[done_mask]
            zysk_suma = 0.0
            czas_pracy_suma_min = 0
            czas_trasy_suma_min = 0

            for _, r_done in done_records.iterrows():
                kw_str = str(r_done.get('Kwota przyjęta', '') or r_done.get('Kwota', '0')).replace('zł', '').replace(' ', '').replace(',', '.').strip()
                try:
                    zysk_suma += float(kw_str)
                except Exception:
                    pass
                try:
                    czas_pracy_suma_min += int(float(str(r_done.get('Czas_pracy_min', '0') or '0')))
                except Exception:
                    pass
                try:
                    czas_trasy_suma_min += int(float(str(r_done.get('Czas_trasy_min', '0') or '0')))
                except Exception:
                    pass

            stat_rows.append({
                "Serwisant": tech_name,
                "Zlecenia w trakcie": int(w_trakcie_cnt),
                "Zlecenia zrealizowane": int(zrealizowane_cnt),
                "Podsumowanie zysków (zł)": round(zysk_suma, 2),
                "Czas pracy na instalacjach": f"{czas_pracy_suma_min // 60}h {czas_pracy_suma_min % 60}m ({czas_pracy_suma_min} min)",
                "Czas w trasie": f"{czas_trasy_suma_min // 60}h {czas_trasy_suma_min % 60}m ({czas_trasy_suma_min} min)"
            })

        stats_df = pd.DataFrame(stat_rows)
        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        col_m1.metric("Łącznie w trakcie", int(stats_df["Zlecenia w trakcie"].sum()))
        col_m2.metric("Łącznie zrealizowane", int(stats_df["Zlecenia zrealizowane"].sum()))
        col_m3.metric("Suma zysków (od 08.10)", f"{stats_df['Podsumowanie zysków (zł)'].sum():,.2f} zł")
        col_m4.metric("Zgłoszenia w poczekalni", int((active_df['Poczekalnia'].astype(str).str.lower() == 'true').sum()))

        st.dataframe(stats_df, use_container_width=True, hide_index=True)

    # --- ZAKŁADKA 2: BIEŻĄCE ZLECENIA ---
    with tab_z2:
        st.subheader(f"📋 Wszystkie bieżące zlecenia serwisowe (od {CUTOFF_DATE})")
        st.dataframe(active_df.drop(columns=['Klient_ID'], errors='ignore'), use_container_width=True, hide_index=True)

    # --- ZAKŁADKA 3: PEŁNE ARCHIWUM ---
    with tab_z3:
        st.subheader(f"🗄️ Archiwum zleceń (przed {CUTOFF_DATE} — łącznie {len(archive_df)} rekordów)")
        arch_q = st.text_input("Przeszukaj archiwum (nazwisko, miejscowość, telefon, serwisant):").strip().lower()
        view_arch = archive_df
        if arch_q:
            view_arch = archive_df[
                archive_df['Nazwisko'].str.lower().str.contains(arch_q, na=False) |
                archive_df['Miejscowość'].str.lower().str.contains(arch_q, na=False) |
                archive_df['Numer telefonu'].str.lower().str.contains(arch_q, na=False) |
                archive_df['SERWISANT'].str.lower().str.contains(arch_q, na=False)
            ]
        st.dataframe(view_arch.drop(columns=['Klient_ID'], errors='ignore'), use_container_width=True, hide_index=True)
