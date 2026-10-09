import streamlit as st
import pandas as pd
import datetime
import os
import urllib.parse

st.set_page_config(
    page_title="PROMAG SERWIS - System Obsługi",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Katalogi na dane i załączniki
ATTACHMENTS_DIR = "zalaczniki"
NOWE_FILE = "nowe_zgloszenia.csv"
os.makedirs(ATTACHMENTS_DIR, exist_ok=True)

# SZTYWNA DATA STARTU NOWEGO SYSTEMU (ODCIĘCIE HISTORII)
CUTOFF_DATE = "2026-10-08"

# Style responsywne pod komputery i smartfony
st.markdown("""
<style>
    .mobile-card {
        background-color: #f8f9fa;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 14px;
        margin-bottom: 12px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.04);
    }
    .badge {
        display: inline-block;
        padding: 4px 10px;
        border-radius: 12px;
        font-size: 0.75rem;
        font-weight: 700;
        text-transform: uppercase;
        margin-right: 5px;
    }
    .badge-solar { background-color: #feebc8; color: #c05621; }
    .badge-foto { background-color: #fefcbf; color: #975a16; }
    .badge-pompa { background-color: #bee3f8; color: #2b6cb0; }
    .badge-magazyn { background-color: #e9d8fd; color: #6b46c1; }
</style>
""", unsafe_allow_html=True)

# Konfiguracja okręgów pocztowych (0-9)
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

def wyswietl_multimedia(ticket_nr):
    ticket_folder = os.path.join(ATTACHMENTS_DIR, str(ticket_nr))
    if os.path.exists(ticket_folder):
        files = os.listdir(ticket_folder)
        if files:
            st.markdown("📷 **Załączone materiały (zdjęcia / filmiki):**")
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

# Ładowanie danych z obsługą trwałego pliku CSV
def load_all_data():
    file_path = "Serwisowa posortowana.xlsx"
    base_df = pd.DataFrame()
    if os.path.exists(file_path):
        try:
            base_df = pd.read_excel(file_path, sheet_name="Wszystkie posortowane")
        except Exception:
            base_df = pd.read_excel(file_path, sheet_name=0)

    if os.path.exists(NOWE_FILE):
        try:
            nowe_df = pd.read_csv(NOWE_FILE, dtype=str)
            df = pd.concat([nowe_df, base_df], ignore_index=True)
            df = df.drop_duplicates(subset=['nr.'], keep='first')
        except Exception:
            df = base_df
    else:
        df = base_df

    cols_to_str = ['nr.', 'INSTALACJA', 'PRZYJMUJĄCY', 'Imię', 'Nazwisko', 'Numer telefonu', 
                   'Kod pocztowy', 'Miejscowość', 'Adres', 'Rejon Polski', 'STATUS ZGŁOSZENIA', 
                   'SERWISANT', 'OPIS OD KLIENTA', 'OPIS', 'INFORMACJE DODATKOWE', 'Kwota', 'Płatność ', 'E-mail']
    for c in cols_to_str:
        if c in df.columns:
            df[c] = df[c].fillna('').astype(str)

    if 'DATA ZGŁOSZENIA' in df.columns:
        df['DATA ZGŁOSZENIA'] = df['DATA ZGŁOSZENIA'].fillna('').astype(str)

    for opt_col in ['Zużyte materiały', 'Czynności dodatkowe', 'Ilość przyjazdów', 'Kwota przyjęta', 
                    'Komentarz serwisanta', 'Email handlowca', 'Etap oferty', 'Data oferty']:
        if opt_col not in df.columns:
            df[opt_col] = ''
        else:
            df[opt_col] = df[opt_col].fillna('').astype(str)

    if 'Wymaga oferty' not in df.columns:
        df['Wymaga oferty'] = False

    df['Klient_ID'] = df.apply(get_clean_client_id, axis=1)
    return df

# Trwały zapis pojedynczego zgłoszenia do pliku CSV
def save_ticket_persistently(ticket_dict):
    df_single = pd.DataFrame([ticket_dict])
    if os.path.exists(NOWE_FILE):
        existing_nowe = pd.read_csv(NOWE_FILE, dtype=str)
        if str(ticket_dict.get('nr.', '')) in existing_nowe['nr.'].astype(str).values:
            existing_nowe = existing_nowe[existing_nowe['nr.'].astype(str) != str(ticket_dict.get('nr.', ''))]
        updated = pd.concat([df_single, existing_nowe], ignore_index=True)
        updated.to_csv(NOWE_FILE, index=False, encoding='utf-8-sig')
    else:
        df_single.to_csv(NOWE_FILE, index=False, encoding='utf-8-sig')

# Trwały zapis wielu zaktualizowanych zgłoszeń (np. po edycji danych klienta)
def save_multiple_tickets_persistently(updated_df_rows):
    if updated_df_rows.empty:
        return
    if os.path.exists(NOWE_FILE):
        existing_nowe = pd.read_csv(NOWE_FILE, dtype=str)
        updated_nrs = updated_df_rows['nr.'].astype(str).values
        existing_nowe = existing_nowe[~existing_nowe['nr.'].astype(str).isin(updated_nrs)]
        merged = pd.concat([updated_df_rows, existing_nowe], ignore_index=True)
        merged.to_csv(NOWE_FILE, index=False, encoding='utf-8-sig')
    else:
        updated_df_rows.to_csv(NOWE_FILE, index=False, encoding='utf-8-sig')

if 'df' not in st.session_state:
    st.session_state.df = load_all_data()

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

if 'user_role' not in st.session_state:
    st.session_state.user_role = None

if 'tech_pin' not in st.session_state:
    st.session_state.tech_pin = None

if 'selected_client_id' not in st.session_state:
    st.session_state.selected_
