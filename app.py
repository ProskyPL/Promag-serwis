import streamlit as st
import pandas as pd
import datetime
import os
import shutil
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
EXCEL_FILE = "Serwisowa posortowana.xlsx"
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
    .badge-status-started { background-color: #c6f6d5; color: #22543d; }
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

# Ładowanie danych z obsługą trwałego pliku CSV i Excela
def load_all_data():
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
                    'Brakujące materiały', 'Szczegóły oferty', 'Tabliczki_OK']:
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

# Bezpieczny zapis całego DataFrame do pliku Excel z kopią zapasową
def sync_excel_file(df_to_save):
    if df_to_save.empty:
        return
    try:
        if os.path.exists(EXCEL_FILE):
            shutil.copyfile(EXCEL_FILE, f"{EXCEL_FILE}.bak")
        
        clean_save_df = df_to_save.drop(columns=['Klient_ID'], errors='ignore')
        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
            clean_save_df.to_excel(writer, sheet_name="Wszystkie posortowane", index=False)
    except Exception as e:
        print(f"Błąd zapisu Excel: {e}")

# Podwójny trwały zapis pojedynczego zgłoszenia (CSV + Excel)
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

# Podwójny trwały zapis wielu rekordów (CSV + Excel)
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
    st.session_state.selected_client_id = None

if 'show_new_ticket_modal' not in st.session_state:
    st.session_state.show_new_ticket_modal = False

if 'show_new_client_modal' not in st.session_state:
    st.session_state.show_new_client_modal = False

if 'show_edit_client_modal' not in st.session_state:
    st.session_state.show_edit_client_modal = False

if 'active_ticket_nr' not in st.session_state:
    st.session_state.active_ticket_nr = None

# --- EKRAN STARTOWY: WYBÓR ROLI ---
if not st.session_state.user_role:
    st.markdown("<h2 style='text-align: center; margin-top: 2rem;'>⚡ System Zgłoszeń Serwisowych</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #4a5568;'>Wybierz moduł, z którego chcesz skorzystać:</p>", unsafe_allow_html=True)
    st.write("")
    
    col_r1, col_r2 = st.columns(2)
    with col_r1:
        st.info("### 🖥️ Panel Koordynatora")
        st.write("Wyszukiwarka klientów, pełna historia zgłoszeń ze statusami, prywatne notatki, edycja danych, zdjęcia/wideo, lejek ofertowy.")
        if st.button("Wejdź jako Koordynator", use_container_width=True, type="primary"):
            st.session_state.user_role = "koordynator"
            st.rerun()

    with col_r2:
        st.success("### 📱 Panel Serwisanta")
        st.write("Zlecenia od 08.10.2026 w Twoim rejonie, rozpoczęcie pracy, zamknięcie z procedurą weryfikacji, pełna historia i wyszukiwarka klientów.")
        if st.button("Wejdź jako Serwisant (PIN)", use_container_width=True):
            st.session_state.user_role = "serwisant"
            st.rerun()
    st.stop()

# =========================================================================
# 1. MODUŁ KOORDYNATORA
# =========================================================================
if st.session_state.user_role == "koordynator":
    df = st.session_state.df

    st.sidebar.title("⚡ Koordynator")
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

    # Monitor pilnych ofert handlowych zgłoszonych z terenu przez serwisantów (w pełni bezpieczne sprawdzenie logiczne)
    is_wymaga_oferty = df['Wymaga oferty'].astype(str).str.lower().isin(['true', '1', 'tak'])
    is_pilne_etap = df['Etap oferty'].astype(str).str.contains('PILNE', case=False, na=False)
    pilne_oferty_df = df[is_wymaga_oferty & is_pilne_etap]

    if not pilne_oferty_df.empty:
        st.warning(f"🚨 **UWAGA: Masz {len(pilne_oferty_df)} zgłoszeń z terenu wymagających PILNEJ oferty handlowej (modyfikacja/dorobienie instalacji)!** Sprawdź zakładkę *Lejek Ofert Handlowych* poniżej.")

    tab_k1, tab_k2 = st.tabs(["🔍 Baza Klientów & Zgłoszenia", "📑 Lejek Ofert Handlowych"])

    # --- ZAKŁADKA 2: LEJEK OFERTOWY ---
    with tab_k2:
        st.subheader("📑 Zgłoszenia wymagające oferty handlowej")
        st.caption("Poniżej znajdują się zapytania ofertowe oraz zgłoszenia z terenu, gdzie serwisant stwierdził konieczność płatnej modyfikacji/dorobienia instalacji.")

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
                        st.error(f"🚨 **SERWISANT ({o_row.get('SERWISANT', 'Serwisant')}) ZGŁOSIŁ POTRZEBĘ OFERTY Z OBIEKTU!**\n\n**Co klient musi dołożyć / wycenić:** {o_row.get('Szczegóły oferty', 'Brak opisu')}\n\n**Zalecenia serwisanta dla klienta:** {o_row.get('Zalecenia dla klienta', 'Brak')}")

                    c_o1, c_o2 = st.columns(2)
                    with c_o1:
                        st.markdown(f"**Klient:** {o_klient} (📞 {o_tel})")
                        st.markdown(f"**Lokalizacja:** {o_miasto}, {o_row.get('Adres', '')}")
                        st.markdown(f"**Usterka bazowa:** {o_row.get('OPIS OD KLIENTA', '')}")
                        if o_row.get('Komentarz koordynatora', ''):
                            st.info(f"🔒 **Notatka wewnętrzna:** {o_row.get('Komentarz koordynatora')}")
                    with c_o2:
                        st.markdown(f"**Przypisany handlowiec:** `{o_email}`")
                        st.markdown(f"**Data wpisu:** {o_row.get('Data oferty', '')}")
                        st.markdown(f"**Aktualny etap:** **{o_etap}**")

                    wyswietl_multimedia(o_nr)

                    st.write("")
                    col_b1, col_b2 = st.columns(2)

                    if "Czeka" in str(o_etap) or "PILNE" in str(o_etap):
                        with col_b1:
                            if st.button("✅ Oferta wysłana do klienta", key=f"btn_off_sent_{o_nr}", type="primary"):
                                df.at[idx, 'Etap oferty'] = 'Oczekiwanie na akceptację klienta'
                                df.at[idx, 'STATUS ZGŁOSZENIA'] = 'Oczekiwanie na akceptację'
                                st.session_state.df = df
                                save_ticket_persistently(df.loc[idx].to_dict(), current_full_df=st.session_state.df)
                                st.success("Zmieniono etap na: Oczekiwanie na akceptację klienta")
                                st.rerun()

                    elif o_etap == "Oczekiwanie na akceptację klienta":
                        with col_b1:
                            if st.button("🚀 Klient zaakceptował → Przekaż do serwisu", key=f"btn_off_accept_{o_nr}", type="primary"):
                                df.at[idx, 'Etap oferty'] = 'Zaakceptowano (w serwisie)'
                                df.at[idx, 'STATUS ZGŁOSZENIA'] = 'w toku'
                                st.session_state.df = df
                                save_ticket_persistently(df.loc[idx].to_dict(), current_full_df=st.session_state.df)
                                st.success("✅ Oferta zaakceptowana! Zlecenie przekazane do serwisu.")
                                st.rerun()
                        with col_b2:
                            if st.button("❌ Klient odrzucił ofertę", key=f"btn_off_rej_{o_nr}"):
                                df.at[idx, 'Etap oferty'] = 'Odrzucona przez klienta'
                                df.at[idx, 'STATUS ZGŁOSZENIA'] = 'anulowane'
                                save_ticket_persistently(df.loc[idx].to_dict(), current_full_df=st.session_state.df)
                                st.warning("Zgłoszenie anulowane.")
                                st.rerun()

    # --- ZAKŁADKA 1: BAZA KLIENTÓW ---
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
                            "Data zgłoszenia": str(p_row.get('DATA ZGŁOSZENIA', ''))[:16],
                            "Klient": f"{p_row.get('Imię', '')} {p_row.get('Nazwisko', '')}",
                            "Telefon": p_row.get('Numer telefonu', ''),
                            "Miejscowość": p_row.get('Miejscowość', ''),
                            "Adres": p_row.get('Adres', ''),
                            "Instalacja": str(p_row.get('INSTALACJA', '')).upper(),
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
                            "Data zgłoszenia": str(p_row.get('DATA ZGŁOSZENIA', ''))[:16],
                            "Rejon": p_row.get('Rejon Polski', ''),
                            "Klient": f"{p_row.get('Imię', '')} {p_row.get('Nazwisko', '')}",
                            "Telefon": p_row.get('Numer telefonu', ''),
                            "Miejscowość": p_row.get('Miejscowość', ''),
                            "Instalacja": str(p_row.get('INSTALACJA', '')).upper(),
                            "Status": p_row.get('STATUS ZGŁOSZENIA', 'w toku'),
                            "Opis usterki": p_row.get('OPIS OD KLIENTA', '') or p_row.get('OPIS', ''),
                            "ID": p_row['Klient_ID']
                        })
                    t_all_df = pd.DataFrame(t_all_rows)
                    st.dataframe(t_all_df.drop(columns=['ID']), use_container_width=True, hide_index=True)

            filtered_df = df.copy()

        st.markdown("---")
        
        # NAGŁÓWEK WYSZUKIWANIA I PRZYCISK NOWEGO KLIENTA
        c_title, c_add_btn = st.columns([3, 1])
        with c_title:
            st.subheader("1. Baza Klientów (Wyszukaj lub Dodaj)")
        with c_add_btn:
            if st.button("➕ DODAJ NOWEGO KLIENTA", type="primary", use_container_width=True):
                st.session_state.show_new_client_modal = True

        # MODAL / FORMULARZ DODAWANIA NOWEGO KLIENTA
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
                            "Klient_ID": get_clean_client_id({'Numer telefonu': nc_tel, 'Imię': nc_imie, 'Nazwisko': nc_nazwisko})
                        }
                        
                        st.session_state.df = pd.concat([pd.DataFrame([new_client_row]), st.session_state.df], ignore_index=True)
                        save_ticket_persistently(new_client_row, current_full_df=st.session_state.df)
                        
                        st.session_state.selected_client_id = new_client_row['Klient_ID']
                        st.session_state.show_new_client_modal = False
                        st.success(f"✅ Klient {nc_imie} {nc_nazwisko} dodany do bazy i arkusza Excel!")
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
        st.write(f"Znaleziono pasujących klientów: **{len(clients_unique)}** (łącznie pasujących zgłoszeń: **{len(filtered_df)}**)")

        if len(clients_unique) > 0:
            st.markdown("##### Wybierz klienta, aby otworzyć jego kartę:")
            client_choices = []
            id_to_label = {}
            for idx, r in clients_unique.iterrows():
                name = f"{r.get('Imię', '')} {r.get('Nazwisko', '')}".strip() or "Brak nazwiska"
                tel = r.get('Numer telefonu', '')
                miasto = r.get('Miejscowość', '')
                ulica = r.get('Adres', '')
                cid = r['Klient_ID']
                lbl = f"{name} | 📞 {tel} | 📍 {miasto}, {ulica}"
                client_choices.append({'ID': cid, 'Etykieta': lbl})
                id_to_label[cid] = lbl

            current_choice_idx = 0
            if st.session_state.selected_client_id:
                for i, c in enumerate(client_choices):
                    if c['ID'] == st.session_state.selected_client_id:
                        current_choice_idx = i
                        break

            choice_labels = [c['Etykieta'] for c in client_choices]
            selected_label = st.selectbox(
                "Wybierz klienta z listy:",
                options=choice_labels,
                index=current_choice_idx,
                key="select_client_box"
            )
            
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
                    st.markdown(f"**Liczba zgłoszeń w bazie:** **{len(client_records)}**")

                st.write("")
                btn_c1, btn_c2, _ = st.columns([1.5, 1.5, 3])
                with btn_c1:
                    if st.button("➕ NOWE ZGŁOSZENIE", type="primary", use_container_width=True):
                        st.session_state.show_new_ticket_modal = True
                        st.session_state.show_edit_client_modal = False
                with btn_c2:
                    if st.button("✏️ EDYTUJ DANE KLIENTA", use_container_width=True):
                        st.session_state.show_edit_client_modal = not st.session_state.show_edit_client_modal
                        st.session_state.show_new_ticket_modal = False

                # FORMULARZ EDYCJI DANYCH KLIENTA
                if st.session_state.show_edit_client_modal:
                    st.info("✏️ Edycja danych tego klienta (zmiany zaktualizują bazę i arkusz Excel):")
                    with st.form("form_edytuj_klienta"):
                        ec_c1, ec_c2 = st.columns(2)
                        with ec_c1:
                            ec_imie = st.text_input("Imię:", value=first_record.get('Imię', ''))
                            ec_nazwisko = st.text_input("Nazwisko:*", value=first_record.get('Nazwisko', ''))
                            ec_tel = st.text_input("Numer telefonu:*", value=first_record.get('Numer telefonu', ''))
                            ec_email = st.text_input("E-mail:", value=first_record.get('E-mail', ''))
                        with ec_c2:
                            ec_kod = st.text_input("Kod pocztowy:", value=first_record.get('Kod pocztowy', ''))
                            ec_miasto = st.text_input("Miejscowość:", value=first_record.get('Miejscowość', ''))
                            ec_adres = st.text_input("Adres (ulica i numer):*", value=first_record.get('Adres', ''))
                            
                            all_reg_names = [r['nazwa'] for r in REGION_DEFINITIONS]
                            curr_reg = first_record.get('Rejon Polski', '')
                            reg_idx = all_reg_names.index(curr_reg) if curr_reg in all_reg_names else 0
                            ec_rejon = st.selectbox("Rejon Polski:", all_reg_names, index=reg_idx)

                        ec_sub, ec_canc = st.columns([1, 4])
                        with ec_sub:
                            ec_saved = st.form_submit_button("💾 ZAPISZ ZMIANY W DANYCH KLIENTA", type="primary")
                        with ec_canc:
                            ec_anuluj = st.form_submit_button("Anuluj")

                        if ec_saved:
                            if not ec_nazwisko.strip() or not ec_tel.strip():
                                st.error("❌ Nazwisko i numer telefonu są wymagane!")
                            else:
                                old_id = st.session_state.selected_client_id
                                new_id = get_clean_client_id({'Numer telefonu': ec_tel, 'Imię': ec_imie, 'Nazwisko': ec_nazwisko})
                                
                                mask_client_records = (st.session_state.df['Klient_ID'] == old_id)
                                st.session_state.df.loc[mask_client_records, 'Imię'] = ec_imie.strip()
                                st.session_state.df.loc[mask_client_records, 'Nazwisko'] = ec_nazwisko.strip()
                                st.session_state.df.loc[mask_client_records, 'Numer telefonu'] = ec_tel.strip()
                                st.session_state.df.loc[mask_client_records, 'E-mail'] = ec_email.strip()
                                st.session_state.df.loc[mask_client_records, 'Kod pocztowy'] = ec_kod.strip()
                                st.session_state.df.loc[mask_client_records, 'Miejscowość'] = ec_miasto.strip()
                                st.session_state.df.loc[mask_client_records, 'Adres'] = ec_adres.strip()
                                st.session_state.df.loc[mask_client_records, 'Rejon Polski'] = ec_rejon
                                st.session_state.df.loc[mask_client_records, 'Klient_ID'] = new_id

                                updated_rows = st.session_state.df[st.session_state.df['Klient_ID'] == new_id]
                                save_multiple_tickets_persistently(updated_rows, current_full_df=st.session_state.df)

                                st.session_state.selected_client_id = new_id
                                st.session_state.show_edit_client_modal = False
                                st.success("✅ Dane klienta zostały zaktualizowane w bazie i pliku Excel!")
                                st.rerun()

                        if ec_anuluj:
                            st.session_state.show_edit_client_modal = False
                            st.rerun()

                # PEŁNA LISTA ZGŁOSZEŃ I STATUSÓW KLIENTA
                st.markdown("### 📋 Historia i status zgłoszeń tego klienta:")
                
                summary_data = []
                for _, r_t in client_records.iterrows():
                    stat = str(r_t.get('STATUS ZGŁOSZENIA', '')).strip()
                    etap = str(r_t.get('Etap oferty', '')).strip()
                    
                    display_status = stat
                    if etap:
                        display_status = f"{stat} ({etap})" if stat != etap else etap
                    
                    serw_osoba = r_t.get('SERWISANT', '')
                    if not serw_osoba and r_t.get('Email handlowca', ''):
                        serw_osoba = f"Handlowiec: {r_t.get('Email handlowca')}"
                    elif not serw_osoba:
                        serw_osoba = "Nieprzypisany"

                    summary_data.append({
                        "Nr zgłoszenia": r_t.get('nr.', ''),
                        "Data zgłoszenia": str(r_t.get('DATA ZGŁOSZENIA', ''))[:16],
                        "Typ instalacji": str(r_t.get('INSTALACJA', '')).upper(),
                        "Status zgłoszenia": display_status or "w toku",
                        "Przypisany": serw_osoba,
                        "Opis usterki / zlecenia": r_t.get('OPIS OD KLIENTA', '') or r_t.get('OPIS', ''),
                        "Kwota": f"{r_t.get('Kwota', '')} zł" if str(r_t.get('Kwota', '')).strip() else "—"
                    })

                summary_df = pd.DataFrame(summary_data)
                st.dataframe(summary_df, use_container_width=True, hide_index=True)

                # FORMULARZ NOWEGO ZGŁOSZENIA Z PRYWATNĄ NOTATKĄ KOORDYNATORA
                if st.session_state.show_new_ticket_modal:
                    st.info("📝 Rejestracja nowego zgłoszenia:")
                    
                    next_id = get_next_ticket_id()
                    new_ticket_nr_preview = f"N{next_id}"

                    # 1. Zmiana daty
                    custom_date_enabled = st.checkbox("📅 Zmień datę i czas zgłoszenia (domyślnie: teraz)")
                    now_default = datetime.datetime.now()
                    if custom_date_enabled:
                        col_d1, col_d2 = st.columns(2)
                        with col_d1:
                            wybrana_data = st.date_input("Wybierz datę zgłoszenia:", value=now_default.date())
                        with col_d2:
                            wybrana_godzina = st.time_input("Wybierz godzinę:", value=now_default.time())
                        now_str = f"{wybrana_data} {wybrana_godzina.strftime('%H:%M')}"
                    else:
                        now_str = now_default.strftime("%Y-%m-%d %H:%M")

                    # 2. Załączanie zdjęć i filmików
                    st.markdown("📷 **Załącz zdjęcia lub filmiki:**")
                    uploaded_media = st.file_uploader(
                        "Przeciągnij zdjęcia lub filmiki (JPG, PNG, MP4, MOV...):",
                        type=['png', 'jpg', 'jpeg', 'webp', 'mp4', 'mov', 'avi'],
                        accept_multiple_files=True,
                        key="uploader_new_ticket"
                    )

                    with st.form("form_nowe_zgloszenie"):
                        f_col1, f_col2 = st.columns(2)
                        with f_col1:
                            st.text_input("Numer zgłoszenia (auto):", value=f"#{new_ticket_nr_preview}", disabled=True)
                            forma = st.radio("Forma zgłoszenia:", ["telefon", "SMS", "WhatsApp", "e-mail"], horizontal=True)
                            priorytet = st.selectbox("Priorytet:", ["Standardowy", "Pilny", "Awaria", "Niski"], index=0)
                            
                        with f_col2:
                            st.text_input("Data i godzina rejestracji:", value=now_str, disabled=True)
                            usterka_dotyczy = st.selectbox("Usterka dotyczy:", ["solar", "foto", "pompa", "magazyn"], index=0)
                            producent = st.text_input("Producent instalacji:", placeholder="np. Hewalex, Deye, Fronius, Sofar...")

                        opis_zgloszenia = st.text_area("Opis usterki / zapytania od klienta (widoczne dla serwisanta):", placeholder="Wpisz treść problemu...")

                        # PRYWATNA NOTATKA KOORDYNATORA (NIEWIDOCZNA DLA SERWISANTA)
                        st.markdown("🔒 **Wewnętrzny komentarz koordynatora (POUFNE — niewidoczne dla serwisanta):**")
                        komentarz_koordynatora = st.text_area(
                            "Notatka biura / koordynatora:",
                            placeholder="Wpisz uwagi biurowe, ustalenia cenowe, historię rozmów itp. (serwisant tego NIE zobaczy)...",
                            key="notatka_koordynatora_input"
                        )

                        st.markdown("---")
                        wymaga_oferty = st.checkbox("💼 Zgłoszenie wymaga oferty handlowej (przekaż do handlowca)")
                        email_handlowca = ""
                        if wymaga_oferty:
                            email_handlowca = st.text_input("Wpisz adres e-mail handlowca:", placeholder="np. handlowiec@twojafirma.pl")

                        c_save, c_cancel = st.columns([1, 4])
                        with c_save:
                            submitted = st.form_submit_button("💾 ZAPISZ ZGŁOSZENIE (NA STAŁE)", type="primary")
                        with c_cancel:
                            cancel = st.form_submit_button("Anuluj")

                        if submitted:
                            new_ticket_nr = f"N{next_id}"
                            etap_startowy = "Czeka na ofertę handlowca" if wymaga_oferty else ""
                            status_startowy = "Czeka na ofertę handlowca" if wymaga_oferty else "w toku"

                            if uploaded_media:
                                target_dir = os.path.join(ATTACHMENTS_DIR, str(new_ticket_nr))
                                os.makedirs(target_dir, exist_ok=True)
                                for file in uploaded_media:
                                    f_dest = os.path.join(target_dir, file.name)
                                    with open(f_dest, "wb") as f_out:
                                        f_out.write(file.getbuffer())

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
                                "SERWISANT": "",
                                "OPIS OD KLIENTA": opis_zgloszenia,
                                "OPIS": "",
                                "INFORMACJE DODATKOWE": f"Forma: {forma} | Priorytet: {priorytet} | Producent: {producent}",
                                "Kwota": "",
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
                                "Klient_ID": first_record['Klient_ID']
                            }

                            st.session_state.df = pd.concat([pd.DataFrame([new_row_dict]), st.session_state.df], ignore_index=True)
                            save_ticket_persistently(new_row_dict, current_full_df=st.session_state.df)

                            st.session_state.show_new_ticket_modal = False
                            st.success(f"✅ Zgłoszenie #{new_ticket_nr} zapisane w bazie i pliku Excel!")
                            st.rerun()

                        if cancel:
                            st.session_state.show_new_ticket_modal = False
                            st.rerun()

                # SZCZEGÓŁOWE ARCHIWUM Z PODGLĄDEM ZDJĘĆ/WIDEO I POUFNEJ NOTATKI
                st.markdown("#### 🔍 Szczegóły poszczególnych zleceń:")
                for i, (_, row) in enumerate(client_records.iterrows(), 1):
                    nr_t = str(row.get('nr.', 'Brak nr'))
                    with st.expander(f"Zgłoszenie nr: {nr_t} | Typ: {row.get('INSTALACJA', '')} | Data: {str(row.get('DATA ZGŁOSZENIA', ''))[:10]} | Status: {row.get('STATUS ZGŁOSZENIA', '')}", expanded=(i==1)):
                        c1, c2 = st.columns(2)
                        with c1:
                            st.markdown(f"**Data zgłoszenia:** {row.get('DATA ZGŁOSZENIA', '')}")
                            st.markdown(f"**Serwisant:** {row.get('SERWISANT', 'Nieprzypisany')}")
                            st.markdown(f"**Kwota:** {row.get('Kwota', '')} zł")
                        with c2:
                            st.markdown(f"**Data serwisu:** {row.get('DATA SERWISU', 'Brak')}")
                            st.markdown(f"**Płatność:** {row.get('Płatność ', '')}")
                            st.markdown(f"**Status:** {row.get('STATUS ZGŁOSZENIA', '')}")

                        st.markdown(f"**Opis od klienta:** {row.get('OPIS OD KLIENTA', 'Brak')}")
                        st.markdown(f"**Opis techniczny:** {row.get('OPIS', 'Brak')}")
                        
                        if row.get('Komentarz koordynatora', ''):
                            st.info(f"🔒 **Poufna notatka koordynatora:** {row.get('Komentarz koordynatora')}")
                        
                        if row.get('Komentarz serwisanta', ''):
                            st.write(f"🛠️ **Raport serwisanta:** {row.get('Komentarz serwisanta')}")
                        if row.get('Zużyte materiały', ''):
                            st.write(f"📦 **Zużyte materiały:** {row.get('Zużyte materiały')}")
                        if row.get('Zalecenia dla klienta', ''):
                            st.write(f"💡 **Zalecenia dla klienta:** {row.get('Zalecenia dla klienta')}")
                        if row.get('Szczegóły oferty', ''):
                            st.warning(f"💼 **Wymagana oferta (zgłoszona z terenu):** {row.get('Szczegóły oferty')}")
                        if row.get('Przyczyna nieukończenia', ''):
                            st.error(f"⚠️ **Przyczyna nieukończenia:** {row.get('Przyczyna nieukończenia')} | Braki: {row.get('Brakujące materiały', '')}")

                        wyswietl_multimedia(nr_t)

# =========================================================================
# 2. MODUŁ SERWISANTA (MOBILNY Z PIN-EM)
# =========================================================================
elif st.session_state.user_role == "serwisant":
    df = st.session_state.df

    if not st.session_state.tech_pin:
        col_w, _ = st.columns([1, 4])
        with col_w:
            if st.button("⬅️ Wróć do wyboru"):
                st.session_state.user_role = None
                st.rerun()

        st.markdown("<h2 style='text-align: center; margin-top: 1rem;'>🔒 Logowanie Serwisanta</h2>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center; color: #718096;'>Wprowadź swój kod PIN rejonu:</p>", unsafe_allow_html=True)

        with st.form("form_pin"):
            pin_input = st.text_input("KOD PIN (4 cyfry):", type="password", max_chars=4, placeholder="np. 1111")
            submit_pin = st.form_submit_button("ZALOGUJ DO MOJEGO REJONU", type="primary")

            if submit_pin:
                if pin_input in PIN_CONFIG:
                    st.session_state.tech_pin = pin_input
                    st.success("Zalogowano pomyślnie!")
                    st.rerun()
                else:
                    st.error("❌ Błędny PIN. Wprowadź poprawny kod.")

        with st.expander("ℹ️ Kody PIN dla prezentacji (ściągawka)"):
            st.write("• **1110** - Rejon 0 (Mazowsze - Leszek)")
            st.write("• **1111** - Rejon 2 (Lubelskie/Świętokrzyskie - Stanisław)")
            st.write("• **1112** - Rejon 6 (Wielkopolskie - Wojtek)")
            st.write("• **1113** - Rejon 8 (Pomorze/Kujawy - Mikołaj)")
            st.write("• **1114** - Rejon 9 (Łódzkie - Artur)")
            st.write("• **9999** - Wszystkie rejony")
        st.stop()

    cfg = PIN_CONFIG[st.session_state.tech_pin]
    rejon_kod = cfg["rejon_kod"]
    rejon_nazwa = cfg["nazwa_rejonu"]
    domyslny_tech = cfg["domyslny_serwisant"]

    col_h1, col_h2 = st.columns([3, 1])
    with col_h1:
        st.markdown(f"**📍 {rejon_nazwa}**")
        st.caption(f"Zalogowany serwisant: **{domyslny_tech}**")
    with col_h2:
        if st.button("Wyloguj"):
            st.session_state.tech_pin = None
            st.session_state.active_ticket_nr = None
            st.rerun()

    # -------------------------------------------------------------------------
    # JEŚLI KLIKNIĘTO ZAMKNIĘCIE ZLECENIA -> WYŚWIETL DEDYKOWANY PEŁNY FORMULARZ
    # -------------------------------------------------------------------------
    if st.session_state.active_ticket_nr:
        curr_rows = df[df['nr.'].astype(str) == str(st.session_state.active_ticket_nr)]
        
        if curr_rows.empty:
            st.error("Nie znaleziono wybranego zlecenia w bazie.")
            if st.button("⬅️ Wróć do listy zleceń"):
                st.session_state.active_ticket_nr = None
                st.rerun()
            st.stop()

        ticket = curr_rows.iloc[0]
        
        if st.button("⬅️ Wróć do listy zleceń bez zapisywania"):
            st.session_state.active_ticket_nr = None
            st.rerun()

        st.markdown(f"### 🛠️ Rozliczenie / Zamknięcie zlecenia #{ticket['nr.']}")
        st.caption(f"Klient: **{ticket['Imię']} {ticket['Nazwisko']}** | 📍 {ticket['Miejscowość']}, {ticket['Adres']}")

        poradzil_sobie = st.radio(
            "Czy poradziłeś sobie z tym zleceniem?*",
            ["TAK — Usterka została usunięta", "NIE — Zlecenie nieukończone / Wymaga powrotu"],
            key=f"tech_outcome_radio_{ticket['nr.']}"
        )

        is_success = "TAK" in poradzil_sobie

        klient_id_curr = ticket.get('Klient_ID', '')
        klient_ma_juz_tabliczki = (
            (df['Klient_ID'] == klient_id_curr) & 
            (df['Tabliczki_OK'].astype(str).str.lower().isin(['true', '1', 'tak']))
        ).any()

        with st.form(f"form_rozliczenie_{ticket['nr.']}"):
            if is_success:
                st.success("✅ **Zlecenie zakończone sukcesem:**")
                opis_prac = st.text_area(
                    "1. Opis tego, co zostało zrobione / wykonane czynności:*",
                    value=ticket.get('Komentarz serwisanta', '') or ticket.get('OPIS', ''),
                    placeholder="np. Odpowietrzono układ, wymieniono pompę obiegową..."
                )
                zuzyte_czesci = st.text_area(
                    "2. Lista wykorzystanych części i materiałów:",
                    value=ticket.get('Zużyte materiały', ''),
                    placeholder="np. Pompa Wilo, zawór zwrotny..."
                )
                zalecenia_klient = st.text_area(
                    "3. Zalecenia dla klienta (OBOWIĄZKOWE):*",
                    value=ticket.get('Zalecenia dla klienta', ''),
                    placeholder="np. Kontrola ciśnienia na manometrze za 48h..."
                )
                przyczyna_nieuk = ""
                brakujace_mat = ""
                status_domyslny_idx = 0
            else:
                st.error("⚠️ **Zlecenie nie zostało ukończone:**")
                przyczyna_nieuk = st.text_area(
                    "1. Dokładny opis przyczyny:*",
                    value=ticket.get('Przyczyna nieukończenia', '') or ticket.get('Komentarz serwisanta', ''),
                    placeholder="np. Uszkodzona elektronika inwertera..."
                )
                brakujace_mat = st.text_area(
                    "2. Lista brakujących części do zamówienia:*",
                    value=ticket.get('Brakujące materiały', ''),
                    placeholder="np. Nowa płyta główna inwertera..."
                )
                zalecenia_klient = st.text_area(
                    "3. Zalecenia dla klienta (OBOWIĄZKOWE):*",
                    value=ticket.get('Zalecenia dla klienta', ''),
                    placeholder="np. Wyłączono bezpiecznik główny..."
                )
                opis_prac = ""
                zuzyte_czesci = ""
                status_domyslny_idx = 1

            st.markdown("---")
            st.markdown("##### 📸 Zdjęcia z obiektu & Tabliczki znamionowe")

            if klient_ma_juz_tabliczki:
                st.success("✅ Tabliczki znamionowe tego klienta są już w bazie.")
                potwierdzenie_tabliczek = True
            else:
                st.warning("📌 Brak tabliczek znamionowych w bazie tego klienta. Wykonaj ich zdjęcia!")
                potwierdzenie_tabliczek = st.checkbox(
                    "✅ POTWIERDZAM: Wykonałem i załączam zdjęcia tabliczek znamionowych."
                )

            tech_media = st.file_uploader(
                "Załącz zdjęcia:",
                type=['png', 'jpg', 'jpeg', 'webp', 'mp4'],
                accept_multiple_files=True,
                key=f"tech_upload_{ticket['nr.']}"
            )

            st.markdown("---")
            potrzebna_oferta = st.checkbox(
                "⚠️ WYMAGANA OFERTA HANDLOWA DLA KLIENTA (dorobienie / modyfikacja)",
                value=bool(ticket.get('Wymaga oferty', False))
            )
            szczegoly_oferty = ""
            if potrzebna_oferta:
                szczegoly_oferty = st.text_area(
                    "Co klient musi dołożyć / co wycenić:*",
                    value=ticket.get('Szczegóły oferty', ''),
                    placeholder="np. Wycena dodatkowego bufora 200L..."
                )

            st.markdown("---")
            f_col1, f_col2 = st.columns(2)
            with f_col1:
                kwota_pobrana = st.text_input(
                    "Kwota pobrana (zł):",
                    value=ticket.get('Kwota przyjęta', '') or ticket.get('Kwota', '') or ("0" if not is_success else "")
                )
                przyjazdy = st.selectbox("Ilość przyjazdów:", ["1", "2", "3", "Więcej"], index=0)
            with f_col2:
                forma_plat = st.selectbox(
                    "Forma płatności:",
                    ["Gotówka", "Przelew", "Gwarancja (0 zł)", "Czeka na przelew"],
                    index=0
                )
                status_opcje = ["zakończone", "do ponownego przyjazdu (brak części)", "w toku (wymaga części)", "anulowane"]
                status_nowy = st.selectbox("Status zlecenia:", status_opcje, index=status_domyslny_idx)

            zapisz = st.form_submit_button("💾 ZAPISZ RAPORT I ZAMKNIJ", type="primary", use_container_width=True)

            if zapisz:
                if not zalecenia_klient.strip():
                    st.error("❌ Pole 'Zalecenia dla klienta' jest obowiązkowe!")
                elif not potwierdzenie_tabliczek:
                    st.error("❌ Musisz potwierdzić sprawdzenie tabliczek znamionowych!")
                elif is_success and not opis_prac.strip():
                    st.error("❌ Opis wykonanych czynności jest wymagany!")
                elif not is_success and (not przyczyna_nieuk.strip() or not brakujace_mat.strip()):
                    st.error("❌ Przyczyna oraz lista braków są wymagane!")
                elif potrzebna_oferta and not szczegoly_oferty.strip():
                    st.error("❌ Wpisz szczegóły potrzebnej oferty!")
                else:
                    idx = curr_rows.index[0]
                    if is_success:
                        df.at[idx, 'Komentarz serwisanta'] = opis_prac
                        df.at[idx, 'Zalecenia dla klienta'] = zalecenia_klient.strip()
                        df.at[idx, 'Zużyte materiały'] = zuzyte_czesci.strip()
                        df.at[idx, 'Przyczyna nieukończenia'] = ""
                        df.at[idx, 'Brakujące materiały'] = ""
                    else:
                        df.at[idx, 'Komentarz serwisanta'] = przyczyna_nieuk
                        df.at[idx, 'Przyczyna nieukończenia'] = przyczyna_nieuk.strip()
                        df.at[idx, 'Brakujące materiały'] = brakujace_mat.strip()
                        df.at[idx, 'Zalecenia dla klienta'] = zalecenia_klient.strip()
                        df.at[idx, 'Zużyte materiały'] = ""

                    if potwierdzenie_tabliczek or klient_ma_juz_tabliczki:
                        mask_c = (df['Klient_ID'] == klient_id_curr)
                        df.loc[mask_c, 'Tabliczki_OK'] = 'True'
                        df.at[idx, 'Tabliczki_OK'] = 'True'

                    if tech_media:
                        target_dir = os.path.join(ATTACHMENTS_DIR, str(ticket['nr.']))
                        os.makedirs(target_dir, exist_ok=True)
                        for file in tech_media:
                            with open(os.path.join(target_dir, file.name), "wb") as f_out:
                                f_out.write(file.getbuffer())

                    if potrzebna_oferta:
                        df.at[idx, 'Wymaga oferty'] = True
                        df.at[idx, 'Etap oferty'] = '🚨 PILNE: Zgłoszenie z serwisu (wymaga wyceny modyfikacji)'
                        df.at[idx, 'Szczegóły oferty'] = szczegoly_oferty.strip()
                        df.at[idx, 'Data oferty'] = datetime.date.today().strftime("%Y-%m-%d")

                    df.at[idx, 'Ilość przyjazdów'] = przyjazdy
                    df.at[idx, 'Kwota przyjęta'] = kwota_pobrana
                    df.at[idx, 'Płatność '] = forma_plat
                    df.at[idx, 'STATUS ZGŁOSZENIA'] = status_nowy
                    df.at[idx, 'DATA SERWISU'] = datetime.date.today().strftime("%Y-%m-%d")
                    df.at[idx, 'SERWISANT'] = domyslny_tech

                    st.session_state.df = df
                    save_ticket_persistently(df.loc[idx].to_dict(), current_full_df=st.session_state.df)

                    st.session_state.active_ticket_nr = None
                    st.success("✅ Raport został zapisany!")
                    st.rerun()

        # Izolacja formularza – zapobiega renderowaniu listy w tle
        st.stop()

    # Filtrowanie rejonu
    if rejon_kod == "ALL":
        tech_df = df.copy()
    else:
        mask_rejon = (
            df['Rejon Polski'].str.contains(f"Rejon {rejon_kod}", case=False, na=False) |
            df['Kod pocztowy'].str.startswith(f"{rejon_kod}")
        )
        tech_df = df[mask_rejon]

    # Ukrycie spraw w trakcie ofertowania handlowego
    ready_for_service_mask = (
        (tech_df['Etap oferty'] != 'Czeka na ofertę handlowca') &
        (tech_df['Etap oferty'] != 'Oczekiwanie na akceptację klienta')
    )
    tech_df = tech_df[ready_for_service_mask]

    # ODCIĘCIE HISTORII: Serwisant widzi w swoich bieżących kolejkach WYŁĄCZNIE sprawy od CUTOFF_DATE (08.10.2026)
    tech_recent_mask = tech_df['DATA ZGŁOSZENIA'].str[:10] >= CUTOFF_DATE
    tech_operational_df = tech_df[tech_recent_mask]

    is_done = tech_operational_df['STATUS ZGŁOSZENIA'].str.lower().str.contains('zakończ|anulow', na=False)
    otwarte_df = tech_operational_df[~is_done]
    zakonczone_df = tech_operational_df[is_done]

    # TRZY ZAKŁADKI: 1. Do zrobienia, 2. Zakończone, 3. Wyszukiwarka klienta i całej historii
    tab1, tab2, tab3 = st.tabs([
        f"⏳ Do zrobienia ({len(otwarte_df)})", 
        f"✅ Zakończone ({len(zakonczone_df)})",
        "🔍 Szukaj klienta (Cała historia)"
    ])

    def render_ticket_card(row, is_open=True):
        nr = row.get('nr.', '')
        typ = row.get('INSTALACJA', '').lower()
        imie_nazwisko = f"{row.get('Imię', '')} {row.get('Nazwisko', '')}".strip() or "Brak danych"
        tel = str(row.get('Numer telefonu', '')).strip()
        tel_clean = "".join([c for c in tel.split(',')[0].split('/')[0] if c.isdigit()])
        adres_full = f"{row.get('Kod pocztowy', '')} {row.get('Miejscowość', '')}, {row.get('Adres', '')}".strip(', ')
        map_query = urllib.parse.quote(f"{row.get('Adres', '')}, {row.get('Miejscowość', '')}")
        curr_status = str(row.get('STATUS ZGŁOSZENIA', 'w toku')).strip()

        badge_type = "badge-solar" if "solar" in typ else ("badge-foto" if "foto" in typ else ("badge-pompa" if "pompa" in typ else "badge-magazyn"))

        st.markdown(f"""
        <div class="mobile-card">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
                <span style="font-weight:bold; font-size:1.1rem; color:#1a202c;">#{nr}</span>
                <div>
                    <span class="badge {badge_type}">{typ.upper() or 'SERWIS'}</span>
                    <span class="badge" style="background:#e2e8f0; color:#2d3748;">{curr_status.upper()}</span>
                </div>
            </div>
            <div style="font-size:1.05rem; font-weight:700; color:#2d3748;">👤 {imie_nazwisko}</div>
            <div style="font-size:0.95rem; color:#4a5568; margin-top:3px;">📍 {adres_full}</div>
            <div style="margin-top:8px; padding:8px; background:#edf2f7; border-radius:8px; font-size:0.9rem; color:#2d3748;">
                <b>Usterka od klienta:</b> {row.get('OPIS OD KLIENTA', 'Brak opisu')}
            </div>
            <div style="display:flex; justify-content:space-between; margin-top:8px; font-size:0.85rem; color:#718096;">
                <span>💰 Wycena: <b>{row.get('Kwota', '0')} zł</b></span>
                <span>📅 {str(row.get('DATA ZGŁOSZENIA', ''))[:10]}</span>
            </div>
        </div>
        """, unsafe_allow_html=True)

        wyswietl_multimedia(nr)

        col_a, col_b = st.columns(2)
        with col_a:
            if tel_clean:
                st.markdown(f'<a href="tel:{tel_clean}" target="_blank" style="text-decoration:none;"><button style="width:100%; height:45px; border-radius:10px; background:#2b6cb0; color:white; border:none; font-weight:bold; cursor:pointer;">📞 Zadzwoń</button></a>', unsafe_allow_html=True)
        with col_b:
            if row.get('Miejscowość', ''):
                st.markdown(f'<a href="https://www.google.com/maps/search/?api=1&query={map_query}" target="_blank" style="text-decoration:none;"><button style="width:100%; height:45px; border-radius:10px; background:#38a169; color:white; border:none; font-weight:bold; cursor:pointer;">🗺️ Nawiguj</button></a>', unsafe_allow_html=True)

        if is_open:
            btn_col1, btn_col2 = st.columns(2)
            
            with btn_col1:
                if curr_status.lower() in ['w trakcie realizacji', 'rozpoczęte']:
                    st.button("⚙️ W TRAKCIE PRAC", key=f"btn_started_{nr}", disabled=True, use_container_width=True)
                else:
                    if st.button("▶️ Rozpocznij zlecenie", key=f"btn_start_{nr}", type="secondary", use_container_width=True):
                        match_idx = df[df['nr.'].astype(str) == str(nr)].index
                        if not match_idx.empty:
                            df.at[match_idx[0], 'STATUS ZGŁOSZENIA'] = 'w trakcie realizacji'
                            df.at[match_idx[0], 'SERWISANT'] = domyslny_tech
                            st.session_state.df = df
                            save_ticket_persistently(df.loc[match_idx[0]].to_dict(), current_full_df=st.session_state.df)
                            st.success(f"Rozpoczęto zlecenie #{nr}!")
                            st.rerun()

            with btn_col2:
                if st.button("✅ Rozlicz / Zakończ", key=f"btn_close_{nr}", type="primary", use_container_width=True):
                    st.session_state.active_ticket_nr = str(nr)
                    st.rerun()

        st.write("")

    with tab1:
        if otwarte_df.empty:
            st.info("🎉 Brak otwartych zleceń od 08.10 w Twoim rejonie!")
        else:
            q = st.text_input("🔍 Filtruj zlecenia (miejscowość / nazwisko):", key="search_open").strip().lower()
            view_df = otwarte_df
            if q:
                view_df = view_df[
                    view_df['Nazwisko'].str.lower().str.contains(q, na=False) |
                    view_df['Miejscowość'].str.lower().str.contains(q, na=False) |
                    view_df['Adres'].str.lower().str.contains(q, na=False)
                ]
            for _, row in view_df.head(30).iterrows():
                render_ticket_card(row, is_open=True)

    with tab2:
        if zakonczone_df.empty:
            st.info("Brak historii zleceń zakończonych od 08.10 w Twoim rejonie.")
        else:
            for _, row in zakonczone_df.head(20).iterrows():
                render_ticket_card(row, is_open=False)

    # ZAKŁADKA 3: WYSZUKIWARKA KLIENTA DLA SERWISANTA (PEŁNA HISTORIA Z ARCHIWUM)
    with tab3:
        st.subheader("🔍 Sprawdź historię klienta (np. pomoc telefoniczna)")
        st.caption("Wyszukaj dowolnego klienta w całej bazie (również zlecenia archiwalne sprzed 08.10.2026 r.).")

        tech_search = st.text_input("Wpisz nazwisko, numer telefonu lub miejscowość klienta:", key="tech_search_client").strip().lower()
        if tech_search:
            t_match = df[
                df['Nazwisko'].str.lower().str.contains(tech_search, na=False) |
                df['Numer telefonu'].str.lower().str.contains(tech_search, na=False) |
                df['Miejscowość'].str.lower().str.contains(tech_search, na=False) |
                df['Imię'].str.lower().str.contains(tech_search, na=False)
            ]
            t_clients = t_match.drop_duplicates(subset=['Klient_ID'])
            st.write(f"Znaleziono pasujących klientów: **{len(t_clients)}**")

            for _, c_row in t_clients.head(10).iterrows():
                c_name = f"{c_row.get('Imię', '')} {c_row.get('Nazwisko', '')}"
                c_id = c_row['Klient_ID']
                c_history = df[df['Klient_ID'] == c_id]

                with st.expander(f"👤 {c_name} | 📞 {c_row.get('Numer telefonu', '')} | 📍 {c_row.get('Miejscowość', '')}, {c_row.get('Adres', '')} (Zgłoszeń: {len(c_history)})"):
                    col_tc1, col_tc2 = st.columns(2)
                    with col_tc1:
                        st.write(f"**Telefon:** {c_row.get('Numer telefonu', '')}")
                        st.write(f"**Adres:** {c_row.get('Kod pocztowy', '')} {c_row.get('Miejscowość', '')}, {c_row.get('Adres', '')}")
                    with col_tc2:
                        st.write(f"**Rejon:** {c_row.get('Rejon Polski', '')}")
                        st.write(f"**E-mail:** {c_row.get('E-mail', 'Brak')}")

                    st.markdown("##### Pełna historia zgłoszeń tego klienta:")
                    for _, h_row in c_history.iterrows():
                        is_arch = "📦 ARCHIWUM (przed 08.10)" if str(h_row.get('DATA ZGŁOSZENIA', ''))[:10] < CUTOFF_DATE else "⚡ NOWE (po 08.10)"
                        st.markdown(f"""
                        <div style="background:#edf2f7; padding:10px; border-radius:8px; margin-bottom:8px; font-size:0.9rem;">
                            <b>#{h_row.get('nr.', '')}</b> | Typ: <b>{h_row.get('INSTALACJA', '').upper()}</b> | Data: {str(h_row.get('DATA ZGŁOSZENIA', ''))[:10]} | Status: <b>{h_row.get('STATUS ZGŁOSZENIA', '')}</b> | <i>{is_arch}</i><br>
                            <b>Usterka:</b> {h_row.get('OPIS OD KLIENTA', '') or h_row.get('OPIS', 'Brak opisu')}<br>
                            <b>Serwisant:</b> {h_row.get('SERWISANT', 'Nieprzypisany')} | <b>Zużyte części:</b> {h_row.get('Zużyte materiały', 'Brak')}
                        </div>
                        """, unsafe_allow_html=True)
