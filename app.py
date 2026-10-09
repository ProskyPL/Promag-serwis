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

    # Wczytanie trwale zapisanych nowych zgłoszeń z dysku
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

# Trwały zapis do pliku CSV na serwerze
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

if 'df' not in st.session_state:
    st.session_state.df = load_all_data()

def get_next_ticket_id():
    existing_nrs = st.session_state.df['nr.'].astype(str).tolist()
    max_id = 0
    for n in existing_nrs:
        clean = n.replace('N', '').replace('#', '').strip()
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
        st.write("Wyszukiwarka klientów, pełna historia zgłoszeń ze statusami, załączanie zdjęć/wideo, lejek ofertowy handlowców.")
        if st.button("Wejdź jako Koordynator", use_container_width=True, type="primary"):
            st.session_state.user_role = "koordynator"
            st.rerun()

    with col_r2:
        st.success("### 📱 Panel Serwisanta")
        st.write("Widok mobilny pod telefon: zatwierdzone zlecenia w Twoim rejonie, nawigacja Google Maps, zdjęcia/filmy usterki, rozliczenie (dostęp na PIN).")
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
        st.rerun()

    st.sidebar.markdown("---")
    st.sidebar.subheader("📍 Rejony Pocztowe")
    st.sidebar.caption(f"Sortowanie wg zgłoszeń od **{CUTOFF_DATE}**:")

    # Filtrujemy zgłoszenia bieżące od sztywnej daty odcięcia CUTOFF_DATE
    is_active_since_cutoff = df['DATA ZGŁOSZENIA'].str[:10] >= CUTOFF_DATE

    region_stats = []
    for r_def in REGION_DEFINITIONS:
        r_kod = r_def['kod']
        mask_r = (
            df['Rejon Polski'].str.contains(f"Rejon {r_kod}", case=False, na=False) |
            df['Kod pocztowy'].str.startswith(r_kod)
        )
        total_in_r = mask_r.sum()
        # Zgłoszenia od daty odcięcia w tym rejonie
        active_in_r = (mask_r & is_active_since_cutoff).sum()
        # W tym oczekujące (niezakończone) od daty odcięcia
        pending_in_r = (mask_r & is_active_since_cutoff & ~df['STATUS ZGŁOSZENIA'].str.lower().str.contains('zakończ|anulow', na=False)).sum()
        
        region_stats.append({
            'kod': r_kod,
            'krotka': r_def['krotka'],
            'pelna': r_def['nazwa'],
            'active': active_in_r,
            'pending': pending_in_r,
            'total': total_in_r
        })

    # Sortowanie rejonów MALEJĄCO wg liczby spraw od daty odcięcia
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

    tab_k1, tab_k2 = st.tabs(["🔍 Baza Klientów & Zgłoszenia", "📑 Lejek Ofert Handlowych"])

    # --- ZAKŁADKA 2: LEJEK OFERTOWY ---
    with tab_k2:
        st.subheader("📑 Zgłoszenia wymagające oferty handlowej")
        st.caption("Sprawy w tym lejku NIE są widoczne u serwisanta, dopóki klient nie zaakceptuje oferty.")

        oferty_df = df[df['Wymaga oferty'] == True]

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

                with st.expander(f"Zgłoszenie #{o_nr} | Klient: {o_klient} ({o_miasto}) | Status: {o_etap}", expanded=True):
                    c_o1, c_o2 = st.columns(2)
                    with c_o1:
                        st.markdown(f"**Klient:** {o_klient} (📞 {o_tel})")
                        st.markdown(f"**Lokalizacja:** {o_miasto}, {o_row.get('Adres', '')}")
                        st.markdown(f"**Usterka / Zapytanie:** {o_row.get('OPIS OD KLIENTA', '')}")
                    with c_o2:
                        st.markdown(f"**Handlowiec:** `{o_email}`")
                        st.markdown(f"**Data wpisu:** {o_row.get('Data oferty', '')}")
                        st.markdown(f"**Aktualny etap:** **{o_etap}**")

                    wyswietl_multimedia(o_nr)

                    st.write("")
                    col_b1, col_b2 = st.columns(2)

                    if o_etap == "Czeka na ofertę handlowca":
                        with col_b1:
                            if st.button("✅ Oferta wysłana do klienta", key=f"btn_off_sent_{o_nr}", type="primary"):
                                df.at[idx, 'Etap oferty'] = 'Oczekiwanie na akceptację klienta'
                                df.at[idx, 'STATUS ZGŁOSZENIA'] = 'Oczekiwanie na akceptację'
                                save_ticket_persistently(df.loc[idx].to_dict())
                                st.session_state.df = df
                                st.success("Zmieniono etap na: Oczekiwanie na akceptację klienta")
                                st.rerun()

                    elif o_etap == "Oczekiwanie na akceptację klienta":
                        with col_b1:
                            if st.button("🚀 Klient zaakceptował → Przekaż do serwisu", key=f"btn_off_accept_{o_nr}", type="primary"):
                                df.at[idx, 'Etap oferty'] = 'Zaakceptowano (w serwisie)'
                                df.at[idx, 'STATUS ZGŁOSZENIA'] = 'w toku'
                                save_ticket_persistently(df.loc[idx].to_dict())
                                st.session_state.df = df
                                st.success("✅ Oferta zaakceptowana! Zlecenie przekazane do serwisu.")
                                st.rerun()
                        with col_b2:
                            if st.button("❌ Klient odrzucił ofertę", key=f"btn_off_rej_{o_nr}"):
                                df.at[idx, 'Etap oferty'] = 'Odrzucona przez klienta'
                                df.at[idx, 'STATUS ZGŁOSZENIA'] = 'anulowane'
                                save_ticket_persistently(df.loc[idx].to_dict())
                                st.session_state.df = df
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
            
            # Pokaż zgłoszenia w tym rejonie od CUTOFF_DATE
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
        st.subheader("1. Wyszukaj klienta w bazie")
        col_s1, col_s2 = st.columns([3, 1])
        with col_s1:
            search_query = st.text_input(
                "Wpisz nazwisko, imię, numer telefonu, miejscowość lub adres:",
                placeholder="np. Kowalski, 602 123 456, Radom, Topolowa...",
            ).strip().lower()

        with col_s2:
            if st.button("🔄 Resetuj wybór", use_container_width=True):
                st.session_state.selected_client_id = None
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
            for idx, r in clients_unique.head(50).iterrows():
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
                if st.button("➕ DODAJ NOWE ZGŁOSZENIE DLA TEGO KLIENTA", type="primary"):
                    st.session_state.show_new_ticket_modal = True

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

                # FORMULARZ NOWEGO ZGŁOSZENIA Z TRWAŁYM ZAPISEM
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

                        opis_zgloszenia = st.text_area("Opis usterki / zapytania od klienta:", placeholder="Wpisz treść problemu...")

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

                            # Zapis plików multimedialnych
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
                                "Wymaga oferty": str(wymaga_oferty),
                                "Email handlowca": email_handlowca,
                                "Etap oferty": etap_startowy,
                                "Data oferty": now_str[:10],
                                "Klient_ID": first_record['Klient_ID']
                            }

                            # FIZYCZNY ZAPIS DO PLIKU CSV
                            save_ticket_persistently(new_row_dict)

                            # Aktualizacja sesji
                            st.session_state.df = pd.concat([pd.DataFrame([new_row_dict]), st.session_state.df], ignore_index=True)
                            st.session_state.show_new_ticket_modal = False
                            st.success(f"✅ Zgłoszenie #{new_ticket_nr} zapisane trwale na dysku!")
                            st.rerun()

                        if cancel:
                            st.session_state.show_new_ticket_modal = False
                            st.rerun()

                # SZCZEGÓŁOWE ARCHIWUM Z PODGLĄDEM ZDJĘĆ/WIDEO
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
                        st.markdown(f"**Informacje dodatkowe:** {row.get('INFORMACJE DODATKOWE', 'Brak')}")

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
        st.caption(f"Serwisant: **{domyslny_tech}**")
    with col_h2:
        if st.button("Wyloguj"):
            st.session_state.tech_pin = None
            st.session_state.active_ticket_nr = None
            st.rerun()

    if rejon_kod == "ALL":
        tech_df = df.copy()
    else:
        mask_rejon = (
            df['Rejon Polski'].str.contains(f"Rejon {rejon_kod}", case=False, na=False) |
            df['Kod pocztowy'].str.startswith(f"{rejon_kod}")
        )
        tech_df = df[mask_rejon]

    ready_for_service_mask = (
        (tech_df['Etap oferty'] != 'Czeka na ofertę handlowca') &
        (tech_df['Etap oferty'] != 'Oczekiwanie na akceptację klienta')
    )
    tech_df = tech_df[ready_for_service_mask]

    is_done = tech_df['STATUS ZGŁOSZENIA'].str.lower().str.contains('zakończ|anulow', na=False)
    otwarte_df = tech_df[~is_done]
    zakonczone_df = tech_df[is_done]

    tab1, tab2 = st.tabs([f"⏳ Do zrobienia ({len(otwarte_df)})", f"✅ Zakończone ({len(zakonczone_df)})"])

    if st.session_state.active_ticket_nr:
        curr_rows = df[df['nr.'] == str(st.session_state.active_ticket_nr)]
        if not curr_rows.empty:
            ticket = curr_rows.iloc[0]
            st.markdown("---")
            st.markdown(f"#### 🛠️ Raport zlecenia #{ticket['nr.']}")
            st.caption(f"Klient: {ticket['Imię']} {ticket['Nazwisko']} | {ticket['Miejscowość']}, {ticket['Adres']}")

            with st.form("form_rozliczenie_serwisant"):
                st.markdown("##### 1. Co zostało zrobione?")
                komentarz_serw = st.text_area(
                    "Komentarz serwisanta (diagnoza / co naprawiono):",
                    value=ticket.get('Komentarz serwisanta', '') or ticket.get('OPIS', ''),
                    placeholder="np. Odpowietrzono układ, wymieniono naczynie przeponowe..."
                )
                materialy = st.text_input(
                    "Zużyte materiały:",
                    value=ticket.get('Zużyte materiały', ''),
                    placeholder="np. Pompka Wilo, 10L glikol, naczynie 18L..."
                )
                czynnosci = st.text_input(
                    "Czynności dodatkowe:",
                    value=ticket.get('Czynności dodatkowe', ''),
                    placeholder="np. Przegląd kotłowni, odpowietrzenie..."
                )
                przyjazdy = st.selectbox("Ilość przyjazdów:", ["1", "2", "3", "Więcej"], index=0)

                st.markdown("##### 2. Rozliczenie i status")
                f_col1, f_col2 = st.columns(2)
                with f_col1:
                    kwota_pobrana = st.text_input(
                        "Kwota przyjęta (zł):",
                        value=ticket.get('Kwota przyjęta', '') or ticket.get('Kwota', '')
                    )
                with f_col2:
                    forma_plat = st.selectbox(
                        "Forma płatności:",
                        ["Gotówka", "Przelew", "Gwarancja (0 zł)", "Czeka na przelew"],
                        index=0
                    )

                status_nowy = st.selectbox(
                    "Nowy status:",
                    ["zakończone", "w toku (wymaga części)", "do ponownego przyjazdu"],
                    index=0
                )

                c_save, c_back = st.columns(2)
                with c_save:
                    zapisz = st.form_submit_button("💾 ZAPISZ RAPORT", type="primary")
                with c_back:
                    anuluj = st.form_submit_button("Wróć bez zmian")

                if zapisz:
                    idx = curr_rows.index[0]
                    df.at[idx, 'Komentarz serwisanta'] = komentarz_serw
                    df.at[idx, 'Zużyte materiały'] = materialy
                    df.at[idx, 'Czynności dodatkowe'] = czynnosci
                    df.at[idx, 'Ilość przyjazdów'] = przyjazdy
                    df.at[idx, 'Kwota przyjęta'] = kwota_pobrana
                    df.at[idx, 'Płatność '] = forma_plat
                    df.at[idx, 'STATUS ZGŁOSZENIA'] = status_nowy
                    df.at[idx, 'DATA SERWISU'] = datetime.date.today().strftime("%Y-%m-%d")
                    df.at[idx, 'SERWISANT'] = domyslny_tech

                    # TRWAŁY ZAPIS RAPORTU
                    save_ticket_persistently(df.loc[idx].to_dict())

                    st.session_state.df = df
                    st.session_state.active_ticket_nr = None
                    st.success("✅ Raport zapisany!")
                    st.rerun()

                if anuluj:
                    st.session_state.active_ticket_nr = None
                    st.rerun()

    def render_ticket_card(row, is_open=True):
        nr = row.get('nr.', '')
        typ = row.get('INSTALACJA', '').lower()
        imie_nazwisko = f"{row.get('Imię', '')} {row.get('Nazwisko', '')}".strip() or "Brak danych"
        tel = str(row.get('Numer telefonu', '')).strip()
        tel_clean = "".join([c for c in tel.split(',')[0].split('/')[0] if c.isdigit()])
        adres_full = f"{row.get('Kod pocztowy', '')} {row.get('Miejscowość', '')}, {row.get('Adres', '')}".strip(', ')
        map_query = urllib.parse.quote(f"{row.get('Adres', '')}, {row.get('Miejscowość', '')}")

        badge_type = "badge-solar" if "solar" in typ else ("badge-foto" if "foto" in typ else ("badge-pompa" if "pompa" in typ else "badge-magazyn"))

        st.markdown(f"""
        <div class="mobile-card">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
                <span style="font-weight:bold; font-size:1.1rem; color:#1a202c;">#{nr}</span>
                <span class="badge {badge_type}">{typ.upper() or 'SERWIS'}</span>
            </div>
            <div style="font-size:1.05rem; font-weight:700; color:#2d3748;">👤 {imie_nazwisko}</div>
            <div style="font-size:0.95rem; color:#4a5568; margin-top:3px;">📍 {adres_full}</div>
            <div style="margin-top:8px; padding:8px; background:#edf2f7; border-radius:8px; font-size:0.9rem; color:#2d3748;">
                <b>Usterka:</b> {row.get('OPIS OD KLIENTA', 'Brak opisu')}
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
            if st.button(f"📝 Uzupełnij raport / Zakończ #{nr}", key=f"btn_m_{nr}", type="primary"):
                st.session_state.active_ticket_nr = nr
                st.rerun()

        st.write("")

    with tab1:
        if otwarte_df.empty:
            st.info("🎉 Brak otwartych zleceń w Twoim rejonie!")
        else:
            q = st.text_input("🔍 Szukaj w rejonie (miejscowość / nazwisko):", key="search_open").strip().lower()
            view_df = otwarte_df
            if q:
                view_df = view_df[
                    view_df['Nazwisko'].str.lower().str.contains(q, na=False) |
                    view_df['Miejscowość'].str.lower().str.contains(q, na=False) |
                    view_df['Adres'].str.lower().str.contains(q, na=False)
                ]
            for _, row in view_df.head(25).iterrows():
                render_ticket_card(row, is_open=True)

    with tab2:
        if zakonczone_df.empty:
            st.info("Brak historii zleceń zakończonych w tym rejonie.")
        else:
            for _, row in zakonczone_df.head(20).iterrows():
                render_ticket_card(row, is_open=False)
