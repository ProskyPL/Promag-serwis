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

# Style responsywne pod komputery i telefony
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

# Kody PIN i przypisane rejony
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

@st.cache_data
def load_data():
    file_path = "Serwisowa posortowana.xlsx"
    if os.path.exists(file_path):
        try:
            df = pd.read_excel(file_path, sheet_name="Wszystkie posortowane")
        except Exception:
            df = pd.read_excel(file_path, sheet_name=0)
    else:
        df = pd.DataFrame()

    cols_to_str = ['nr.', 'INSTALACJA', 'PRZYJMUJĄCY', 'Imię', 'Nazwisko', 'Numer telefonu', 
                   'Kod pocztowy', 'Miejscowość', 'Adres', 'Rejon Polski', 'STATUS ZGŁOSZENIA', 
                   'SERWISANT', 'OPIS OD KLIENTA', 'OPIS', 'INFORMACJE DODATKOWE', 'Kwota', 'Płatność ', 'E-mail']
    for c in cols_to_str:
        if c in df.columns:
            df[c] = df[c].fillna('').astype(str)

    if 'Zużyte materiały' not in df.columns:
        df['Zużyte materiały'] = ''
    if 'Czynności dodatkowe' not in df.columns:
        df['Czynności dodatkowe'] = ''
    if 'Ilość przyjazdów' not in df.columns:
        df['Ilość przyjazdów'] = '1'
    if 'Kwota przyjęta' not in df.columns:
        df['Kwota przyjęta'] = df['Kwota']
    if 'Komentarz serwisanta' not in df.columns:
        df['Komentarz serwisanta'] = ''

    df['Klient_ID'] = df.apply(
        lambda r: f"{r.get('Numer telefonu', '')}_{r.get('Imię', '')}_{r.get('Nazwisko', '')}".strip('_'), 
        axis=1
    )
    return df

if 'df' not in st.session_state:
    st.session_state.df = load_data()

if 'user_role' not in st.session_state:
    st.session_state.user_role = None

if 'tech_pin' not in st.session_state:
    st.session_state.tech_pin = None

if 'nowe_zgloszenia' not in st.session_state:
    st.session_state.nowe_zgloszenia = []

if 'next_ticket_id' not in st.session_state:
    st.session_state.next_ticket_id = 1

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
        st.write("Wyszukiwarka klientów, pełna historia zleceń, podział na regiony, rejestracja nowych zgłoszeń.")
        if st.button("Wejdź jako Koordynator", use_container_width=True, type="primary"):
            st.session_state.user_role = "koordynator"
            st.rerun()

    with col_r2:
        st.success("### 📱 Panel Serwisanta")
        st.write("Widok mobilny pod telefon: zadania w terenie, nawigacja Google Maps, szybki telefon, rozliczenie (dostęp na PIN).")
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
    st.sidebar.subheader("📍 Regiony Pocztowe")
    st.sidebar.caption("Sortowanie malejąco wg oczekujących")

    pending_mask = ~df['STATUS ZGŁOSZENIA'].str.lower().str.contains('zakończ', na=False)
    pending_df = df[pending_mask]
    counts_per_region = pending_df['Rejon Polski'].value_counts()

    all_regions = ["Wszystkie rejony"] + list(counts_per_region.index)
    region_options = {"Wszystkie rejony": f"🌐 Wszystkie ({len(pending_df)} oczekujących)"}
    for r in counts_per_region.index:
        region_options[r] = f"{r} ({counts_per_region[r]})"

    selected_region_raw = st.sidebar.radio(
        "Filtruj wg rejonu:",
        options=all_regions,
        format_func=lambda x: region_options.get(x, x)
    )

    st.title("🖥️ Pulpit Koordynatora")
    st.subheader("1. Wyszukaj klienta w bazie")
    col_s1, col_s2 = st.columns([3, 1])
    with col_s1:
        search_query = st.text_input(
            "Wpisz nazwisko, imię, numer telefonu, miejscowość lub adres:",
            placeholder="np. Kowalski, 602 123 456, Radom, Topolowa...",
        ).strip().lower()

    with col_s2:
        if st.button("🔄 Resetuj filtry", use_container_width=True):
            st.session_state.selected_client_id = None
            st.rerun()

    filtered_df = df.copy()
    if selected_region_raw != "Wszystkie rejony":
        filtered_df = filtered_df[filtered_df['Rejon Polski'] == selected_region_raw]

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
    st.write(f"Znaleziono klientów: **{len(clients_unique)}** (łącznie zgłoszeń: **{len(filtered_df)}**)")

    if len(clients_unique) > 0:
        st.markdown("##### Wyniki — wybierz klienta z listy poniżej:")
        client_choices = []
        for idx, r in clients_unique.head(50).iterrows():
            name = f"{r.get('Imię', '')} {r.get('Nazwisko', '')}".strip() or "Brak nazwiska"
            tel = r.get('Numer telefonu', '')
            miasto = r.get('Miejscowość', '')
            ulica = r.get('Adres', '')
            client_choices.append({
                'ID': r['Klient_ID'],
                'Etykieta': f"{name} | 📞 {tel} | 📍 {miasto}, {ulica} ({r.get('Rejon Polski', '')})"
            })

        choice_labels = [c['Etykieta'] for c in client_choices]
        selected_idx = 0
        if st.session_state.selected_client_id:
            for i, c in enumerate(client_choices):
                if c['ID'] == st.session_state.selected_client_id:
                    selected_idx = i
                    break

        selected_label = st.selectbox("Lista klientów:", options=choice_labels, index=selected_idx)
        for c in client_choices:
            if c['Etykieta'] == selected_label:
                st.session_state.selected_client_id = c['ID']
                break

    if st.session_state.selected_client_id:
        client_records = df[df['Klient_ID'] == st.session_state.selected_client_id]
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
            st.markdown(f"**Historia wizyt:** {len(client_records)}")

        st.write("")
        if st.button("➕ DODAJ NOWE ZGŁOSZENIE DLA TEGO KLIENTA", type="primary"):
            st.session_state.show_new_ticket_modal = True

        if st.session_state.show_new_ticket_modal:
            st.info("📝 Rejestracja nowego zgłoszenia:")
            with st.form("form_nowe_zgloszenie"):
                f_col1, f_col2 = st.columns(2)
                with f_col1:
                    st.text_input("Numer zgłoszenia (auto):", value=f"#{st.session_state.next_ticket_id}", disabled=True)
                    forma = st.radio("Forma zgłoszenia:", ["telefon", "SMS", "WhatsApp", "e-mail"], horizontal=True)
                    priorytet = st.selectbox("Priorytet:", ["Standardowy", "Pilny", "Awaria", "Niski"], index=0)
                with f_col2:
                    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
                    st.text_input("Data i godzina:", value=now_str, disabled=True)
                    usterka_dotyczy = st.selectbox("Usterka dotyczy:", ["solar", "foto", "pompa", "magazyn"], index=0)
                    producent = st.text_input("Producent instalacji:", placeholder="np. Hewalex, Deye, Fronius, Sofar...")

                opis_zgloszenia = st.text_area("Opis usterki zgłaszanej przez klienta:")
                
                c_save, c_cancel = st.columns([1, 4])
                with c_save:
                    submitted = st.form_submit_button("💾 Zapisz zgłoszenie", type="primary")
                with c_cancel:
                    cancel = st.form_submit_button("Anuluj")

                if submitted:
                    new_entry = {
                        "Numer": f"#{st.session_state.next_ticket_id}",
                        "Data i godzina": now_str,
                        "Klient": f"{first_record.get('Imię', '')} {first_record.get('Nazwisko', '')}",
                        "Telefon": first_record.get('Numer telefonu', ''),
                        "Miejscowość": first_record.get('Miejscowość', ''),
                        "Rejon": first_record.get('Rejon Polski', ''),
                        "Forma": forma,
                        "Priorytet": priorytet,
                        "Dotyczy": usterka_dotyczy,
                        "Producent": producent,
                        "Opis": opis_zgloszenia,
                        "Status": "Nowe (oczekujące)"
                    }
                    st.session_state.nowe_zgloszenia.append(new_entry)
                    st.session_state.next_ticket_id += 1
                    st.session_state.show_new_ticket_modal = False
                    st.success("✅ Pomyślnie zarejestrowano zgłoszenie!")
                    st.rerun()

                if cancel:
                    st.session_state.show_new_ticket_modal = False
                    st.rerun()

        st.markdown("#### 📂 Informacje Archiwalne:")
        for i, (_, row) in enumerate(client_records.iterrows(), 1):
            with st.expander(f"Zgłoszenie nr: {row.get('nr.', 'Brak nr')} | Typ: {row.get('INSTALACJA', '')} | Data: {str(row.get('DATA ZGŁOSZENIA', ''))[:10]} | Status: {row.get('STATUS ZGŁOSZENIA', '')}", expanded=(i==1)):
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

    is_done = tech_df['STATUS ZGŁOSZENIA'].str.lower().str.contains('zakończ', na=False)
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
