import streamlit as st
import pandas as pd
import datetime
import os

st.set_page_config(page_title="Panel Koordynatora Serwisu", layout="wide", initial_sidebar_state="expanded")

# Inicjalizacja bazy w sesji
@st.cache_data
def load_data():
    file_path = "Serwisowa posortowana.xlsx"
    if os.path.exists(file_path):
        try:
            df = pd.read_excel(file_path, sheet_name="Wszystkie posortowane")
        except Exception:
            df = pd.read_excel(file_path, sheet_name=0)
    else:
        df = pd.DataFrame(columns=[
            'nr.', 'INSTALACJA', 'DATA ZGŁOSZENIA', 'PRZYJMUJĄCY', 'Imię', 'Nazwisko', 
            'Numer telefonu', 'E-mail', 'Kod pocztowy', 'Miejscowość', 'Adres', 'Rejon Polski',
            'OPIS OD KLIENTA', 'OPIS', 'INFORMACJE DODATKOWE', 'SERWISANT', 'DATA SERWISU',
            'STATUS ZGŁOSZENIA', 'Kwota', 'Płatność '
        ])
    
    text_cols = ['Imię', 'Nazwisko', 'Numer telefonu', 'Miejscowość', 'Adres', 'Rejon Polski', 'STATUS ZGŁOSZENIA']
    for c in text_cols:
        if c in df.columns:
            df[c] = df[c].fillna('').astype(str)
            
    df['Klient_ID'] = df.apply(
        lambda r: f"{r.get('Numer telefonu', '')}_{r.get('Imię', '')}_{r.get('Nazwisko', '')}".strip('_'), 
        axis=1
    )
    return df

if 'df' not in st.session_state:
    st.session_state.df = load_data()

if 'nowe_zgloszenia' not in st.session_state:
    st.session_state.nowe_zgloszenia = []

if 'next_ticket_id' not in st.session_state:
    st.session_state.next_ticket_id = 1

if 'selected_client_id' not in st.session_state:
    st.session_state.selected_client_id = None

if 'show_new_ticket_modal' not in st.session_state:
    st.session_state.show_new_ticket_modal = False

df = st.session_state.df

# --- PASEK BOCZNY: REGIONY POCZTOWE ---
st.sidebar.title("📍 Regiony Pocztowe")
st.sidebar.caption("Sortowanie malejąco wg oczekujących spraw")

pending_mask = ~df['STATUS ZGŁOSZENIA'].str.lower().str.contains('zakończ', na=False)
pending_df = df[pending_mask]
counts_per_region = pending_df['Rejon Polski'].value_counts()

all_regions = ["Wszystkie rejony"] + list(counts_per_region.index)
region_options = {"Wszystkie rejony": f"🌐 Wszystkie rejony ({len(pending_df)} oczekujących)"}
for r in counts_per_region.index:
    region_options[r] = f"{r} ({counts_per_region[r]})"

selected_region_raw = st.sidebar.radio(
    "Wybierz rejon do filtrowania:",
    options=all_regions,
    format_func=lambda x: region_options.get(x, x)
)

# --- PANEL GŁÓWNY ---
st.title("🛠️ Panel Koordynatora Serwisu")
st.write("Szybkie wyszukiwanie klientów, wgląd w historię interwencji i rejestracja nowych zgłoszeń.")

# 1. Wyszukiwarka
st.subheader("1. Wyszukaj klienta")
col_s1, col_s2 = st.columns([3, 1])
with col_s1:
    search_query = st.text_input(
        "Wpisz nazwisko, imię, numer telefonu, miejscowość lub adres:",
        placeholder="np. Kowalski, 602 123 456, Radom, Topolowa...",
    ).strip().lower()

with col_s2:
    if st.button("🔄 Wyczyść filtry i wybór", use_container_width=True):
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
st.write(f"Znaleziono klientów: **{len(clients_unique)}** (łącznie zgłoszeń w filtrze: **{len(filtered_df)}**)")

if len(clients_unique) > 0:
    st.markdown("##### Wyniki wyszukiwania — wybierz klienta z listy poniżej:")
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

    selected_label = st.selectbox(
        "Wybierz klienta z listy:",
        options=choice_labels,
        index=selected_idx
    )
    
    for c in client_choices:
        if c['Etykieta'] == selected_label:
            st.session_state.selected_client_id = c['ID']
            break

# 2. Widok wybranego klienta
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
        st.markdown(f"**🌐 Rejon Polski:** {first_record.get('Rejon Polski', '')}")
        st.markdown(f"**Liczba interwencji w bazie:** {len(client_records)}")

    st.write("")
    if st.button("➕ DODAJ NOWE ZGŁOSZENIE DLA TEGO KLIENTA", type="primary"):
        st.session_state.show_new_ticket_modal = True

    # 3. Formularz nowego zgłoszenia
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
                st.text_input("Data i godzina rejestracji (auto):", value=now_str, disabled=True)
                usterka_dotyczy = st.selectbox("Usterka dotyczy:", ["solar", "foto", "pompa", "magazyn"], index=0)
                producent = st.text_input("Producent instalacji:", placeholder="np. Hewalex, Deye, Fronius, Sofar...")

            opis_zgloszenia = st.text_area("Opis usterki / zgłoszenia od klienta:", placeholder="Opisz problem...")
            
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
                st.success("✅ Pomyślnie zapisano nowe zgłoszenie!")
                st.rerun()

            if cancel:
                st.session_state.show_new_ticket_modal = False
                st.rerun()

    # Archiwum klienta
    st.markdown("#### 📂 Informacje Archiwalne:")
    for i, (_, row) in enumerate(client_records.iterrows(), 1):
        with st.expander(f"Zgłoszenie archiwalne nr: {row.get('nr.', 'Brak nr')} | Typ: {row.get('INSTALACJA', '')} | Status: {row.get('STATUS ZGŁOSZENIA', '')}", expanded=(i==1)):
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
