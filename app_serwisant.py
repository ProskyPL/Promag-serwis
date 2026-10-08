import streamlit as st
import pandas as pd
import datetime
import os
import urllib.parse

st.set_page_config(
    page_title="Panel Serwisanta",
    page_icon="🔧",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Styl dopasowany pod smartfony
st.markdown("""
<style>
    .block-container {
        padding-top: 1rem;
        padding-bottom: 2rem;
        padding-left: 0.8rem;
        padding-right: 0.8rem;
    }
    .stButton>button {
        width: 100%;
        border-radius: 12px;
        height: 3.2rem;
        font-size: 1.05rem;
        font-weight: 600;
        margin-top: 0.3rem;
        margin-bottom: 0.3rem;
    }
    .mobile-card {
        background-color: #f8f9fa;
        border: 1px solid #e2e8f0;
        border-radius: 14px;
        padding: 14px;
        margin-bottom: 12px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .badge {
        display: inline-block;
        padding: 4px 10px;
        border-radius: 12px;
        font-size: 0.75rem;
        font-weight: 700;
        text-transform: uppercase;
    }
    .badge-solar { background-color: #feebc8; color: #c05621; }
    .badge-foto { background-color: #fefcbf; color: #975a16; }
    .badge-pompa { background-color: #bee3f8; color: #2b6cb0; }
    .badge-magazyn { background-color: #e9d8fd; color: #6b46c1; }
</style>
""", unsafe_allow_html=True)

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
                   'SERWISANT', 'OPIS OD KLIENTA', 'OPIS', 'INFORMACJE DODATKOWE', 'Kwota', 'Płatność ']
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
        
    return df

if 'df_serv' not in st.session_state:
    st.session_state.df_serv = load_data()

if 'active_ticket_nr' not in st.session_state:
    st.session_state.active_ticket_nr = None

df = st.session_state.df_serv

# --- 1. WYBÓR SERWISANTA ---
st.markdown("### 📱 Serwis Mobilny")

known_techs = ["Wszyscy", "Wojtek", "Leszek", "Piotr", "Stanisław", "Artur", "Mikołaj", "Hubert", "Piasta", "Dominik", "Alex"]
wybrany_serwisant = st.selectbox("👷 Zalogowany jako serwisant:", known_techs, index=1)

if wybrany_serwisant != "Wszyscy":
    tech_df = df[df['SERWISANT'].str.lower().str.contains(wybrany_serwisant.lower(), na=False)]
else:
    tech_df = df.copy()

is_done = tech_df['STATUS ZGŁOSZENIA'].str.lower().str.contains('zakończ', na=False)
otwarte_df = tech_df[~is_done]
zakonczone_df = tech_df[is_done]

tab1, tab2 = st.tabs([f"⏳ Do zrobienia ({len(otwarte_df)})", f"✅ Zakończone ({len(zakonczone_df)})"])

# --- FORMULARZ RAPORTU / ROZLICZENIA ---
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
                "Komentarz serwisanta (diagnoza / naprawa):",
                value=ticket.get('Komentarz serwisanta', '') or ticket.get('OPIS', ''),
                placeholder="np. Odpowietrzono układ, wymieniono naczynie przeponowe..."
            )
            
            materialy = st.text_input(
                "Zużyte materiały:",
                value=ticket.get('Zużyte materiały', ''),
                placeholder="np. Pompka Wilo, 10L glikol..."
            )
            
            czynnosci = st.text_input(
                "Czynności dodatkowe:",
                value=ticket.get('Czynności dodatkowe', ''),
                placeholder="np. Odpowietrzenie grzejników..."
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
                    "Płatność:",
                    ["Gotówka", "Przelew", "Gwarancja (0 zł)", "Czeka na przelew"],
                    index=0
                )
                
            status_nowy = st.selectbox(
                "Status zlecenia:",
                ["zakończone", "w toku (wymaga części)", "do ponownego przyjazdu"],
                index=0
            )

            c_save, c_back = st.columns(2)
            with c_save:
                zapisz = st.form_submit_button("💾 ZAPISZ RAPORT", type="primary")
            with c_back:
                anuluj = st.form_submit_button("Wróć")

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
                
                st.session_state.df_serv = df
                st.session_state.active_ticket_nr = None
                st.success("✅ Raport zapisany!")
                st.rerun()

            if anuluj:
                st.session_state.active_ticket_nr = None
                st.rerun()

# --- WIDOK POJEDYNCZEJ KARTY ---
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
        if st.button(f"📝 Uzupełnij raport / Zakończ #{nr}", key=f"btn_close_{nr}", type="primary"):
            st.session_state.active_ticket_nr = nr
            st.rerun()
            
    st.write("")

# --- ZAKŁADKI ---
with tab1:
    if otwarte_df.empty:
        st.info("🎉 Brak otwartych zleceń dla wybranego serwisanta!")
    else:
        q = st.text_input("🔍 Szukaj zlecenia (miejscowość / nazwisko):", key="search_open").strip().lower()
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
        st.info("Brak historii zleceń zakończonych.")
    else:
        for _, row in zakonczone_df.head(20).iterrows():
            render_ticket_card(row, is_open=False)
