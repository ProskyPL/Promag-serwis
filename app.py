# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import datetime
import os

st.set_page_config(
    page_title="PROMAG OZE - Serwis",
    page_icon="🌿",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    .stButton > button { width: 100%; min-height: 48px; font-size: 15px !important; font-weight: 700 !important; border-radius: 10px !important; margin-bottom: 6px; }
    .promag-header { background: linear-gradient(135deg, #0a2540 0%, #10375c 100%); color: #ffffff; padding: 16px 22px; border-radius: 14px; border-bottom: 4px solid #16a34a; margin-bottom: 14px; }
    .kpi-card { background: linear-gradient(135deg, #0a2540 0%, #154360 100%); color: #ffffff; padding: 16px; border-radius: 14px; border-left: 6px solid #22c55e; margin-bottom: 10px; }
    .install-badge { background-color: #f0f9ff; border: 1px solid #bae6fd; border-left: 5px solid #0284c7; color: #0c4a6e; padding: 12px; border-radius: 10px; margin-top: 8px; margin-bottom: 10px; font-size: 14px; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="promag-header">
    <span style="font-size:24px; font-weight:800;">🌿 PROMAG OZE - System Serwisowy</span><br>
    <span style="font-size:13px; color:#86efac;">Fotowoltaika • Pompy Ciepła • Kolektory Słoneczne • Magazyny Energii</span>
</div>
""", unsafe_allow_html=True)

LOCAL_EXCEL_FILE = "Serwisowa posortowana.xlsx"

REQUIRED_COLS = [
    "ID_Zgloszenia", "Klient", "Telefon", "Kod_Pocztowy", "Miasto", "Adres",
    "Typ_Instalacji", "Opis_Zgloszenia", "Status", "Serwisant", "Kwota_Przewidywana"
]

def load_data():
    if not os.path.exists(LOCAL_EXCEL_FILE):
        df_init = pd.DataFrame(columns=REQUIRED_COLS)
        try: df_init.to_excel(LOCAL_EXCEL_FILE, index=False)
        except Exception: pass
        return df_init
    try:
        df = pd.read_excel(LOCAL_EXCEL_FILE)
        for c in REQUIRED_COLS:
            if c not in df.columns: df[c] = ""
        return df[REQUIRED_COLS]
    except Exception:
        return pd.DataFrame(columns=REQUIRED_COLS)

def save_data(df):
    st.session_state["df_tickets"] = df.copy()
    try:
        df.to_excel(LOCAL_EXCEL_FILE, index=False)
    except Exception:
        pass

if "df_tickets" not in st.session_state:
    st.session_state["df_tickets"] = load_data()

df = st.session_state["df_tickets"]

# Proste logowanie PIN (Koordynator: 1234, Serwisant: 4829, Zarząd: 9999)
PIN_KOORDYNATOR = "1234"
PIN_SERWISANT = "4829"
PIN_ZARZAD = "9999"

if "logged_role" not in st.session_state:
    st.session_state["logged_role"] = None
    st.session_state["logged_user"] = None

if st.session_state["logged_role"] is None:
    col_l1, col_l2 = st.columns([1, 1.2])
    with col_l1:
        with st.container(border=True):
            st.subheader("🔐 Logowanie do Systemu PIN")
            pin_input = st.text_input("Wpisz kod PIN:", type="password", placeholder="np. 1234 lub 4829")
            if st.button("🚀 Wejdź do aplikacji", type="primary"):
                p_clean = pin_input.strip()
                if p_clean == PIN_KOORDYNATOR:
                    st.session_state["logged_role"] = "koordynator"
                    st.session_state["logged_user"] = "Koordynator"
                    st.rerun()
                elif p_clean == PIN_ZARZAD:
                    st.session_state["logged_role"] = "zarzad"
                    st.session_state["logged_user"] = "Zarząd"
                    st.rerun()
                elif p_clean == PIN_SERWISANT:
                    st.session_state["logged_role"] = "serwisant"
                    st.session_state["logged_user"] = "Leszek"
                    st.rerun()
                else:
                    st.error("❌ Nieznany kod PIN. Spróbuj: 1234 (koordynator), 4829 (serwisant Leszek), 9999 (zarząd).")
    with col_l2:
        with st.container(border=True):
            st.subheader("💡 Informacje o systemie")
            st.markdown("""
            * **Koordynator (PIN: `1234`)**: Przyjmowanie zgłoszeń, planowanie tras, oddelegowania.
            * **Serwisant Leszek (PIN: `4829`)**: Mobilny panel obsługi wizyt, części, zdjęcia, podpis.
            * **Zarząd (PIN: `9999`)**: Podsumowania finansowe i prowizje.
            """)
    st.stop()

# Pasek górny po zalogowaniu
top1, top2, top3 = st.columns([2, 1, 0.5])
with top1:
    st.markdown(f"**Zalogowany użytkownik:** `{st.session_state['logged_user']}` ({st.session_state['logged_role'].capitalize()})")
with top3:
    if st.button("🚪 Wyloguj"):
        st.session_state["logged_role"] = None
        st.session_state["logged_user"] = None
        st.rerun()

st.divider()

role = st.session_state["logged_role"]

if role == "koordynator":
    st.subheader("➕ Przyjęcie Nowego Zgłoszenia Serwisowego")
    with st.container(border=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            k_name = st.text_input("Imię i nazwisko klienta:")
            k_tel = st.text_input("Numer telefonu:")
        with c2:
            k_kod = st.text_input("Kod pocztowy (np. 05-500):")
            k_miasto = st.text_input("Miejscowość:")
        with c3:
            k_adres = st.text_input("Ulica i nr domu:")
            k_kwota = st.number_input("Szacowana kwota (zł):", 0.0, 50000.0, 350.0, 50.0)
        
        k_opis = st.text_area("Opis usterki / zakres prac dla serwisanta:")

        if st.button("🌿 ZAPISZ ZGŁOSZENIE W BAZIE", type="primary"):
            if not k_name.strip():
                st.error("Podaj imię i nazwisko klienta!")
            else:
                new_id = f"ZGL-{len(df) + 1001}"
                new_row = {
                    "ID_Zgloszenia": new_id,
                    "Klient": k_name.strip(),
                    "Telefon": k_tel.strip(),
                    "Kod_Pocztowy": k_kod.strip(),
                    "Miasto": k_miasto.strip(),
                    "Adres": k_adres.strip(),
                    "Typ_Instalacji": "Pompa Ciepła / PV",
                    "Opis_Zgloszenia": k_opis.strip(),
                    "Status": "Nowe",
                    "Serwisant": "Leszek",
                    "Kwota_Przewidywana": float(k_kwota)
                }
                df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                save_data(df)
                st.success(f"✅ Zapisano pomyślnie zgłoszenie {new_id} dla klienta {k_name}!")
                st.rerun()

    st.divider()
    st.subheader("📋 Lista Wszystkich Zgłoszeń Serwisowych")
    st.dataframe(df, use_container_width=True, hide_index=True)

elif role == "serwisant":
    st.subheader(f"🛠️ Panel Mobilny Serwisanta – {st.session_state['logged_user']}")
    my_tickets = df[df["Serwisant"] == st.session_state["logged_user"]]
    if my_tickets.empty:
        st.info("Brak przypisanych zleceń na Twoim profilu.")
    else:
        for idx, row in my_tickets.iterrows():
            with st.container(border=True):
                st.markdown(f"### {row['ID_Zgloszenia']} | {row['Klient']} | Tel: `{row['Telefon']}`")
                st.markdown(f"📍 **Adres:** `{row['Kod_Pocztowy']} {row['Miasto']}, {row['Adres']}`")
                st.markdown(f"📝 **Zadanie:** {row['Opis_Zgloszenia']}")
                st.markdown(f"💰 **Wycena:** {row['Kwota_Przewidywana']} zł | **Status:** `{row['Status']}`")
                
                if row['Status'] != "Zakończone":
                    if st.button(f"✅ Zakończ wizytę i rozlicz ({row['ID_Zgloszenia']})", key=f"finish_{row['ID_Zgloszenia']}"):
                        df.loc[idx, "Status"] = "Zakończone"
                        save_data(df)
                        st.success("Wizyta została zamknięta!")
                        st.rerun()

elif role == "zarzad":
    st.subheader("📊 Panel Zarządu – Podsumowanie Finansowe")
    total_rev = float(df["Kwota_Przewidywana"].sum())
    total_count = len(df)
    closed_count = len(df[df["Status"] == "Zakończone"])
    
    m1, m2, m3 = st.columns(3)
    m1.metric("Wszystkich zgłoszeń", total_count)
    m2.metric("Zamknięte wizyty", closed_count)
    m3.metric("Szacowany obrót", f"{total_rev:,.0f} zł")
    
    st.divider()
    st.dataframe(df, use_container_width=True, hide_index=True)
