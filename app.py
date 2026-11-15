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

st.title("🌿 PROMAG OZE - System Serwisowy")

LOCAL_EXCEL_FILE = "Serwisowa posortowana.xlsx"

if not os.path.exists(LOCAL_EXCEL_FILE):
    df_init = pd.DataFrame(columns=[
        "ID_Zgloszenia", "Klient", "Telefon", "Adres", "Status", "Kwota_Przewidywana"
    ])
    df_init.to_excel(LOCAL_EXCEL_FILE, index=False)

df = pd.read_excel(LOCAL_EXCEL_FILE)

st.subheader("📋 Lista zgłoszeń w systemie")
st.dataframe(df, use_container_width=True)

with st.form("add_ticket_form"):
    st.subheader("➕ Dodaj nowe zgłoszenie")
    c1, c2 = st.columns(2)
    with c1:
        klient_in = st.text_input("Imię i nazwisko klienta:")
        telefon_in = st.text_input("Telefon:")
    with c2:
        adres_in = st.text_input("Adres / Miasto:")
        kwota_in = st.number_input("Szacowana kwota (zł):", 0.0, 50000.0, 300.0, 50.0)
    
    submitted = st.form_submit_button("Zapisz zgłoszenie", type="primary")
    if submitted:
        if not klient_in.strip():
            st.error("Podaj imię i nazwisko klienta!")
        else:
            new_id = f"ZGL-{len(df) + 1001}"
            new_row = pd.DataFrame([{
                "ID_Zgloszenia": new_id,
                "Klient": klient_in.strip(),
                "Telefon": telefon_in.strip(),
                "Adres": adres_in.strip(),
                "Status": "Nowe",
                "Kwota_Przewidywana": float(kwota_in)
            }])
            df = pd.concat([df, new_row], ignore_index=True)
            df.to_excel(LOCAL_EXCEL_FILE, index=False)
            st.success(f"✅ Dodano zgłoszenie {new_id} pomyślnie!")
            st.rerun()
