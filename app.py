import streamlit as st
import pandas as pd
import numpy as np
import datetime
import os
import json
import base64
import math
import re
from io import BytesIO
from PIL import Image
from streamlit_drawable_canvas import st_canvas
import gspread
from google.oauth2.service_account import Credentials

# ==========================================
# 1. KONFIGURACJA STRONY I RESPONSYWNY CSS
# ==========================================
st.set_page_config(
    page_title="PROMAG OZE - System Serwisowy",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    .stButton > button {
        width: 100%;
        min-height: 50px;
        font-size: 15px !important;
        font-weight: 600 !important;
        border-radius: 10px !important;
        margin-bottom: 6px;
    }
    .kpi-card {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        color: #ffffff;
        padding: 15px;
        border-radius: 12px;
        border-left: 6px solid #10b981;
        margin-bottom: 10px;
    }
    .kpi-card-warn { border-left: 6px solid #f59e0b; }
    .kpi-title { font-size: 12px; text-transform: uppercase; color: #94a3b8; }
    .kpi-value { font-size: 22px; font-weight: 800; color: #f8fafc; }
    .kpi-sub { font-size: 13px; color: #38bdf8; margin-top: 4px; }
    .day-card {
        background-color: #1e293b;
        color: #f8fafc;
        border-radius: 12px;
        padding: 15px;
        margin-bottom: 10px;
        border-left: 6px solid #38bdf8;
    }
    .day-card-today {
        border-left: 6px solid #10b981;
        background: linear-gradient(135deg, #064e3b 0%, #0f172a 100%);
    }
    .client-box {
        background-color: #f8fafc;
        border: 2px solid #cbd5e1;
        border-radius: 12px;
        padding: 18px;
        margin-top: 12px;
        margin-bottom: 12px;
        color: #0f172a;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. STAŁE, CENNIK I BAZA WIEDZY (DEYE, GOODWE, ARGOM)
# ==========================================
GOOGLE_SHEET_NAME = "Serwisowa posortowana"
LOCAL_EXCEL_FILE = "Serwisowa posortowana.xlsx"
PROFILES_FILE = "profiles.json"
LOGO_FILE = "logo.png"

PIN_KOORDYNATOR = "1234"
PIN_ZARZAD = "9999"

DEFAULT_TECHS = {
    "Leszek": {
        "pin": "4829",
        "pelne_dane": "Leszek - Serwisant OZE",
        "telefon": "",
        "umowa": "Umowa o pracę",
        "wynagrodzenie_bazowe": 6500.0,
        "bhp_data": "2026-01-15",
        "bhp_protokol": "BHP/01/2026 - Ważne",
        "medycyna_data": "2026-01-10",
        "medycyna_zakres": "Praca pow. 3m, SEP E+D, prawo jazdy kat. B",
        "urlopy_l4": [],
        "notatki_poufne": "Doświadczony w pompach ciepła i falownikach hybrydowych."
    },
    "Mikołaj": {
        "pin": "7315",
        "pelne_dane": "Mikołaj - Serwisant OZE",
        "telefon": "",
        "umowa": "B2B",
        "wynagrodzenie_bazowe": 7000.0,
        "bhp_data": "2026-02-01",
        "bhp_protokol": "BHP/02/2026 - Ważne",
        "medycyna_data": "2026-01-25",
        "medycyna_zakres": "Praca pow. 3m, SEP E+D, F-gazy",
        "urlopy_l4": [],
        "notatki_poufne": ""
    },
    "Wojtek": {
        "pin": "5942",
        "pelne_dane": "Wojtek - Serwisant OZE",
        "telefon": "",
        "umowa": "Umowa o pracę",
        "wynagrodzenie_bazowe": 6000.0,
        "bhp_data": "2026-03-01",
        "bhp_protokol": "BHP/03/2026 - Ważne",
        "medycyna_data": "2026-02-20",
        "medycyna_zakres": "Praca na wysokości, SEP E, kat. B",
        "urlopy_l4": [],
        "notatki_poufne": ""
    },
    "Serwisant 4": {
        "pin": "8164",
        "pelne_dane": "Wakat (Serwisant 4) - zmień w panelu Kadry",
        "telefon": "",
        "umowa": "Umowa zlecenie",
        "wynagrodzenie_bazowe": 6000.0,
        "bhp_data": "2026-01-01",
        "bhp_protokol": "Brak",
        "medycyna_data": "2026-01-01",
        "medycyna_zakres": "Do uzupełnienia",
        "urlopy_l4": [],
        "notatki_poufne": ""
    }
}

COMMISSION_TIER_1 = 15000.0
COMMISSION_RATE_1 = 0.03
COMMISSION_TIER_2 = 30000.0
COMMISSION_RATE_2 = 0.04

CENNIK_BAZOWY = {
    "Przegląd/Konserwacja PC (Powietrze)": 650.0,
    "Przegląd/Konserwacja PC (Grunt)": 850.0,
    "Przegląd instalacji PV (<10 kWp)": 550.0,
    "Przegląd instalacji PV (>10 kWp)": 800.0,
    "Diagnostyka / Awaria Pompy Ciepła": 750.0,
    "Diagnostyka / Awaria Falownika PV": 600.0,
    "Wymiana glikolu + odpowietrzenie układu": 1120.0,
    "Odpowietrzenie układu solarnego / CO": 840.0,
    "Czyszczenie filtrów i separatora magnetycznego": 400.0,
    "Pomiary elektryczne + Protokół": 650.0,
    "Serwis Gwarancyjny (Rozliczenie wewnętrzne)": 0.0
}

DIAGNOSTYKA_BAZA = {
    "Falownik Deye (Hybrydowy / Stringowy)": {
        "F08 - GFDI_Relay_Failure (Błąd przekaźnika uziemienia / N-PE)": (
            "CO TO ZNACZY: Falownik wykrył problem z uziemieniem obudowy lub przebicie między N a PE.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Wyłącz falownik (AC i DC). Weź miernik uniwersalny (V AC) i zmierz napięcie między szyną N a PE w rozdzielnicy.\n"
            "2. Jeśli napięcie N-PE > 5V-10V: słabe uziemienie budynku lub brak podziału PEN. Sprawdź przewód PE od falownika do bednarki.\n"
            "3. W menu Deye (Advanced Function -> kod 7777) sprawdź opcję 'Signal ISLAND MODE' (gdy podpięty jest obwód Backup/Load)."
        ),
        "F13 - Working_Mode_Change (Reset trybu pracy)": (
            "CO TO ZNACZY: Zmiana parametrów sieci lub trybu pracy.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Odczekaj 3 minuty – falownik zazwyczaj sam odliczy czas i wystartuje.\n"
            "2. Jeśli wisi na stałe: wyłącz AC, przełącz hebel DC na OFF, wyłącz baterię. Odczekaj 2 minuty aż zgaśnie ekran. Włącz w kolejności: Bateria -> DC -> AC."
        ),
        "F18 - Tz_Ac_OverCurr_Fault (Zwarcie lub przetężenie po stronie AC / LOAD)": (
            "CO TO ZNACZY: Nagły skok prądu na wyjściu AC lub w obwodzie awaryjnym LOAD.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Wyłącz bezpiecznik obwodu LOAD (wyjście awaryjne) oraz odepnij wyjście awaryjne LOAD (Backup).\n"
            "2. Uruchom falownik na samym DC i baterii (bez obciążenia domu na LOAD).\n"
            "3. Jeśli ruszy bez błędu F18: w instalacji domowej podpiętej pod LOAD jest zwarcie lub za duże urządzenie (np. pompa ciepła).\n"
            "4. Jeśli błąd F18 wyskakuje przy odpiętych wtyczkach AC i LOAD: uszkodzony mostek IGBT – zgłoś demontaż na gwarancję."
        ),
        "F20 - Tz_Dc_OverCurr_Fault (Przetężenie po stronie DC)": (
            "CO TO ZNACZY: Prąd z paneli PV lub baterii przekroczył limit.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Wyłącz hebel DC, rozepnij wtyczki MC4 wchodzące do falownika.\n"
            "2. Zmierz miernikiem napięcie każdego stringu (Uoc) oraz prąd zwarciowy (Isc).\n"
            "3. Sprawdź czy pod jedno MPPT nie podpięto równolegle zbyt wielu stringów."
        ),
        "F22 - Tz_EmergStop_Fault (Aktywowany wyłącznik PPOŻ / Brak zwory RSD)": (
            "CO TO ZNACZY: Otwarty obwód wyłącznika awaryjnego na płycie sterującej.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Odkręć dolną pokrywę komunikacyjną falownika Deye.\n"
            "2. Znajdź kostkę RSD / Emergency Stop / Dry Contact.\n"
            "3. Jeśli klient nie ma przycisku PPOŻ: wypadła lub poluzowała się fabryczna zworka z drutu łącząca dwa piny RSD. Dokręć ją płaskim śrubokrętem.\n"
            "4. Jeśli jest przycisk PPOŻ: sprawdź miernikiem na ciągłość obwodu (piszczyk), czy przycisk nie jest wciśnięty."
        ),
        "F23 / F24 - Tz_GFCI_OC / DC_Insulation_Fault (Prąd upływu / Błąd izolacji DC)": (
            "CO TO ZNACZY: Przebicie prądu z paneli PV lub kabli solarnych przez wilgoć do uziemienia (konstrukcji dachu).\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Wyłącz AC i DC. Wypnij wszystkie stringi z falownika.\n"
            "2. Zmierz miernikiem napięcie DC między Plusem stringu a uziemieniem (PE) oraz Minusem a PE.\n"
            "3. Jeśli napięcie do ziemi jest stałe i wysokie (np. 200V-400V) i nie spada: masz zwarcie kabla lub złącza MC4 do dachu!\n"
            "4. Wypnij wkładki ochronnika przepięć (SPD) DC w skrzynce – jeśli błąd zniknie, wymień zawilgoconą wkładkę SPD. Jeśli nie – szukaj na dachu złącza MC4 leżącego w wodzie."
        ),
        "F29 - Parallel_Comm_Fault (Błąd komunikacji równoległej Master/Slave)": (
            "CO TO ZNACZY: W menu włączono pracę kaskadową wielu falowników, a wisi tylko jeden.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Wejdź w: Advanced Function (Ustawienia zaawansowane) -> kod 7777.\n"
            "2. Odznacz ptaszek przy 'Parallel', 'Master' oraz 'Slave'. Zatwierdź OK."
        ),
        "F35 - AC_NoUtility_Fault (Brak sieci energetycznej na zaciskach GRID)": (
            "CO TO ZNACZY: Do kostki GRID w falowniku nie dociera napięcie 3-fazowe.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Zmierz miernikiem napięcia na śrubach kostki GRID: L1-N, L2-N, L3-N (ma być 210-250V) oraz międzyfazowo (400V).\n"
            "2. Jeśli brakuje fazy: sprawdź bezpiecznik przedlicznikowy i nadprądowy w rozdzielnicy.\n"
            "3. Jeśli napięcia są obecne, a nadal jest F35: zamień miejscami przewody L1 i L2 (błąd kolejności wirowania faz)."
        ),
        "F56 / F58 - DC_VoltLow / BMS_Comm_Fault (Rozładowana bateria / Brak komunikacji BMS)": (
            "CO TO ZNACZY: BMS odciął baterię lub brak komunikacji po kablu CAN/RS485.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Zmierz napięcie V DC na śrubach BAT+ i BAT-. Jeśli 0V: włącz hebel na baterii i przytrzymaj przycisk RST/SW przez 5 sekund (wybudzenie Black Start).\n"
            "2. Sprawdź kabel RJ45: musi być wpięty w port BMS w falowniku i port CAN w baterii.\n"
            "3. W menu Battery Setting sprawdź 'Lithium Mode' (zazwyczaj 00 dla CAN). Jeśli brak kabla, przełącz tymczasowo na 'Use Batt V', aby podładować baterię."
        )
    },
    "Falownik GoodWe (SDT G2 / MS / DNS / ET Plus Hybryda)": {
        "Isolation Failure (Kod 16 - Błąd rezystancji izolacji DC < 1 MOhm)": (
            "CO TO ZNACZY: Zawilgocone złącze MC4, przetarty kabel na dachu lub uszkodzony ochronnik SPD DC.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Wyłącz AC i DC, wypnij wszystkie stringi. Podpinaj po jednym stringu i sprawdź, który wywołuje błąd.\n"
            "2. Zmierz rezystancję izolacji Riso miernikiem na 1000V DC (norma >= 1.0 MOhm). Sprawdź wkładki SPD DC w rozdzielnicy."
        ),
        "Vac Failure / Grid Overvoltage (Napięcie sieci AC > 253V)": (
            "CO TO ZNACZY: Zbyt wysokie napięcie w sieci energetycznej wyłącza falownik.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Zmierz napięcie L1-N, L2-N, L3-N. Dokręć wszystkie śruby w rozdzielnicy PV i głównej (luźny przewód N podbija napięcie!).\n"
            "2. Połącz się z falownikiem przez aplikację SolarGo (hasło instalatora: goodwe2010).\n"
            "3. Sprawdź kod kraju (Poland EN50549) i włącz funkcję mocy biernej Q(U) w zakładce Advanced Settings -> Reactive Power."
        ),
        "Utility Loss (Brak sieci AC na wtyczce falownika)": (
            "PROCEDURA KROK PO KROKU:\n"
            "1. Wyłącz bezpiecznik AC, wyjmij wtyczkę AC z dołu falownika GoodWe.\n"
            "2. Rozkręć wtyczkę i sprawdź, czy któryś przewód fazowy lub N nie wysunął się z tulejki zaciskowej."
        ),
        "Relay Check Failure (Błąd testu przekaźników AC)": (
            "PROCEDURA KROK PO KROKU:\n"
            "1. Zmierz napięcie między N a PE (powinno być bliskie 0V).\n"
            "2. Zrób pełny restart falownika (AC i DC wyłączone na 5 minut). Jeśli nadal klika i wyrzuca Relay Failure – uszkodzony przekaźnik na płycie (wymiana gwarancyjna)."
        ),
        "BMS Loss / Battery Communication Failure (Seria hybrydowa GoodWe ET)": (
            "PROCEDURA KROK PO KROKU:\n"
            "1. W aplikacji SolarGo sprawdź, czy wybrano dokładny model podłączonego magazynu energii (np. Pylontech Force H2, BYD HVS, GoodWe Lynx Home F).\n"
            "2. Sprawdź, czy na końcu magistrali CAN w ostatnim module bateryjnym jest wpięty **rezystor terminujący (zworka 120 $\Omega$)**.\n"
            "3. Zrestartuj najpierw falownik, a następnie kolumnę bateryjną."
        )
    },
    "Pompa Ciepła ARGOM / Układ Hydrauliczny i Solarny": {
        "E03 / E08 / Flow Switch Error (Błąd przepływu wody / zładu)": (
            "CO TO ZNACZY: Zbyt wolny przepływ wody/glikolu przez wymiennik – ryzyko zamrożenia wymiennika.\n"
            "PROCEDURA KROK PO KROKU (OBOWIĄZKOWA CHECKLISTA):\n"
            "1. MANOMETR: Sprawdź ciśnienie na zegarze. Jeśli < 1.0 bar -> dopuść wody/glikolu do **1.5 - 1.8 bar**.\n"
            "2. FILTRY (90% PRZYCZYN!): Zakręć zawory przed i za filtrem siatkowym (skośnym) i separatorem magnetycznym. Odkręć korek (podstaw wiadro!), wyjmij sitko i trzpień magnetyczny. Wypłucz czarny szlam pod wodą, zamontuj i odpowietrz.\n"
            "3. PRZEPŁYW PRZEZ PĘTLE: Odkręć na stałe min. 2 pętle podłogówki / grzejniki (klient nie może zamknąć wszystkich obiegów siłownikami!).\n"
            "4. POMPA OBIEGOWA: Sprawdź czy pompa obiegowa świeci i pracuje oraz czy pod łopatkę czujnika przepływu (Flow Switch) nie podeszły pakuły/teflon."
        ),
        "E01 / P01 - High Pressure Protection (Błąd wysokiego ciśnienia freonu HP)": (
            "CO TO ZNACZY: Ciepło ze sprężarki nie jest odbierane przez wodę.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Jeśli wyskakuje przy grzaniu CWU (bojlera): obniż temperaturę zadaną CWU ze sprężarki na 48-50 st.C. Sprawdź siłownik zaworu 3-drogowego (czy trzpień nie zaciął się w połowie) oraz odpowietrz wężownicę bojlera.\n"
            "2. Jeśli wyskakuje przy grzaniu CO: wyczyść filtr siatkowy/magnetyczny i zwiększ bieg pompy obiegowej na wyższy."
        ),
        "E02 / P02 - Low Pressure Protection (Błąd niskiego ciśnienia LP / Oblodzenie)": (
            "CO TO ZNACZY: Za niskie ciśnienie ssania – brak odbioru ciepła z powietrza lub wyciek czynnika (F-gaz).\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Jeśli parownik to bryła lodu: polej go ciepłą wodą (NIE KŁUJ ŚRUBOKRĘTEM!), udrożnij odpływ tacy ociekowej, sprawdź grzałkę tacy i włącz ręczny Defrost.\n"
            "2. Jeśli parownik jest czysty i wentylator kręci: podepnij manometry chłodnicze i sprawdź szczelność połączeń kielichowych detektorem gazu."
        ),
        "E04 / P04 - High Discharge Temp (Przegrzanie tłoczenia sprężarki > 115 st.C)": (
            "PROCEDURA KROK PO KROKU:\n"
            "1. Sprawdź wtyczkę cewki zaworu rozprężnego (EEV) na płycie głównej agregatu (czy po resecie zasilania słychać tykanie EEV).\n"
            "2. Jeśli zawór EEV działa: w układzie brakuje czynnika chłodniczego (nieszczelność). Zlokalizuj wyciek, zrób próbę azotową, próżnię i napełnij wagowo."
        ),
        "E05..E15 / T1..T4 Error (Błąd czujnika temperatury NTC)": (
            "PROCEDURA KROK PO KROKU:\n"
            "1. Wejdź w podgląd temperatur na sterowniku – znajdź czujnik pokazujący skrajną wartość (np. -39 st.C, 150 st.C lub kreski).\n"
            "2. Wypnij wtyczkę tego czujnika z płyty i zmierz opór miernikiem (kOhm). Przy 20-25 st.C sprawny czujnik NTC 10k ma mieć ok. 10-12.5 kOhm. Jeśli pokazuje przerwę (OL) lub 0.0 -> wymień czujnik."
        ),
        "E09 / Comm Error (Brak komunikacji sterownika z jednostką zewnętrzną)": (
            "PROCEDURA KROK PO KROKU:\n"
            "1. Sprawdź przewód komunikacyjny RS485 (zaciski A i B). Upewnij się, że A idzie do A, a B do B.\n"
            "2. Sprawdź zasilanie 230V/400V na listwie agregatu zewnętrznego."
        )
    }
}

POSTAL_PREFIX_COORDS = {
    "00": (52.23, 21.01), "01": (52.25, 20.95), "02": (52.18, 21.00), "03": (52.27, 21.08),
    "04": (52.22, 21.15), "05": (52.17, 20.80), "06": (52.88, 20.61), "07": (52.60, 21.80),
    "08": (52.16, 22.28), "09": (52.55, 19.70), "10": (53.78, 20.48), "11": (54.03, 21.00),
    "12": (53.60, 21.50), "13": (53.35, 19.90), "14": (53.85, 19.80), "15": (53.13, 23.16),
    "16": (53.65, 23.10), "17": (52.75, 23.40), "18": (53.18, 22.06), "19": (53.83, 22.35),
    "20": (51.25, 22.57), "21": (51.65, 22.80), "22": (51.05, 23.47), "23": (50.80, 22.30),
    "24": (51.40, 21.97), "25": (50.87, 20.63), "26": (51.40, 21.15), "27": (50.85, 21.45),
    "28": (50.45, 20.70), "29": (50.80, 19.95), "30": (50.06, 19.94), "31": (50.08, 20.00),
    "32": (50.00, 19.65), "33": (50.01, 20.99), "34": (49.65, 19.95), "35": (50.04, 22.00),
    "36": (50.25, 21.95), "37": (50.00, 22.68), "38": (49.68, 21.76), "39": (50.29, 21.42),
    "40": (50.26, 19.02), "41": (50.30, 18.95), "42": (50.70, 19.15), "43": (49.95, 18.95),
    "44": (50.20, 18.60), "45": (50.67, 17.92), "46": (50.75, 18.10), "47": (50.35, 18.20),
    "48": (50.47, 17.33), "49": (50.75, 17.50), "50": (51.11, 17.04), "51": (51.15, 17.10),
    "52": (51.07, 17.05), "53": (51.08, 16.98), "54": (51.14, 16.92), "55": (51.05, 17.25),
    "56": (51.40, 17.35), "57": (50.44, 16.65), "58": (50.84, 16.30), "59": (51.20, 16.16),
    "60": (52.41, 16.93), "61": (52.39, 16.95), "62": (52.35, 17.50), "63": (51.76, 17.80),
    "64": (52.75, 16.60), "65": (51.94, 15.51), "66": (52.25, 15.25), "67": (51.66, 16.08),
    "68": (51.63, 15.14), "69": (52.35, 14.65), "70": (53.43, 14.55), "71": (53.46, 14.50),
    "72": (53.75, 14.90), "73": (53.34, 15.05), "74": (53.00, 14.60), "75": (54.19, 16.18),
    "76": (54.46, 17.03), "77": (53.85, 17.25), "78": (53.70, 16.20), "80": (54.35, 18.65),
    "81": (54.52, 18.53), "82": (54.10, 19.05), "83": (54.10, 18.45), "84": (54.60, 18.20),
    "85": (53.12, 18.01), "86": (53.45, 18.55), "87": (53.01, 18.60), "88": (52.80, 18.25),
    "89": (53.58, 17.85), "90": (51.76, 19.46), "91": (51.80, 19.42), "92": (51.75, 19.55),
    "93": (51.71, 19.47), "94": (51.74, 19.38), "95": (51.88, 19.35), "96": (51.96, 20.15),
    "97": (51.40, 19.70), "98": (51.60, 18.95), "99": (52.10, 19.25)
}

# ==========================================
# 3. FUNKCJE POMOCNICZE (TELEFONY, GSHEETS, BAZA)
# ==========================================
def extract_phone_numbers(raw_phone_str, raw_client_str=""):
    combined = f"{raw_phone_str} {raw_client_str}"
    candidates = re.findall(r'(?:\+48)?\s*\d{3}\s*\d{3}\s*\d{3}|\d{9,11}|\d{2}\s*\d{3}\s*\d{2}\s*\d{2}', combined)
    phones = []
    for c in candidates:
        digits = re.sub(r'\D', '', c)
        if len(digits) >= 9:
            short_num = digits[-9:]
            if short_num not in phones:
                phones.append(short_num)
    if not phones:
        digits_all = re.findall(r'\d{9}', re.sub(r'\s+', '', combined))
        for d in digits_all:
            if d not in phones:
                phones.append(d)
    return phones

def get_gsheet_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    try:
        if "gcp_service_account" in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
            creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
            return gspread.authorize(creds)
        elif os.path.exists("credentials.json"):
            creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
            return gspread.authorize(creds)
    except Exception:
        return None
    return None

def load_config_and_profiles():
    profiles = DEFAULT_TECHS.copy()
    logo_b64 = None
    if os.path.exists(PROFILES_FILE):
        try:
            with open(PROFILES_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
                profiles = saved.get("techs", DEFAULT_TECHS)
                logo_b64 = saved.get("logo_b64", None)
        except Exception:
            pass
    client = get_gsheet_client()
    if client:
        try:
            sh = client.open(GOOGLE_SHEET_NAME)
            try:
                ws = sh.worksheet("Config_System")
                raw_json = ws.acell("A1").value
                raw_logo = ws.acell("B1").value
                if raw_json:
                    profiles = json.loads(raw_json)
                if raw_logo:
                    logo_b64 = raw_logo
            except gspread.WorksheetNotFound:
                ws = sh.add_worksheet(title="Config_System", rows="10", cols="5")
                ws.update_acell("A1", json.dumps(profiles, ensure_ascii=False))
        except Exception:
            pass
    return profiles, logo_b64

def save_config_and_profiles(profiles, logo_b64=None):
    if logo_b64 is None and "logo_b64" in st.session_state:
        logo_b64 = st.session_state["logo_b64"]
    payload = {"techs": profiles, "logo_b64": logo_b64}
    try:
        with open(PROFILES_FILE, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
    client = get_gsheet_client()
    if client:
        try:
            sh = client.open(GOOGLE_SHEET_NAME)
            try:
                ws = sh.worksheet("Config_System")
            except gspread.WorksheetNotFound:
                ws = sh.add_worksheet(title="Config_System", rows="10", cols="5")
            ws.update_acell("A1", json.dumps(profiles, ensure_ascii=False))
            if logo_b64:
                ws.update_acell("B1", logo_b64)
        except Exception:
            pass

def render_brand_logo(max_width=240, centered=False):
    logo_b64 = st.session_state.get("logo_b64", None)
    if logo_b64:
        align_style = "display: block; margin-left: auto; margin-right: auto;" if centered else ""
        st.markdown(
            f'<img src="data:image/png;base64,{logo_b64}" style="max-width:{max_width}px; max-height:95px; object-fit:contain; {align_style} margin-bottom:10px;" alt="Logo Firmy">',
            unsafe_allow_html=True
        )
    elif os.path.exists(LOGO_FILE):
        st.image(LOGO_FILE, width=max_width)
    else:
        if centered:
            st.markdown("<h3 style='text-align:center; color:#0ea5e9; margin-bottom:5px;'>⚡ PROMAG SERWIS OZE</h3>", unsafe_allow_html=True)
        else:
            st.markdown("<h4 style='color:#0ea5e9; margin-bottom:5px;'>⚡ PROMAG SERWIS OZE</h4>", unsafe_allow_html=True)

REQUIRED_COLS = [
    "ID_Zgloszenia", "Klient", "Adres", "Kod_Pocztowy", "Miasto", "Telefon",
    "Typ_Instalacji", "Wymagany_Przeglad", "Opis_Zgloszenia", "Priorytet_Krytyczny", "Typ_Serwisu",
    "Serwisant", "Data_Planowana", "Status", "Naprawione_1_Wizyta",
    "Kwota_Przewidywana", "Kwota_Koncowa", "Czas_Pracy_Min",
    "Wyniki_Przegladu", "Zdjecia_Potwierdzone", "Zalecenia_Klienta",
    "Notatka_Serwisanta", "Data_Zamkniecia", "Podpis_Zlozony"
]

NUMERIC_COLS = ["Kwota_Przewidywana", "Kwota_Koncowa", "Czas_Pracy_Min"]

def ensure_columns(df):
    if df is None or df.empty:
        empty_df = pd.DataFrame(columns=REQUIRED_COLS)
        for nc in NUMERIC_COLS:
            empty_df[nc] = empty_df[nc].astype(float)
        return empty_df

    clean_cols = []
    cols_to_keep = []
    for idx, c in enumerate(df.columns):
        c_str = str(c).strip()
        if c_str != "" and not c_str.lower().startswith("unnamed"):
            cols_to_keep.append(idx)
            clean_cols.append(c_str)

    df = df.iloc[:, cols_to_keep].copy()
    df.columns = clean_cols
    df = df.loc[:, ~df.columns.duplicated()].copy()

    col_map = {}
    used_targets = set(df.columns)

    for c in df.columns:
        cl = c.lower()
        target = None
        if ("klient" in cl or "nazwisko" in cl or "imie" in cl or "imię" in cl or "dane" in cl):
            target = "Klient"
        elif ("adres" in cl or "ulica" in cl):
            target = "Adres"
        elif ("kod" in cl and "poczt" in cl):
            target = "Kod_Pocztowy"
        elif ("miasto" in cl or "miejscowo" in cl):
            target = "Miasto"
        elif ("tel" in cl):
            target = "Telefon"
        elif ("opis_od_klienta" in cl or "opis od" in cl or "usterk" in cl):
            target = "Opis_Zgloszenia"
        elif (cl == "kwota"):
            target = "Kwota_Przewidywana"
        elif ("data_serwisu" in cl or "data ser" in cl):
            target = "Data_Planowana"

        if target and target not in used_targets:
            col_map[c] = target
            used_targets.add(target)

    if col_map:
        df = df.rename(columns=col_map)

    df = df.loc[:, ~df.columns.duplicated()].copy()
    df = df.reset_index(drop=True)

    for rc in REQUIRED_COLS:
        if rc not in df.columns:
            if rc == "ID_Zgloszenia":
                df[rc] = [f"ZGL-{i+1001}" for i in range(len(df))]
            elif rc == "Status":
                df[rc] = "Nowe"
            elif rc in ["Priorytet_Krytyczny", "Wymagany_Przeglad", "Zdjecia_Potwierdzone", "Podpis_Zlozony"]:
                df[rc] = "Nie"
            elif rc == "Typ_Serwisu":
                df[rc] = "Płatny"
            elif rc == "Naprawione_1_Wizyta":
                df[rc] = "Tak"
            elif rc == "Kwota_Przewidywana":
                df[rc] = 650.0
            elif rc in ["Kwota_Koncowa", "Czas_Pracy_Min"]:
                df[rc] = 0.0
            else:
                df[rc] = ""

    extra_cols = [c for c in df.columns if c not in REQUIRED_COLS]
    df = df[REQUIRED_COLS + extra_cols].copy()

    for col in df.columns:
        if col in NUMERIC_COLS:
            default_val = 650.0 if col == "Kwota_Przewidywana" else 0.0
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(default_val).astype(float)
        else:
            df[col] = df[col].fillna("").astype(str).replace(["None", "nan", "NaT", "<NA>"], "")

    for idx, r in df.iterrows():
        adr = str(r["Adres"])
        if str(r["Kod_Pocztowy"]).strip() == "":
            m_zip = re.search(r'\b(\d{2}-\d{3})\b', adr)
            if m_zip:
                df.at[idx, "Kod_Pocztowy"] = m_zip.group(1)
        if str(r["Telefon"]).strip() == "":
            found_p = extract_phone_numbers("", str(r["Klient"]))
            if found_p:
                df.at[idx, "Telefon"] = " / ".join(found_p)
        st_val = str(r["Status"]).strip().lower()
        if "zako" in st_val:
            df.at[idx, "Status"] = "Zakończone"
        elif st_val == "" or st_val == "brak":
            df.at[idx, "Status"] = "Nowe"

    return df.reset_index(drop=True)

def load_tickets():
    client = get_gsheet_client()
    df = pd.DataFrame()
    if client:
        try:
            sh = client.open(GOOGLE_SHEET_NAME)
            ws = sh.get_worksheet(0)
            raw_values = ws.get_all_values()
            if raw_values and len(raw_values) > 1:
                headers = raw_values[0]
                rows = raw_values[1:]
                df = pd.DataFrame(rows, columns=headers)
        except Exception:
            pass

    if df.empty and os.path.exists(LOCAL_EXCEL_FILE):
        try:
            df = pd.read_excel(LOCAL_EXCEL_FILE)
        except Exception:
            pass

    if df.empty:
        df = pd.DataFrame([
            {
                "ID_Zgloszenia": "ZGL-1001", "Klient": "Jan Kowalski", "Adres": "ul. Słoneczna 12",
                "Kod_Pocztowy": "05-800", "Miasto": "Pruszków", "Telefon": "500100200 / 600700800",
                "Typ_Instalacji": "Pompa Ciepła ARGOM", "Wymagany_Przeglad": "Tak",
                "Opis_Zgloszenia": "Brak grzania CO - błąd przepływu E03 + przegląd roczny",
                "Priorytet_Krytyczny": "Tak", "Typ_Serwisu": "Płatny", "Serwisant": "Leszek",
                "Data_Planowana": str(datetime.date.today()), "Status": "Zaplanowane",
                "Naprawione_1_Wizyta": "Tak", "Kwota_Przewidywana": 850.0, "Kwota_Koncowa": 0.0,
                "Czas_Pracy_Min": 0.0, "Wyniki_Przegladu": "", "Zdjecia_Potwierdzone": "Nie",
                "Zalecenia_Klienta": "", "Notatka_Serwisanta": "", "Data_Zamkniecia": "", "Podpis_Zlozony": "Nie"
            }
        ])

    return ensure_columns(df)

def save_tickets(df):
    clean_df = ensure_columns(df)
    st.session_state["df_tickets"] = clean_df.copy()
    try:
        clean_df.to_excel(LOCAL_EXCEL_FILE, index=False)
    except Exception:
        pass
    client = get_gsheet_client()
    if client:
        try:
            sh = client.open(GOOGLE_SHEET_NAME)
            ws = sh.get_worksheet(0)
            ws.clear()
            df_str = clean_df.fillna("").astype(str)
            ws.update([df_str.columns.values.tolist()] + df_str.values.tolist())
        except Exception as e:
            st.warning(f"Zapisano lokalnie. Synchronizacja z Google Sheets napotkała problem: {e}")

# ==========================================
# 4. ALGORYTM PLANOWANIA TRAS (Z WYBOREM SERWISANTA I KODU)
# ==========================================
def extract_coords_from_row(row):
    raw_text = f"{row.get('Kod_Pocztowy', '')} {row.get('Adres', '')} {row.get('Miasto', '')}"
    match = re.search(r'\b(\d{2})-\d{3}\b', str(raw_text))
    if match:
        prefix = match.group(1)
        if prefix in POSTAL_PREFIX_COORDS:
            full_zip = re.sub(r'\D', '', match.group(0))
            offset = (int(full_zip[-2:]) - 50) * 0.003
            base_lat, base_lon = POSTAL_PREFIX_COORDS[prefix]
            return base_lat + offset, base_lon + offset
    return 52.23, 21.01

def haversine_km(lat1, lon1, lat2, lon2):
    R = 6371.0
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlambda = math.radians(lon2 - lon1)
    a = math.sin(dphi / 2.0)**2 + math.cos(phi1) * math.cos(phi2) * math.sin(dlambda / 2.0)**2
    return 2 * R * math.atan2(math.sqrt(a), math.sqrt(1 - a))

def is_tech_available(tech_name, target_date_str, profiles):
    tech_data = profiles.get(tech_name, {})
    absences = tech_data.get("urlopy_l4", [])
    for ab in absences:
        start_d = ab.get("od", "")
        end_d = ab.get("do", "")
        if start_d <= target_date_str <= end_d:
            return False
    return True

def generate_smart_routes(df, profiles, start_date, days_ahead=5, target_tech="Wszyscy"):
    if target_tech != "Wszyscy":
        pending = df[(df["Status"].isin(["Nowe", "Do zaplanowania", ""])) & (df["Serwisant"].isin(["", target_tech]))].copy()
    else:
        pending = df[df["Status"].isin(["Nowe", "Do zaplanowania", ""])].copy()

    if pending.empty:
        return [], "Brak nowych zgłoszeń spełniających kryteria do zaplanowania."

    coords = pending.apply(extract_coords_from_row, axis=1)
    pending["lat"] = [c[0] for c in coords]
    pending["lon"] = [c[1] for c in coords]
    pending["is_crit"] = pending.apply(
        lambda r: 1 if str(r.get("Priorytet_Krytyczny", "")).lower() == "tak" or "awaria" in str(r.get("Opis_Zgloszenia", "")).lower() else 0,
        axis=1
    )

    pending = pending.sort_values(by=["is_crit", "Kwota_Przewidywana"], ascending=[False, False])
    used_indices = set()
    valid_clusters = []

    for idx, seed in pending.iterrows():
        if idx in used_indices:
            continue
        cluster_indices = [idx]
        for cand_idx, cand in pending.iterrows():
            if cand_idx == idx or cand_idx in used_indices:
                continue
            if len(cluster_indices) >= 4:
                break
            dist = haversine_km(seed["lat"], seed["lon"], cand["lat"], cand["lon"])
            if dist <= 50.0:
                cluster_indices.append(cand_idx)

        cluster_df = pending.loc[cluster_indices]
        total_rev = float(cluster_df["Kwota_Przewidywana"].sum())
        crit_count = int(cluster_df["is_crit"].sum())

        if total_rev >= 1200.0 or crit_count >= 2:
            for c_i in cluster_indices:
                used_indices.add(c_i)
            cities = [m for m in cluster_df["Miasto"].astype(str).unique() if m.strip()]
            if not cities:
                cities = [str(a)[:20] for a in cluster_df["Adres"].astype(str).unique()[:2]]
            valid_clusters.append({
                "indices": cluster_indices,
                "total_rev": total_rev,
                "crit_count": crit_count,
                "miasta": ", ".join(cities),
                "count": len(cluster_indices)
            })

    if not valid_clusters:
        return [], "Żaden z klastrów nie spełnił warunku min. 1200 zł lub min. 2 zgłoszeń krytycznych w promieniu 50 km."

    tech_names = [target_tech] if target_tech != "Wszyscy" else list(profiles.keys())
    proposals = []
    cluster_ptr = 0

    for d_offset in range(days_ahead):
        curr_date = start_date + datetime.timedelta(days=d_offset)
        curr_date_str = str(curr_date)
        for t_name in tech_names:
            if cluster_ptr >= len(valid_clusters):
                break
            if is_tech_available(t_name, curr_date_str, profiles):
                cl = valid_clusters[cluster_ptr]
                cl["serwisant"] = t_name
                cl["data"] = curr_date_str
                proposals.append(cl)
                cluster_ptr += 1

    return proposals, f"Wygenerowano {len(proposals)} opłacalnych tras (dla: {target_tech})."

def calculate_tech_metrics(df, tech_name):
    tech_df = df[df["Serwisant"] == tech_name]
    closed_df = tech_df[tech_df["Status"] == "Zakończone"].copy()
    total_closed = len(closed_df)
    if total_closed > 0:
        ftfr_count = len(closed_df[closed_df["Naprawione_1_Wizyta"] == "Tak"])
        ftfr_pct = round((ftfr_count / total_closed) * 100.0, 1)
    else:
        ftfr_pct = 100.0

    paid_count = len(closed_df[closed_df["Typ_Serwisu"] == "Płatny"])
    warranty_count = len(closed_df[closed_df["Typ_Serwisu"] == "Gwarancyjny"])

    current_ym = datetime.date.today().strftime("%Y-%m")
    closed_df["ym"] = closed_df["Data_Zamkniecia"].astype(str).str.slice(0, 7)
    monthly_df = closed_df[(closed_df["ym"] == current_ym) | (closed_df["ym"] == "")]
    monthly_revenue = float(monthly_df["Kwota_Koncowa"].sum())

    missing_tier_1 = max(0.0, COMMISSION_TIER_1 - monthly_revenue)
    missing_tier_2 = max(0.0, COMMISSION_TIER_2 - monthly_revenue)

    if monthly_revenue >= COMMISSION_TIER_2:
        commission_pln = monthly_revenue * COMMISSION_RATE_2
        active_rate_str = "4% (Próg II zdobyty!)"
    elif monthly_revenue >= COMMISSION_TIER_1:
        commission_pln = monthly_revenue * COMMISSION_RATE_1
        active_rate_str = "3% (Próg I zdobyty!)"
    else:
        commission_pln = 0.0
        active_rate_str = "0% (W drodze do Progu I)"

    return {
        "total_closed": total_closed,
        "ftfr_pct": ftfr_pct,
        "paid_count": paid_count,
        "warranty_count": warranty_count,
        "monthly_revenue": monthly_revenue,
        "missing_tier_1": missing_tier_1,
        "missing_tier_2": missing_tier_2,
        "commission_pln": commission_pln,
        "active_rate_str": active_rate_str
    }

# ==========================================
# 5. INICJALIZACJA SESJI I TRWAŁE LOGOWANIE
# ==========================================
if "profiles" not in st.session_state or "logo_b64" not in st.session_state:
    p_loaded, l_loaded = load_config_and_profiles()
    st.session_state["profiles"] = p_loaded
    st.session_state["logo_b64"] = l_loaded

if "df_tickets" not in st.session_state:
    st.session_state["df_tickets"] = load_tickets()
else:
    st.session_state["df_tickets"] = ensure_columns(st.session_state["df_tickets"])

if "logged_role" not in st.session_state:
    st.session_state["logged_role"] = None
if "logged_user" not in st.session_state:
    st.session_state["logged_user"] = None
if "active_day" not in st.session_state:
    st.session_state["active_day"] = None
if "active_ticket_id" not in st.session_state:
    st.session_state["active_ticket_id"] = None

if st.session_state["logged_role"] is None:
    qp_user = st.query_params.get("user", None)
    qp_token = st.query_params.get("token", None)
    
    if qp_user == "Koordynator" and qp_token == PIN_KOORDYNATOR:
        st.session_state["logged_role"] = "koordynator"
        st.session_state["logged_user"] = "Koordynator"
    elif qp_user == "Zarząd" and qp_token == PIN_ZARZAD:
        st.session_state["logged_role"] = "zarzad"
        st.session_state["logged_user"] = "Zarząd"
    elif qp_user in st.session_state["profiles"]:
        expected_pin = str(st.session_state["profiles"][qp_user].get("pin", ""))
        if qp_token == expected_pin:
            st.session_state["logged_role"] = "serwisant"
            st.session_state["logged_user"] = qp_user

if st.session_state["logged_role"] is None:
    render_brand_logo(max_width=280, centered=True)
    st.markdown("<p style='text-align:center; color:#64748b;'>Mobilny System Zarządzania Serwisem OZE (Telefon / Tablet)</p>", unsafe_allow_html=True)
    st.divider()

    col_login, col_onboard = st.columns([1.2, 1])
    with col_login:
        st.subheader("🔐 Logowanie do systemu")
        tech_list = list(st.session_state["profiles"].keys())
        user_choice = st.selectbox(
            "Wybierz swój profil:",
            ["--- Wybierz z listy ---"] + tech_list + ["Koordynator Serwisu", "Zarząd (Raporty)"]
        )
        pin_input = st.text_input("Wpisz 4-cyfrowy kod PIN:", type="password", max_chars=8)
        remember_me = st.checkbox("🔒 Zapamiętaj mnie na tym urządzeniu (nie wylogowuj po nocy)", value=True)

        if st.button("🚀 Zaloguj się", type="primary"):
            if user_choice == "--- Wybierz z listy ---":
                st.error("Wybierz swój profil z listy powyżej.")
            elif user_choice == "Koordynator Serwisu":
                if pin_input == PIN_KOORDYNATOR:
                    st.session_state["logged_role"] = "koordynator"
                    st.session_state["logged_user"] = "Koordynator"
                    if remember_me:
                        st.query_params["user"] = "Koordynator"
                        st.query_params["token"] = PIN_KOORDYNATOR
                    st.rerun()
                else:
                    st.error("Nieprawidłowy kod PIN Koordynatora.")
            elif user_choice == "Zarząd (Raporty)":
                if pin_input == PIN_ZARZAD:
                    st.session_state["logged_role"] = "zarzad"
                    st.session_state["logged_user"] = "Zarząd"
                    if remember_me:
                        st.query_params["user"] = "Zarząd"
                        st.query_params["token"] = PIN_ZARZAD
                    st.rerun()
                else:
                    st.error("Nieprawidłowy kod PIN Zarządu.")
            elif user_choice in st.session_state["profiles"]:
                expected_pin = str(st.session_state["profiles"][user_choice].get("pin", ""))
                if pin_input == expected_pin:
                    st.session_state["logged_role"] = "serwisant"
                    st.session_state["logged_user"] = user_choice
                    st.session_state["active_day"] = None
                    st.session_state["active_ticket_id"] = None
                    if remember_me:
                        st.query_params["user"] = user_choice
                        st.query_params["token"] = expected_pin
                    st.rerun()
                else:
                    st.error("Nieprawidłowy kod PIN serwisanta.")

    with col_onboard:
        st.subheader("🎓 Dla nowych pracowników")
        st.info("Rozpoczynasz pracę w zespole serwisowym? Sprawdź standardy obsługi i wymogi dokumentacji zdjęciowej.")
        if st.button("📘 Otwórz Szkolenie Wstępne (Onboarding)"):
            st.session_state["show_onboarding"] = not st.session_state.get("show_onboarding", False)
        if st.session_state.get("show_onboarding", False):
            with st.container(border=True):
                st.markdown("""
                ### Standard Wizyty Serwisowej PROMAG OZE
                1. **Kontakt i kultura pracy** – zadzwoń lub wyślij SMS z aplikacji przed dojazdem.
                2. **Obowiązkowe zdjęcia** – zawsze wykonaj zdjęcie tabliczki znamionowej urządzenia oraz instalacji przed i po naprawie. Bez potwierdzenia zdjęć system nie pozwoli zamknąć zlecenia!
                3. **Diagnostyka i Przegląd** – korzystaj z wbudowanej Ścieżki Diagnostycznej (Deye, GoodWe, Argom) oraz walidatora pomiarów glikolu (refraktometr + pasek pH).
                """)
    st.stop()

top_c1, top_c2, top_c3 = st.columns([2, 2, 1])
with top_c1:
    render_brand_logo(max_width=200, centered=False)
with top_c2:
    st.markdown(f"**Zalogowany jako:** `{st.session_state['logged_user']}`")
with top_c3:
    if st.button("🚪 Wyloguj"):
        st.query_params.clear()
        st.session_state["logged_role"] = None
        st.session_state["logged_user"] = None
        st.session_state["active_day"] = None
        st.session_state["active_ticket_id"] = None
        st.rerun()

st.divider()

# ==========================================
# 6. PANEL SERWISANTA
# ==========================================
if st.session_state["logged_role"] == "serwisant":
    tech_name = st.session_state["logged_user"]
    df = st.session_state["df_tickets"]
    metrics = calculate_tech_metrics(df, tech_name)

    m_col1, m_col2, m_col3 = st.columns([1.4, 1.4, 1])
    with m_col1:
        if metrics["missing_tier_1"] > 0:
            headline = f"Brakuje: {metrics['missing_tier_1']:,.0f} zł do prowizji 3%"
            subline = f"Próg I (15 000 zł) | Do progu 4% (30 000 zł) brakuje: {metrics['missing_tier_2']:,.0f} zł"
        elif metrics["missing_tier_2"] > 0:
            headline = f"🔥 Masz już 3%! Brakuje: {metrics['missing_tier_2']:,.0f} zł do 4%"
            subline = f"Wypracowany obrót: {metrics['monthly_revenue']:,.0f} zł (Premia: {metrics['commission_pln']:,.0f} zł)"
        else:
            headline = "🏆 Maksymalny próg 4% zdobyty!"
            subline = f"Wypracowany obrót: {metrics['monthly_revenue']:,.0f} zł (Premia: {metrics['commission_pln']:,.0f} zł)"

        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-title">🎯 Twój Cel Prowizyjny (Bieżący Miesiąc)</div>
            <div class="kpi-value">{headline}</div>
            <div class="kpi-sub">{subline}</div>
        </div>
        """, unsafe_allow_html=True)

    with m_col2:
        st.markdown(f"""
        <div class="kpi-card kpi-card-warn">
            <div class="kpi-title">⚡ Skuteczność 1. Wizyty (First-Time Fix)</div>
            <div class="kpi-value">{metrics['ftfr_pct']}% naprawionych od ręki</div>
            <div class="kpi-sub">Obsłużone: {metrics['total_closed']} (Płatne: {metrics['paid_count']} | Gwarancyjne: {metrics['warranty_count']})</div>
        </div>
        """, unsafe_allow_html=True)

    with m_col3:
        st.write("")
        if st.button("🩺 ŚCIEŻKA DIAGNOSTYCZNA / KODY BŁĘDÓW"):
            st.session_state["show_kb"] = not st.session_state.get("show_kb", False)

    progress_val = min(1.0, float(metrics["monthly_revenue"] / COMMISSION_TIER_2))
    st.progress(progress_val, text=f"Postęp obrotu w tym miesiącu: {metrics['monthly_revenue']:,.0f} zł / 30 000 zł ({metrics['active_rate_str']})")

    if st.session_state.get("show_kb", False):
        with st.container(border=True):
            st.subheader("🩺 Interaktywna Ścieżka Diagnostyczna (Deye / GoodWe / Argom)")
            kb_cat = st.selectbox("Wybierz urządzenie:", list(DIAGNOSTYKA_BAZA.keys()), key="kb_cat_top")
            kb_err = st.selectbox("Wybierz kod błędu lub objaw z wyświetlacza:", list(DIAGNOSTYKA_BAZA[kb_cat].keys()), key="kb_err_top")
            st.info(DIAGNOSTYKA_BAZA[kb_cat][kb_err])

    st.divider()

    my_tickets = df[(df["Serwisant"] == tech_name) & (df["Status"] != "Zakończone")].copy()

    if my_tickets.empty:
        st.success("🎉 Wszystkie zaplanowane dla Ciebie zlecenia zostały zrealizowane!")
        st.session_state["active_day"] = None
        st.session_state["active_ticket_id"] = None
    else:
        today_str = str(datetime.date.today())
        my_tickets["Data_Grupa"] = my_tickets["Data_Planowana"].astype(str).str.strip()
        my_tickets.loc[my_tickets["Data_Grupa"] == "", "Data_Grupa"] = today_str

        if st.session_state["active_day"] is None:
            st.subheader("📅 Twoje zaplanowane dni wyjazdowe (Trasy)")
            st.caption("Zlecenia są pogrupowane na pojedyncze dni. Kliknij **Rozpocznij dzień**, aby odblokować listę klientów z danego dnia.")

            unique_days = sorted(my_tickets["Data_Grupa"].unique())
            for day_str in unique_days:
                day_cluster = my_tickets[my_tickets["Data_Grupa"] == day_str]
                pts_count = len(day_cluster)
                cities_list = ", ".join([m for m in day_cluster["Miasto"].astype(str).unique() if m.strip() != ""])
                if not cities_list:
                    cities_list = ", ".join([str(a)[:22] for a in day_cluster["Adres"].astype(str).unique()[:3]])
                est_rev = float(day_cluster["Kwota_Przewidywana"].sum())
                crit_cnt = len(day_cluster[day_cluster["Priorytet_Krytyczny"] == "Tak"])

                is_today = (day_str == today_str)
                card_class = "day-card day-card-today" if is_today else "day-card"
                badge_today = "🟢 DZISIAJ" if is_today else "🗓️ ZAPLANOWANA TRASA"

                with st.container(border=True):
                    c_day_info, c_day_btn = st.columns([2.2, 1])
                    with c_day_info:
                        st.markdown(f"""
                        <div class="{card_class}">
                            <div style="font-size:13px; color:#38bdf8; font-weight:700;">{badge_today}</div>
                            <div style="font-size:22px; font-weight:800; margin-top:2px;">📅 Data trasy: {day_str}</div>
                            <div style="font-size:15px; margin-top:6px;">📍 <b>Lokalizacje:</b> {cities_list}</div>
                            <div style="font-size:14px; color:#cbd5e1; margin-top:4px;">
                                🛠️ Punktów: <b>{pts_count}</b> &nbsp;|&nbsp;
                                💰 Szacowany obrót: <b>{est_rev:,.0f} zł</b> &nbsp;|&nbsp;
                                🔥 Krytyczne: <b>{crit_cnt}</b>
                            </div>
                        </div>
                        """, unsafe_allow_html=True)

                        with st.expander(f"👁️ Podgląd punktów w trasie na dzień {day_str} ({pts_count})"):
                            for idx_pt, (_, r_pt) in enumerate(day_cluster.iterrows(), 1):
                                crit_tag = " 🔥 [AWARIA]" if r_pt["Priorytet_Krytyczny"] == "Tak" else ""
                                rev_tag = " 🔍 [PRZEGLĄD]" if r_pt.get("Wymagany_Przeglad") == "Tak" else ""
                                st.markdown(f"**{idx_pt}. {r_pt['Adres']}** ({r_pt['Miasto']}) – *{r_pt['Typ_Instalacji']}*{crit_tag}{rev_tag}")

                    with c_day_btn:
                        st.write("")
                        if st.button(f"🚀 ROZPOCZNIJ DZIEŃ ({day_str})", key=f"start_day_{day_str}", type="primary"):
                            st.session_state["active_day"] = day_str
                            st.session_state["active_ticket_id"] = None
                            st.rerun()

        elif st.session_state["active_day"] is not None and st.session_state["active_ticket_id"] is None:
            act_day = st.session_state["active_day"]
            day_tickets = my_tickets[my_tickets["Data_Grupa"] == act_day].copy()

            head_c1, head_c2 = st.columns([2.2, 1])
            with head_c1:
                st.subheader(f"🚚 Aktywny dzień pracy: {act_day}")
                st.caption("Kolejność odwiedzin ustalasz sam. Zadzwoń lub wyślij SMS do klienta i kliknij **Rozpocznij zlecenie**.")
            with head_c2:
                if st.button("⬅️ Zmień dzień / Wróć do harmonogramu"):
                    st.session_state["active_day"] = None
                    st.session_state["active_ticket_id"] = None
                    st.rerun()

            if day_tickets.empty:
                st.success(f"🎉 Wszystkie zlecenia zaplanowane na dzień **{act_day}** zostały zakończone!")
                if st.button("📅 Wróć do wyboru kolejnego dnia", type="primary"):
                    st.session_state["active_day"] = None
                    st.rerun()
            else:
                for idx_card, (t_idx, t_row) in enumerate(day_tickets.iterrows(), 1):
                    t_id = t_row["ID_Zgloszenia"]
                    is_crit = (str(t_row["Priorytet_Krytyczny"]) == "Tak")
                    is_przeglad = (str(t_row.get("Wymagany_Przeglad", "Nie")) == "Tak")
                    status_badge = "🟡 W TRAKCIE" if t_row["Status"] == "W trakcie" else "⏳ OCZEKUJE"
                    crit_badge = " 🔥 AWARIA" if is_crit else ""
                    przeg_badge = " 🔍 PRZEGLĄD" if is_przeglad else ""

                    with st.container(border=True):
                        st.markdown(f"#### Punkt #{idx_card}: **{t_row['Klient']}** `{status_badge}` {crit_badge} {przeg_badge}")
                        col_d1, col_d2, col_d3 = st.columns([1.4, 1.3, 1.1])

                        with col_d1:
                            st.markdown(f"📍 **Adres:** `{t_row['Kod_Pocztowy']} {t_row['Miasto']}, {t_row['Adres']}`")
                            st.markdown(f"🔧 **Urządzenie:** {t_row['Typ_Instalacji']} ({t_row['Typ_Serwisu']})")
                            st.markdown(f"📝 **Opis:** {t_row['Opis_Zgloszenia']}")

                        with col_d2:
                            phones = extract_phone_numbers(str(t_row["Telefon"]), str(t_row["Klient"]))
                            if phones:
                                for p_idx, p_num in enumerate(phones, 1):
                                    pc1, pc2 = st.columns(2)
                                    with pc1:
                                        st.link_button(f"📞 Zadzwoń ({p_num})", f"tel:{p_num}")
                                    with pc2:
                                        sms_txt = f"Dzien dobry, tu serwis PROMAG OZE. Dzisiaj realizujemy zgloszenie pod adresem {t_row['Adres']}. O ktorej godzinie bedzie Pan/Pani w domu?"
                                        st.link_button(f"💬 SMS ({p_num})", f"sms:{p_num}?body={sms_txt.replace(' ', '%20')}")
                            else:
                                st.warning("Brak wykrytego numeru telefonu w karcie.")

                            maps_q = f"{t_row['Adres']} {t_row['Kod_Pocztowy']} {t_row['Miasto']}".replace(" ", "+")
                            st.link_button("🗺️ Nawiguj w Google Maps", f"https://www.google.com/maps/search/?api=1&query={maps_q}")

                        with col_d3:
                            st.markdown(f"💰 **Szacowana kwota:** `{float(t_row['Kwota_Przewidywana']):,.0f} zł`")
                            if st.button(f"🛠️ ROZPOCZNIJ ZLECENIE ({t_id})", key=f"start_ticket_{t_id}", type="primary"):
                                df.loc[t_idx, "Status"] = "W trakcie"
                                save_tickets(df)
                                st.session_state["active_ticket_id"] = t_id
                                st.rerun()

        elif st.session_state["active_ticket_id"] is not None:
            selected_id = st.session_state["active_ticket_id"]
            act_day = st.session_state["active_day"]

            matching_rows = df[df["ID_Zgloszenia"] == selected_id]
            if matching_rows.empty:
                st.session_state["active_ticket_id"] = None
                st.rerun()

            t_idx = matching_rows.index[0]
            t_row = df.loc[t_idx]

            nav_c1, nav_c2 = st.columns([2.2, 1])
            with nav_c1:
                st.subheader(f"🛠️ Obsługa zlecenia: {selected_id} (Dzień: {act_day})")
            with nav_c2:
                if st.button("⬅️ Wróć do listy klientów z tego dnia"):
                    st.session_state["active_ticket_id"] = None
                    st.rerun()

            c_info1, c_info2 = st.columns([1.3, 1])
            with c_info1:
                st.markdown(f"### 👤 Klient: **{t_row['Klient']}**")
                st.markdown(f"📍 **Adres:** `{t_row['Kod_Pocztowy']} {t_row['Miasto']}, {t_row['Adres']}`")
                st.markdown(f"🔧 **Urządzenie:** `{t_row['Typ_Instalacji']}` | **Typ:** `{t_row['Typ_Serwisu']}`")
                st.markdown(f"📝 **Opis zgłoszenia:** {t_row['Opis_Zgloszenia']}")

            with c_info2:
                phones = extract_phone_numbers(str(t_row["Telefon"]), str(t_row["Klient"]))
                for p_num in phones:
                    p_col_a, p_col_b = st.columns(2)
                    with p_col_a:
                        st.link_button(f"📞 Zadzwoń: {p_num}", f"tel:{p_num}")
                    with p_col_b:
                        sms_body = f"Dzien dobry, tu serwis PROMAG OZE. Jestem w drodze pod adres {t_row['Adres']}."
                        st.link_button(f"💬 SMS: {p_num}", f"sms:{p_num}?body={sms_body.replace(' ', '%20')}")
                maps_query = f"{t_row['Adres']} {t_row['Kod_Pocztowy']} {t_row['Miasto']}".replace(" ", "+")
                st.link_button("🗺️ Nawiguj w Google Maps", f"https://www.google.com/maps/search/?api=1&query={maps_query}")

            st.divider()

            with st.expander("🩺 NIE WIESZ CO SIĘ DZIEJE Z URZĄDZENIEM? Kliknij, aby otworzyć Ścieżkę Diagnostyczną", expanded=False):
                inst_lower = str(t_row["Typ_Instalacji"]).lower()
                default_cat_idx = 0
                if "goodwe" in inst_lower:
                    default_cat_idx = 1
                elif "pomp" in inst_lower or "argom" in inst_lower or "solar" in inst_lower:
                    default_cat_idx = 2

                diag_cat = st.selectbox("Wybierz diagnozowany sprzęt:", list(DIAGNOSTYKA_BAZA.keys()), index=default_cat_idx, key="diag_cat_ticket")
                diag_err = st.selectbox("Wybierz kod błędu lub objaw usterki:", list(DIAGNOSTYKA_BAZA[diag_cat].keys()), key="diag_err_ticket")
                st.warning(DIAGNOSTYKA_BAZA[diag_cat][diag_err])

            przeglad_summary_lines = []
            need_inspection = (
                str(t_row.get("Wymagany_Przeglad", "Nie")) == "Tak"
                or "przegl" in str(t_row["Opis_Zgloszenia"]).lower()
            )
            enable_inspection = st.checkbox(
                "🔍 Wykonuję przegląd techniczny urządzenia (Wypełnij protokół pomiarowy i walidator norm)",
                value=need_inspection
            )

            if enable_inspection:
                with st.container(border=True):
                    st.markdown("### 📋 Protokół Przeglądu Urządzenia i Walidacja Norm na Żywo")
                    insp_type = st.radio(
                        "Wybierz rodzaj kontrolowanej instalacji:",
                        ["Pompa Ciepła / Kolektory Słoneczne (Hydraulika + Glikol)", "Fotowoltaika PV / Magazyn Energii (Pomiary DC/AC)"],
                        horizontal=True
                    )

                    if insp_type.startswith("Pompa"):
                        p_c1, p_c2 = st.columns(2)
                        with p_c1:
                            cisnienie_co = st.number_input("Ciśnienie statyczne układu [bar] (Norma: 1.2 - 2.2 bar):", min_value=0.0, max_value=6.0, step=0.1, value=1.5)
                            delta_t = st.number_input("Delta temperatur Zasilanie-Powrot [st.C] (Norma: 4.0 - 8.0 st.C):", min_value=0.0, max_value=30.0, step=0.5, value=5.5)
                            filtr_czystosc = st.selectbox("Stan filtra siatkowego i separatora magnetycznego:", ["Czysty / Wyczyszczono", "Silnie zanieczyszczony szlamem (Wyczyszczono)"])
                        with p_c2:
                            glikol_temp = st.number_input("Temperatura krzepnięcia glikolu z refraktometru [st.C] (Norma: -15 do -28 st.C):", min_value=-50.0, max_value=0.0, step=1.0, value=-20.0)
                            glikol_ph = st.number_input("Odczyn pH zładu / glikolu z paska pH (Norma: 7.5 - 9.0 pH):", min_value=3.0, max_value=12.0, step=0.2, value=8.0)
                            szczelnosc_fgaz = st.selectbox("Kontrola szczelności układu i naczynia przeponowego:", ["Szczelny / Ciśnienie OK", "Wykryto usterkę / Dopompowano naczynie"])

                        if cisnienie_co < 1.0:
                            st.error("🚨 ALERT: Zbyt niskie ciśnienie (< 1.0 bar)! Ryzyko błędu przepływu E03 i kawitacji pompy. PROCEDURA: Sprawdź szczelność, zweryfikuj naczynie przeponowe i dopuść zładu do 1.5 bar.")
                        elif cisnienie_co > 2.5:
                            st.error("🚨 ALERT: Zbyt wysokie ciśnienie (> 2.5 bar)! Ryzyko otwarcia zaworu bezpieczeństwa 3 bar. PROCEDURA: Uszkodzona przepona naczynia wzbiorczego lub brak poduszki powietrznej. Upuść ciśnienie do 1.5 bar.")

                        if delta_t > 9.0:
                            st.error("🚨 ALERT: Zbyt wysoka Delta T (> 9 st.C) – zdławiony przepływ przez wymiennik! PROCEDURA: Wyczyść filtr siatkowy i magnetyczny, odkręć rotametry i zwiększ bieg pompy obiegowej.")

                        if glikol_temp > -15.0:
                            st.error(f"🚨 ALERT GLIKOL ({glikol_temp} st.C): Zbyt słabe stężenie glikolu! Ryzyko zamarznięcia i rozerwania wymiennika. PROCEDURA: Nie dolewaj wody! Wtłocz koncentrat glikolu propylenowego i wymieszaj obieg do uzyskania min. -18 st.C.")
                        elif glikol_temp < -32.0:
                            st.warning(f"⚠️ OSTRZEŻENIE GLIKOL ({glikol_temp} st.C): Zbyt gęsty glikol! Duży opór dla pompy obiegowej (błąd E03). Rozcieńcz lekko wodą demineralizowaną do ok. -22 st.C.")

                        if glikol_ph < 7.0:
                            st.error(f"🚨 ALERT pH ({glikol_ph}): Zakwaszenie glikolu (pH < 7.0)! Płyn trawi uszczelki i miedź w wymienniku. PROCEDURA: Konieczna całkowita wymiana glikolu na nowy z inhibitorem korozji i płukanie układu.")
                        elif glikol_ph > 9.5:
                            st.warning(f"⚠️ OSTRZEŻENIE pH ({glikol_ph}): Zbyt wysokie pH (> 9.5) – ryzyko korozji aluminium/mosiądzu.")

                        przeglad_summary_lines.append(
                            f"PRZEGLĄD PC/SOLAR: Ciśnienie={cisnienie_co} bar | DeltaT={delta_t} C | Glikol={glikol_temp} C | pH={glikol_ph} | Filtr={filtr_czystosc} | Układ={szczelnosc_fgaz}"
                        )
                    else:
                        pv_c1, pv_c2 = st.columns(2)
                        with pv_c1:
                            riso_val = st.number_input("Rezystancja izolacji Riso [MOhm] (Norma: >= 1.0 MOhm):", min_value=0.0, max_value=500.0, step=0.5, value=50.0)
                            u_max_ac = st.number_input("Maksymalne napięcie fazy AC [V] (Norma: 207 - 253 V):", min_value=180.0, max_value=280.0, step=1.0, value=238.0)
                        with pv_c2:
                            uziemienie_val = st.number_input("Rezystancja uziemienia [Ohm] (Norma: <= 10 Ohm):", min_value=0.1, max_value=100.0, step=0.5, value=4.5)
                            styki_stan = st.selectbox("Kontrola zacisków AC/DC i ochronników SPD:", ["Dokręcone / Sprawne", "Wykryto luźne styki / Przegrzanie (Dokręcono/Wymieniono)"])

                        if riso_val < 1.0:
                            st.error("🚨 ALERT Riso (< 1.0 MOhm): Uszkodzona izolacja DC! PROCEDURA: Rozepnij stringi, sprawdź wkładki SPD DC w rozdzielnicy oraz złącza MC4 na dachu.")
                        if u_max_ac >= 251.0:
                            st.warning("⚠️ OSTRZEŻENIE AC (>= 251V): Napięcie bliskie progu wyłączenia falownika (253V). PROCEDURA: Dokręć zaciski N i L w rozdzielnicy oraz włącz tryb mocy biernej Q(U) w falowniku.")
                        if uziemienie_val > 10.0:
                            st.error("🚨 ALERT UZIEMIENIE (> 10 Ohm): Zbyt wysoka rezystancja uziemienia! Ryzyko błędu F08 i braku ochrony przepięciowej. PROCEDURA: Sprawdź złącze kontrolne i rozbuduj uziom szpilkowy.")

                        przeglad_summary_lines.append(
                            f"PRZEGLĄD PV: Riso={riso_val} MOhm | Max AC={u_max_ac} V | Uziemienie={uziemienie_val} Ohm | Styki={styki_stan}"
                        )

            st.divider()

            st.subheader("📸 Obowiązkowa dokumentacja zdjęciowa z instalacji")
            st.info("Wgraj zdjęcia z telefonu/tabletu: **1) Tabliczka znamionowa urządzenia**, **2) Widok instalacji / usterki przed i po naprawie**.")
            uploaded_photos = st.file_uploader(
                "Dodaj zdjęcia z galerii lub aparatu (minimum 1 zdjęcie wymagane):",
                type=["png", "jpg", "jpeg"],
                accept_multiple_files=True,
                key=f"photos_{selected_id}"
            )
            if uploaded_photos:
                st.success(f"Załączono zdjęć: {len(uploaded_photos)}")

            confirm_photos = st.checkbox(
                "✅ Oświadczam i potwierdzam, że dodałem wszystkie wymagane zdjęcia z instalacji (w tym czytelne zdjęcie tabliczki znamionowej urządzenia).",
                key=f"conf_photo_{selected_id}"
            )

            st.divider()

            st.subheader("🧾 Rozliczenie wizyty i wykonanych prac")
            calc_c1, calc_c2 = st.columns(2)
            with calc_c1:
                selected_services = st.multiselect(
                    "Wykonane czynności z cennika:",
                    options=list(CENNIK_BAZOWY.keys()),
                    default=[]
                )
                extra_parts = st.number_input("Koszt użytych części / materiałów (zł):", min_value=0.0, step=50.0, value=0.0)
                km_dojazd = st.number_input("Dojazd (km ponad ryczałt):", min_value=0, step=10, value=0)

            with calc_c2:
                work_minutes = st.number_input(
                    "Czas pracy na miejscu (minuty):",
                    min_value=15,
                    step=15,
                    value=int(float(t_row["Czas_Pracy_Min"])) if float(t_row["Czas_Pracy_Min"]) > 0 else 60
                )
                typ_serwisu_real = st.selectbox("Kwalifikacja zlecenia:", ["Płatny", "Gwarancyjny"], index=0 if t_row["Typ_Serwisu"] != "Gwarancyjny" else 1)
                fixed_first_visit = st.radio(
                    "Czy usterka została w pełni rozwiązana podczas tej wizyty?",
                    ["Tak", "Nie (Wymagana kolejna wizyta / zamówienie części)"],
                    horizontal=True
                )

            base_sum = sum(CENNIK_BAZOWY[s] for s in selected_services)
            if base_sum == 0 and typ_serwisu_real == "Płatny":
                base_sum = float(t_row["Kwota_Przewidywana"])
            auto_total = 0.0 if typ_serwisu_real == "Gwarancyjny" else (base_sum + extra_parts + (km_dojazd * 2.5))

            final_price = st.number_input("💰 Kwota końcowa do rozliczenia (zł brutto):", min_value=0.0, step=50.0, value=float(auto_total))
            internal_note = st.text_area("Notatka wewnętrzna dla Koordynatora (niewidoczna dla klienta):", value=str(t_row["Notatka_Serwisanta"]))

            st.markdown('<div class="client-box">', unsafe_allow_html=True)
            render_brand_logo(max_width=220, centered=False)
            st.markdown("### 📄 Protokół Odbioru Serwisu i Zalecenia dla Klienta")
            st.markdown(f"**Klient:** {t_row['Klient']} | **Adres:** {t_row['Kod_Pocztowy']} {t_row['Miasto']}, {t_row['Adres']} | **Data:** {datetime.date.today()}")
            if przeglad_summary_lines:
                st.markdown(f"**Wyniki pomiarów przeglądu:** `{przeglad_summary_lines[0]}`")

            client_recommendations = st.text_area(
                "Zalecenia po-serwisowe dla Klienta (widoczne na protokole):",
                value=str(t_row["Zalecenia_Klienta"]),
                placeholder="Wpisz zalecenia dotyczące eksploatacji instalacji, ciśnienia w układzie lub terminu kolejnego przeglądu..."
            )
            st.markdown(f"#### Do zapłaty / Rozliczenie: **{final_price:,.2f} zł** ({typ_serwisu_real})")
            st.markdown("**Podpis Klienta (potwierdzenie wykonania prac i zapoznania się z zaleceniami):**")

            canvas_result = st_canvas(
                fill_color="rgba(255, 255, 255, 0)",
                stroke_width=3,
                stroke_color="#0f172a",
                background_color="#ffffff",
                height=180,
                width=620,
                drawing_mode="freedraw",
                key=f"canvas_{selected_id}"
            )
            st.markdown('</div>', unsafe_allow_html=True)

            if st.button("✅ ZAKOŃCZ ZLECENIE I WRÓĆ DO TRASY Z TEGO DNIA", type="primary"):
                if not confirm_photos or not uploaded_photos:
                    st.error("🚨 BLOKADA ZAMKNIĘCIA: Musisz załączyć zdjęcie tabliczki znamionowej / instalacji oraz zaznaczyć pole potwierdzenia wykonania wszystkich wymaganych zdjęć!")
                else:
                    df.loc[t_idx, "Status"] = "Zakończone"
                    df.loc[t_idx, "Typ_Serwisu"] = typ_serwisu_real
                    df.loc[t_idx, "Naprawione_1_Wizyta"] = "Tak" if fixed_first_visit.startswith("Tak") else "Nie"
                    df.loc[t_idx, "Kwota_Koncowa"] = float(final_price)
                    df.loc[t_idx, "Czas_Pracy_Min"] = float(work_minutes)
                    df.loc[t_idx, "Wyniki_Przegladu"] = "; ".join(przeglad_summary_lines)
                    df.loc[t_idx, "Zdjecia_Potwierdzone"] = f"Tak ({len(uploaded_photos)} zdj.)"
                    df.loc[t_idx, "Zalecenia_Klienta"] = str(client_recommendations)
                    df.loc[t_idx, "Notatka_Serwisanta"] = str(internal_note)
                    df.loc[t_idx, "Data_Zamkniecia"] = str(datetime.date.today())

                    podpisane = "Nie"
                    try:
                        if canvas_result is not None and getattr(canvas_result, "image_data", None) is not None:
                            podpisane = "Tak"
                    except Exception:
                        podpisane = "Nie"

                    df.loc[t_idx, "Podpis_Zlozony"] = podpisane
                    save_tickets(df)
                    st.session_state["active_ticket_id"] = None
                    st.success("Zlecenie zamknięte! Wracasz do pozostałych klientów z dzisiejszej trasy.")
                    st.rerun()

# ==========================================
# 7. PANEL KOORDYNATORA (NOWY PLANER TRAS Z WYBOREM ZLECEŃ "Z PALCA")
# ==========================================
elif st.session_state["logged_role"] == "koordynator":
    df = st.session_state["df_tickets"]
    profiles = st.session_state["profiles"]

    tab_plan, tab_kadry, tab_zlecenia, tab_brand = st.tabs([
        "🗺️ Harmonogram i Planowanie Tras",
        "👷 Profilowanie Serwisantów (Kadry i Wyniki)",
        "📋 Baza Zgłoszeń i Edycja",
        "🎨 Branding (Logo) i Plik Excel"
    ])

    with tab_plan:
        st.subheader("🗺️ Planowanie wyjazdów dla serwisanta (Algorytm + Wybór Ręczny)")
        st.markdown("Wybierz serwisanta, podaj interesujący Cię kod pocztowy (lub pozostaw puste), a następnie zaznacz zlecenia z okolicy, które chcesz mu przypisać na dany dzień (w tym tematy krytyczne, które musisz wcisnąć poza algorytmem).")

        p_col1, p_col2, p_col3 = st.columns(3)
        with p_col1:
            sel_tech_plan = st.selectbox("Wybierz serwisanta:", list(profiles.keys()), key="plan_tech_sel")
        with p_col2:
            target_postal = st.text_input("Wpisz kod pocztowy okolicy (np. 05-800 lub 05):", value="", placeholder="np. 05-800")
        with p_col3:
            plan_date = st.date_input("Data planowanej trasy:", value=datetime.date.today() + datetime.timedelta(days=1))

        # Filtrowanie dostępnych nowych zgłoszeń
        pending_df = df[df["Status"].isin(["Nowe", "Do zaplanowania", ""])].copy()
        if target_postal.strip() != "":
            mask_postal = pending_df["Kod_Pocztowy"].astype(str).str.contains(target_postal.strip(), case=False, na=False) | \
                          pending_df["Adres"].astype(str).str.contains(target_postal.strip(), case=False, na=False) | \
                          pending_df["Miasto"].astype(str).str.contains(target_postal.strip(), case=False, na=False)
            filtered_pending = pending_df[mask_postal]
            if filtered_pending.empty:
                st.warning(f"Brak nowych zgłoszeń pasujących do zapytania '{target_postal}'. Poniżej wyświetlam wszystkie nierozliczone zgłoszenia z bazy:")
                filtered_pending = pending_df
        else:
            filtered_pending = pending_df

        st.markdown(f"#### 📍 Dostępne zlecenia do zaplanowania ({len(filtered_pending)} poz.):")
        st.caption("Zaznacz ptaszkiem w kolumnie **Wybierz** te zlecenia, które chcesz wysłać wybranemu serwisantowi na dany dzień.")

        if not filtered_pending.empty:
            # Tworzymy tabelę interaktywną z polem wyboru
            display_plan_df = filtered_pending[["ID_Zgloszenia", "Klient", "Kod_Pocztowy", "Miasto", "Adres", "Typ_Instalacji", "Priorytet_Krytyczny", "Kwota_Przewidywana", "Opis_Zgloszenia"]].copy()
            display_plan_df.insert(0, "Wybierz", False)
            
            edited_plan_df = st.data_editor(
                display_plan_df,
                use_container_width=True,
                hide_index=True,
                key="editor_manual_routes"
            )

            selected_ids = edited_plan_df[edited_plan_df["Wybierz"] == True]["ID_Zgloszenia"].tolist()

            col_btn_auto, col_btn_man = st.columns(2)
            with col_btn_auto:
                if st.button("🚀 Uruchom Auto-Planer klastrów (50km / 1200zł)", type="secondary"):
                    proposals, msg = generate_smart_routes(df, profiles, plan_date, days_ahead=3, target_tech=sel_tech_plan)
                    st.session_state["route_proposals"] = proposals
                    st.session_state["route_msg"] = msg
                    st.rerun()

            with col_btn_man:
                if st.button(f"✅ PRZYPISZ ZAZNACZONE ZLECENIA ({len(selected_ids)}) DLA: {sel_tech_plan}", type="primary"):
                    if not selected_ids:
                        st.error("Nie zaznaczono żadnego zlecenia do przypisania!")
                    else:
                        for tid in selected_ids:
                            idx_match = df[df["ID_Zgloszenia"] == tid].index[0]
                            df.loc[idx_match, "Serwisant"] = sel_tech_plan
                            df.loc[idx_match, "Data_Planowana"] = str(plan_date)
                            df.loc[idx_match, "Status"] = "Zaplanowane"
                        save_tickets(df)
                        st.success(f"Pomyślnie przypisano {len(selected_ids)} zleceń do serwisanta {sel_tech_plan} na dzień {plan_date}!")
                        st.rerun()
        else:
            st.info("Brak nowych zleceń w bazie oczekujących na zaplanowanie.")

        # Obsługa propozycji z auto-planera
        if "route_msg" in st.session_state and st.session_state["route_msg"]:
            st.info(st.session_state["route_msg"])

        proposals = st.session_state.get("route_proposals", [])
        if proposals:
            st.markdown("### 🤖 Propozycje z inteligentnego Auto-Planera:")
            for i, prop in enumerate(proposals):
                with st.container(border=True):
                    st.markdown(
                        f"#### 📅 Dzień: `{prop['data']}` | 👷 Serwisant: **{prop['serwisant']}** | "
                        f"📍 Punkty: **{prop['count']}/4** ({prop['miasta']}) | "
                        f"💰 Przewidywany przychód: **{prop['total_rev']:,.0f} zł** | 🔥 Krytyczne: **{prop['crit_count']}**"
                    )
                    sub_df = df.loc[prop["indices"], ["ID_Zgloszenia", "Klient", "Kod_Pocztowy", "Miasto", "Adres", "Typ_Instalacji", "Opis_Zgloszenia", "Kwota_Przewidywana"]]
                    st.dataframe(sub_df, use_container_width=True, hide_index=True)

            if st.button("✅ ZATWIERDŹ WSZYSTKIE PROPOZYCJE AUTO-PLANERA", type="primary"):
                for prop in proposals:
                    for idx in prop["indices"]:
                        df.loc[idx, "Serwisant"] = prop["serwisant"]
                        df.loc[idx, "Data_Planowana"] = prop["data"]
                        df.loc[idx, "Status"] = "Zaplanowane"
                save_tickets(df)
                st.session_state["route_proposals"] = []
                st.success("Wszystkie wygenerowane trasy zostały zatwierdzone!")
                st.rerun()

    with tab_kadry:
        st.subheader("👷 Kartoteka Kadrowa, Uprawnienia i Wyniki Serwisantów")
        selected_tech = st.selectbox("Wybierz serwisanta do edycji profilu:", list(profiles.keys()))
        t_prof = profiles[selected_tech]
        t_metrics = calculate_tech_metrics(df, selected_tech)

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Obsłużone zgłoszenia", t_metrics["total_closed"], f"Płatne: {t_metrics['paid_count']} | Gwar.: {t_metrics['warranty_count']}")
        k2.metric("Naprawione przy 1. wizycie", f"{t_metrics['ftfr_pct']}%")
        k3.metric("Obrót w tym miesiącu", f"{t_metrics['monthly_revenue']:,.0f} zł", t_metrics["active_rate_str"])
        k4.metric("Brakuje do progów (3% / 4%)", f"I: {t_metrics['missing_tier_1']:,.0f} zł", f"II: {t_metrics['missing_tier_2']:,.0f} zł")

        st.divider()

        with st.form(f"form_profile_{selected_tech}"):
            fc1, fc2, fc3 = st.columns(3)
            with fc1:
                new_full_name = st.text_input("Pełne imię, nazwisko i stanowisko:", value=t_prof.get("pelne_dane", ""))
                new_phone = st.text_input("Telefon służbowy:", value=t_prof.get("telefon", ""))
            with fc2:
                new_pin = st.text_input("Stałe 4-cyfrowe hasło (PIN):", value=str(t_prof.get("pin", "")))
                umowa_opts = ["Umowa o pracę", "B2B", "Umowa zlecenie", "Inna"]
                curr_umowa = t_prof.get("umowa", "Umowa o pracę")
                new_umowa = st.selectbox("Rodzaj umowy:", umowa_opts, index=umowa_opts.index(curr_umowa) if curr_umowa in umowa_opts else 0)
            with fc3:
                new_base_sal = st.number_input("Wynagrodzenie zasadnicze (zł):", min_value=0.0, step=100.0, value=float(t_prof.get("wynagrodzenie_bazowe", 6000.0)))

            bh1, bh2 = st.columns(2)
            with bh1:
                new_bhp_date = st.text_input("Data szkolenia BHP (RRRR-MM-DD):", value=t_prof.get("bhp_data", ""))
                new_bhp_prot = st.text_input("Numer i status protokołu BHP:", value=t_prof.get("bhp_protokol", ""))
            with bh2:
                new_med_date = st.text_input("Data badań medycyny pracy (RRRR-MM-DD):", value=t_prof.get("medycyna_data", ""))
                new_med_scope = st.text_area("Zakres badań medycyny pracy:", value=t_prof.get("medycyna_zakres", ""))

            new_secret_notes = st.text_area("Poufne notatki Koordynatora / Zarządu:", value=t_prof.get("notatki_poufne", ""))

            if st.form_submit_button("💾 Zapisz kartotekę serwisanta", type="primary"):
                profiles[selected_tech]["pelne_dane"] = new_full_name
                profiles[selected_tech]["telefon"] = new_phone
                profiles[selected_tech]["pin"] = new_pin
                profiles[selected_tech]["umowa"] = new_umowa
                profiles[selected_tech]["wynagrodzenie_bazowe"] = new_base_sal
                profiles[selected_tech]["bhp_data"] = new_bhp_date
                profiles[selected_tech]["bhp_protokol"] = new_bhp_prot
                profiles[selected_tech]["medycyna_data"] = new_med_date
                profiles[selected_tech]["medycyna_zakres"] = new_med_scope
                profiles[selected_tech]["notatki_poufne"] = new_secret_notes
                save_config_and_profiles(profiles)
                st.success(f"Zapisano profil serwisanta {selected_tech}!")

        st.markdown("#### 🏖️ Urlopy i Zwolnienia Lekarskie (L4)")
        u_col1, u_col2, u_col3, u_col4 = st.columns(4)
        with u_col1:
            ab_type = st.selectbox("Typ nieobecności:", ["Urlop wypoczynkowy", "Zwolnienie lekarskie (L4)", "Szkolenie / Inne"])
        with u_col2:
            ab_from = st.date_input("Od dnia:", value=datetime.date.today(), key="ab_from")
        with u_col3:
            ab_to = st.date_input("Do dnia:", value=datetime.date.today() + datetime.timedelta(days=3), key="ab_to")
        with u_col4:
            st.write("")
            if st.button("➕ Dodaj nieobecność"):
                profiles[selected_tech].setdefault("urlopy_l4", []).append({"typ": ab_type, "od": str(ab_from), "do": str(ab_to)})
                save_config_and_profiles(profiles)
                st.rerun()

        if t_prof.get("urlopy_l4"):
            st.table(pd.DataFrame(t_prof["urlopy_l4"]))
            if st.button("🗑️ Wyczyść historię urlopów/L4 tego serwisanta"):
                profiles[selected_tech]["urlopy_l4"] = []
                save_config_and_profiles(profiles)
                st.rerun()

    with tab_zlecenia:
        col_add, col_edit = st.columns(2)
        with col_add:
            st.subheader("➕ Dodaj nowe zgłoszenie")
            with st.expander("Rozwiń formularz nowego zgłoszenia", expanded=False):
                with st.form("new_ticket_form"):
                    n_klient = st.text_input("Imię i nazwisko / Firma klienta:")
                    n_tel = st.text_input("Telefon (możesz wpisać dwa po ukośniku np. 500111222 / 600333444):")
                    n_kod = st.text_input("Kod pocztowy (XX-XXX):", placeholder="np. 05-800")
                    n_miasto = st.text_input("Miejscowość:")
                    n_adres = st.text_input("Ulica i numer:")
                    n_typ_inst = st.selectbox("Urządzenie / Typ instalacji:", [
                        "Pompa Ciepła ARGOM", "Falownik Deye (PV / Hybryda)", "Falownik GoodWe (PV / Hybryda)",
                        "Kolektory Słoneczne (Solar)", "Pompa Ciepła + PV", "Magazyn Energii", "Inne"
                    ])
                    n_przeglad = st.selectbox("Czy wykonać pełny przegląd urządzenia (Protokół Przeglądu)?", ["Nie", "Tak"])
                    n_kryt = st.selectbox("Czy zgłoszenie krytyczne (Awaria)?", ["Nie", "Tak"])
                    n_typ_serw = st.selectbox("Rodzaj serwisu:", ["Płatny", "Gwarancyjny"])
                    n_kwota = st.number_input("Przewidywany przychód (zł):", min_value=0.0, step=50.0, value=650.0)
                    n_tech = st.selectbox("Przypisz serwisanta (lub Auto-Planer):", ["Auto-Planer"] + list(profiles.keys()))
                    n_data_plan = st.date_input("Data planowana (przy ręcznym przypisaniu):", value=datetime.date.today())
                    n_opis = st.text_area("Opis usterki / zakres prac:")

                    if st.form_submit_button("➕ Zapisz nowe zgłoszenie", type="primary"):
                        new_id = f"ZGL-{len(df) + 1001}"
                        new_row = {
                            "ID_Zgloszenia": new_id, "Klient": n_klient, "Adres": n_adres,
                            "Kod_Pocztowy": n_kod, "Miasto": n_miasto, "Telefon": n_tel,
                            "Typ_Instalacji": n_typ_inst, "Wymagany_Przeglad": n_przeglad,
                            "Opis_Zgloszenia": n_opis, "Priorytet_Krytyczny": n_kryt,
                            "Typ_Serwisu": n_typ_serw,
                            "Serwisant": "" if n_tech == "Auto-Planer" else n_tech,
                            "Data_Planowana": "" if n_tech == "Auto-Planer" else str(n_data_plan),
                            "Status": "Nowe" if n_tech == "Auto-Planer" else "Zaplanowane",
                            "Naprawione_1_Wizyta": "Tak", "Kwota_Przewidywana": float(n_kwota),
                            "Kwota_Koncowa": 0.0, "Czas_Pracy_Min": 0.0, "Wyniki_Przegladu": "",
                            "Zdjecia_Potwierdzone": "Nie", "Zalecenia_Klienta": "",
                            "Notatka_Serwisanta": "", "Data_Zamkniecia": "", "Podpis_Zlozony": "Nie"
                        }
                        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                        save_tickets(df)
                        st.success(f"Dodano zgłoszenie {new_id}!")
                        st.rerun()

        with col_edit:
            st.subheader("✏️ Szybka edycja wybranego zgłoszenia")
            with st.expander("Rozwiń formularz edycji zgłoszenia", expanded=False):
                if not df.empty:
                    edit_opts = [f"{r['ID_Zgloszenia']} | {r['Miasto']} - {r['Klient']}" for _, r in df.iterrows()]
                    sel_edit = st.selectbox("Wybierz zgłoszenie do edycji:", edit_opts)
                    sel_edit_id = sel_edit.split(" | ")[0]
                    e_idx = df[df["ID_Zgloszenia"] == sel_edit_id].index[0]
                    e_row = df.loc[e_idx]

                    with st.form("quick_edit_ticket_form"):
                        e_klient = st.text_input("Klient:", value=str(e_row["Klient"]))
                        e_tel = st.text_input("Telefon (np. dwa numery po /):", value=str(e_row["Telefon"]))
                        e_kod = st.text_input("Kod pocztowy:", value=str(e_row["Kod_Pocztowy"]))
                        e_miasto = st.text_input("Miasto:", value=str(e_row["Miasto"]))
                        e_adres = st.text_input("Adres:", value=str(e_row["Adres"]))
                        e_inst = st.text_input("Urządzenie / Instalacja:", value=str(e_row["Typ_Instalacji"]))
                        e_przeg = st.selectbox("Wymagany przegląd z protokołem:", ["Nie", "Tak"], index=1 if str(e_row.get("Wymagany_Przeglad", "Nie")) == "Tak" else 0)

                        tech_choices = [""] + list(profiles.keys())
                        curr_tech = str(e_row["Serwisant"]) if str(e_row["Serwisant"]) in tech_choices else ""
                        e_tech = st.selectbox("Przypisany serwisant:", tech_choices, index=tech_choices.index(curr_tech))

                        status_choices = ["Nowe", "Zaplanowane", "W trakcie", "Zakończone"]
                        curr_stat = str(e_row["Status"]) if str(e_row["Status"]) in status_choices else "Nowe"
                        e_status = st.selectbox("Status:", status_choices, index=status_choices.index(curr_stat))

                        e_data_plan = st.text_input("Data planowana (RRRR-MM-DD):", value=str(e_row["Data_Planowana"]))
                        e_kryt = st.selectbox("Krytyczne:", ["Nie", "Tak"], index=1 if str(e_row["Priorytet_Krytyczny"]) == "Tak" else 0)
                        e_kwota = st.number_input("Kwota przewidywana (zł):", min_value=0.0, step=50.0, value=float(e_row["Kwota_Przewidywana"]))
                        e_opis = st.text_area("Opis zgłoszenia:", value=str(e_row["Opis_Zgloszenia"]))

                        if st.form_submit_button("💾 Zapisz zmiany w zgłoszeniu", type="primary"):
                            df.loc[e_idx, "Klient"] = e_klient
                            df.loc[e_idx, "Telefon"] = e_tel
                            df.loc[e_idx, "Kod_Pocztowy"] = e_kod
                            df.loc[e_idx, "Miasto"] = e_miasto
                            df.loc[e_idx, "Adres"] = e_adres
                            df.loc[e_idx, "Typ_Instalacji"] = e_inst
                            df.loc[e_idx, "Wymagany_Przeglad"] = e_przeg
                            df.loc[e_idx, "Serwisant"] = e_tech
                            df.loc[e_idx, "Status"] = e_status
                            df.loc[e_idx, "Data_Planowana"] = e_data_plan
                            df.loc[e_idx, "Priorytet_Krytyczny"] = e_kryt
                            df.loc[e_idx, "Kwota_Przewidywana"] = float(e_kwota)
                            df.loc[e_idx, "Opis_Zgloszenia"] = e_opis
                            save_tickets(df)
                            st.success(f"Zaktualizowano zgłoszenie {sel_edit_id}!")
                            st.rerun()

        st.divider()
        st.subheader("📋 Wykaz wszystkich zgłoszeń w bazie")
        f_col1, f_col2, f_col3 = st.columns(3)
        with f_col1:
            filter_status = st.multiselect("Filtruj po statusie:", ["Nowe", "Zaplanowane", "W trakcie", "Zakończone"], default=["Nowe", "Zaplanowane", "W trakcie", "Zakończone"])
        with f_col2:
            filter_tech = st.multiselect("Filtruj po serwisancie:", [""] + list(profiles.keys()), default=[])
        with f_col3:
            search_q = st.text_input("🔍 Szukaj (Klient / Miasto / Adres / ID):")

        view_df = df.copy()
        if filter_status:
            view_df = view_df[view_df["Status"].isin(filter_status)]
        if filter_tech:
            view_df = view_df[view_df["Serwisant"].isin(filter_tech)]
        if search_q:
            mask = (
                view_df["Klient"].astype(str).str.contains(search_q, case=False, na=False) |
                view_df["Miasto"].astype(str).str.contains(search_q, case=False, na=False) |
                view_df["Adres"].astype(str).str.contains(search_q, case=False, na=False) |
                view_df["ID_Zgloszenia"].astype(str).str.contains(search_q, case=False, na=False)
            )
            view_df = view_df[mask]

        st.dataframe(view_df, use_container_width=True, hide_index=True)

    with tab_brand:
        st.subheader("🎨 Branding Aplikacji – Wgraj Logo Firmy")
        uploaded_logo = st.file_uploader("Wybierz plik graficzny z logo (PNG / JPG):", type=["png", "jpg", "jpeg"])
        if uploaded_logo is not None:
            img = Image.open(uploaded_logo)
            img.thumbnail((400, 120))
            buf = BytesIO()
            img.save(buf, format="PNG")
            b64_str = base64.b64encode(buf.getvalue()).decode("utf-8")
            st.session_state["logo_b64"] = b64_str
            with open(LOGO_FILE, "wb") as f:
                f.write(buf.getvalue())
            save_config_and_profiles(profiles, b64_str)
            st.success("Logo zostało zapisane!")
            st.rerun()

        st.divider()
        st.subheader("📥 Import / Podmiana bazy z pliku Excel (.xlsx)")
        uploaded_excel = st.file_uploader("Wgraj plik Excel z bazą zgłoszeń:", type=["xlsx"])
        if uploaded_excel is not None:
            if st.button("⚠️ Wczytaj dane z tego pliku Excel i zapisz w Google Sheets"):
                new_df = ensure_columns(pd.read_excel(uploaded_excel))
                save_tickets(new_df)
                st.success("Baza zgłoszeń została zaimportowana!")
                st.rerun()

# ==========================================
# 8. PANEL ZARZĄDU
# ==========================================
elif st.session_state["logged_role"] == "zarzad":
    df = st.session_state["df_tickets"]
    profiles = st.session_state["profiles"]

    st.subheader("📊 Panel Zarządu – Raporty Operacyjne i Finansowe")
    rep_tab_day, rep_tab_month = st.tabs(["📅 Raport Dzienny z Pracy Serwisantów", "📈 Raport Miesięczny i Prowizje"])

    with rep_tab_day:
        sel_date = st.date_input("Wybierz dzień raportu:", value=datetime.date.today())
        sel_date_str = str(sel_date)

        day_df = df[(df["Data_Zamkniecia"] == sel_date_str) | (df["Data_Planowana"] == sel_date_str)].copy()
        closed_day = day_df[day_df["Status"] == "Zakończone"]

        dc1, dc2, dc3, dc4 = st.columns(4)
        dc1.metric("Przychód w wybranym dniu", f"{closed_day['Kwota_Koncowa'].sum():,.2f} zł")
        dc2.metric("Zlecenia zamknięte dzisiaj", f"{len(closed_day)} / {len(day_df)}")
        ftfr_day = round((len(closed_day[closed_day["Naprawione_1_Wizyta"] == "Tak"]) / len(closed_day) * 100.0), 1) if len(closed_day) > 0 else 100.0
        dc3.metric("Skuteczność 1. wizyty (Dzień)", f"{ftfr_day}%")
        dc4.metric("Łączny czas pracy w terenie", f"{int(closed_day['Czas_Pracy_Min'].sum())} min")

        day_rows = []
        for t_name in profiles.keys():
            t_day = closed_day[closed_day["Serwisant"] == t_name]
            day_rows.append({
                "Serwisant": t_name,
                "Zamknięte zlecenia": len(t_day),
                "Serwisy Płatne": len(t_day[t_day["Typ_Serwisu"] == "Płatny"]),
                "Serwisy Gwarancyjne": len(t_day[t_day["Typ_Serwisu"] == "Gwarancyjny"]),
                "Naprawione od ręki (1. wizyta)": len(t_day[t_day["Naprawione_1_Wizyta"] == "Tak"]),
                "Czas pracy (min)": int(t_day["Czas_Pracy_Min"].sum()),
                "Przychód (zł)": float(t_day["Kwota_Koncowa"].sum())
            })
        st.dataframe(pd.DataFrame(day_rows), use_container_width=True, hide_index=True)
        st.dataframe(day_df, use_container_width=True, hide_index=True)

    with rep_tab_month:
        months_avail = sorted(list(set(
            [datetime.date.today().strftime("%Y-%m")] +
            [str(x)[:7] for x in df["Data_Zamkniecia"].unique() if len(str(x)) >= 7]
        )), reverse=True)
        sel_month = st.selectbox("Wybierz miesiąc rozliczeniowy (RRRR-MM):", months_avail)
        m_df = df[(df["Status"] == "Zakończone") & (df["Data_Zamkniecia"].astype(str).str.startswith(sel_month))].copy()

        mc1, mc2, mc3, mc4 = st.columns(4)
        mc1.metric("Przychód miesięczny ogółem", f"{m_df['Kwota_Koncowa'].sum():,.2f} zł")
        mc2.metric("Obsłużone zgłoszenia", len(m_df))
        mc3.metric("Płatne vs Gwarancyjne", f"{len(m_df[m_df['Typ_Serwisu']=='Płatny'])} / {len(m_df[m_df['Typ_Serwisu']=='Gwarancyjny'])}")
        ftfr_m = round((len(m_df[m_df["Naprawione_1_Wizyta"] == "Tak"]) / len(m_df) * 100.0), 1) if len(m_df) > 0 else 100.0
        mc4.metric("Skuteczność 1. wizyty (Miesiąc)", f"{ftfr_m}%")

        month_summary = []
        for t_name, t_prof in profiles.items():
            tm = calculate_tech_metrics(df, t_name)
            base_sal = float(t_prof.get("wynagrodzenie_bazowe", 0.0))
            month_summary.append({
                "Serwisant": t_name,
                "Umowa": t_prof.get("umowa", ""),
                "Zlecenia ogółem": tm["total_closed"],
                "Płatne": tm["paid_count"],
                "Gwarancyjne": tm["warranty_count"],
                "Skuteczność 1. wizyty (%)": f"{tm['ftfr_pct']}%",
                "Obrót miesięczny (zł)": round(tm["monthly_revenue"], 2),
                "Brakuje do 3% (15 tys.)": round(tm["missing_tier_1"], 2),
                "Brakuje do 4% (30 tys.)": round(tm["missing_tier_2"], 2),
                "Podstawa (zł)": base_sal,
                "Należna prowizja (zł)": round(tm["commission_pln"], 2),
                "Łącznie do wypłaty (zł)": round(base_sal + tm["commission_pln"], 2)
            })
        st.dataframe(pd.DataFrame(month_summary), use_container_width=True, hide_index=True)
