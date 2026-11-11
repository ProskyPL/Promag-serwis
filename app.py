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
# 1. KONFIGURACJA STRONY I RESPONSYWNY CSS (TELEFON + TABLET)
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
        min-height: 52px;
        font-size: 16px !important;
        font-weight: 600 !important;
        border-radius: 10px !important;
        margin-bottom: 6px;
        transition: all 0.2s;
    }
    .stSelectbox, .stTextInput, .stNumberInput, .stDateInput {
        font-size: 16px !important;
    }
    .kpi-card {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        color: #ffffff;
        padding: 16px;
        border-radius: 12px;
        border-left: 6px solid #10b981;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.15);
        margin-bottom: 12px;
    }
    .kpi-card-warn {
        border-left: 6px solid #f59e0b;
    }
    .kpi-title {
        font-size: 13px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        color: #94a3b8;
        margin-bottom: 4px;
    }
    .kpi-value {
        font-size: 24px;
        font-weight: 800;
        color: #f8fafc;
    }
    .kpi-sub {
        font-size: 13px;
        color: #38bdf8;
        margin-top: 4px;
    }
    .day-card {
        background-color: #1e293b;
        color: #f8fafc;
        border-radius: 12px;
        padding: 16px;
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
    .diag-box {
        background-color: #0f172a;
        color: #f8fafc;
        border-left: 6px solid #38bdf8;
        border-radius: 10px;
        padding: 16px;
        margin-top: 8px;
        margin-bottom: 12px;
    }
    @media (max-width: 768px) {
        .kpi-value { font-size: 20px; }
        .block-container { padding-top: 1.5rem !important; padding-left: 0.8rem !important; padding-right: 0.8rem !important; }
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. STAŁE, CENNIK, MAPA KODÓW I BAZA DIAGNOSTYCZNA (DEYE, GOODWE, ARGOM)
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
        "medycyna_zakres": "Praca powyżej 3m, pomiary elektryczne do 1kV, prawo jazdy kat. B",
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
        "medycyna_zakres": "Praca na wysokości pow. 3m, SEP E+D, F-gazy",
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
        "pelne_dane": "Wakat (Serwisant 4) - zmień imię w panelu Kadry",
        "telefon": "",
        "umowa": "Umowa zlecenie",
        "wynagrodzenie_bazowe": 6000.0,
        "bhp_data": "2026-01-01",
        "bhp_protokol": "Brak",
        "medycyna_data": "2026-01-01",
        "medycyna_zakres": "Do uzupełnienia",
        "urlopy_l4": [],
        "notatki_poufne": "Konto przygotowane dla 4. serwisanta."
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
    "Czyszczenie filtrów i układu CO": 400.0,
    "Pomiary elektryczne + Protokół": 650.0,
    "Modernizacja / Rozbudowa rozdzielnicy": 950.0,
    "Wymiana glikolu / płukanie układu": 1200.0,
    "Serwis Gwarancyjny (Rozliczenie wewnętrzne)": 0.0
}

# Szczegółowa baza diagnostyczna i kodów błędów ("krok po kroku")
DIAGNOSTIC_DB = {
    "Pompa Ciepła Argom (Monoblok / Split)": {
        "E03 / E08 (Flow Switch Error - Błąd przepływu wody)": {
            "przyczyna": "Woda lub glikol krąży zbyt wolno przez wymiennik płytowy pompy ciepła. Sterownik odcina sprężarkę, aby nie zamrozić i nie rozerwać wymiennika.",
            "kroki": [
                "KROK 1 (Manometr): Spójrz na manometr układu CO. Jeśli ciśnienie jest poniżej 1.0 bar – dopuść wody/glikolu do 1.5 - 1.8 bar.",
                "KROK 2 (Filtry - 90% wszystkich awarii!): Zakręć zawory odcinające przed i za filtrem siatkowym (skośnym) oraz separatorem magnetycznym. Podstaw wiadro, odkręć korek, wyjmij sitko i trzpień magnetyczny. Wypłucz czarny szlam (magnetyt) pod bieżącą wodą, zamontuj z powrotem, odkręć zawory i odpowietrz układ.",
                "KROK 3 (Zamknięte obiegi): Jeśli instalacja nie ma bufora, a głowice na grzejnikach lub siłowniki podłogówki pozamykały wszystkie pętle – pompa nie ma gdzie tłoczyć wody! Odkręć min. 2 pętle na stałe lub wyreguluj zawór nadmiarowo-upustowy (bypass).",
                "KROK 4 (Pompa obiegowa): Przyłóż dłoń lub śrubokręt do korpusu pompy obiegowej (Wilo/Grundfos). Sprawdź czy świeci się dioda i czy wirnik pracuje. Jeśli pompa stoi – odpowietrz ją lub sprawdź miernikiem czy z płyty głównej wychodzi zasilanie 230V AC.",
                "KROK 5 (Czujnik przepływu - łopatka): Jeśli rury są gorące i przepływ fizycznie jest, wykręć czujnik przepływu (Flow Switch) i sprawdź, czy pod łopatkę nie weszły pakuły lub taśma teflonowa z montażu."
            ]
        },
        "E01 / P01 (High Pressure - Błąd wysokiego ciśnienia HP)": {
            "przyczyna": "Sprężarka tłoczy gorący gaz, ale woda w wymienniku nie odbiera tego ciepła. Ciśnienie freonu przekracza limit bezpieczeństwa i presostat HP wyłącza układ.",
            "kroki": [
                "KROK 1 (Ustal kiedy wybija błąd): Sprawdź czy błąd E01/P01 wyskakuje przy grzaniu domu (CO), czy przy grzaniu bojlera (CWU).",
                "KROK 2 (Jeśli wybija na CWU): Sprawdź zadaną temperaturę CWU na sterowniku. Jeśli klient ustawił 58-60°C na samej sprężarce – obniż nastawę dla sprężarki na 48-50°C, a dogrzewanie powyżej 50°C (np. Antylegionella) ustaw na grzałkę elektryczną.",
                "KROK 3 (Zawór 3-drogowy): Zdejmij siłownik z zaworu 3-drogowego (CO/CWU) i sprawdź, czy trzpień zaworu nie zaciął się w połowie drogi (zablokowany przepływ).",
                "KROK 4 (Odpowietrzenie wężownicy CWU): Odpowietrz górny króciec wężownicy przy zasobniku CWU.",
                "KROK 5 (Jeśli wybija na CO): Wyczyść filtr siatkowy i magnetyczny (jak w błędzie E03) oraz zwiększ bieg pompy obiegowej, aby różnica temperatur (Delta T) między zasilaniem a powrotem wynosiła 4-7°C."
            ]
        },
        "E02 / P02 (Low Pressure - Błąd niskiego ciśnienia LP)": {
            "przyczyna": "Zbyt niskie ciśnienie po stronie ssawnej sprężarki – brak odbioru ciepła z powietrza zewnętrznego (oblodzenie/wentylator) lub ubytek czynnika chłodniczego (wyciek gazu).",
            "kroki": [
                "KROK 1 (Kontrola parownika): Sprawdź tylny wymiennik lamelowy jednostki zewnętrznej. Jeśli jest całkowicie skuty lodem – polej go ciepłą wodą (NIGDY nie kłuj lodu śrubokrętem, bo przebijesz rurki miedziane!).",
                "KROK 2 (Taca ociekowa i Defrost): Udrożnij otwór odpływowy skroplin w dolnej tacy, sprawdź czy grzeje kabel grzewczy tacy i uruchom ręczne odszranianie (Defrost) ze sterownika.",
                "KROK 3 (Wentylator): Sprawdź, czy podczas pracy sprężarki obraca się wentylator jednostki zewnętrznej.",
                "KROK 4 (Kontrola szczelności F-gaz): Jeśli parownik jest czysty i wentylator kręci, a błąd LP wraca po kilku minutach – podepnij manometry chłodnicze i sprawdź śrubunki/kielichy detektorem gazu (szukaj tłustych plam oleju chłodniczego)."
            ]
        },
        "E04 / P04 (High Discharge Temp - Zbyt wysoka temp. tłoczenia sprężarki >110°C)": {
            "przyczyna": "Przegrzewanie głowicy sprężarki spowodowane zablokowanym elektronicznym zaworem rozprężnym (EEV) lub zbyt małą ilością czynnika chłodniczego w układzie.",
            "kroki": [
                "KROK 1 (Wtyczka EEV): Otwórz pokrywę elektroniki w jednostce zewnętrznej i dociśnij 5/6-pinową wtyczkę cewki zaworu rozprężnego (EEV) na płycie głównej.",
                "KROK 2 (Test cewki EEV): Wyłącz zasilanie pompy z bezpiecznika na 1 minutę i włącz ponownie – nasłuchuj przy zaworze EEV, czy słychać charakterystyczne 'tykanie' pozycjonowania krokowego silnika zaworu. Zmierz omomierzem opór uzwojeń cewki EEV (zazwyczaj kilkadziesiąt Ohm na parę pinów).",
                "KROK 3 (Ilość czynnika): Jeśli cewka i zawór EEV są sprawne, przyczyną jest ubytek freonu. Zlokalizuj nieszczelność, odzyskaj czynnik, wykonaj próbę azotową, próżnię i napełnij układ wagowo zgodnie z tabliczką znamionową Argom."
            ]
        },
        "E05 do E15 (Błędy czujników temperatury NTC - przerwa lub zwarcie)": {
            "przyczyna": "Płyta główna widzi nielogiczną rezystancję czujnika temperatury (np. -40°C przy przerwie w kablu lub +120°C przy zwarciu czujnika).",
            "kroki": [
                "KROK 1 (Podgląd temperatur na ekranie): Wejdź w menu 'Status / Parametry bieżące' na sterowniku Argom i odczytaj wszystkie temperatury po kolei.",
                "KROK 2 (Namierzenie uszkodzonej czujki): Znajdź pozycję pokazującą skrajną wartość (np. -39°C, kreski '--' albo 99°C na zimnej rurze).",
                "KROK 3 (Pomiar miernikiem): Wypnij wtyczkę tego czujnika z płyty głównej, ustaw miernik na pomiar rezystancji (kOhm) i zmierz opór na dwóch pinach czujnika.",
                "KROK 4 (Weryfikacja): W temperaturze pokojowej (ok. 20-25°C) czujnik NTC 10k powinien mieć opór 10 - 12.5 kOhm (a czujnik 5k ok. 5 - 6 kOhm). Jeśli miernik pokazuje 'OL' (przerwa) lub '0.0' (zwarcie) – wymień czujnik na nowy i osadź go w tulei z pastą termoprzewodzącą."
            ]
        },
        "E09 / Comm Error (Błąd komunikacji ekranu / hydroboxa z agregatem zewnętrznym)": {
            "przyczyna": "Brak wymiany danych po magistrali RS485 pomiędzy jednostką wewnętrzną a zewnętrzną.",
            "kroki": [
                "KROK 1 (Zasilanie agregatu): Sprawdź miernikiem na listwie zasilającej w jednostce zewnętrznej, czy agregat w ogóle dostaje 230V / 400V (często wybija osobny bezpiecznik od jednostki zewnętrznej).",
                "KROK 2 (Przewód RS485 A/B): Sprawdź dwużyłowy przewód komunikacyjny (zaciski A i B lub P i Q). Upewnij się, że żyła A idzie dokładnie do A, a B do B oraz że śrubki nie zaciskają izolacji kabla.",
                "KROK 3 (Zakłócenia): Sprawdź, czy kabel komunikacyjny nie biegnie w jednym korytku obok kabli siłowych AC lub kabli DC od fotowoltaiki bez ekranowania."
            ]
        }
    },
    "Falownik Deye (Hybrydowy SG04LP3 / SG01HP3 i Stringowy)": {
        "F08 (GFDI_Relay_Failure - Błąd przekaźnika uziemienia / prądu upływu)": {
            "przyczyna": "Falownik wykrył problem z uziemieniem obudowy lub zbyt wysokie napięcie między przewodem neutralnym (N) a ochronnym (PE).",
            "kroki": [
                "KROK 1 (Pomiar N-PE): Ustaw miernik na V AC i zmierz napięcie między szyną N a szyną PE w rozdzielnicy. Jeśli wynosi powyżej 5V-10V – w budynku jest słabe uziemienie lub brak uziemienia punktu podziału PEN.",
                "KROK 2 (Przewód PE): Sprawdź, czy żółto-zielony przewód uziemiający jest solidnie przykręcony do obudowy falownika i szyny PE.",
                "KROK 3 (Ustawienia wyspowe Deye): Wejdź w menu: Advanced Function (kod 7777) i sprawdź opcję 'Signal ISLAND MODE' (zaznacz ją, jeśli falownik pracuje z wyjściem awaryjnym LOAD i przekaźnik N-PE ma zwierać neutralny w trybie wyspowym)."
            ]
        },
        "F13 (Working_Mode_Change - Zmiana trybu pracy / Reset sieci)": {
            "przyczyna": "Zmiana parametrów sieci lub trybu pracy w menu; falownik przełącza wewnętrzne przekaźniki.",
            "kroki": [
                "KROK 1 (Odczekaj): Poczekaj 3 minuty – zazwyczaj falownik sam odlicza czas startu i wraca do normalnej pracy.",
                "KROK 2 (Twardy restart): Jeśli F13 wisi na stałe: wyłącz bezpiecznik AC (Grid), przełącz hebel DC na OFF, wyłącz baterię. Poczekaj 2 minuty aż ekran całkowicie zgaśnie. Włącz ponownie w kolejności: 1. Bateria -> 2. DC -> 3. AC."
            ]
        },
        "F18 (Tz_Ac_OverCurr_Fault - Przetężenie / Zwarcie po stronie AC)": {
            "przyczyna": "Nagły skok prądu na wyjściu AC lub przeciążenie/zwarcie w obwodzie rezerwowym LOAD (Backup).",
            "kroki": [
                "KROK 1 (Odłączenie wyjścia LOAD): Wyłącz bezpiecznik zasilający obwód awaryjny LOAD (Backup) oraz bezpiecznik sieciowy GRID.",
                "KROK 2 (Test bez obciążenia): Uruchom falownik na samej baterii i panelach PV (z całkowicie odpiętymi wtyczkami AC Grid i Load).",
                "KROK 3 (Diagnoza): Jeśli bez kabli AC falownik startuje czysto i daje 230V na zaciskach LOAD – w instalacji domowej klienta na wyjściu Backup jest zwarcie albo podpięto tam za duże obciążenie (np. pompa ciepła lub indukcja na 1 fazie). Jeśli błąd F18 wyskakuje nawet przy odpiętych kablach AC – uszkodzony moduł IGBT w falowniku (zgłoś wymianę gwarancyjną)."
            ]
        },
        "F20 (Tz_Dc_OverCurr_Fault - Przetężenie po stronie DC)": {
            "przyczyna": "Prąd z paneli PV lub z baterii przekroczył maksymalny prąd wejściowy falownika.",
            "kroki": [
                "KROK 1: Wyłącz hebel DC i wypnij stringi PV.",
                "KROK 2: Zmierz napięcie każdego stringu (Uoc) oraz prąd cęgami DC. Upewnij się, że nie połączono równolegle dwóch stringów pod pojedyncze wejście MPPT o niskim limicie prądowym.",
                "KROK 3: W menu Battery Setting zmniejsz maksymalny prąd ładowania/rozładowania baterii (np. do 50A) i zrestartuj falownik."
            ]
        },
        "F22 (Tz_EmergStop_Fault - Wyłącznik awaryjny PPOŻ / Brak zwory RSD)": {
            "przyczyna": "Otwarty obwód na stykach wyłącznika bezpieczeństwa (Rapid Shutdown / Emergency Stop) na płycie sterującej Deye.",
            "kroki": [
                "KROK 1: Odkręć dolną pokrywę komory przyłączeniowej falownika Deye.",
                "KROK 2: Zlokalizuj zieloną wielopinową kostkę sterującą i piny opisane w instrukcji jako RSD / Emergency Stop (zazwyczaj fabrycznie połączone krótkim mostkiem z drutu).",
                "KROK 3: Jeśli klient NIE MA zewnętrznego przycisku PPOŻ – sprawdź czy fabryczna zworka (mostek) nie wypadła lub nie poluzowała się w gnieździe. Dokręć ją płaskim wkrętakiem.",
                "KROK 4: Jeśli klient MA przycisk PPOŻ – sprawdź miernikiem na 'piszczyku' czy przycisk nie został wciśnięty przez klienta lub czy przewód do przycisku nie jest przerwany."
            ]
        },
        "F23 (Tz_GFCI_OC_Fault - Prąd upływu powyżej 300 mA)": {
            "przyczyna": "Upływ prądu z modułów PV lub kabli solarnych przez wilgoć do konstrukcji dachowej / ziemi.",
            "kroki": [
                "KROK 1: Wypnij wszystkie stringi PV z falownika. Podłącz tylko String 1 i zobacz czy falownik ruszy. Następnie podłącz sam String 2 – w ten sposób od razu wiesz, który obwód na dachu ma przebicie.",
                "KROK 2: Na wadliwym stringu zmierz rezystancję izolacji (Riso) miernikiem przy 1000V DC (Plus do PE, Minus do PE). Szukaj wtyczki MC4 leżącej w wodzie na membranie dachowej lub przetartego kabla pod klemą.",
                "KROK 3: Jeśli wybija zewnętrzna różnicówka (RCD) w rozdzielnicy – upewnij się, że zamontowano RCD o prądzie 300mA (a nie domową 30mA!)."
            ]
        },
        "F24 (DC_Insulation_Fault - Niska rezystancja izolacji DC Riso)": {
            "przyczyna": "Zwarcie bieguna dodatniego (+) lub ujemnego (-) stringu PV do uziemionej ramy paneli lub uszkodzony ochronnik przepięć SPD DC.",
            "kroki": [
                "KROK 1 (Test ochronników SPD DC): Wyłącz DC, otwórz rozdzielnicę DC i wyjmij wkładki ochronników przepięć (SPD). Włącz falownik – jeśli błąd F24 zniknął, zawilgocona wkładka SPD zwierała do ziemi (wymień wkładkę).",
                "KROK 2 (Pomiar napięć do ziemi): Wypnij stringi z falownika. Zmierz miernikiem napięcie DC między PLUSEM stringu a szyną PE (uziemieniem) oraz MINUSEM a PE.",
                "KROK 3 (Lokalizacja zwarcia): W zdrowym stringu napięcie do ziemi szybko spada w okolice kilku-kilkunastu woltów. Jeśli na jednym biegunie trzyma sztywno np. 250V-400V, a na drugim 0V – masz twarde zwarcie kabla lub złącza MC4 do konstrukcji na dachu."
            ]
        },
        "F29 (Parallel_Comm_Fault - Błąd komunikacji pracy równoległej)": {
            "przyczyna": "W menu falownika włączono tryb pracy kaskadowej kilku falowników (Master/Slave), mimo że zamontowany jest tylko jeden falownik.",
            "kroki": [
                "KROK 1: Wejdź na ekranie dotykowym Deye w koło zębate -> Advanced Function (hasło: 7777).",
                "KROK 2: Przejdź do sekcji 'Parallel' i ODZNACZ ptaszek przy 'Parallel', 'Master' oraz 'Slave'.",
                "KROK 3: Zatwierdź przyciskiem OK w prawym dolnym rogu ekranu – błąd F29 zniknie natychmiast."
            ]
        },
        "F35 (AC_NoUtility_Fault - Brak sieci AC / Brak fazy lub złe wirowanie faz)": {
            "przyczyna": "Falownik nie widzi poprawnego napięcia 3-fazowego na kostce GRID lub fazy L1, L2, L3 są podłączone w złej kolejności.",
            "kroki": [
                "KROK 1 (Pomiar na śrubach GRID): Zmierz miernikiem napięcia bezpośrednio na śrubach kostki GRID wewnątrz falownika: L1-N, L2-N, L3-N (musi być 210-250V) oraz L1-L2, L2-L3, L1-L3 (ok. 400V). Jeśli brakuje fazy – sprawdź bezpiecznik w rozdzielnicy i przy liczniku.",
                "KROK 2 (Kolejność faz / Wirowanie): Falowniki 3-fazowe Deye wymagają zgodnego wirowania faz! Jeśli napięcia są obecne, a nadal wisi F35 – wyłącz bezpiecznik AC i zamień miejscami przewody L1 oraz L2 na kostce GRID w falowniku."
            ]
        },
        "F56 / F58 (DC_VoltLow_Fault / BMS_Comm_Fault - Problem z baterią / komunikacją BMS)": {
            "przyczyna": "Bateria rozładowana poniżej progu odcięcia, wyłączony stycznik BMS lub brak komunikacji po kablu CAN/RS485 między Deye a magazynem energii.",
            "kroki": [
                "KROK 1 (Pomiar napięcia na klemach BAT): Zmierz napięcie DC na grubych śrubach BAT+ i BAT- w falowniku. Jeśli jest 0V – włącz bezpiecznik na baterii i przytrzymaj przycisk SW/RST na magazynie przez 5 sekund (wybudzenie Black Start).",
                "KROK 2 (Weryfikacja portu BMS): Sprawdź czy skrętka RJ45 od baterii jest wpięta w falowniku w gniazdo opisane jako 'BMS' (częsty błąd to wpięcie w gniazdo Meter lub DRM).",
                "KROK 3 (Kod protokołu Lithium Mode): Wejdź w Battery Setting -> zaznacz 'Lithium' -> kliknij strzałkę w dół i sprawdź 'Lithium Mode': wpisz '00' (dla protokołu CAN Pylontech/Deye) lub '12' (dla RS485).",
                "KROK 4 (Awaryjne doładowanie): Jeśli bateria jest głęboko rozładowana i nie podaje komunikacji, przełącz tymczasowo w Battery Setting na 'Use Batt V' (praca napięciowa), zaznacz ładowanie z sieci 'Grid Charge', podładuj baterię przez 30 minut i przełącz z powrotem na 'Lithium'."
            ]
        }
    },
    "Falownik GoodWe (Serie SDT G2, MS, DNS, ET Plus Hybrydowy)": {
        "Isolation Failure (Kod 16 - Błąd rezystancji izolacji DC < 1 MOhm)": {
            "przyczyna": "Przebicie izolacji kabla solarnego, złącza MC4 lub ochronnika SPD po stronie DC do uziemienia.",
            "kroki": [
                "KROK 1: Wyłącz AC i DC. Wypnij wszystkie stringi z dolnej obudowy GoodWe.",
                "KROK 2: Podłącz tylko String 1 i włącz falownik. Jeśli ruszy bez błędu – uszkodzenie jest na Stringu 2 (i odwrotnie).",
                "KROK 3: Sprawdź wkładki ochronnika przepięć (SPD) DC w skrzynce zabezpieczeń oraz zmierz napięcie DC+ do PE i DC- do PE na uszkodzonym stringu.",
                "KROK 4: Sprawdź czy przewód PE we wtyczce AC falownika jest dobrze dokręcony."
            ]
        },
        "Vac Failure / Grid Overvoltage (Napięcie sieci poza zakresem > 253V AC)": {
            "przyczyna": "Napięcie w sieci energetycznej przekroczyło 253V (średnia 10-min) lub 260V (chwilowe), co powoduje wyłączenie falownika.",
            "kroki": [
                "KROK 1 (Pomiar i styk): Zmierz napięcia L1-N, L2-N, L3-N. Dokręć wszystkie śruby w rozdzielnicy PV oraz rozdzielnicy głównej (szczególnie przewód neutralny N!).",
                "KROK 2 (Aplikacja SolarGo): Połącz się z falownikiem przez aplikację SolarGo (Wi-Fi / Bluetooth). Zaloguj się jako Instalator (hasło domyślne: goodwe2010).",
                "KROK 3 (Aktywacja mocy biernej Q(U)): Wejdź w Ustawienia Zaawansowane -> Regulacja Mocy Biernej (Reactive Power) -> włącz krzywą Q(U). Falownik zacznie dociążać sieć mocą bierną indukcyjną powyżej 248V i sam obniży napięcie na fazach o 3-6V!"
            ]
        },
        "Utility Loss / No Grid (Brak sieci elektrycznej AC)": {
            "kroki": [
                "KROK 1: Wyłącz bezpiecznik AC w rozdzielnicy i wypnij dużą wtyczkę AC z dołu falownika GoodWe.",
                "KROK 2: Rozkręć obudowę wtyczki AC i sprawdź, czy któryś z przewodów (L1, L2, L3, N) nie wysunął się z zacisku śrubowego (bardzo częsta przyczyna w złączach wtykowych GoodWe!).",
                "KROK 3: Zmierz napięcie fazowe na stykach wtyczki (musi być 230V na każdej fazie względem N)."
            ],
            "przyczyna": "Brak napięcia na jednej lub wszystkich fazach na wtyczce AC falownika."
        },
        "Relay Check Failure (Błąd autotestu przekaźników sieciowych)": {
            "przyczyna": "Zwarcie między przewodem N a PE w instalacji budynku lub sklejony przekaźnik na płycie AC falownika.",
            "kroki": [
                "KROK 1: Zmierz napięcie między N a PE oraz sprawdź czy w rozdzielnicy nie pomylono przewodu N z PE po stronie za wyłącznikiem różnicowoprądowym.",
                "KROK 2: Odłącz całkowicie zasilanie AC i DC od falownika na pełne 5 minut i uruchom ponownie.",
                "KROK 3: Jeśli napięcia L1/L2/L3-N wynoszą równo ~230V, a falownik przy odliczaniu stuka przekaźnikami i wyrzuca Relay Check Failure – uszkodzony przekaźnik wewnętrzny (zgłoś do serwisu GoodWe)."
            ]
        },
        "PV Over Voltage (Zbyt wysokie napięcie wejściowe DC > 1000V/1100V)": {
            "przyczyna": "Zbyt duża liczba paneli połączonych w jeden szereg (string) – w chłodny dzień napięcie przekracza limit falownika.",
            "kroki": [
                "KROK 1 (BHP!): Nie rozpinaj wtyczek MC4 pod obciążeniem! Wyłącz AC i hebel DC.",
                "KROK 2: Zmierz miernikiem napięcie otwartego obwodu (Uoc) na wtyczkach MC4.",
                "KROK 3: Jeśli napięcie przekracza 980V – rozdziel string na dwa wejścia MPPT lub odepnij 1-2 moduły z szeregu."
            ]
        },
        "BMS Loss / Battery Comm Failure (Seria hybrydowa GoodWe ET)": {
            "przyczyna": "Brak komunikacji CAN między falownikiem GoodWe ET a magazynem energii.",
            "kroki": [
                "KROK 1: Otwórz aplikację SolarGo -> Ustawienia baterii -> sprawdź czy wybrano dokładny model podłączonego magazynu energii.",
                "KROK 2: Sprawdź czy użyto oryginalnego kabla komunikacyjnego CAN z kompletu falownika/baterii oraz czy na ostatnim module bateryjnym wpięta jest zworka terminująca (120 Ohm).",
                "KROK 3: Zrestartuj kolumnę bateryjną oraz falownik."
            ]
        }
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
# 3. FUNKCJE POMOCNICZE (TELEFONY, DIAGNOSTYKA, GOOGLE SHEETS, LOGO)
# ==========================================
def extract_phone_numbers(raw_phone_str):
    """Wyciąga wszystkie numery telefonów z tekstu (obsługuje 2 lub więcej numerów po przecinku/ukośniku/spacji)."""
    if not raw_phone_str:
        return []
    text = str(raw_phone_str)
    # Dzielimy po popularnych separatorach wielu numerów
    parts = re.split(r'[;/,|lub]+', text)
    phones = []
    for p in parts:
        digits = re.sub(r'[^\d+]', '', p)
        if len(digits) >= 9:
            phones.append(digits)
    # Jeśli ktoś wpisał dwa 9-cyfrowe numery ciągiem ze spacją bez ukośnika (np. "500100200 600200300")
    if len(phones) == 1 and len(re.sub(r'\D', '', phones[0])) >= 18:
        all_d = re.sub(r'\D', '', phones[0])
        phones = [all_d[i:i+9] for i in range(0, len(all_d), 9) if len(all_d[i:i+9]) == 9]
    return list(dict.fromkeys(phones))

def render_phone_and_sms_buttons(raw_phone_str, address_str, key_prefix=""):
    """Generuje dla każdego wykrytego numeru parę przycisków: 📞 Zadzwoń oraz 💬 SMS."""
    phones = extract_phone_numbers(raw_phone_str)
    if not phones:
        st.markdown("📞 **Telefon:** `Brak numeru w bazie`")
        return

    sms_default = f"Dzien dobry, tu serwis PROMAG OZE. Kontaktuje sie w sprawie zgloszenia serwisowego pod adresem {address_str}."
    sms_encoded = sms_default.replace(" ", "%20")

    for idx, ph in enumerate(phones, 1):
        label_suffix = f" ({idx})" if len(phones) > 1 else ""
        c_call, c_sms = st.columns(2)
        with c_call:
            st.link_button(f"📞 Zadzwoń{label_suffix}: {ph}", f"tel:{ph}", use_container_width=True)
        with c_sms:
            st.link_button(f"💬 SMS{label_suffix}: {ph}", f"sms:{ph}?body={sms_encoded}", use_container_width=True)

def render_diagnostic_wizard(default_install_type="", ticket_key="global"):
    """Wyświetla interaktywną ścieżkę diagnostyczną i bazę kodów błędów krok po kroku."""
    st.markdown('<div class="diag-box">', unsafe_allow_html=True)
    st.markdown("### 🩺 Interaktywna Ścieżka Diagnostyczna i Kody Błędów")
    st.caption("Wybierz urządzenie oraz kod błędu lub objaw z wyświetlacza, aby otrzymać dokładną procedurę pomiarową i naprawczą krok po kroku.")

    cat_options = list(DIAGNOSTIC_DB.keys())
    default_idx = 0
    if "fotowoltaika" in str(default_install_type).lower() or "pv" in str(default_install_type).lower():
        default_idx = 1

    sel_cat = st.selectbox("1️⃣ Wybierz markę / typ urządzenia:", cat_options, index=default_idx, key=f"diag_cat_{ticket_key}")
    codes_dict = DIAGNOSTIC_DB[sel_cat]
    sel_code = st.selectbox("2️⃣ Wybierz kod błędu lub objaw usterki:", list(codes_dict.keys()), key=f"diag_code_{ticket_key}")

    diag_data = codes_dict[sel_code]
    st.warning(f"**🔍 Przyczyna techniczna:** {diag_data['przyczyna']}")
    st.markdown("#### 🛠️ Procedura diagnostyczna i naprawcza (zrób po kolei):")
    for step_txt in diag_data["kroki"]:
        st.markdown(f"- {step_txt}")

    st.markdown('</div>', unsafe_allow_html=True)
    return f"[{sel_cat} -> {sel_code}]"

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

# ==========================================
# 4. PANCERNE WCZYTYWANIE I CZYSZCZENIE BAZY ZGŁOSZEŃ
# ==========================================
REQUIRED_COLS = [
    "ID_Zgloszenia", "Klient", "Adres", "Kod_Pocztowy", "Miasto", "Telefon",
    "Typ_Instalacji", "Opis_Zgloszenia", "Wymagany_Przeglad", "Priorytet_Krytyczny", "Typ_Serwisu",
    "Serwisant", "Data_Planowana", "Status", "Naprawione_1_Wizyta",
    "Kwota_Przewidywana", "Kwota_Koncowa", "Czas_Pracy_Min",
    "Zalecenia_Klienta", "Notatka_Serwisanta", "Wyniki_Przegladu", "Data_Zamkniecia", "Podpis_Zlozony"
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
        if ("klient" in cl or "nazwisko" in cl or "imie" in cl or "imię" in cl or "inwestor" in cl):
            target = "Klient"
        elif ("adres" in cl or "ulica" in cl):
            target = "Adres"
        elif ("kod" in cl and "poczt" in cl):
            target = "Kod_Pocztowy"
        elif ("miasto" in cl or "miejscowo" in cl):
            target = "Miasto"
        elif ("tel" in cl):
            target = "Telefon"
        elif ("opis" in cl or "usterk" in cl or "uwag" in cl or "zgłosz" in cl or "zglosz" in cl):
            target = "Opis_Zgloszenia"

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
            elif rc == "Wymagany_Przeglad":
                df[rc] = [
                    "Tak" if "przegl" in str(r.get("Opis_Zgloszenia", "")).lower() else "Nie"
                    for _, r in df.iterrows()
                ]
            elif rc == "Priorytet_Krytyczny":
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

    empty_id_mask = df["ID_Zgloszenia"].str.strip() == ""
    if empty_id_mask.any():
        df.loc[empty_id_mask, "ID_Zgloszenia"] = [f"ZGL-{i+1001}" for i in df.index[empty_id_mask]]

    empty_status_mask = df["Status"].str.strip() == ""
    if empty_status_mask.any():
        df.loc[empty_status_mask, "Status"] = "Nowe"

    empty_przeglad_mask = df["Wymagany_Przeglad"].str.strip() == ""
    if empty_przeglad_mask.any():
        df.loc[empty_przeglad_mask, "Wymagany_Przeglad"] = "Nie"

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
                "Kod_Pocztowy": "05-800", "Miasto": "Pruszków", "Telefon": "500100200 / 600999888",
                "Typ_Instalacji": "Pompa Ciepła Argom", "Opis_Zgloszenia": "Brak grzania CO - błąd E03 przepływu (Awaria)",
                "Wymagany_Przeglad": "Nie", "Priorytet_Krytyczny": "Tak", "Typ_Serwisu": "Płatny", "Serwisant": "Leszek",
                "Data_Planowana": str(datetime.date.today()), "Status": "Zaplanowane",
                "Naprawione_1_Wizyta": "Tak", "Kwota_Przewidywana": 850.0, "Kwota_Koncowa": 0.0,
                "Czas_Pracy_Min": 0.0, "Zalecenia_Klienta": "", "Notatka_Serwisanta": "", "Wyniki_Przegladu": "",
                "Data_Zamkniecia": "", "Podpis_Zlozony": "Nie"
            },
            {
                "ID_Zgloszenia": "ZGL-1002", "Klient": "Anna Nowak", "Adres": "ul. Lipowa 4",
                "Kod_Pocztowy": "05-840", "Miasto": "Brwinów", "Telefon": "600200300",
                "Typ_Instalacji": "Pompa Ciepła Argom + PV Deye", "Opis_Zgloszenia": "Przegląd roczny przed sezonem + badanie glikolu",
                "Wymagany_Przeglad": "Tak", "Priorytet_Krytyczny": "Nie", "Typ_Serwisu": "Płatny", "Serwisant": "Leszek",
                "Data_Planowana": str(datetime.date.today()), "Status": "Zaplanowane",
                "Naprawione_1_Wizyta": "Tak", "Kwota_Przewidywana": 650.0, "Kwota_Koncowa": 0.0,
                "Czas_Pracy_Min": 0.0, "Zalecenia_Klienta": "", "Notatka_Serwisanta": "", "Wyniki_Przegladu": "",
                "Data_Zamkniecia": "", "Podpis_Zlozony": "Nie"
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
# 5. ALGORYTM PLANOWANIA TRAS I KLASTRÓW
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

def generate_smart_routes(df, profiles, start_date, days_ahead=5):
    pending = df[df["Status"].isin(["Nowe", "Do zaplanowania", ""])].copy()
    if pending.empty:
        return [], "Brak nowych, nieprzypisanych zgłoszeń w bazie."

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
            valid_clusters.append({
                "indices": cluster_indices,
                "total_rev": total_rev,
                "crit_count": crit_count,
                "miasta": ", ".join([m for m in cluster_df["Miasto"].astype(str).unique() if m]),
                "count": len(cluster_indices)
            })

    if not valid_clusters:
        return [], "Żaden z klastrów nie spełnił warunku minimum 1200 zł lub min. 2 zgłoszeń krytycznych w promieniu 50 km."

    tech_names = list(profiles.keys())
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

    return proposals, f"Wygenerowano {len(proposals)} opłacalnych tras (klastrów <=50 km)."

# ==========================================
# 6. OBLICZANIE WYNIKÓW I PROWIZJI SERWISANTA
# ==========================================
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
# 7. INICJALIZACJA STANU SESJI
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

# ==========================================
# 8. EKRAN GŁÓWNY LOGOWANIA + SZKOLENIE WSTĘPNE
# ==========================================
if st.session_state["logged_role"] is None:
    render_brand_logo(max_width=280, centered=True)
    st.markdown("<p style='text-align:center; color:#64748b; margin-top:-5px;'>Mobilny System Zarządzania Serwisem OZE (Telefon / Tablet)</p>", unsafe_allow_html=True)
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

        if st.button("🚀 Zaloguj się", type="primary"):
            if user_choice == "--- Wybierz z listy ---":
                st.error("Wybierz swój profil z listy powyżej.")
            elif user_choice == "Koordynator Serwisu":
                if pin_input == PIN_KOORDYNATOR:
                    st.session_state["logged_role"] = "koordynator"
                    st.session_state["logged_user"] = "Koordynator"
                    st.rerun()
                else:
                    st.error("Nieprawidłowy kod PIN Koordynatora.")
            elif user_choice == "Zarząd (Raporty)":
                if pin_input == PIN_ZARZAD:
                    st.session_state["logged_role"] = "zarzad"
                    st.session_state["logged_user"] = "Zarząd"
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
                    st.rerun()
                else:
                    st.error("Nieprawidłowy kod PIN serwisanta.")

    with col_onboard:
        st.subheader("🎓 Dla nowych pracowników")
        st.info("Rozpoczynasz pracę w naszym zespole serwisowym? Zapoznaj się z procedurami bezpieczeństwa i standardem obsługi klienta.")
        if st.button("📘 Otwórz Szkolenie Wstępne (Onboarding)"):
            st.session_state["show_onboarding"] = not st.session_state.get("show_onboarding", False)

        if st.session_state.get("show_onboarding", False):
            with st.container(border=True):
                st.markdown("### Moduł Szkolenia Wstępnego")
                st.markdown("""
                1. **Standard wizyty serwisowej** – weryfikacja układu, dokumentacja zdjęciowa przed i po, kultura pracy.
                2. **Procedury BHP przy OZE** – bezpieczne odłączanie strony DC/AC w instalacjach PV oraz praca z czynnikiem chłodniczym w pompach ciepła.
                3. **Obsługa aplikacji mobilnej** – każde zlecenie kończymy czytelnymi zaleceniami oraz podpisem klienta na ekranie tabletu/telefonu.
                
                *(Szczegółowe materiały wideo oraz testy wdrożeniowe zostaną opublikowane w tym module w kolejnym etapie).*
                """)

    st.stop()

# Pasek górny po zalogowaniu
top_c1, top_c2, top_c3 = st.columns([2, 2, 1])
with top_c1:
    render_brand_logo(max_width=200, centered=False)
with top_c2:
    st.markdown(f"**Zalogowany jako:** `{st.session_state['logged_user']}`")
with top_c3:
    if st.button("🚪 Wyloguj"):
        st.session_state["logged_role"] = None
        st.session_state["logged_user"] = None
        st.session_state["active_day"] = None
        st.session_state["active_ticket_id"] = None
        st.rerun()

st.divider()

# ==========================================
# 9. PANEL SERWISANTA (POGRUPOWANE DNI -> ROZPOCZNIJ DZIEŃ -> ROZPOCZNIJ ZLECENIE)
# ==========================================
if st.session_state["logged_role"] == "serwisant":
    tech_name = st.session_state["logged_user"]
    df = st.session_state["df_tickets"]
    metrics = calculate_tech_metrics(df, tech_name)

    # --- PASEK MOTYWACYJNY NA GÓRZE EKRANU SERWISANTA ---
    m_col1, m_col2, m_col3 = st.columns([1.4, 1.4, 1])
    with m_col1:
        if metrics["missing_tier_1"] > 0:
            headline = f"Brakuje: {metrics['missing_tier_1']:,.0f} zł do prowizji 3%"
            subline = f"Próg I (15 000 zł) | Do progu 4% (30 000 zł) brakuje: {metrics['missing_tier_2']:,.0f} zł"
        elif metrics["missing_tier_2"] > 0:
            headline = f"🔥 Masz już 3%! Brakuje: {metrics['missing_tier_2']:,.0f} zł do 4%"
            subline = f"Wypracowany obrót: {metrics['monthly_revenue']:,.0f} zł (Premia: {metrics['commission_pln']:,.0f} zł)"
        else:
            headline = f"🏆 Maksymalny próg 4% zdobyty!"
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
        if st.button("📚 BAZA WIEDZY / KODY BŁĘDÓW"):
            st.session_state["show_kb"] = not st.session_state.get("show_kb", False)

    progress_val = min(1.0, float(metrics["monthly_revenue"] / COMMISSION_TIER_2))
    st.progress(progress_val, text=f"Postęp obrotu w tym miesiącu: {metrics['monthly_revenue']:,.0f} zł / 30 000 zł ({metrics['active_rate_str']})")

    if st.session_state.get("show_kb", False):
        with st.expander("📚 Baza Wiedzy Technicznej OZE (Deye / GoodWe / Argom)", expanded=True):
            render_diagnostic_wizard(default_install_type="Pompa Ciepła Argom", ticket_key="top_kb")

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

        # =========================================================================
        # ETAP 1: WIDOK POGRUPOWANYCH DNI (GDY SERWISANT NIE KLIKNĄŁ "ROZPOCZNIJ DZIEŃ")
        # =========================================================================
        if st.session_state["active_day"] is None:
            st.subheader("📅 Twoje zaplanowane dni wyjazdowe (Trasy)")
            st.caption("Zlecenia są pogrupowane w optymalne trasy dzienne. Kliknij **Rozpocznij dzień**, aby odblokować obsługę klientów zaplanowanych na dany dzień.")

            unique_days = sorted(my_tickets["Data_Grupa"].unique())

            for day_str in unique_days:
                day_cluster = my_tickets[my_tickets["Data_Grupa"] == day_str]
                pts_count = len(day_cluster)
                cities_list = ", ".join([m for m in day_cluster["Miasto"].astype(str).unique() if m.strip() != ""])
                if not cities_list:
                    cities_list = "Lokalizacje wg adresów"
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
                            <div style="font-size:15px; margin-top:6px;">📍 <b>Miejscowości:</b> {cities_list}</div>
                            <div style="font-size:14px; color:#cbd5e1; margin-top:4px;">
                                🛠️ Punktów na trasie: <b>{pts_count}</b> &nbsp;|&nbsp;
                                💰 Przewidywany obrót: <b>{est_rev:,.0f} zł</b> &nbsp;|&nbsp;
                                🔥 Krytyczne: <b>{crit_cnt}</b>
                            </div>
                        </div>
                        """, unsafe_allow_html=True)

                        with st.expander(f"👁️ Podgląd punktów w trasie na dzień {day_str} ({pts_count})"):
                            for idx_pt, (_, r_pt) in enumerate(day_cluster.iterrows(), 1):
                                crit_tag = " 🔥 [AWARIA]" if r_pt["Priorytet_Krytyczny"] == "Tak" else ""
                                prz_tag = " 🔍 [PRZEGLĄD]" if str(r_pt.get("Wymagany_Przeglad", "")) == "Tak" else ""
                                st.markdown(
                                    f"**{idx_pt}. {r_pt['Miasto']}** ({r_pt['Kod_Pocztowy']}), {r_pt['Adres']} – "
                                    f"*{r_pt['Typ_Instalacji']}*{crit_tag}{prz_tag}"
                                )

                    with c_day_btn:
                        st.write("")
                        st.write("")
                        if st.button(f"🚀 ROZPOCZNIJ DZIEŃ ({day_str})", key=f"start_day_{day_str}", type="primary"):
                            st.session_state["active_day"] = day_str
                            st.session_state["active_ticket_id"] = None
                            st.rerun()

        # =========================================================================
        # ETAP 2: AKTYWNY DZIEŃ – WYBÓR KLIENTA Z DANEGO DNIA
        # =========================================================================
        elif st.session_state["active_day"] is not None and st.session_state["active_ticket_id"] is None:
            act_day = st.session_state["active_day"]
            day_tickets = my_tickets[my_tickets["Data_Grupa"] == act_day].copy()

            head_c1, head_c2 = st.columns([2.2, 1])
            with head_c1:
                st.subheader(f"🚚 Aktywny dzień pracy: {act_day}")
                st.caption("Kolejność i godziny wizyt ustalasz samodzielnie. Zadzwoń lub napisz SMS do klienta, a po przyjeździe kliknij **Rozpocznij zlecenie**.")
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
                    is_insp = (str(t_row.get("Wymagany_Przeglad", "Nie")) == "Tak")
                    status_badge = "🟡 W TRAKCIE REALIZACJI" if t_row["Status"] == "W trakcie" else "⏳ OCZEKUJE NA WIZYTĘ"
                    crit_badge = " 🔥 KRYTYCZNE / AWARIA" if is_crit else ""
                    insp_badge = " 🔍 WYMAGANY PRZEGLĄD" if is_insp else ""

                    with st.container(border=True):
                        st.markdown(f"#### Punkt #{idx_card}: **{t_row['Miasto']}** – {t_row['Klient']} `{status_badge}`{crit_badge}{insp_badge}")
                        col_d1, col_d2, col_d3 = st.columns([1.4, 1.4, 1.1])

                        with col_d1:
                            st.markdown(f"📍 **Adres:** `{t_row['Kod_Pocztowy']} {t_row['Miasto']}, {t_row['Adres']}`")
                            st.markdown(f"🔧 **Urządzenie:** {t_row['Typ_Instalacji']} ({t_row['Typ_Serwisu']})")
                            st.markdown(f"📝 **Zgłoszenie:** {t_row['Opis_Zgloszenia']}")

                        with col_d2:
                            render_phone_and_sms_buttons(t_row["Telefon"], f"{t_row['Miasto']}, {t_row['Adres']}", key_prefix=f"day_{t_id}")
                            maps_q = f"{t_row['Adres']} {t_row['Kod_Pocztowy']} {t_row['Miasto']}".replace(" ", "+")
                            st.link_button("🗺️ Nawiguj (Google Maps)", f"https://www.google.com/maps/search/?api=1&query={maps_q}", use_container_width=True)

                        with col_d3:
                            st.markdown(f"💰 **Szacowana kwota:** `{float(t_row['Kwota_Przewidywana']):,.0f} zł`")
                            st.write("")
                            if st.button(f"🛠️ ROZPOCZNIJ ZLECENIE ({t_id})", key=f"start_ticket_{t_id}", type="primary"):
                                df.loc[t_idx, "Status"] = "W trakcie"
                                save_tickets(df)
                                st.session_state["active_ticket_id"] = t_id
                                st.rerun()

        # =========================================================================
        # ETAP 3: AKTYWNE ZLECENIE (DIAGNOSTYKA + PROTOKÓŁ PRZEGLĄDU + ROZLICZENIE)
        # =========================================================================
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

            c_info1, c_info2 = st.columns([1.3, 1.2])
            with c_info1:
                st.markdown(f"### 👤 Klient: **{t_row['Klient']}**")
                st.markdown(f"📍 **Adres:** `{t_row['Kod_Pocztowy']} {t_row['Miasto']}, {t_row['Adres']}`")
                st.markdown(f"🔧 **Urządzenie:** **{t_row['Typ_Instalacji']}** | **Typ:** `{t_row['Typ_Serwisu']}`")
                st.markdown(f"📝 **Opis zgłoszenia:** {t_row['Opis_Zgloszenia']}")

            with c_info2:
                render_phone_and_sms_buttons(t_row["Telefon"], f"{t_row['Miasto']}, {t_row['Adres']}", key_prefix=f"act_{selected_id}")
                maps_query = f"{t_row['Adres']} {t_row['Kod_Pocztowy']} {t_row['Miasto']}".replace(" ", "+")
                st.link_button("🗺️ Nawiguj w Google Maps", f"https://www.google.com/maps/search/?api=1&query={maps_query}", use_container_width=True)

            st.divider()

            # --- 1. ŚCIEŻKA DIAGNOSTYCZNA DLA TEGO URZĄDZENIA ---
            diag_col1, diag_col2 = st.columns([1.5, 1])
            with diag_col1:
                st.markdown("#### 🩺 Nieoczywista usterka lub kod błędu na urządzeniu?")
            with diag_col2:
                if st.button("🩺 URUCHOM ŚCIEŻKĘ DIAGNOSTYCZNĄ KROK PO KROKU", key=f"btn_diag_{selected_id}"):
                    st.session_state[f"open_diag_{selected_id}"] = not st.session_state.get(f"open_diag_{selected_id}", False)

            selected_diag_summary = ""
            if st.session_state.get(f"open_diag_{selected_id}", False):
                selected_diag_summary = render_diagnostic_wizard(
                    default_install_type=str(t_row["Typ_Instalacji"]),
                    ticket_key=selected_id
                )

            st.divider()

            # --- 2. KALKULATOR CZYNNOŚCI I ROZLICZENIE ---
            st.subheader("🧾 Wykonane czynności i rozliczenie")
            calc_c1, calc_c2 = st.columns(2)
            with calc_c1:
                selected_services = st.multiselect(
                    "Wykonane czynności z cennika:",
                    options=list(CENNIK_BAZOWY.keys()),
                    default=[]
                )
                extra_parts = st.number_input("Koszt użytych części / materiałów (zł):", min_value=0.0, step=50.0, value=0.0)
                km_dojazd = st.number_input("Dojazd (km w obie strony powyżej ryczałtu):", min_value=0, step=10, value=0)

            with calc_c2:
                work_minutes = st.number_input(
                    "Czas pracy na miejscu (minuty):",
                    min_value=15,
                    step=15,
                    value=int(float(t_row["Czas_Pracy_Min"])) if float(t_row["Czas_Pracy_Min"]) > 0 else 60
                )
                typ_serwisu_real = st.selectbox("Kwalifikacja zlecenia:", ["Płatny", "Gwarancyjny"], index=0 if t_row["Typ_Serwisu"] != "Gwarancyjny" else 1)
                fixed_first_visit = st.radio(
                    "Czy usterka / zgłoszenie zostało w pełni rozwiązane podczas tej wizyty?",
                    ["Tak", "Nie (Wymagana kolejna wizyta / zamówienie części)"],
                    horizontal=True
                )

            # --- 3. PROTOKÓŁ PRZEGLĄDU URZĄDZENIA + POMIARY GLIKOLU (REFRAKTOMETR + PASEK PH) I WALIDACJA NA ŻYWO ---
            is_inspection_needed = (
                str(t_row.get("Wymagany_Przeglad", "Nie")) == "Tak"
                or "przegl" in str(t_row.get("Opis_Zgloszenia", "")).lower()
                or any("Przegląd" in s for s in selected_services)
            )

            inspection_summary_lines = []
            auto_recommendations_from_inspection = []

            st.markdown("#### 🔍 Protokół Przeglądu Technicznego Urządzenia")
            enable_inspection_form = st.checkbox(
                "Wypełnij protokół pomiarowy przeglądu instalacji (z automatyczną kontrolą norm)",
                value=is_inspection_needed,
                key=f"chk_insp_{selected_id}"
            )

            if enable_inspection_form:
                inst_type_lower = str(t_row["Typ_Instalacji"]).lower()
                show_pc = ("pompa" in inst_type_lower or "pc" in inst_type_lower or "argom" in inst_type_lower or "inne" in inst_type_lower)
                show_pv = ("pv" in inst_type_lower or "foto" in inst_type_lower or "deye" in inst_type_lower or "goodwe" in inst_type_lower or "magazyn" in inst_type_lower)
                if not show_pc and not show_pv:
                    show_pc, show_pv = True, True

                with st.container(border=True):
                    if show_pc:
                        st.markdown("##### 🌡️ A. Pomiary Pompy Ciepła + Badanie Glikolu (Refraktometr & Pasek pH)")
                        with st.expander("ℹ️ Jak poprawnie pobrać próbkę glikolu na refraktometr i pasek pH? (Instrukcja)", expanded=False):
                            st.markdown("""
                            1. **Pobranie próbki:** Upuść 2–3 krople płynu z zaworu do czystego naczynia (nie bierz pierwszej kropli z samego zaworu, bo bywa zabrudzona smarem lub osadem).
                            2. **Refraktometr:** Przetrzyj pryzmat czystą chusteczką, nanieś kroplę płynu i zamknij płytkę tak, aby nie było pęcherzyków powietrza. Odczytaj granicę niebiesko-białą na skali **Propylene Glycol (°C)**.
                            3. **Pasek pH:** Zanurz pasek w próbce na 2 sekundy, strzepnij nadmiar płynu i po 15 sekundach porównaj kolor ze skalą na opakowaniu.
                            """)

                        pc1, pc2, pc3 = st.columns(3)
                        with pc1:
                            p_co_bar = st.number_input("Ciśnienie statyczne układu CO [bar] (Norma: 1.2 - 2.2):", min_value=0.0, max_value=6.0, value=1.5, step=0.1, key=f"p_co_{selected_id}")
                            p_naczynie_bar = st.number_input("Ciśnienie naczynia przeponowego [bar] (Norma: 0.8 - 1.5):", min_value=0.0, max_value=4.0, value=1.2, step=0.1, key=f"p_nacz_{selected_id}")
                        with pc2:
                            t_zas = st.number_input("Temp. zasilania CO [°C]:", min_value=10.0, max_value=80.0, value=35.0, step=0.5, key=f"t_zas_{selected_id}")
                            t_pow = st.number_input("Temp. powrotu CO [°C]:", min_value=10.0, max_value=80.0, value=30.0, step=0.5, key=f"t_pow_{selected_id}")
                            delta_t = round(abs(t_zas - t_pow), 1)
                            st.caption(f"Wyliczona Delta T (ΔT): **{delta_t} °C** (Norma: 3.0 - 8.0 °C)")
                        with pc3:
                            glikol_temp = st.number_input("Temp. krzepnięcia glikolu (Refraktometr) [°C] (Norma: -15 do -30):", min_value=-50.0, max_value=0.0, value=-22.0, step=1.0, key=f"glik_t_{selected_id}")
                            glikol_ph = st.number_input("Odczyn pH zładu (Pasek pH) (Norma: 7.5 - 9.0):", min_value=3.0, max_value=13.0, value=8.0, step=0.1, key=f"glik_ph_{selected_id}")

                        pc_c4, pc_c5 = st.columns(2)
                        with pc_c4:
                            filtr_stan = st.selectbox("Stan filtra siatkowego i magnetycznego:", ["Czysty / Wyczyyszczono podczas przeglądu", "Silnie zanieczyszczony magnetytem (wymaga płukania)"], key=f"f_stan_{selected_id}")
                        with pc_c5:
                            fgaz_szczelnosc = st.selectbox("Kontrola szczelności układu chłodniczego i tacy skroplin:", ["Szczelny, odpływ skroplin drożny", "Wykryto nieszczelność / niedrożny odpływ"], key=f"fgaz_{selected_id}")

                        # --- WALIDACJA PARAMETRÓW POMPY CIEPŁA NA ŻYWO ---
                        if p_co_bar < 1.1:
                            st.error("🚨 **ODCHYLENIE OD NORMY (Ciśnienie CO < 1.1 bar):** Ryzyko błędu przepływu (E03/E08) oraz zapowietrzenia i kawitacji pompy obiegowej!\n\n"
                                     "**Procedura naprawy:** Sprawdź szczelność śrubunków, zweryfikuj ciśnienie wstępne naczynia przeponowego i dopuść zładu do wartości 1.5 - 1.8 bar.")
                            auto_recommendations_from_inspection.append(f"Skorygowano ciśnienie układu CO z {p_co_bar} bar do normy roboczej.")
                        elif p_co_bar > 2.4:
                            st.error("🚨 **ODCHYLENIE OD NORMY (Ciśnienie CO > 2.4 bar):** Ryzyko otwarcia zaworu bezpieczeństwa (3 bar) i zalania kotłowni!\n\n"
                                     "**Procedura naprawy:** Sprawdź naczynie przeponowe (pęknięta membrana lub brak powietrza pod wentylem) oraz szczelność zaworu dopuszczającego wodę. Upuść ciśnienie do 1.5 bar.")

                        if delta_t > 8.5:
                            st.error(f"🚨 **ODCHYLENIE OD NORMY (Delta T = {delta_t}°C > 8.5°C):** Zbyt wolny odbiór ciepła ze skraplacza (zdławiony przepływ)! Ryzyko błędu wysokiego ciśnienia HP (E01).\n\n"
                                     "**Procedura naprawy:** Wyczyść filtr siatkowy i separator magnetyczny, odkręć przepływomierze (rotametry) na rozdzielaczu i zwiększ bieg pompy obiegowej.")
                        elif delta_t < 2.0:
                            st.warning(f"⚠️ **UWAGA (Delta T = {delta_t}°C < 2.0°C):** Zbyt szybki przepływ krótki (tzw. podmieszanie na buforze lub otwarty bypass) albo sprężarka nie pracuje z pełną mocą.")

                        if glikol_temp > -15.0:
                            st.error(f"🚨 **ALARM GLIKOLOWY (Krzepnięcie = {glikol_temp}°C, norma minimum -15°C):** Zbyt rozcieńczony glikol! Przy awarii zasilania na mrozie dojdzie do zamarznięcia i rozerwania wymiennika płytowego!\n\n"
                                     "**Procedura naprawy krok po kroku:**\n"
                                     "1. Zakaz dolewania czystej wody do układu!\n"
                                     "2. Podłącz pompę serwisową z koncentratem glikolu propylenowego.\n"
                                     "3. Upuść część rozcieńczonego zładu i wtłocz koncentrat, wymieszaj obieg przez 15 minut i zmierz ponownie refraktometrem aż uzyskasz min. -18°C do -25°C.")
                            auto_recommendations_from_inspection.append(f"Wykryto zbyt słabe stężenie glikolu ({glikol_temp}°C) - wymagane dosycenie koncentratem glikolu propylenowego do min. -20°C.")
                        elif glikol_temp < -33.0:
                            st.warning(f"⚠️ **OSTRZEŻENIE GLIKOLOWE (Krzepnięcie = {glikol_temp}°C < -33°C):** Zbyt gęsty glikol! Duża lepkość płynu dławi pompę obiegową (ryzyko błędu przepływu E03) i pogarsza wymianę ciepła.\n\n"
                                       "**Procedura naprawy:** Upuść niewielką ilość zładu, dopuść wody demineralizowanej i wymieszaj układ do uzyskania ok. -22°C.")

                        if glikol_ph < 7.0:
                            st.error(f"🚨 **ALARM CHEMICZNY pH (pH = {glikol_ph} < 7.0):** Glikol uległ zakwaszeniu (degradacja termiczna)! Płyn stracił właściwości antykorozyjne i trawi wymiennik lutowany miedzią oraz uszczelnienia!\n\n"
                                     "**Procedura naprawy krok po kroku:**\n"
                                     "1. Zleć całkowite spuszczenie starego glikolu z układu.\n"
                                     "2. Wypłucz instalację czystą wodą, wyczyść separator magnetyczny.\n"
                                     "3. Napełnij układ świeżym glikolem propylenowym z pakietem inhibitorów korozji (pH 7.5 - 9.0).")
                            auto_recommendations_from_inspection.append(f"Wykryto zakwaszenie czynnika grzewczego (pH {glikol_ph}) - zalecana pełna wymiana glikolu i płukanie układu.")
                        elif glikol_ph > 9.5:
                            st.warning(f"⚠️ **OSTRZEŻENIE pH (pH = {glikol_ph} > 9.5):** Odczyn zbyt zasadowy (ryzyko korozji elementów aluminiowych i mosiężnych). Zastosuj korektor pH lub wymień część zładu na wodę demineralizowaną.")

                        inspection_summary_lines.append(
                            f"POMPA CIEPŁA: Ciśnienie CO={p_co_bar} bar, Naczynie={p_naczynie_bar} bar, Delta T={delta_t}°C (Zas:{t_zas}°C/Pow:{t_pow}°C), "
                            f"Glikol Refraktometr={glikol_temp}°C, Odczyn pH={glikol_ph}, Filtr={filtr_stan}, Układ={fgaz_szczelnosc}."
                        )

                    if show_pv:
                        st.markdown("##### ☀️ B. Pomiary Instalacji Fotowoltaicznej (PV) i Magazynu Energii")
                        pv1, pv2, pv3 = st.columns(3)
                        with pv1:
                            r_iso = st.number_input("Rezystancja izolacji stringów Riso [MOhm] (Norma >= 1.0):", min_value=0.0, max_value=500.0, value=50.0, step=0.5, key=f"riso_{selected_id}")
                            u_pe_ohm = st.number_input("Rezystancja uziemienia [Ohm] (Norma <= 10.0):", min_value=0.1, max_value=100.0, value=4.5, step=0.5, key=f"upe_{selected_id}")
                        with pv2:
                            v_l1 = st.number_input("Napięcie fazy L1-N [V] (Norma: 207-253V):", min_value=0.0, max_value=300.0, value=234.0, step=1.0, key=f"vl1_{selected_id}")
                            v_l2 = st.number_input("Napięcie fazy L2-N [V] (Norma: 207-253V):", min_value=0.0, max_value=300.0, value=235.0, step=1.0, key=f"vl2_{selected_id}")
                            v_l3 = st.number_input("Napięcie fazy L3-N [V] (Norma: 207-253V):", min_value=0.0, max_value=300.0, value=234.0, step=1.0, key=f"vl3_{selected_id}")
                        with pv3:
                            styki_stan = st.selectbox("Kontrola dokręcenia zacisków AC/DC i SPD:", ["Wszystkie styki dokręcone, SPD sprawne", "Wykryto luźne styki / ślady przegrzania"], key=f"styki_{selected_id}")
                            bms_stan = st.selectbox("Stan Magazynu Energii / SOH:", ["Nie dotyczy / Brak magazynu", "Sprawny (SOH > 90%, komunikacja CAN OK)", "Wymaga diagnostyki BMS"], key=f"bms_st_{selected_id}")

                        if r_iso < 1.0:
                            st.error(f"🚨 **ALARM IZOLACJI DC (Riso = {r_iso} MOhm < 1.0 MOhm):** Przebicie strony DC do uziemienia (błąd F24 w Deye / Isolation Failure w GoodWe)!\n\n"
                                     "**Procedura naprawy:** Wypnij wkładki ochronnika SPD DC i sprawdź czy Riso wraca do normy. Jeśli nie – rozepnij stringi i szukaj zawilgoconego złącza MC4 lub przetartego kabla pod konstrukcją na dachu.")
                            auto_recommendations_from_inspection.append(f"Niska rezystancja izolacji DC ({r_iso} MOhm) - wymagana weryfikacja złączy MC4 i ochronników SPD.")

                        if max(v_l1, v_l2, v_l3) >= 251.0:
                            st.warning(f"⚠️ **WYSOKIE NAPIĘCIE SIECI AC ({max(v_l1, v_l2, v_l3)} V):** Napięcie bliskie progu wyłączenia falownika (253V)!\n\n"
                                       "**Procedura naprawy:** Dokręć zaciski przewodu N i faz w rozdzielnicy. Włącz w ustawieniach falownika Deye/GoodWe regulację mocy biernej Q(U).")
                            auto_recommendations_from_inspection.append("Aktywowano charakterystykę mocy biernej Q(U) ze względu na podwyższone napięcie sieci AC.")

                        if u_pe_ohm > 10.0:
                            st.error(f"🚨 **ODCHYLENIE UZIEMIENIA ({u_pe_ohm} Ohm > 10 Ohm):** Zbyt wysoka rezystancja uziomu! Ryzyko błędu F08 oraz braku skuteczności ochronników przepięć SPD.\n\n"
                                     "**Procedura naprawy:** Sprawdź połączenie śrubowe w złączu kontrolnym, oczyść bednarkę z korozji i rozbuduj uziom szpilkowy.")

                        if "luźne" in styki_stan.lower():
                            st.error("🚨 **LUŹNE STYKI / PRZEGRZANIE W ROZDZIELNICY:** Natychmiast dokręć wszystkie zaciski śrubowe wyłączników nadprądowych, RCD i bloków rozdzielczych. Jeśli aparat modułowy ma nadtopioną obudowę – wymień go na nowy!")

                        inspection_summary_lines.append(
                            f"INSTALACJA PV: Riso={r_iso} MOhm, Uziemienie={u_pe_ohm} Ohm, Napięcia AC=({v_l1}/{v_l2}/{v_l3} V), "
                            f"Rozdzielnica={styki_stan}, Magazyn={bms_stan}."
                        )

            inspection_full_text = " | ".join(inspection_summary_lines)

            base_sum = sum(CENNIK_BAZOWY[s] for s in selected_services)
            if base_sum == 0 and typ_serwisu_real == "Płatny":
                base_sum = float(t_row["Kwota_Przewidywana"])
            auto_total = 0.0 if typ_serwisu_real == "Gwarancyjny" else (base_sum + extra_parts + (km_dojazd * 2.5))

            final_price = st.number_input("💰 Kwota końcowa do rozliczenia (zł brutto):", min_value=0.0, step=50.0, value=float(auto_total))

            default_internal_note = str(t_row["Notatka_Serwisanta"])
            if selected_diag_summary and selected_diag_summary not in default_internal_note:
                default_internal_note = f"{default_internal_note} Diagnostyka: {selected_diag_summary}".strip()

            internal_note = st.text_area("Notatka wewnętrzna dla Koordynatora (niewidoczna dla klienta):", value=default_internal_note)

            # --- 4. PROTOKÓŁ DLA KLIENTA Z LOGO FIRMY I PODPISEM ---
            st.markdown('<div class="client-box">', unsafe_allow_html=True)
            render_brand_logo(max_width=220, centered=False)
            st.markdown("### 📄 Protokół Odbioru Serwisu / Przeglądu i Zalecenia dla Klienta")
            st.markdown(f"**Klient:** {t_row['Klient']} | **Adres:** {t_row['Kod_Pocztowy']} {t_row['Miasto']}, {t_row['Adres']} | **Data:** {datetime.date.today()}")
            st.markdown(f"**Urządzenie:** {t_row['Typ_Instalacji']}")

            if inspection_full_text:
                st.info(f"**📊 Wyniki pomiarów z przeglądu:**\n{inspection_full_text}")

            default_client_rec = str(t_row["Zalecenia_Klienta"]).strip()
            if auto_recommendations_from_inspection and not default_client_rec:
                default_client_rec = " ".join(auto_recommendations_from_inspection)

            client_recommendations = st.text_area(
                "Zalecenia po-serwisowe dla Klienta (widoczne na protokole):",
                value=default_client_rec,
                placeholder="Wpisz zalecenia dotyczące eksploatacji instalacji, stężenia glikolu, ciśnienia w układzie lub terminu kolejnego przeglądu..."
            )
            st.markdown(f"#### Do zapłaty / Rozliczenie: **{final_price:,.2f} zł** ({typ_serwisu_real})")
            st.markdown("**Podpis Klienta (potwierdzenie wykonania prac, pomiarów i zapoznania się z zaleceniami):**")

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
                df.loc[t_idx, "Status"] = "Zakończone"
                df.loc[t_idx, "Typ_Serwisu"] = typ_serwisu_real
                df.loc[t_idx, "Naprawione_1_Wizyta"] = "Tak" if fixed_first_visit.startswith("Tak") else "Nie"
                df.loc[t_idx, "Kwota_Koncowa"] = float(final_price)
                df.loc[t_idx, "Czas_Pracy_Min"] = float(work_minutes)
                df.loc[t_idx, "Zalecenia_Klienta"] = str(client_recommendations)
                df.loc[t_idx, "Notatka_Serwisanta"] = str(internal_note)
                df.loc[t_idx, "Wyniki_Przegladu"] = str(inspection_full_text)
                df.loc[t_idx, "Data_Zamkniecia"] = str(datetime.date.today())
                # NOWY ZAPIS (stabilny i zabezpieczony):
podpisane = "Nie"
try:
    if canvas_result is not None and getattr(canvas_result, "image_data", None) is not None:
        podpisane = "Tak"
except Exception:
    podpisane = "Nie"

df.loc[t_idx, "Podpis_Zlozony"] = podpisane


                save_tickets(df)
                st.session_state["active_ticket_id"] = None
                st.success("Zlecenie i protokół zostały zapisane! Wracasz do pozostałych klientów z dzisiejszej trasy.")
                st.rerun()

# ==========================================
# 10. PANEL KOORDYNATORA
# ==========================================
elif st.session_state["logged_role"] == "koordynator":
    df = st.session_state["df_tickets"]
    profiles = st.session_state["profiles"]

    tab_plan, tab_kadry, tab_zlecenia, tab_brand = st.tabs([
        "🗺️ Harmonogram i Auto-Planer Tras",
        "👷 Profilowanie Serwisantów (Kadry i Wyniki)",
        "📋 Baza Zgłoszeń i Edycja",
        "🎨 Branding (Logo) i Plik Excel"
    ])

    with tab_plan:
        st.subheader("🗺️ Inteligentne Planowanie Wyjazdów (Cała Polska)")
        st.markdown("""
        **Zasady algorytmu:**
        * Brak sztywnej rejonizacji – serwisanci wyruszają z domu lub z hotelu z poprzedniego dnia.
        * Grupowanie do **maks. 4 punktów** oddalonych od siebie o **maks. 50 km**.
        * Trasa jest zatwierdzana, jeśli jej przewidywany przychód wynosi **min. 1200 zł** LUB zawiera **min. 2 zgłoszenia krytyczne** (np. awaria pompy ciepła).
        * Automatyczne wykluczanie serwisantów mających wpisany urlop lub L4 w danym dniu.
        """)

        p_col1, p_col2 = st.columns(2)
        with p_col1:
            plan_start_date = st.date_input("Planuj wyjazdy począwszy od dnia:", value=datetime.date.today() + datetime.timedelta(days=1))
        with p_col2:
            plan_days = st.slider("Liczba dni do przodu:", min_value=1, max_value=10, value=5)

        if st.button("🚀 Wygeneruj propozycje tras na kolejne dni", type="primary"):
            proposals, msg = generate_smart_routes(df, profiles, plan_start_date, plan_days)
            st.session_state["route_proposals"] = proposals
            st.session_state["route_msg"] = msg

        if "route_msg" in st.session_state:
            st.info(st.session_state["route_msg"])

        proposals = st.session_state.get("route_proposals", [])
        if proposals:
            for i, prop in enumerate(proposals):
                with st.container(border=True):
                    st.markdown(
                        f"#### 📅 Dzień: `{prop['data']}` | 👷 Serwisant: **{prop['serwisant']}** | "
                        f"📍 Punkty: **{prop['count']}/4** ({prop['miasta']}) | "
                        f"💰 Przewidywany przychód: **{prop['total_rev']:,.0f} zł** | 🔥 Krytyczne: **{prop['crit_count']}**"
                    )
                    sub_df = df.loc[prop["indices"], ["ID_Zgloszenia", "Klient", "Kod_Pocztowy", "Miasto", "Adres", "Typ_Instalacji", "Wymagany_Przeglad", "Priorytet_Krytyczny", "Kwota_Przewidywana"]]
                    st.dataframe(sub_df, use_container_width=True, hide_index=True)

            if st.button("✅ ZATWIERDŹ I PRZYPISZ WSZYSTKIE POWYŻSZE TRASY DO SERWISANTÓW", type="primary"):
                for prop in proposals:
                    for idx in prop["indices"]:
                        df.loc[idx, "Serwisant"] = prop["serwisant"]
                        df.loc[idx, "Data_Planowana"] = prop["data"]
                        df.loc[idx, "Status"] = "Zaplanowane"
                save_tickets(df)
                st.session_state["route_proposals"] = []
                st.success("Trasy zostały przypisane do harmonogramu serwisantów!")
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
            st.markdown("#### 1. Dane osobowe, hasło (PIN) i warunki zatrudnienia")
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
                new_base_sal = st.number_input("Wynagrodzenie zasadnicze / podstawa (zł):", min_value=0.0, step=100.0, value=float(t_prof.get("wynagrodzenie_bazowe", 6000.0)))

            st.markdown("#### 2. Szkolenia BHP oraz Badania Medycyny Pracy")
            bh1, bh2 = st.columns(2)
            with bh1:
                new_bhp_date = st.text_input("Data szkolenia BHP (RRRR-MM-DD):", value=t_prof.get("bhp_data", ""))
                new_bhp_prot = st.text_input("Numer i status protokołu BHP:", value=t_prof.get("bhp_protokol", ""))
            with bh2:
                new_med_date = st.text_input("Data badań medycyny pracy (RRRR-MM-DD):", value=t_prof.get("medycyna_data", ""))
                new_med_scope = st.text_area("Zakres badań medycyny pracy (np. wysokość >3m, SEP, kat. B):", value=t_prof.get("medycyna_zakres", ""))

            st.markdown("#### 3. Poufne notatki Koordynatora / Zarządu")
            new_secret_notes = st.text_area("Poufne notatki (widoczne wyłącznie dla Koordynatora i Zarządu):", value=t_prof.get("notatki_poufne", ""))

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
        u_col1, u_col2, u_col3, u_col4 = st.columns([1, 1, 1, 1])
        with u_col1:
            ab_type = st.selectbox("Typ nieobecności:", ["Urlop wypoczynkowy", "Zwolnienie lekarskie (L4)", "Szkolenie / Inne"])
        with u_col2:
            ab_from = st.date_input("Od dnia:", value=datetime.date.today(), key="ab_from")
        with u_col3:
            ab_to = st.date_input("Do dnia:", value=datetime.date.today() + datetime.timedelta(days=3), key="ab_to")
        with u_col4:
            st.write("")
            if st.button("➕ Dodaj nieobecność"):
                profiles[selected_tech].setdefault("urlopy_l4", []).append({
                    "typ": ab_type,
                    "od": str(ab_from),
                    "do": str(ab_to)
                })
                save_config_and_profiles(profiles)
                st.success("Dodano wpis o nieobecności (Auto-Planer automatycznie pominie te dni).")
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
            with st.expander("Rozwiń formularz dodawania nowego zgłoszenia", expanded=False):
                with st.form("new_ticket_form"):
                    n_klient = st.text_input("Imię i nazwisko / Firma klienta:")
                    n_tel = st.text_input("Telefon (jeśli dwa numery, wpisz po ukośniku np. 500111222 / 600333444):")
                    n_kod = st.text_input("Kod pocztowy (XX-XXX):", placeholder="np. 05-800")
                    n_miasto = st.text_input("Miejscowość:")
                    n_adres = st.text_input("Ulica i numer:")
                    n_typ_inst = st.selectbox(
                        "Urządzenie / Typ instalacji:",
                        [
                            "Pompa Ciepła Argom",
                            "Falownik Deye (PV / Hybryda)",
                            "Falownik GoodWe (PV / Hybryda)",
                            "Pompa Ciepła Argom + PV Deye",
                            "Pompa Ciepła Argom + PV GoodWe",
                            "Magazyn Energii",
                            "Inne urządzenie OZE"
                        ]
                    )
                    n_przeglad = st.selectbox("Czy konserwator ma wykonać PRZEGLĄD urządzenia (Protokół pomiarowy)?", ["Nie", "Tak"])
                    n_kryt = st.selectbox("Czy zgłoszenie krytyczne (np. awaria PC zimą)?", ["Nie", "Tak"])
                    n_typ_serw = st.selectbox("Rodzaj serwisu:", ["Płatny", "Gwarancyjny"])
                    n_kwota = st.number_input("Przewidywany przychód (zł):", min_value=0.0, step=50.0, value=650.0)
                    n_tech = st.selectbox("Przypisz serwisanta (lub Auto-Planer):", ["Auto-Planer"] + list(profiles.keys()))
                    n_data_plan = st.date_input("Data planowana (jeśli przypisujesz ręcznie):", value=datetime.date.today())
                    n_opis = st.text_area("Opis usterki / kod błędu / zakres prac:")

                    if st.form_submit_button("➕ Zapisz nowe zgłoszenie", type="primary"):
                        new_id = f"ZGL-{len(df) + 1001}"
                        new_row = {
                            "ID_Zgloszenia": new_id,
                            "Klient": n_klient,
                            "Adres": n_adres,
                            "Kod_Pocztowy": n_kod,
                            "Miasto": n_miasto,
                            "Telefon": n_tel,
                            "Typ_Instalacji": n_typ_inst,
                            "Opis_Zgloszenia": n_opis,
                            "Wymagany_Przeglad": n_przeglad,
                            "Priorytet_Krytyczny": n_kryt,
                            "Typ_Serwisu": n_typ_serw,
                            "Serwisant": "" if n_tech == "Auto-Planer" else n_tech,
                            "Data_Planowana": "" if n_tech == "Auto-Planer" else str(n_data_plan),
                            "Status": "Nowe" if n_tech == "Auto-Planer" else "Zaplanowane",
                            "Naprawione_1_Wizyta": "Tak",
                            "Kwota_Przewidywana": float(n_kwota),
                            "Kwota_Koncowa": 0.0,
                            "Czas_Pracy_Min": 0.0,
                            "Zalecenia_Klienta": "",
                            "Notatka_Serwisanta": "",
                            "Wyniki_Przegladu": "",
                            "Data_Zamkniecia": "",
                            "Podpis_Zlozony": "Nie"
                        }
                        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                        save_tickets(df)
                        st.success(f"Dodano zgłoszenie {new_id}!")
                        st.rerun()

        with col_edit:
            st.subheader("✏️ Szybka edycja wybranego zgłoszenia")
            with st.expander("Rozwiń formularz edycji istniejącego zgłoszenia", expanded=False):
                if not df.empty:
                    edit_opts = [f"{r['ID_Zgloszenia']} | {r['Miasto']} - {r['Klient']}" for _, r in df.iterrows()]
                    sel_edit = st.selectbox("Wybierz zgłoszenie do edycji:", edit_opts)
                    sel_edit_id = sel_edit.split(" | ")[0]
                    e_idx = df[df["ID_Zgloszenia"] == sel_edit_id].index[0]
                    e_row = df.loc[e_idx]

                    with st.form("quick_edit_ticket_form"):
                        e_klient = st.text_input("Klient:", value=str(e_row["Klient"]))
                        e_tel = st.text_input("Telefon (można wpisać 2 po ukośniku):", value=str(e_row["Telefon"]))
                        e_kod = st.text_input("Kod pocztowy:", value=str(e_row["Kod_Pocztowy"]))
                        e_miasto = st.text_input("Miasto:", value=str(e_row["Miasto"]))
                        e_adres = st.text_input("Adres:", value=str(e_row["Adres"]))
                        e_inst = st.text_input("Urządzenie / Typ instalacji:", value=str(e_row["Typ_Instalacji"]))

                        tech_choices = [""] + list(profiles.keys())
                        curr_tech = str(e_row["Serwisant"]) if str(e_row["Serwisant"]) in tech_choices else ""
                        e_tech = st.selectbox("Przypisany serwisant:", tech_choices, index=tech_choices.index(curr_tech))

                        status_choices = ["Nowe", "Zaplanowane", "W trakcie", "Zakończone"]
                        curr_stat = str(e_row["Status"]) if str(e_row["Status"]) in status_choices else "Nowe"
                        e_status = st.selectbox("Status:", status_choices, index=status_choices.index(curr_stat))

                        e_data_plan = st.text_input("Data planowana (RRRR-MM-DD):", value=str(e_row["Data_Planowana"]))
                        e_przeglad = st.selectbox("Wymagany przegląd (Protokół):", ["Nie", "Tak"], index=1 if str(e_row.get("Wymagany_Przeglad", "Nie")) == "Tak" else 0)
                        e_kryt = st.selectbox("Krytyczne:", ["Nie", "Tak"], index=1 if str(e_row["Priorytet_Krytyczny"]) == "Tak" else 0)
                        e_kwota = st.number_input("Kwota przewidywana (zł):", min_value=0.0)
