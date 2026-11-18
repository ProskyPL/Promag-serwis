# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import datetime
import os
import json
import base64
import zipfile
import math
import re
import random
import smtplib
import urllib.parse
from email.message import EmailMessage
from io import BytesIO
from PIL import Image
from streamlit_drawable_canvas import st_canvas

# Bezpieczny import gspread do trwałego zapisu w Arkuszach Google (Google Sheets)
try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSPREAD_AVAILABLE = True
except ImportError:
    GSPREAD_AVAILABLE = False

# Bezpieczny import wydzielonych modułów (cennik.py, kadry.py, baza_wiedzy.py) z pełnymi fallbackami
try:
    from cennik import (
        STAWKA_RBH_NETTO,
        KOSZT_PALIWA_LITR,
        SR_CZAS_WIZYTY_H,
        RYCZALT_DOJAZD,
        TABELA_CENNIKA,
        wylicz_cene_uslugi,
        pobierz_slownik_cennika,
        oblicz_koszyk_uslug
    )
except ImportError:
    STAWKA_RBH_NETTO = 100.0
    KOSZT_PALIWA_LITR = 8.0
    SR_CZAS_WIZYTY_H = 1.5
    RYCZALT_DOJAZD = 200.0
    TABELA_CENNIKA = [
        {"usluga": "Przegląd roczny Pompy Ciepła (Powietrze)", "cena_materialow": 50.0, "czas_pracy_h": 2.0},
        {"usluga": "Przegląd roczny Pompy Ciepła (Grunt)", "cena_materialow": 50.0, "czas_pracy_h": 2.5},
        {"usluga": "Przegląd instalacji PV + pomiary Riso (<10 kWp)", "cena_materialow": 0.0, "czas_pracy_h": 2.0},
        {"usluga": "Przegląd instalacji PV + pomiary Riso (>10 kWp)", "cena_materialow": 0.0, "czas_pracy_h": 3.5},
        {"usluga": "Wymiana glikolu + płukanie i odpowietrzenie solarów", "cena_materialow": 520.0, "czas_pracy_h": 2.5},
        {"usluga": "Odpowietrzenie układu solarnego / CO", "cena_materialow": 40.0, "czas_pracy_h": 1.5},
        {"usluga": "Wymiana anody magnezowej w zasobniku", "cena_materialow": 120.0, "czas_pracy_h": 1.0},
        {"usluga": "Wymiana naczynia przeponowego solarnego 18L", "cena_materialow": 280.0, "czas_pracy_h": 1.5},
        {"usluga": "Wymiana czujnika temperatury NTC / PT1000", "cena_materialow": 80.0, "czas_pracy_h": 1.0},
        {"usluga": "Wymiana automatycznego odpowietrznika solarnego", "cena_materialow": 90.0, "czas_pracy_h": 1.0},
        {"usluga": "Czyszczenie filtrów siatkowych i separatora magnetycznego", "cena_materialow": 20.0, "czas_pracy_h": 1.0},
        {"usluga": "Diagnostyka awarii Pompy Ciepła / Automatyki CAREL", "cena_materialow": 0.0, "czas_pracy_h": 2.0},
        {"usluga": "Diagnostyka falownika PV / Magazynu Energii + Konfiguracja Wi-Fi", "cena_materialow": 0.0, "czas_pracy_h": 1.5},
        {"usluga": "Wymiana ochronnika przepięć SPD DC / złączy MC4", "cena_materialow": 140.0, "czas_pracy_h": 1.0},
        {"usluga": "Serwis Gwarancyjny (0 zł dla klienta)", "cena_materialow": 0.0, "czas_pracy_h": 0.0}
    ]
    def wylicz_cene_uslugi(wiersz, stawka_rbh=STAWKA_RBH_NETTO):
        if "Gwarancyjny" in str(wiersz.get("usluga", "")):
            return 0.0
        return round(float(wiersz.get("cena_materialow", 0.0)) + (float(wiersz.get("czas_pracy_h", 0.0)) * stawka_rbh), 2)

    def pobierz_slownik_cennika(tabela=None, stawka_rbh=STAWKA_RBH_NETTO):
        if tabela is None:
            tabela = TABELA_CENNIKA
        out = {}
        for w in tabela:
            out[w["usluga"]] = {
                "cena_materialow": float(w.get("cena_materialow", 0.0)),
                "czas_pracy_h": float(w.get("czas_pracy_h", 0.0)),
                "cena_calkowita": wylicz_cene_uslugi(w, stawka_rbh)
            }
        return out

    def oblicz_koszyk_uslug(wybrane_uslugi, tabela=None, stawka_rbh=STAWKA_RBH_NETTO, dolicz_dojazd=True, kwota_dojazdu=RYCZALT_DOJAZD):
        slownik = pobierz_slownik_cennika(tabela, stawka_rbh)
        suma_mat = sum(slownik[u]["cena_materialow"] for u in wybrane_uslugi if u in slownik)
        suma_czas = sum(slownik[u]["czas_pracy_h"] for u in wybrane_uslugi if u in slownik)
        suma_uslug = sum(slownik[u]["cena_calkowita"] for u in wybrane_uslugi if u in slownik)
        czynnosci_opis = [f"{u} ({slownik[u]['cena_calkowita']:.0f} zł)" for u in wybrane_uslugi if u in slownik]
        platne = [u for u in wybrane_uslugi if "Gwarancyjny" not in u]
        dojazd = float(kwota_dojazdu) if (dolicz_dojazd and len(platne) > 0) else 0.0
        razem = round(suma_uslug + dojazd, 2)
        auto_notatka = ("ZAKRES PRAC Z KALKULATORA: " + ", ".join(czynnosci_opis) + (f" + Dojazd ({dojazd:.0f} zł)." if dojazd > 0 else ".")) if czynnosci_opis else ""
        return {
            "suma_materialow": round(suma_mat, 2),
            "suma_czasu_h": round(suma_czas, 2),
            "suma_robocizny": round(suma_czas * stawka_rbh, 2),
            "dojazd": round(dojazd, 2),
            "razem_dla_klienta": razem,
            "auto_notatka": auto_notatka
        }

try:
    from kadry import (
        COMMISSION_TIER_1,
        COMMISSION_RATE_1,
        COMMISSION_TIER_2,
        COMMISSION_RATE_2,
        DOMYSLNE_NARZEDZIA,
        DOMYSLNE_MATERIALY,
        DEFAULT_TECHS,
        ensure_tech_profile_fields,
        get_low_stock_items,
        calculate_tech_metrics
    )
except ImportError:
    COMMISSION_TIER_1 = 15000.0
    COMMISSION_RATE_1 = 0.03
    COMMISSION_TIER_2 = 30000.0
    COMMISSION_RATE_2 = 0.04
    DOMYSLNE_NARZEDZIA = [
        {"nazwa": "Refraktometr optyczny (Glikol)", "ilosc": 1},
        {"nazwa": "Miernik cęgowy DC/AC + Riso", "ilosc": 1},
        {"nazwa": "Pompa elektryczna do napełniania solarów", "ilosc": 1},
        {"nazwa": "Klucz dynamometryczny + końcówki MC4", "ilosc": 1},
        {"nazwa": "Detektor nieszczelności F-gaz + Manometry", "ilosc": 1},
        {"nazwa": "Zestaw pasków pH + Klucze hydrauliczne", "ilosc": 1}
    ]
    DOMYSLNE_MATERIALY = [
        {"nazwa": "Glikol propylenowy 20L (bańka)", "ilosc": 3, "koszt_zakupu": 180.0, "cena_klienta": 320.0, "slowo_klucz": "glikol"},
        {"nazwa": "Anoda magnezowa 3/4 cala", "ilosc": 4, "koszt_zakupu": 45.0, "cena_klienta": 120.0, "slowo_klucz": "anod"},
        {"nazwa": "Naczynie przeponowe solarne 18L", "ilosc": 2, "koszt_zakupu": 140.0, "cena_klienta": 280.0, "slowo_klucz": "naczyni"},
        {"nazwa": "Czujnik temperatury NTC 10k / PT1000", "ilosc": 5, "koszt_zakupu": 20.0, "cena_klienta": 80.0, "slowo_klucz": "czujnik"},
        {"nazwa": "Odpowietrznik automatyczny solarny 180C", "ilosc": 3, "koszt_zakupu": 35.0, "cena_klienta": 90.0, "slowo_klucz": "odpowietrz"},
        {"nazwa": "Zawór bezpieczeństwa 3 bar / 6 bar", "ilosc": 3, "koszt_zakupu": 30.0, "cena_klienta": 85.0, "slowo_klucz": "zawór"},
        {"nazwa": "Ochronnik przepięć SPD DC (wkładka)", "ilosc": 2, "koszt_zakupu": 55.0, "cena_klienta": 140.0, "slowo_klucz": "spd"},
        {"nazwa": "Złącza MC4 oryginał (komplet)", "ilosc": 10, "koszt_zakupu": 8.0, "cena_klienta": 25.0, "slowo_klucz": "mc4"}
    ]
    def _mk_prof(name, pin, umowa, sal):
        return {
            "pin": pin,
            "pelne_dane": f"{name} - Serwisant OZE",
            "telefon": "",
            "umowa": umowa,
            "wynagrodzenie_bazowe": sal,
            "narzedzia": [dict(x) for x in DOMYSLNE_NARZEDZIA],
            "materialy": [dict(x) for x in DOMYSLNE_MATERIALY],
            "flota": {"nr_rej": "WPR 00000", "przebieg_km": 120000, "limit_olej_km": 130000, "data_oc_przeglad": "2027-05-01"}
        }
    DEFAULT_TECHS = {
        "Leszek": _mk_prof("Leszek", "4829", "Umowa o pracę", 6500.0),
        "Mikołaj": _mk_prof("Mikołaj", "7315", "B2B", 7000.0),
        "Wojtek": _mk_prof("Wojtek", "5942", "Umowa o pracę", 6000.0),
        "Piotr": _mk_prof("Piotr", "3190", "Umowa o pracę", 6000.0),
        "Stanisław": _mk_prof("Stanisław", "8164", "Umowa zlecenie", 6000.0)
    }
    def ensure_tech_profile_fields(profiles):
        for t_name, p in profiles.items():
            if "narzedzia" not in p or not isinstance(p["narzedzia"], list):
                p["narzedzia"] = [dict(x) for x in DOMYSLNE_NARZEDZIA]
            if "materialy" not in p or not isinstance(p["materialy"], list):
                p["materialy"] = [dict(x) for x in DOMYSLNE_MATERIALY]
            if "flota" not in p or not isinstance(p["flota"], dict):
                p["flota"] = {"nr_rej": "WPR 00000", "przebieg_km": 120000, "limit_olej_km": 130000, "data_oc_przeglad": "2027-05-01"}
        return profiles

    def get_low_stock_items(tech_profile):
        low = []
        for item in tech_profile.get("materialy", []):
            qty = int(item.get("ilosc", 0))
            if qty <= 1:
                low.append(f"{item.get('nazwa')} ({qty} szt.)")
        return low

    def calculate_tech_metrics(df, tech_name):
        tech_df = df[df["Serwisant"] == tech_name]
        closed_df = tech_df[tech_df["Status"] == "Zakończone"].copy()
        total_closed = len(closed_df)
        ftfr_pct = round((len(closed_df[closed_df["Naprawione_1_Wizyta"] == "Tak"]) / total_closed) * 100.0, 1) if total_closed > 0 else 100.0
        paid_count = len(closed_df[closed_df["Typ_Serwisu"] == "Płatny"])
        warranty_count = len(closed_df[closed_df["Typ_Serwisu"] == "Gwarancyjny"])
        leads_count = len(tech_df[tech_df["Lead_Handlowy"].astype(str).str.strip() != ""])
        unsettled_cash_df = closed_df[(closed_df["Forma_Platnosci"] == "Gotówka pobrana") & (closed_df["Gotowka_Rozliczona"] != "Tak")]
        cash_in_hand = float(unsettled_cash_df["Kwota_Koncowa"].sum())
        current_ym = datetime.date.today().strftime("%Y-%m")
        closed_df["ym"] = closed_df["Data_Zamkniecia"].astype(str).str.slice(0, 7)
        monthly_df = closed_df[(closed_df["ym"] == current_ym) | (closed_df["ym"] == "")]
        monthly_revenue = float(monthly_df["Kwota_Koncowa"].sum())
        monthly_parts_cost = float(monthly_df["Koszt_Materialow"].sum())
        missing_tier_1 = max(0.0, COMMISSION_TIER_1 - monthly_revenue)
        missing_tier_2 = max(0.0, COMMISSION_TIER_2 - monthly_revenue)
        if monthly_revenue >= COMMISSION_TIER_2:
            commission_pln = monthly_revenue * COMMISSION_RATE_2
            active_rate_str = "4% (Próg II zdobyty!)"
            progress_pct = 100
        elif monthly_revenue >= COMMISSION_TIER_1:
            commission_pln = monthly_revenue * COMMISSION_RATE_1
            active_rate_str = "3% (Próg I zdobyty!)"
            progress_pct = min(99, int((monthly_revenue / COMMISSION_TIER_2) * 100))
        else:
            commission_pln = 0.0
            active_rate_str = "W drodze do Progu I (3%)"
            progress_pct = min(99, int((monthly_revenue / COMMISSION_TIER_1) * 100))
        return {
            "total_closed": total_closed, "ftfr_pct": ftfr_pct, "paid_count": paid_count,
            "warranty_count": warranty_count, "leads_count": leads_count, "cash_in_hand": cash_in_hand,
            "monthly_revenue": monthly_revenue, "monthly_parts_cost": monthly_parts_cost,
            "missing_tier_1": missing_tier_1, "missing_tier_2": missing_tier_2,
            "commission_pln": round(commission_pln, 2), "active_rate_str": active_rate_str, "progress_pct": progress_pct
        }

try:
    from baza_wiedzy import (
        DIAGNOSTYKA_BAZA,
        INSTRUKCJE_KONFIGURACJI_TEXT,
        przeszukaj_lokalne_instrukcje,
        przeszukaj_praktyke_i_internet
    )
except ImportError:
    DIAGNOSTYKA_BAZA = {
        "Falownik Deye": {"F08 - GFDI_Relay_Failure": "Sprawdź uziemienie N-PE oraz opcję Signal Island Mode w menu 7777."},
        "Pompa Ciepła ARGOM": {"AL007 / AL137 - Brak przepływu": "Wyczyść filtr siatkowy i separator magnetyczny, sprawdź ciśnienie."}
    }
    INSTRUKCJE_KONFIGURACJI_TEXT = "### Konfiguracja Wi-Fi 2.4 GHz (Deye / GoodWe)"
    def przeszukaj_lokalne_instrukcje(q, f="."):
        return []
    def przeszukaj_praktyke_i_internet(q):
        return [], "https://www.google.com", "https://www.elektroda.pl", "https://www.youtube.com"

# ==========================================
# SILNIK GRYWALIZACJI (SERIA COMBO 🔥 + WATY ⚡ + ODZNAKI 🏅)
# ==========================================
def calculate_gamification_stats(df, tech_name, tech_profile=None):
    """
    Wylicza dla serwisanta:
    - Aktualną serię napraw 'na strzała' (Combo bez poprawek) oraz rekord serii.
    - Wirtualną walutę 'Waty' (⚡) w stylu punktów InPost:
        +10 ⚡ za zamkniętą wizytę za 1. razem (ze zdjęciem i podpisem)
        +25 ⚡ za przywieziony Lead Handlowy
        +50 ⚡ za wykonanie pilnej awarii (Priorytet_Krytyczny == 'Tak')
        +15 ⚡ za wypełniony protokół przeglądu okresowego
        +15 ⚡ za rozliczoną gotówkę w biurze
    - 6 odznak osiągnięć (odblokowane vs do zdobycia).
    """
    tech_df = df[df["Serwisant"] == tech_name].copy()
    finished_df = tech_df[tech_df["Status"].isin(["Zakończone", "Niewykonane", "Do zamówienia części"])].copy()

    # 1. Licznik serii "Na strzała" (Combo)
    current_streak = 0
    best_streak = 0
    temp_s = 0
    for _, r in finished_df.iterrows():
        is_hit = (str(r.get("Status", "")) == "Zakończone" and str(r.get("Naprawione_1_Wizyta", "Tak")) == "Tak")
        if is_hit:
            temp_s += 1
            if temp_s > best_streak:
                best_streak = temp_s
        else:
            temp_s = 0
    current_streak = temp_s

    # 2. Waluta "Waty" (⚡)
    closed_ok = tech_df[(tech_df["Status"] == "Zakończone") & (tech_df["Naprawione_1_Wizyta"] == "Tak")]
    urgent_ok = tech_df[(tech_df["Status"] == "Zakończone") & (tech_df["Priorytet_Krytyczny"] == "Tak")]
    leads_ok = tech_df[tech_df["Lead_Handlowy"].astype(str).str.strip() != ""]
    protocols_ok = tech_df[tech_df["Wyniki_Przegladu"].astype(str).str.contains("PROTOKÓŁ|Pomiary", case=False, na=False)]
    settled_cash_ok = tech_df[(tech_df["Status"] == "Zakończone") & (tech_df["Gotowka_Rozliczona"] == "Tak")]

    waty_base = (
        len(closed_ok) * 10 +
        len(leads_ok) * 25 +
        len(urgent_ok) * 50 +
        len(protocols_ok) * 15 +
        len(settled_cash_ok) * 5
    )
    bonus_waty = int(tech_profile.get("bonus_waty", 0)) if isinstance(tech_profile, dict) else 0
    total_waty = waty_base + bonus_waty

    # 3. Odznaki osiągnięć
    solar_cnt = len(closed_ok[closed_ok["Typ_Instalacji"].astype(str).str.contains("Solar|Kolektor", case=False, na=False)])
    pv_cnt = len(closed_ok[closed_ok["Typ_Instalacji"].astype(str).str.contains("Deye|GoodWe|PV|Falownik", case=False, na=False)])
    pc_cnt = len(closed_ok[closed_ok["Typ_Instalacji"].astype(str).str.contains("Pompa|ARGOM", case=False, na=False)])

    badges = [
        {
            "id": "snajper",
            "ikona": "🎯",
            "nazwa": "Snajper Diagnostyki",
            "opis": "Seria min. 5 napraw z rzędu za 1. podejściem",
            "postep": f"{current_streak}/5 z rzędu",
            "unlocked": current_streak >= 5 or best_streak >= 5
        },
        {
            "id": "glikol",
            "ikona": "💧",
            "nazwa": "Zaklinacz Glikolu",
            "opis": "Min. 3 zamknięte serwisy/przeglądy solarne",
            "postep": f"{solar_cnt}/3 solary",
            "unlocked": solar_cnt >= 3
        },
        {
            "id": "f08",
            "ikona": "⚡",
            "nazwa": "Pogromca F08",
            "opis": "Min. 3 zamknięte serwisy falowników PV / Deye",
            "postep": f"{pv_cnt}/3 falowniki",
            "unlocked": pv_cnt >= 3
        },
        {
            "id": "pompa",
            "ikona": "🌡️",
            "nazwa": "Doktor CAREL",
            "opis": "Min. 3 zamknięte serwisy Pomp Ciepła",
            "postep": f"{pc_cnt}/3 pompy",
            "unlocked": pc_cnt >= 3
        },
        {
            "id": "strazak",
            "ikona": "🚒",
            "nazwa": "Strażak Miesiąca",
            "opis": "Min. 2 pilne awarie opanowane w terenie",
            "postep": f"{len(urgent_ok)}/2 pilne",
            "unlocked": len(urgent_ok) >= 2
        },
        {
            "id": "lowca",
            "ikona": "🦈",
            "nazwa": "Łowca Leadów",
            "opis": "Min. 2 przekazane tematy handlowe z terenu",
            "postep": f"{len(leads_ok)}/2 leady",
            "unlocked": len(leads_ok) >= 2
        }
    ]
    unlocked_cnt = sum(1 for b in badges if b["unlocked"])

    return {
        "current_streak": current_streak,
        "best_streak": best_streak,
        "total_waty": total_waty,
        "badges": badges,
        "unlocked_cnt": unlocked_cnt
    }

# ==========================================
# BAZA 40 HUMORYSTYCZNYCH CYTATÓW SERWISOWYCH ("MYŚL NA DZIŚ")
# ==========================================
MYSLI_NA_DZIS_QUOTES = [
    "„Teoria jest wtedy, gdy nic nie działa, ale wiemy dlaczego. Praktyka – gdy wszystko działa, ale nikt nie wie czemu. U nas łączymy oba: nic nie działa i nikt nie wie czemu.”",
    "„Złota zasada serwisu: jeśli po skręceniu obudowy zostało Ci kilka śrubek, to znaczy, że zoptymalizowałeś konstrukcję fabryczną.”",
    "„Nie ma rzeczy niemożliwych do naprawienia. Są tylko takie, do których trzeba wrócić jutro z większym młotkiem.”",
    "„Jeśli coś nie działa, a powinno – użyj trytytki. Jeśli się rusza, a nie powinno – użyj taśmy izolacyjnej.”",
    "„Najbardziej podejrzane zdanie w historii serwisu: ‚Panie, ja tu nic nie ruszałem, samo tak zrobiło w nocy’.”",
    "„Każda naprawa przewidziana na 15 minut jest o jedną zapieczoną śrubkę od stania się pracą na pół dnia.”",
    "„Instrukcję obsługi czytamy dopiero wtedy, gdy wszystkie inne metody zawiodą, a obudowa jest już rozkręcona.”",
    "„Dobrego serwisanta poznaje się nie po tym, że nigdy nic nie zaleje, ale po tym, jak szybko znajdzie ręcznik papierowy w busie.”",
    "„Prawa Murphy'ego w kotłowni: narzędzie, którego potrzebujesz, zawsze leży na samym dnie skrzynki pod wszystkim innym.”",
    "„Kawa z rana działa lepiej niż reset falownika – przywraca ustawienia fabryczne w 5 minut.”",
    "„Pompa ciepła jest jak kot – robi dokładnie to, co chce, zwłaszcza wtedy, gdy stoisz nad nią z miernikiem.”",
    "„Zapowietrzony układ solarny to jedyna rzecz na świecie, która potrafi bulgotać głośniej niż brzuch przed obiadem.”",
    "„Filtr siatkowy nie jest po to, żeby go czyścić profilaktycznie. Jest po to, żeby przypomnieć o sobie w piątek o 15:30.”",
    "„Glikol ma jedną magiczną właściwość: jedna kropla potrafi pobrudzić pół kotłowni i dwie pary spodni.”",
    "„Ciśnienie w układzie jest jak ciśnienie w pracy – najlepiej, jak trzyma się w zielonym polu.”",
    "„Najkrótsza droga do zrozumienia schematu hydraulicznego prowadzi przez pytanie: ‚Kto to panu tak tutaj połączył?’.”",
    "„Gdy klient mówi, że ‚trochę kapie’, szykuj wiadro i mop.”",
    "„Anoda magnezowa poświęca się dla zasobnika każdego dnia. Bądź jak anoda – ale bez rozpuszczania się w pracy.”",
    "„Odpowietrznik automatyczny działa automatycznie do pierwszego dnia po zakończeniu gwarancji.”",
    "„W każdej kotłowni najciemniej jest dokładnie tam, gdzie producent przykleił tabliczkę znamionową.”",
    "„Słońce świeci dla wszystkich, ale pełną moc na falowniku mają tylko ci, co dobrze dokręcili złącza MC4.”",
    "„Konfiguracja Wi-Fi u klienta to prawdziwy sport ekstremalny – szczególnie gdy hasło do routera jest na kartce z 2018 roku.”",
    "„Błąd komunikacji z loggerem: w 90% przypadków winny jest router odłączony od prądu, ‚bo migał w nocy’.”",
    "„Falownik bez błędu to piękny widok. Falownik z błędem F08 to okazja, żeby przypomnieć sobie gdzie jest uziemienie.”",
    "„Prąd nie ma koloru, ale potrafi nadać bardzo wyrazisty kolor fryzurze, jeśli zapomnisz wyłączyć DC.”",
    "„Zawsze sprawdzaj biegunowość dwa razy – falownik ma krótką pamięć, ale bardzo długą listę błędów.”",
    "„Kiedy aplikacja pokazuje 0 W produkcji w samo południe, sprawdź najpierw, czy nie pada śnieg.”",
    "„Magazyn energii na 100% naładowany = serwisant na 100% zadowolony.”",
    "„Twój bus to Twój zamek. Tylko w jakim innym zamku na zakrętach dzwonią klucze płaskie z bańką glikolu?”",
    "„Nawigacja mówi: ‚Za 200 metrów skręć w pole’. Pamiętaj – masz busa serwisowego, nie traktor!”",
    "„Dzisiejszy cel: 100% napraw za pierwszym podejściem i zero powrotów po zapomniany miernik.”",
    "„Każda zamknięta wizyta przybliża Cię do wyższego progu premii. Licznik bije na Twoją korzyść!”",
    "„Pamiętaj o zdjęciu tabliczki znamionowej – bez niej system nie uwierzy, że tam byłeś, nawet jak zrobisz selfie z pompą.”",
    "„Dobry plan trasy to połowa sukcesu. Druga połowa to brak korków na wylotówce.”",
    "„Jeśli na pace zostało mniej niż 1 sztuka glikolu lub czujnika – daj znać zanim wyjedziesz na koniec województwa.”",
    "„Klient zadowolony, usterka usunięta, podpis zebrany – można jechać dalej jak po swoje!”",
    "„Nawet najtrudniejszy poniedziałek ma tylko tyle godzin, ile zmieści się w harmonogramie.”",
    "„Dziś pracujemy w trybie PRO: minimum klikania, maksimum konkretów, zero zbędnego stresu.”",
    "„Pamiętaj: najdroższa część w busie to ta, której akurat nie zabrałeś z magazynu.”",
    "„Szerokości na trasie! Niech dzisiaj wszystkie zawory odkręcają się lekko, a wszystkie śrubunki trzymają za pierwszym razem.”"
]

DNI_TYGODNIA_PL = {
    0: "Poniedziałek", 1: "Wtorek", 2: "Środa", 3: "Czwartek",
    4: "Piątek", 5: "Sobota", 6: "Niedziela"
}

# ==========================================
# IKONY WEKTOROWE SVG DO DOLNEGO MENU (Z HUMOREM, W STYLU INPOST + PROMAG)
# ==========================================
SVG_ICONS_NAV = {
    "start": """<svg width="46" height="42" viewBox="0 0 64 58" fill="none" xmlns="http://www.w3.org/2000/svg">
        <polygon points="8,28 32,6 56,28" fill="#0a2540" stroke="#0a2540" stroke-width="2"/>
        <polygon points="33,10 51,26 45,30 28,14" fill="#0284c7" stroke="#ffffff" stroke-width="1.5"/>
        <rect x="16" y="28" width="32" height="24" rx="4" fill="#ffffff" stroke="#0a2540" stroke-width="3"/>
        <circle cx="26" cy="38" r="2.5" fill="#0a2540"/>
        <circle cx="38" cy="38" r="2.5" fill="#0a2540"/>
        <path d="M25 45 Q32 51 39 45" stroke="#16a34a" stroke-width="3" stroke-linecap="round" fill="none"/>
    </svg>""",
    "plan": """<svg width="46" height="42" viewBox="0 0 64 58" fill="none" xmlns="http://www.w3.org/2000/svg">
        <polygon points="10,14 25,19 25,50 10,45" fill="#f8fafc" stroke="#0a2540" stroke-width="2.5"/>
        <polygon points="25,19 39,14 39,45 25,50" fill="#e2e8f0" stroke="#0a2540" stroke-width="2.5"/>
        <polygon points="39,14 54,19 54,50 39,45" fill="#f8fafc" stroke="#0a2540" stroke-width="2.5"/>
        <polyline points="17,40 28,24 36,36 46,20" fill="none" stroke="#0284c7" stroke-width="3" stroke-dasharray="4 3"/>
        <circle cx="17" cy="40" r="4" fill="#16a34a"/>
        <circle cx="46" cy="18" r="6" fill="#ef4444" stroke="#0a2540" stroke-width="2"/>
    </svg>""",
    "przeglad": """<svg width="46" height="42" viewBox="0 0 64 58" fill="none" xmlns="http://www.w3.org/2000/svg">
        <rect x="14" y="10" width="36" height="44" rx="5" fill="#ffffff" stroke="#0a2540" stroke-width="3"/>
        <rect x="23" y="5" width="18" height="9" rx="3" fill="#38bdf8" stroke="#0a2540" stroke-width="2.2"/>
        <polyline points="22,31 29,39 43,21" fill="none" stroke="#16a34a" stroke-width="4.5" stroke-linecap="round" stroke-linejoin="round"/>
        <line x1="22" y1="46" x2="42" y2="46" stroke="#94a3b8" stroke-width="3" stroke-linecap="round"/>
    </svg>""",
    "bus": """<svg width="46" height="42" viewBox="0 0 64 58" fill="none" xmlns="http://www.w3.org/2000/svg">
        <circle cx="27" cy="12" r="4.5" fill="#f59e0b" stroke="#0a2540" stroke-width="2"/>
        <rect x="8" y="16" width="34" height="27" rx="4" fill="#0a2540" stroke="#0a2540" stroke-width="2"/>
        <rect x="38" y="23" width="18" height="20" rx="4" fill="#16a34a" stroke="#0a2540" stroke-width="2"/>
        <polygon points="43,26 53,26 54,33 43,33" fill="#bae6fd"/>
        <text x="24" y="33" fill="#86efac" font-size="9" font-weight="bold" font-family="Arial" text-anchor="middle">OZE</text>
        <circle cx="19" cy="44" r="6" fill="#334155" stroke="#ffffff" stroke-width="2"/>
        <circle cx="46" cy="44" r="6" fill="#334155" stroke="#ffffff" stroke-width="2"/>
    </svg>"""
}

def render_bottom_nav_bar(active_tab_key):
    """Renderuje wizualny pasek dolnego menu z 4 kafelkami SVG i przyciskami zmiany widoku."""
    st.markdown("<div style='margin-top: 18px;'></div>", unsafe_allow_html=True)
    with st.container(border=True):
        cols = st.columns(4)
        nav_defs = [
            ("start", "Start", "🏠 Start"),
            ("plan", "Plan", "🗺️ Plan"),
            ("przeglad", "Przegląd", "📋 Przegląd"),
            ("bus", "Bus", "🚐 Bus")
        ]
        for idx, (k_nav, title_nav, btn_lbl) in enumerate(nav_defs):
            is_act = (active_tab_key == k_nav)
            bg_c = "#dcfce7" if is_act else "#ffffff"
            brd_c = "2px solid #16a34a" if is_act else "1px solid #e2e8f0"
            txt_c = "#15803d" if is_act else "#334155"
            with cols[idx]:
                st.markdown(
                    f"""<div style="background:{bg_c}; border:{brd_c}; border-radius:14px; padding:8px 4px 4px 4px; text-align:center; margin-bottom:6px;">
                        {SVG_ICONS_NAV[k_nav]}
                        <div style="font-size:13px; font-weight:800; color:{txt_c}; margin-top:2px;">{title_nav}</div>
                    </div>""",
                    unsafe_allow_html=True
                )
                if st.button(btn_lbl, key=f"bnav_{k_nav}_{idx}", type="primary" if is_act else "secondary"):
                    st.session_state["tech_nav_tab"] = k_nav
                    if k_nav != "plan":
                        st.session_state["active_ticket_id"] = None
                    st.rerun()

# ==========================================
# PEŁNA TABLICA KODÓW POCZTOWYCH POLSKI (00-99) + MIASTA + PROCEDURY PRZEGLĄDÓW
# ==========================================
FULL_POSTAL_PREFIX_COORDS = {
    "00": (52.23, 21.01), "01": (52.25, 20.95), "02": (52.18, 21.00), "03": (52.26, 21.08),
    "04": (52.22, 21.15), "05": (52.10, 20.85), "06": (52.88, 20.62), "07": (52.65, 21.75),
    "08": (52.16, 22.28), "09": (52.55, 19.70), "10": (53.78, 20.48), "11": (53.95, 21.15),
    "12": (53.58, 21.35), "13": (53.35, 19.95), "14": (53.75, 19.85), "15": (53.13, 23.16),
    "16": (53.65, 23.10), "17": (52.75, 23.45), "18": (53.18, 22.08), "19": (53.85, 22.45),
    "20": (51.25, 22.57), "21": (51.65, 22.85), "22": (51.05, 23.40), "23": (50.85, 22.30),
    "24": (51.35, 22.00), "25": (50.87, 20.63), "26": (51.40, 21.15), "27": (50.85, 21.45),
    "28": (50.45, 20.75), "29": (50.75, 20.05), "30": (50.06, 19.94), "31": (50.08, 20.02),
    "32": (50.00, 19.75), "33": (50.01, 20.98), "34": (49.65, 19.95), "35": (50.04, 21.99),
    "36": (50.25, 21.95), "37": (50.05, 22.65), "38": (49.68, 21.75), "39": (50.35, 21.45),
    "40": (50.26, 19.02), "41": (50.30, 18.95), "42": (50.70, 19.15), "43": (49.95, 18.95),
    "44": (50.20, 18.60), "45": (50.67, 17.92), "46": (50.75, 18.15), "47": (50.35, 18.15),
    "48": (50.45, 17.45), "49": (50.75, 17.50), "50": (51.11, 17.04), "51": (51.15, 17.10),
    "52": (51.06, 17.05), "53": (51.08, 16.98), "54": (51.14, 16.92), "55": (51.05, 17.25),
    "56": (51.45, 17.35), "57": (50.45, 16.65), "58": (50.80, 16.30), "59": (51.20, 16.15),
    "60": (52.41, 16.93), "61": (52.40, 16.95), "62": (52.25, 17.65), "63": (51.85, 17.45),
    "64": (52.75, 16.65), "65": (51.94, 15.51), "66": (52.35, 15.25), "67": (51.68, 15.95),
    "68": (51.65, 15.10), "69": (52.35, 14.65), "70": (53.43, 14.55), "71": (53.46, 14.52),
    "72": (53.75, 14.95), "73": (53.25, 15.35), "74": (53.05, 14.65), "75": (54.19, 16.18),
    "76": (54.45, 16.85), "77": (53.85, 17.35), "78": (53.85, 15.95), "79": (53.95, 16.35),
    "80": (54.35, 18.65), "81": (54.52, 18.53), "82": (54.05, 19.15), "83": (54.05, 18.25),
    "84": (54.60, 18.20), "85": (53.12, 18.01), "86": (53.45, 18.55), "87": (53.01, 18.60),
    "88": (52.80, 18.25), "89": (53.55, 17.75), "90": (51.76, 19.46), "91": (51.80, 19.42),
    "92": (51.75, 19.55), "93": (51.72, 19.45), "94": (51.74, 19.38), "95": (51.90, 19.35),
    "96": (52.00, 20.25), "97": (51.40, 19.70), "98": (51.55, 18.85), "99": (52.15, 19.40)
}

PRECISE_SUBPREFIX_COORDS = {
    "05-5": (52.07, 21.02), "05-8": (52.16, 20.75), "05-0": (52.25, 20.75),
    "05-1": (52.42, 20.92), "05-2": (52.35, 21.25), "05-3": (52.18, 21.55),
    "05-4": (52.12, 21.28), "05-6": (51.86, 20.86), "96-3": (52.05, 20.44),
    "96-5": (52.23, 20.24), "96-1": (51.95, 20.15), "96-2": (51.76, 20.25)
}

CITY_FALLBACK_COORDS = {
    "warszawa": (52.23, 21.01), "piaseczno": (52.07, 21.02), "pruszków": (52.17, 20.80),
    "pruszkow": (52.17, 20.80), "brwinów": (52.14, 20.72), "brwinow": (52.14, 20.72),
    "grodzisk": (52.11, 20.63), "nadarzyn": (52.09, 20.81), "tarczyn": (51.98, 20.84),
    "grójec": (51.86, 20.86), "grojec": (51.86, 20.86), "żyrardów": (52.05, 20.44),
    "zyrardow": (52.05, 20.44), "sochaczew": (52.23, 20.24), "błonie": (52.20, 20.62),
    "blonie": (52.20, 20.62), "ożarów": (52.21, 20.80), "marki": (52.32, 21.10),
    "obywatelska": (52.23, 21.01), "wołomin": (52.35, 21.24), "wolomin": (52.35, 21.24),
    "legionowo": (52.40, 20.93), "otwock": (52.11, 21.26), "mińsk": (52.18, 21.57),
    "minsk": (52.18, 21.57), "siedlce": (52.17, 22.29), "płock": (52.55, 19.70),
    "plock": (52.55, 19.70), "ciechanów": (52.88, 20.62), "ostrołęka": (53.08, 21.57),
    "radom": (51.40, 21.15), "kielce": (50.87, 20.63), "olsztyn": (53.78, 20.48),
    "białystok": (53.13, 23.16), "bialystok": (53.13, 23.16), "lublin": (51.25, 22.57),
    "łódź": (51.76, 19.46), "lodz": (51.76, 19.46), "skierniewice": (51.95, 20.15),
    "łowicz": (52.11, 19.95), "lowicz": (52.11, 19.95), "nowa iwiczna": (52.08, 21.00),
    "bobrowiec": (52.06, 20.97), "ustanów": (52.00, 21.04), "jakubów": (52.21, 21.68)
}

PROCEDURY_PRZEGLADOW = {
    "Kolektory Słoneczne (Solar)": {
        "nazwa": "☀️ Procedura Przeglądu Układu Solarnego",
        "checklist": [
            "Pomiar temperatury zamarzania glikolu refraktometrem optycznym (norma min. -25°C)",
            "Badanie odczynu pH glikolu paskiem testowym (norma pH >= 7.0)",
            "Kontrola ciśnienia wstępnego w naczyniu przeponowym solarnym (na zimno)",
            "Odpowietrzenie układu solarnego na grupie pompowej i odpowietrzniku dachowym",
            "Kontrola przepływu na rotametrze (l/min) oraz pracy pompy obiegowej",
            "Sprawdzenie stanu anody magnezowej w zasobniku CWU oraz zaworu bezpieczeństwa (6 bar)",
            "Odczyt czujników temperatury T1 (kolektor) i T2/T3 (zasobnik) na sterowniku"
        ],
        "pola_pomiarowe": [
            ("Temp. krzepnięcia glikolu (°C)", "-28°C"),
            ("Odczyn pH glikolu", "7.5"),
            ("Ciśnienie robocze układu (bar)", "2.0 bar"),
            ("Ciśnienie naczynia przeponowego (bar)", "1.5 bar")
        ]
    },
    "Pompa Ciepła ARGOM": {
        "nazwa": "🌡️ Procedura Przeglądu Pompy Ciepła (ARGOM / CAREL)",
        "checklist": [
            "Czyszczenie filtra siatkowego oraz separatora magnetycznego na powrocie CO",
            "Kontrola ciśnienia zładu wody kotłowej (norma 1.5 - 2.0 bar) i odpowietrzenie",
            "Kontrola szczelności układu chłodniczego detektorem F-gaz",
            "Czyszczenie lameli parownika jednostki zewnętrznej oraz udrożnienie tacy ociekowej",
            "Sprawdzenie historii alarmów na sterowniku CAREL (AL007/AL137) oraz delty T (T_zasilania - T_powrotu)",
            "Kontrola dokręcenia styków elektrycznych zasilania sprężarki i grzałek szczytowych",
            "Kontrola anody magnezowej / tytanowej w zasobniku CWU"
        ],
        "pola_pomiarowe": [
            ("Ciśnienie układu CO (bar)", "1.8 bar"),
            ("Delta T (Zasilanie / Powrót °C)", "38°C / 33°C"),
            ("Stan filtrów i separatora", "Wyczyszczone"),
            ("Szczelność F-gaz", "Szczelny (OK)")
        ]
    },
    "Falownik Deye (PV / Hybryda)": {
        "nazwa": "⚡ Procedura Przeglądu Instalacji PV i Falownika Deye",
        "checklist": [
            "Pomiar rezystancji izolacji Riso stringów DC (norma > 1 MOhm) oraz napięcia Voc",
            "Kontrola złączy MC4 oraz dokręcenie zacisków w rozdzielnicy DC/AC kluczem dynamometrycznym",
            "Kontrola wkładek ochronników przepięć SPD po stronie DC oraz AC",
            "Sprawdzenie uziemienia PE oraz opcji Signal Island Mode w menu 7777 (ochrona przed F08)",
            "Czyszczenie radiatorów falownika i aktualizacja statusu loggera Wi-Fi (Solarman / Deye Cloud)",
            "Test przełączania zasilania awaryjnego ( port LOAD / Backup )"
        ],
        "pola_pomiarowe": [
            ("Pomiar Riso DC (MOhm)", "> 20 MOhm"),
            ("Napięcie Voc String 1 / String 2 (V)", "580 V / 575 V"),
            ("Stan SPD AC/DC oraz MC4", "Sprawne / Dokręcone"),
            ("Status komunikacji Wi-Fi", "Online")
        ]
    },
    "Falownik GoodWe (PV / Hybryda)": {
        "nazwa": "⚡ Procedura Przeglądu Instalacji PV i Falownika GoodWe",
        "checklist": [
            "Pomiar rezystancji izolacji Riso stringów DC oraz napięcia Voc",
            "Kontrola złączy MC4 oraz rozdzielnicy AC/DC (ochronniki SPD, wyłącznik nadprądowy)",
            "Weryfikacja komunikacji z licznikiem Smart Meter GM3000 (kierunek przekładników CT)",
            "Kontrola połączenia Wi-Fi 2.4 GHz (aplikacja SolarGo / SEMS Portal)",
            "Odczyt rejestru błędów i kontrola temperatury pracy falownika"
        ],
        "pola_pomiarowe": [
            ("Pomiar Riso DC (MOhm)", "> 20 MOhm"),
            ("Napięcie Voc String 1 / String 2 (V)", "600 V / 590 V"),
            ("Komunikacja Smart Meter", "Poprawna (OK)"),
            ("Status SEMS / Wi-Fi", "Online")
        ]
    },
    "Magazyn Energii": {
        "nazwa": "🔋 Procedura Przeglądu Magazynu Energii",
        "checklist": [
            "Odczyt parametrów BMS: SOC (poziom naładowania) oraz SOH (kondycja ogniw %)",
            "Sprawdzenie balansu napięciowego między modułami bateryjnymi",
            "Kontrola przewodów prądowych DC oraz komunikacji CAN / RS485 z falownikiem",
            "Test pracy w trybie awaryjnym (zanik zasilania sieciowego)",
            "Kontrola temperatury otoczenia i warunków wentylacji w pomieszczeniu"
        ],
        "pola_pomiarowe": [
            ("Parametr SOH / SOC (%)", "100% / 85%"),
            ("Napięcie baterii (V)", "52.4 V"),
            ("Komunikacja BMS CAN/RS485", "OK"),
            ("Test Backup (Zanik sieci)", "Pozytywny")
        ]
    }
}

def get_procedure_for_type(inst_type):
    if inst_type in PROCEDURY_PRZEGLADOW:
        return PROCEDURY_PRZEGLADOW[inst_type]
    if "Pompa" in str(inst_type):
        return PROCEDURY_PRZEGLADOW["Pompa Ciepła ARGOM"]
    if "Solar" in str(inst_type) or "Kolektor" in str(inst_type):
        return PROCEDURY_PRZEGLADOW["Kolektory Słoneczne (Solar)"]
    if "Deye" in str(inst_type) or "PV" in str(inst_type):
        return PROCEDURY_PRZEGLADOW["Falownik Deye (PV / Hybryda)"]
    return PROCEDURY_PRZEGLADOW["Pompa Ciepła ARGOM"]

# ==========================================
# GENERATOR SZABLONÓW SMS
# ==========================================
def build_sms_content(mode, ticket_id, client_name="", tech_name="", plan_date="", kwota=0.0):
    if mode == "zwykle_24h":
        return (
            f"Dzien dobry! Przyjelismy Twoje zgloszenie serwisowe nr {ticket_id} w PROMAG OZE. "
            f"Dokladny termin wizyty konserwatora wyznaczymy i przekazaemy w ciagu 24 godzin. "
            f"Pozdrawiamy, Dzial Serwisu PROMAG."
        )
    elif mode == "montaz_5dni":
        return (
            f"Dzien dobry! Twoje zgloszenie nr {ticket_id} w PROMAG OZE zostalo przekazane do Dzialu Montazu. "
            f"Termin przyjazdu ekipy monterskiej zostanie wyznaczony w ciagu maksymalnie 5 dni roboczych. "
            f"Pozdrawiamy, PROMAG."
        )
    elif mode == "handlowiec_3dni":
        return (
            f"Dzien dobry! Twoje zapytanie nr {ticket_id} w PROMAG OZE zostalo przekazane do naszego doradcy handlowego. "
            f"Otrzymasz oferte lub kontakt telefoniczny w ciagu maksymalnie 3 dni roboczych. "
            f"Pozdrawiamy, PROMAG."
        )
    elif mode == "zaplanowano_wizyte":
        kw_txt = f" Szacowany koszt: {float(kwota):.0f} zl." if float(kwota) > 0 else ""
        return (
            f"Dzien dobry! Informujemy, ze wizyta serwisanta PROMAG OZE (zgloszenie {ticket_id}) "
            f"zostala zaplanowana na dzien {plan_date}. Konserwator {tech_name} skontaktuje sie telefonicznie przed przyjazdem.{kw_txt} "
            f"Pozdrawiamy, PROMAG."
        )
    return ""

def make_sms_uri(phone_raw, msg_text):
    phones = extract_phone_numbers(str(phone_raw), "")
    p_clean = phones[0] if phones else re.sub(r'\D', '', str(phone_raw))[-9:]
    if not p_clean:
        return None
    encoded = urllib.parse.quote(msg_text)
    return f"sms:+48{p_clean}?body={encoded}"

# ==========================================
# 1. KONFIGURACJA STRONY I STYLÓW CSS (STYL INPOST + PROMAG + GRYWALIZACJA)
# ==========================================
GOOGLE_SHEET_NAME = "Serwisowa posortowana"
LOCAL_EXCEL_FILE = "Serwisowa posortowana.xlsx"
MONTAZE_CSV_FILE = "skonczone.csv"
MONTAZE_PARQUET_CACHE = "skonczone_clean.parquet"
PROFILES_FILE = "profiles.json"
LOGO_FILE = "logo.png"

PIN_KOORDYNATOR = "1234"
PIN_ZARZAD = "9999"

DEFAULT_BURN_L_100KM = 9.0
DEFAULT_SPEED_KMH = 60.0
ROAD_TORTUOSITY = 1.3
TARGET_NET_PROFIT = 1200.0

DEFAULT_EMAIL_CONFIG = {
    "smtp_server": "smtp.gmail.com",
    "smtp_port": 587,
    "sender_email": "",
    "sender_password": "",
    "email_montaz": "montaz@promag-gaz.pl",
    "email_technik": "technik@promag-gaz.pl",
    "handlowcy": [
        {"nazwa": "Handlowiec 1 (Pompy / PV)", "email": "handlowiec1@promag-gaz.pl", "rejon": "Cała Polska"},
        {"nazwa": "Handlowiec 2 (Magazyny / Modernizacje)", "email": "handlowiec2@promag-gaz.pl", "rejon": "Mazowieckie"}
    ]
}

MIESIACE_PL = {
    1: "Styczeń", 2: "Luty", 3: "Marzec", 4: "Kwiecień",
    5: "Maj", 6: "Czerwiec", 7: "Lipiec", 8: "Sierpień",
    9: "Wrzesień", 10: "Październik", 11: "Listopad", 12: "Grudzień"
}

def get_early_favicon():
    if os.path.exists(LOGO_FILE):
        try:
            return Image.open(LOGO_FILE)
        except Exception:
            pass
    if os.path.exists(PROFILES_FILE):
        try:
            with open(PROFILES_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
                b64 = saved.get("logo_b64", None)
                if b64:
                    return Image.open(BytesIO(base64.b64decode(b64)))
        except Exception:
            pass
    return "🌿"

st.set_page_config(
    page_title="Promag Serwis",
    page_icon=get_early_favicon(),
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    .stButton > button {
        width: 100%;
        min-height: 50px;
        font-size: 15px !important;
        font-weight: 700 !important;
        border-radius: 12px !important;
        margin-bottom: 6px;
        transition: all 0.2s ease-in-out;
    }
    .promag-header {
        background: linear-gradient(135deg, #0a2540 0%, #10375c 100%);
        color: #ffffff;
        padding: 16px 22px;
        border-radius: 16px;
        border-bottom: 4px solid #16a34a;
        margin-bottom: 14px;
    }
    .inpost-welcome-card {
        background: linear-gradient(135deg, #0a2540 0%, #154360 100%);
        color: #ffffff;
        padding: 20px 22px;
        border-radius: 18px;
        border: 2px solid #16a34a;
        box-shadow: 0 6px 16px rgba(10, 37, 64, 0.16);
        margin-bottom: 14px;
    }
    .waty-pill {
        background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%);
        color: #ffffff;
        padding: 6px 14px;
        border-radius: 24px;
        font-size: 15px;
        font-weight: 900;
        display: inline-block;
        box-shadow: 0 2px 8px rgba(245, 158, 11, 0.35);
    }
    .streak-pill {
        background: #fef2f2;
        color: #dc2626;
        border: 1.5px solid #fca5a5;
        padding: 3px 10px;
        border-radius: 16px;
        font-size: 12.5px;
        font-weight: 800;
        display: inline-block;
        margin-top: 4px;
    }
    .inpost-tile {
        background: #ffffff;
        color: #0a2540;
        padding: 18px 16px;
        border-radius: 18px;
        border: 1.5px solid #cbd5e1;
        box-shadow: 0 4px 12px rgba(15, 23, 42, 0.06);
        margin-bottom: 12px;
        min-height: 156px;
    }
    .inpost-tile-label {
        font-size: 12px;
        font-weight: 800;
        text-transform: uppercase;
        letter-spacing: 0.6px;
        color: #64748b;
        margin-bottom: 4px;
    }
    .inpost-tile-val-green {
        font-size: 28px;
        font-weight: 900;
        color: #16a34a;
        line-height: 1.15;
    }
    .inpost-tile-val-blue {
        font-size: 30px;
        font-weight: 900;
        color: #0284c7;
        line-height: 1.15;
    }
    .inpost-tile-val-dark {
        font-size: 26px;
        font-weight: 900;
        color: #0a2540;
        line-height: 1.15;
    }
    .inpost-tile-val-amber {
        font-size: 24px;
        font-weight: 900;
        color: #d97706;
        line-height: 1.15;
    }
    .inpost-tile-sub {
        font-size: 13.5px;
        font-weight: 600;
        color: #334155;
        margin-top: 6px;
    }
    .badge-card-unlocked {
        background: linear-gradient(135deg, #f0fdf4 0%, #dcfce7 100%);
        border: 2px solid #16a34a;
        border-radius: 14px;
        padding: 12px 10px;
        text-align: center;
        margin-bottom: 8px;
        min-height: 118px;
    }
    .badge-card-locked {
        background: #f8fafc;
        border: 1.5px dashed #cbd5e1;
        border-radius: 14px;
        padding: 12px 10px;
        text-align: center;
        margin-bottom: 8px;
        opacity: 0.68;
        min-height: 118px;
    }
    .day-card {
        background: #ffffff;
        color: #0a2540;
        border: 1.5px solid #cbd5e1;
        border-radius: 16px;
        padding: 16px;
        margin-bottom: 10px;
        border-left: 7px solid #0284c7;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    }
    .day-card-urgent {
        background: #ffffff;
        color: #0a2540;
        border: 2px solid #ef4444;
        border-radius: 16px;
        padding: 16px;
        margin-bottom: 10px;
        border-left: 8px solid #ef4444;
        box-shadow: 0 4px 12px rgba(239,68,68,0.12);
    }
    .install-badge {
        background-color: #f0f9ff;
        border: 1px solid #bae6fd;
        border-left: 5px solid #0284c7;
        color: #0c4a6e;
        padding: 12px;
        border-radius: 12px;
        margin-top: 8px;
        margin-bottom: 10px;
        font-size: 14px;
    }
    .sms-box {
        background-color: #f0fdf4;
        border: 2px solid #16a34a;
        border-radius: 14px;
        padding: 14px;
        margin-top: 10px;
        margin-bottom: 12px;
        color: #064e3b;
    }
    .client-box {
        background-color: #ffffff;
        border: 2px solid #16a34a;
        border-radius: 16px;
        padding: 18px;
        margin-top: 12px;
        margin-bottom: 12px;
        color: #0a2540;
    }
    .onboard-card {
        background: #f8fafc;
        border: 1px solid #cbd5e1;
        border-left: 6px solid #16a34a;
        border-radius: 12px;
        padding: 16px;
        margin-bottom: 12px;
        color: #0f172a;
    }
    .step-progress-bar {
        display: flex;
        justify-content: space-between;
        align-items: center;
        background: #ffffff;
        padding: 12px 18px;
        border-radius: 14px;
        border: 1px solid #cbd5e1;
        margin-bottom: 14px;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. KONWERTER DATY + KODÓW POCZTOWYCH + PARSER BAZY MONTAŻY
# ==========================================
def extract_postal_code(text_str):
    s = str(text_str)
    m = re.search(r'\b(\d{2})\s*[-\s]\s*(\d{3})\b', s)
    if m:
        return f"{m.group(1)}-{m.group(2)}"
    return ""

def convert_week_or_date_to_month_year(raw_val):
    s = str(raw_val).strip()
    if not s or s.lower() in ["nan", "none", "nat", ""]:
        return "Brak daty"
    m_iso = re.search(r'\b(20\d{2})[-/.](0?[1-9]|1[0-2])[-/.](\d{1,2})\b', s)
    if m_iso:
        yr, mn = int(m_iso.group(1)), int(m_iso.group(2))
        return f"{MIESIACE_PL.get(mn, '')} {yr} ({yr}-{mn:02d})"
    m_pl = re.search(r'\b(\d{1,2})[-/.](0?[1-9]|1[0-2])[-/.](20\d{2})\b', s)
    if m_pl:
        mn, yr = int(m_pl.group(2)), int(m_pl.group(3))
        return f"{MIESIACE_PL.get(mn, '')} {yr} ({yr}-{mn:02d})"
    m_wk1 = re.search(r'\b([1-9]|[1-4]\d|5[0-3])\s*[/._-]\s*(20[1-3]\d)\b', s)
    if m_wk1:
        wk, yr = int(m_wk1.group(1)), int(m_wk1.group(2))
        try:
            d = datetime.date.fromisocalendar(yr, min(52, max(1, wk)), 4)
            return f"{MIESIACE_PL[d.month]} {d.year} ({d.year}-{d.month:02d})"
        except Exception:
            return f"Rok {yr}"
    m_wk2 = re.search(r'\b(20[1-3]\d)\s*[/._-]?[wWtT]?\s*([1-9]|[1-4]\d|5[0-3])\b', s)
    if m_wk2:
        yr, wk = int(m_wk2.group(1)), int(m_wk2.group(2))
        try:
            d = datetime.date.fromisocalendar(yr, min(52, max(1, wk)), 4)
            return f"{MIESIACE_PL[d.month]} {d.year} ({d.year}-{d.month:02d})"
        except Exception:
            return f"Rok {yr}"
    if s.isdigit() and 1 <= int(s) <= 53:
        wk = int(s)
        d = datetime.date.fromisocalendar(2024, min(52, wk), 4)
        return f"{MIESIACE_PL[d.month]} (Tydz. {wk})"
    return s

def extract_phone_numbers(raw_phone_str, raw_client_str=""):
    combined = f"{raw_phone_str} {raw_client_str}"
    combined = re.sub(r'\bNIP\s*[:.\-]?\s*(?:\d[\s\-]*){10}\b', ' ', combined, flags=re.IGNORECASE)
    combined = re.sub(r'\bPESEL\s*[:.\-]?\s*\d{11}\b', ' ', combined, flags=re.IGNORECASE)
    candidates = re.findall(r'(?:\+48)?\s*\d{3}[\s-]?\d{3}[\s-]?\d{3}|\d{2}[\s-]?\d{3}[\s-]?\d{2}[\s-]?\d{2}|\d{9,11}', combined)
    phones = []
    for c in candidates:
        digits = re.sub(r'\D', '', c)
        if len(digits) == 11 and digits.startswith("48"):
            digits = digits[2:]
        if len(digits) == 9 and not digits.startswith("0"):
            if digits not in phones:
                phones.append(digits)
    return phones

def clean_and_parse_montaze_df(raw_df):
    if raw_df is None or raw_df.empty:
        return pd.DataFrame(columns=["Klient", "Telefon", "Kod_Pocztowy", "Miasto", "Adres", "Data_Montazu", "Opis_Instalacji", "Search_Label"])

    expected_ready = {"Klient", "Telefon", "Kod_Pocztowy", "Miasto", "Adres", "Data_Montazu", "Opis_Instalacji", "Search_Label"}
    if expected_ready.issubset(set(raw_df.columns)):
        raw_df = raw_df.fillna("").copy()
        for idx, r in raw_df.iterrows():
            if not str(r["Kod_Pocztowy"]).strip():
                found_z = extract_postal_code(f"{r['Adres']} {r['Miasto']}")
                if found_z:
                    raw_df.at[idx, "Kod_Pocztowy"] = found_z
        raw_df["Search_Label"] = (
            raw_df["Klient"].astype(str) + " | Tel: " + raw_df["Telefon"].astype(str) +
            " | Adres: " + raw_df["Kod_Pocztowy"].astype(str) + " " + raw_df["Miasto"].astype(str) + " " + raw_df["Adres"].astype(str) +
            " | Montaż: " + raw_df["Data_Montazu"].astype(str)
        )
        return raw_df

    raw_df = raw_df.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)

    if any(str(c).startswith("Unnamed:") for c in raw_df.columns[:5]) and len(raw_df) > 1:
        first_row_vals = [str(v).strip() for v in raw_df.iloc[0].values]
        if any("klient" in v.lower() or "adres" in v.lower() for v in first_row_vals):
            new_cols = []
            counts = {}
            for idx, val in enumerate(raw_df.iloc[0].values):
                clean_v = re.sub(r'\s+', ' ', str(val).strip())
                if not clean_v or clean_v.lower() == 'nan':
                    clean_v = f"Kolumna_{idx}"
                if clean_v in counts:
                    counts[clean_v] += 1
                    clean_v = f"{clean_v}_{counts[clean_v]}"
                else:
                    counts[clean_v] = 1
                new_cols.append(clean_v)
            raw_df = raw_df.iloc[1:].reset_index(drop=True)
            raw_df.columns = new_cols

    cols = [re.sub(r'\s+', ' ', str(c).strip()) for c in raw_df.columns]
    raw_df.columns = cols

    if "Klient" in cols and any("Adres" in c for c in cols):
        adr_col = next((c for c in cols if "Adres" in c), cols[2] if len(cols) > 2 else "Klient")
        tydz_col = next((c for c in cols if "Tydzień realizacji" in c or "Tydzien realizacji" in c), None)
        plan_col = next((c for c in cols if "Planowany termin" in c), None)
        umowa_data_cols = [c for c in cols if any(k in c.lower() for k in ["data podpisania", "data wprowadzenia", "termin zakończenia"])]
        opis_zest_cols = [c for c in cols if any(k in c.lower() for k in ["opis zestawu", "opis z umowy", "dodatkowe podzespoły", "uruchomione instalacje", "dodatkowe opisy", "nazwa instalatora"])]

        records = []
        for _, row in raw_df.iterrows():
            raw_klient = str(row.get("Klient", "")).strip()
            if not raw_klient or raw_klient.lower() in ["nan", "none", "klient"]:
                continue

            raw_adr = str(row.get(adr_col, "")).strip()
            if raw_adr.lower() in ["nan", "none"]:
                raw_adr = ""

            phones = extract_phone_numbers(raw_adr, raw_klient)
            tel_str = " / ".join(phones)
            klient_clean = re.sub(r'\bNIP\s*[:.\-]?\s*(?:\d[\s\-]*){10}\b', '', raw_klient, flags=re.IGNORECASE)
            klient_clean = re.sub(r'(?:\+48)?\s*\d{3}[\s-]?\d{3}[\s-]?\d{3}|\d{2}[\s-]?\d{3}[\s-]?\d{2}[\s-]?\d{2}|\d{9,11}', '', klient_clean).strip(" ,-/") or raw_klient

            kod_p = extract_postal_code(raw_adr) or extract_postal_code(raw_klient)
            adr_no_zip = re.sub(r'\b\d{2}\s*[-\s]\s*\d{3}\b', '', raw_adr).strip(" ,")
            if "," in adr_no_zip:
                parts = [p.strip() for p in adr_no_zip.split(",", 1)]
                miasto, ulica = parts[0], parts[1]
            elif " ul." in adr_no_zip.lower() or " ul " in adr_no_zip.lower():
                idx_ul = adr_no_zip.lower().find(" ul")
                miasto = adr_no_zip[:idx_ul].strip(" ,")
                ulica = adr_no_zip[idx_ul:].strip(" ,")
            else:
                miasto, ulica = adr_no_zip, raw_adr

            raw_umowa_str = " ".join([str(row.get(uc, "")) for uc in umowa_data_cols])
            yr_m = re.search(r'\b(20[0-2]\d)\b', raw_umowa_str)
            yr_val = int(yr_m.group(1)) if yr_m else None
            wk_str = str(row.get(tydz_col, "") if tydz_col else "").strip()
            if not wk_str or wk_str.lower() in ["nan", "none"]:
                wk_str = str(row.get(plan_col, "") if plan_col else "").strip()

            wk_m = re.search(r'\b([1-9]|[1-4]\d|5[0-3])\b', wk_str)
            if wk_m and yr_val:
                wk_num = int(wk_m.group(1))
                try:
                    d_iso = datetime.date.fromisocalendar(yr_val, min(52, max(1, wk_num)), 4)
                    data_montazu = f"{MIESIACE_PL[d_iso.month]} {d_iso.year} ({d_iso.year}-{d_iso.month:02d})"
                except Exception:
                    data_montazu = convert_week_or_date_to_month_year(raw_umowa_str)
            else:
                data_montazu = convert_week_or_date_to_month_year(raw_umowa_str if yr_val else wk_str)

            op_list = []
            for oc in opis_zest_cols:
                v_op = str(row.get(oc, "")).strip()
                if v_op and v_op.lower() not in ["nan", "none", "-", "0", "0,00 zł"]:
                    op_list.append(f"{oc}: {v_op}")
            opis_inst = " | ".join(op_list)

            records.append({
                "Klient": klient_clean,
                "Telefon": tel_str,
                "Kod_Pocztowy": kod_p,
                "Miasto": miasto,
                "Adres": ulica,
                "Data_Montazu": data_montazu,
                "Opis_Instalacji": opis_inst,
                "Search_Label": f"{klient_clean} | Tel: {tel_str} | Adres: {kod_p} {miasto}, {ulica} | Montaż: {data_montazu}"
            })
        return pd.DataFrame(records)

    return pd.DataFrame(columns=["Klient", "Telefon", "Kod_Pocztowy", "Miasto", "Adres", "Data_Montazu", "Opis_Instalacji", "Search_Label"])

@st.cache_data(show_spinner="Wczytywanie bazy wszystkich klientów...")
def load_montaze_database():
    if os.path.exists(MONTAZE_PARQUET_CACHE):
        try:
            pq_df = pd.read_parquet(MONTAZE_PARQUET_CACHE)
            return clean_and_parse_montaze_df(pq_df)
        except Exception:
            pass
    for candidate in ["skonczone_gotowe.csv", "skonczone_lekkie.csv", MONTAZE_CSV_FILE, "Promag/skonczone_lekkie.csv", "Promag/skonczone.csv", "skończone.csv", "skończone.xlsx"]:
        if os.path.exists(candidate):
            try:
                if candidate.endswith(".xlsx"):
                    raw_df = pd.read_excel(candidate)
                else:
                    try:
                        raw_df = pd.read_csv(candidate, sep=None, engine="python", on_bad_lines="skip")
                    except Exception:
                        raw_df = pd.read_csv(candidate, encoding="cp1250", sep=None, engine="python", on_bad_lines="skip")
                clean_df = clean_and_parse_montaze_df(raw_df)
                try:
                    clean_df.to_parquet(MONTAZE_PARQUET_CACHE, index=False)
                except Exception:
                    pass
                return clean_df
            except Exception:
                pass
    return pd.DataFrame(columns=["Klient", "Telefon", "Kod_Pocztowy", "Miasto", "Adres", "Data_Montazu", "Opis_Instalacji", "Search_Label"])

def generate_google_contacts_batches(df_montaze, df_tickets=None, batch_size=2500):
    contacts_by_phone = {}
    email_regex = re.compile(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}')

    def process_record(raw_name, raw_phone, extra_text=""):
        combined_text = f"{raw_name} {raw_phone} {extra_text}"
        phones = extract_phone_numbers(raw_phone, f"{raw_name} {extra_text}")
        if not phones:
            return
        emails = email_regex.findall(combined_text)
        email_val = emails[0].lower() if emails else ""

        clean_name = re.sub(r'\bNIP\s*[:.\-]?\s*(?:\d[\s\-]*){10}\b', ' ', str(raw_name), flags=re.IGNORECASE)
        clean_name = re.sub(r'(?:\+48)?\s*\d{3}[\s-]?\d{3}[\s-]?\d{3}|\d{2}[\s-]?\d{3}[\s-]?\d{2}[\s-]?\d{2}|\d{9,11}', '', clean_name)
        clean_name = email_regex.sub('', clean_name)
        clean_name = re.sub(r'\b(tel\.?|kom\.?|nr)\b', '', clean_name, flags=re.IGNORECASE)
        clean_name = re.sub(r'[\(\)\[\]\/\\,-]+', ' ', clean_name)
        clean_name = re.sub(r'\s+', ' ', clean_name).strip()
        if not clean_name or clean_name.lower() in ["nan", "none", "klient", "brak"]:
            clean_name = f"Klient PROMAG {phones[0]}"

        main_phone = f"+48{phones[0]}"
        sec_phone = f"+48{phones[1]}" if len(phones) > 1 else ""
        parts = clean_name.split(" ", 1)
        given_name = parts[0]
        family_name = parts[1] if len(parts) > 1 else ""

        if main_phone not in contacts_by_phone:
            contacts_by_phone[main_phone] = {
                "Name": clean_name, "Given Name": given_name, "Family Name": family_name,
                "E-mail 1 - Type": "Home" if email_val else "", "E-mail 1 - Value": email_val,
                "Phone 1 - Type": "Mobile", "Phone 1 - Value": main_phone,
                "Phone 2 - Type": "Home" if sec_phone else "", "Phone 2 - Value": sec_phone,
                "Group Membership": "* myContacts ::: Klienci PROMAG"
            }
        else:
            existing = contacts_by_phone[main_phone]
            if not existing["E-mail 1 - Value"] and email_val:
                existing["E-mail 1 - Type"] = "Home"
                existing["E-mail 1 - Value"] = email_val
            if not existing["Phone 2 - Value"] and sec_phone:
                existing["Phone 2 - Type"] = "Home"
                existing["Phone 2 - Value"] = sec_phone

    if df_montaze is not None and not df_montaze.empty:
        for _, r in df_montaze.iterrows():
            process_record(r.get("Klient", ""), r.get("Telefon", ""), r.get("Opis_Instalacji", ""))

    if df_tickets is not None and not df_tickets.empty:
        for _, r in df_tickets.iterrows():
            process_record(r.get("Klient", ""), r.get("Telefon", ""), r.get("Opis_Zgloszenia", ""))

    all_contacts_df = pd.DataFrame(list(contacts_by_phone.values()))
    if all_contacts_df.empty:
        return [], b"", 0

    batches = []
    total = len(all_contacts_df)
    zip_buf = BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for start_idx in range(0, total, batch_size):
            chunk = all_contacts_df.iloc[start_idx:start_idx + batch_size]
            batch_num = (start_idx // batch_size) + 1
            fname = f"Kontakty_Google_PROMAG_Paczka_{batch_num}.csv"
            csv_bytes = chunk.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
            batches.append((fname, csv_bytes, len(chunk)))
            zf.writestr(fname, csv_bytes)

    return batches, zip_buf.getvalue(), total

# ==========================================
# 3. FUNKCJE GPS I PLANERA KORYTARZOWEGO
# ==========================================
def resolve_postal_or_city_coords(postal_str, city_str="", address_str=""):
    combined = f"{postal_str} {city_str} {address_str}"
    found_zip = extract_postal_code(combined)
    if found_zip:
        sub4 = found_zip[:4]
        pref2 = found_zip[:2]
        last_digits = int(found_zip[-3:])
        offset_lat = ((last_digits // 10) - 50) * 0.0025
        offset_lon = ((last_digits % 10) - 5) * 0.004
        if sub4 in PRECISE_SUBPREFIX_COORDS:
            b_lat, b_lon = PRECISE_SUBPREFIX_COORDS[sub4]
            return b_lat + offset_lat, b_lon + offset_lon, found_zip
        if pref2 in FULL_POSTAL_PREFIX_COORDS:
            b_lat, b_lon = FULL_POSTAL_PREFIX_COORDS[pref2]
            return b_lat + offset_lat, b_lon + offset_lon, found_zip

    m2 = re.search(r'^\s*(\d{2})\s*$', str(postal_str))
    if m2 and m2.group(1) in FULL_POSTAL_PREFIX_COORDS:
        b_lat, b_lon = FULL_POSTAL_PREFIX_COORDS[m2.group(1)]
        return b_lat, b_lon, f"{m2.group(1)}-000"

    low_comb = combined.lower()
    for city_k, (c_lat, c_lon) in CITY_FALLBACK_COORDS.items():
        if city_k in low_comb:
            return c_lat, c_lon, found_zip

    return None, None, found_zip

def extract_coords_from_row(row):
    lat, lon, found_zip = resolve_postal_or_city_coords(
        row.get("Kod_Pocztowy", ""),
        row.get("Miasto", ""),
        row.get("Adres", "")
    )
    return lat, lon, found_zip

def get_coords_for_postal(postal_str):
    lat, lon, _ = resolve_postal_or_city_coords(postal_str, postal_str, "")
    if lat is not None and lon is not None:
        return lat, lon
    return 52.08, 20.85

def haversine_km(lat1, lon1, lat2, lon2):
    R = 6371.0
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlambda = math.radians(lon2 - lon1)
    a = math.sin(dphi / 2.0)**2 + math.cos(phi1) * math.cos(phi2) * math.sin(dlambda / 2.0)**2
    return 2 * R * math.atan2(math.sqrt(a), math.sqrt(1 - a))

def distance_from_segment_km(p_lat, p_lon, a_lat, a_lon, b_lat, b_lon):
    lat_scale = 111.32
    lon_scale = 111.32 * math.cos(math.radians((a_lat + b_lat) / 2.0))
    px, py = (p_lon - a_lon) * lon_scale, (p_lat - a_lat) * lat_scale
    bx, by = (b_lon - a_lon) * lon_scale, (b_lat - a_lat) * lat_scale
    seg_len_sq = bx * bx + by * by
    if seg_len_sq == 0:
        return math.hypot(px, py)
    t = max(0.0, min(1.0, (px * bx + py * by) / seg_len_sq))
    proj_x, proj_y = t * bx, t * by
    return math.hypot(px - proj_x, py - proj_y)

def estimate_route_economics(points_coords, base_coords, gross_revenue, parts_cost=0.0,
                             hourly_rate=STAWKA_RBH_NETTO, fuel_price=KOSZT_PALIWA_LITR,
                             burn_l_100km=DEFAULT_BURN_L_100KM, onsite_hours_per_client=SR_CZAS_WIZYTY_H):
    if not points_coords:
        return {
            "road_km": 0.0, "drive_hours": 0.0, "onsite_hours": 0.0,
            "fuel_cost": 0.0, "labor_drive_cost": 0.0, "labor_onsite_cost": 0.0,
            "total_cost": 0.0, "net_profit": 0.0
        }
    remaining = list(points_coords)
    curr_lat, curr_lon = base_coords
    straight_km = 0.0
    while remaining:
        nearest_idx = min(range(len(remaining)), key=lambda i: haversine_km(curr_lat, curr_lon, remaining[i][0], remaining[i][1]))
        nxt_lat, nxt_lon = remaining.pop(nearest_idx)
        straight_km += haversine_km(curr_lat, curr_lon, nxt_lat, nxt_lon)
        curr_lat, curr_lon = nxt_lat, nxt_lon
    straight_km += haversine_km(curr_lat, curr_lon, base_coords[0], base_coords[1])

    road_km = straight_km * ROAD_TORTUOSITY
    drive_hours = road_km / DEFAULT_SPEED_KMH
    onsite_hours = len(points_coords) * onsite_hours_per_client

    fuel_cost = (road_km / 100.0) * burn_l_100km * fuel_price
    labor_drive_cost = drive_hours * hourly_rate
    labor_onsite_cost = onsite_hours * hourly_rate
    total_cost = fuel_cost + labor_drive_cost + labor_onsite_cost + parts_cost
    net_profit = gross_revenue - total_cost

    return {
        "road_km": round(road_km, 1),
        "drive_hours": round(drive_hours, 1),
        "onsite_hours": round(onsite_hours, 1),
        "fuel_cost": round(fuel_cost, 2),
        "labor_drive_cost": round(labor_drive_cost, 2),
        "labor_onsite_cost": round(labor_onsite_cost, 2),
        "total_cost": round(total_cost, 2),
        "net_profit": round(net_profit, 2)
    }

def send_system_email(email_cfg, to_emails, subject, body_text, attachments=None):
    sender = email_cfg.get("sender_email", "").strip()
    password = email_cfg.get("sender_password", "").strip()
    server_host = email_cfg.get("smtp_server", "smtp.gmail.com").strip()
    port = int(email_cfg.get("smtp_port", 587))

    if not sender or not password or not to_emails:
        return False, "Brak skonfigurowanego konta SMTP w zakładce Ustawienia."

    to_list = [x.strip() for x in (to_emails.split(",") if isinstance(to_emails, str) else to_emails) if str(x).strip()]
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = ", ".join(to_list)
    msg.set_content(body_text)

    if attachments:
        for att in attachments:
            try:
                msg.add_attachment(att.getvalue(), maintype="application", subtype="octet-stream", filename=att.name)
            except Exception:
                pass

    try:
        if port == 465:
            with smtplib.SMTP_SSL(server_host, port, timeout=10) as server:
                server.login(sender, password)
                server.send_message(msg)
        else:
            with smtplib.SMTP(server_host, port, timeout=10) as server:
                server.starttls()
                server.login(sender, password)
                server.send_message(msg)
        return True, f"Wysłano e-mail do: {', '.join(to_list)}"
    except Exception as e:
        return False, f"Błąd SMTP: {e}"

def load_config_and_profiles():
    profiles = DEFAULT_TECHS.copy()
    logo_b64 = None
    email_cfg = DEFAULT_EMAIL_CONFIG.copy()
    cennik_tab = [dict(x) for x in TABELA_CENNIKA]
    stawka_rbh = STAWKA_RBH_NETTO
    dojazd_kw = RYCZALT_DOJAZD

    if os.path.exists(PROFILES_FILE):
        try:
            with open(PROFILES_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
                profiles = saved.get("techs", DEFAULT_TECHS)
                logo_b64 = saved.get("logo_b64", None)
                email_cfg = saved.get("email_cfg", DEFAULT_EMAIL_CONFIG)
                cennik_tab = saved.get("cennik_tab", TABELA_CENNIKA)
                stawka_rbh = float(saved.get("stawka_rbh", STAWKA_RBH_NETTO))
                dojazd_kw = float(saved.get("dojazd_kw", RYCZALT_DOJAZD))
        except Exception:
            pass
    return ensure_tech_profile_fields(profiles), logo_b64, email_cfg, cennik_tab, stawka_rbh, dojazd_kw

def save_config_and_profiles(profiles, logo_b64=None, email_cfg=None, cennik_tab=None, stawka_rbh=None, dojazd_kw=None):
    if logo_b64 is None:
        logo_b64 = st.session_state.get("logo_b64", None)
    if email_cfg is None:
        email_cfg = st.session_state.get("email_cfg", DEFAULT_EMAIL_CONFIG)
    if cennik_tab is None:
        cennik_tab = st.session_state.get("cennik_tab", TABELA_CENNIKA)
    if stawka_rbh is None:
        stawka_rbh = st.session_state.get("stawka_rbh", STAWKA_RBH_NETTO)
    if dojazd_kw is None:
        dojazd_kw = st.session_state.get("dojazd_kw", RYCZALT_DOJAZD)

    profiles = ensure_tech_profile_fields(profiles)
    payload = {
        "techs": profiles, "logo_b64": logo_b64, "email_cfg": email_cfg,
        "cennik_tab": cennik_tab, "stawka_rbh": float(stawka_rbh), "dojazd_kw": float(dojazd_kw)
    }
    try:
        with open(PROFILES_FILE, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def render_brand_logo(max_width=220, centered=False):
    logo_b64 = st.session_state.get("logo_b64", None)
    if logo_b64:
        align_style = "display: block; margin-left: auto; margin-right: auto;" if centered else ""
        st.markdown(
            f'<img src="data:image/png;base64,{logo_b64}" style="max-width:{max_width}px; max-height:85px; object-fit:contain; {align_style} margin-bottom:8px;" alt="PROMAG OZE">',
            unsafe_allow_html=True
        )
    elif os.path.exists(LOGO_FILE):
        st.image(LOGO_FILE, width=max_width)
    else:
        align = "text-align:center;" if centered else ""
        st.markdown(
            f"<div class='promag-header' style='{align}'><span style='font-size:22px; font-weight:800;'>🌿 PROMAG SERWIS OZE</span><br><span style='font-size:13px; color:#86efac;'>Fotowoltaika • Pompy Ciepła • Kolektory Słoneczne • Magazyny Energii</span></div>",
            unsafe_allow_html=True
        )

# ==========================================
# 4. TRWAŁY ZAPIS I ODCZYT (GOOGLE SHEETS + LOCAL EXCEL FALLBACK)
# ==========================================
REQUIRED_COLS = [
    "ID_Zgloszenia", "Klient", "Telefon", "Kod_Pocztowy", "Miasto", "Adres",
    "Typ_Instalacji", "Opis_Zgloszenia", "Status", "Serwisant", "Data_Planowana",
    "Kwota_Przewidywana", "Zgoda_Klienta_Kwota", "Data_Montazu", "Opis_Instalacji_Montaz",
    "Priorytet_Krytyczny", "Typ_Serwisu", "Wymagany_Przeglad", "Naprawione_1_Wizyta",
    "Kwota_Koncowa", "Koszt_Materialow", "Forma_Platnosci", "Gotowka_Rozliczona",
    "Czas_Pracy_Min", "Wyniki_Przegladu", "Zdjecia_Potwierdzone", "Zalecenia_Klienta",
    "Notatka_Serwisanta", "Powod_Niewykonania", "Data_Zamkniecia", "Podpis_Zlozony",
    "Do_Oddzwonienia", "Data_Oddzwonienia", "Potwierdzony_Tel", "Trudny_Klient",
    "Potrzebne_Czesci", "Lead_Handlowy", "Oddelegowano_Dzial", "Oddelegowano_Email",
    "Data_Oddelegowania", "Notatka_Dla_Dzialu", "Otrzymano_Oferte_Termin", "Ostatnie_Przypomnienie"
]

NUMERIC_COLS = ["Kwota_Przewidywana", "Zgoda_Klienta_Kwota", "Kwota_Koncowa", "Koszt_Materialow", "Czas_Pracy_Min"]

def get_gsheet_worksheet():
    if not GSPREAD_AVAILABLE:
        return None
    try:
        if "gcp_service_account" in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
        elif "connections" in st.secrets and "gsheets" in st.secrets["connections"]:
            creds_dict = dict(st.secrets["connections"]["gsheets"])
        else:
            return None
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        client = gspread.authorize(creds)
        sh = client.open(GOOGLE_SHEET_NAME)
        return sh.get_worksheet(0)
    except Exception:
        return None

def ensure_columns(df):
    if df is None or df.empty:
        empty_df = pd.DataFrame(columns=REQUIRED_COLS)
        for nc in NUMERIC_COLS:
            empty_df[nc] = empty_df[nc].astype(float)
        return empty_df

    clean_cols, cols_to_keep = [], []
    for idx, c in enumerate(df.columns):
        c_str = str(c).strip()
        if c_str != "" and not c_str.lower().startswith("unnamed"):
            cols_to_keep.append(idx)
            clean_cols.append(c_str)

    df = df.iloc[:, cols_to_keep].copy()
    df.columns = clean_cols
    df = df.loc[:, ~df.columns.duplicated()].copy().reset_index(drop=True)

    col_map = {}
    used_targets = set(df.columns)
    for c in df.columns:
        cl = c.lower()
        target = None
        if cl in ["dane", "klient", "imię i nazwisko", "imie i nazwisko"]:
            target = "Klient"
        elif "opis_od_klienta" in cl or "opis od klienta" in cl or cl == "opis":
            target = "Opis_Zgloszenia"
        elif "opis_serwisowy" in cl or "informacje_dodatkowe" in cl:
            target = "Notatka_Serwisanta"
        elif cl == "kwota":
            target = "Kwota_Przewidywana"
        elif "data_serwisu" in cl or "data serwisu" in cl:
            target = "Data_Planowana"
        elif "platnosc" in cl or "płatność" in cl:
            target = "Forma_Platnosci"
        if target and target not in used_targets:
            col_map[c] = target
            used_targets.add(target)

    if col_map:
        df = df.rename(columns=col_map)

    if "Opis_Zgloszenia" not in df.columns and "Opis_Od_Klienta" in df.columns:
        df["Opis_Zgloszenia"] = df["Opis_Od_Klienta"].fillna("").astype(str)
    if "Opis_Serwisowy" in df.columns:
        opis_serw = df["Opis_Serwisowy"].fillna("").astype(str)
        if "Opis_Zgloszenia" in df.columns:
            df["Opis_Zgloszenia"] = (df["Opis_Zgloszenia"].fillna("").astype(str) + " " + opis_serw).str.strip()
        else:
            df["Opis_Zgloszenia"] = opis_serw

    for rc in REQUIRED_COLS:
        if rc not in df.columns:
            if rc == "ID_Zgloszenia":
                df[rc] = [f"ZGL-{i+1001}" for i in range(len(df))]
            elif rc == "Status":
                df[rc] = "Nowe"
            elif rc in ["Priorytet_Krytyczny", "Wymagany_Przeglad", "Zdjecia_Potwierdzone", "Podpis_Zlozony", "Do_Oddzwonienia", "Potwierdzony_Tel", "Trudny_Klient", "Otrzymano_Oferte_Termin"]:
                df[rc] = "Nie"
            elif rc in ["Naprawione_1_Wizyta", "Gotowka_Rozliczona"]:
                df[rc] = "Tak"
            elif rc == "Typ_Serwisu":
                df[rc] = "Płatny"
            elif rc == "Forma_Platnosci":
                df[rc] = "Przelew"
            elif rc in NUMERIC_COLS:
                df[rc] = 0.0
            else:
                df[rc] = ""

    df = df[REQUIRED_COLS].copy()
    for col in df.columns:
        if col in NUMERIC_COLS:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0).astype(float)
        else:
            df[col] = df[col].fillna("").astype(str).replace(["None", "nan", "NaT", "<NA>"], "")

    for idx, r in df.iterrows():
        raw_kl = str(r["Klient"]).strip()
        raw_tel = str(r["Telefon"]).strip()
        if not raw_tel and raw_kl:
            found_p = extract_phone_numbers("", raw_kl)
            if found_p:
                df.at[idx, "Telefon"] = " / ".join(found_p)
                clean_k = re.sub(r'(?:\+48)?\s*\d{3}[\s-]?\d{3}[\s-]?\d{3}|\d{2}[\s-]?\d{3}[\s-]?\d{2}[\s-]?\d{2}|\d{9,11}', '', raw_kl).strip(" ,-/")
                if clean_k:
                    df.at[idx, "Klient"] = clean_k

        raw_adr = f"{r['Adres']} {r['Miasto']}".strip()
        if not str(r["Kod_Pocztowy"]).strip() and raw_adr:
            found_z = extract_postal_code(raw_adr)
            if found_z:
                df.at[idx, "Kod_Pocztowy"] = found_z
        if not str(r["Miasto"]).strip() and raw_adr:
            adr_no_z = re.sub(r'\b\d{2}\s*[-\s]\s*\d{3}\b', '', str(r["Adres"])).strip(" ,")
            if "," in adr_no_z:
                df.at[idx, "Miasto"] = adr_no_z.split(",", 1)[0].strip()
            elif " ul" in adr_no_z.lower():
                df.at[idx, "Miasto"] = adr_no_z[:adr_no_z.lower().find(" ul")].strip(" ,")

        st_val = str(r["Status"]).strip().lower()
        if "zako" in st_val:
            df.at[idx, "Status"] = "Zakończone"
        elif "częśc" in st_val or "czesc" in st_val:
            df.at[idx, "Status"] = "Do zamówienia części"
        elif "niewyk" in st_val or "nie uda" in st_val:
            df.at[idx, "Status"] = "Niewykonane"
        elif "oddeleg" in st_val:
            df.at[idx, "Status"] = "Oddelegowane"
        elif st_val in ["", "brak", "nan"]:
            df.at[idx, "Status"] = "Nowe"

    return df.reset_index(drop=True)

def load_tickets():
    ws = get_gsheet_worksheet()
    if ws is not None:
        try:
            records = ws.get_all_records()
            if records:
                df_gs = ensure_columns(pd.DataFrame(records))
                st.session_state["storage_mode"] = "Google Sheets (Chmura - Trwały Zapis)"
                try:
                    df_gs.to_excel(LOCAL_EXCEL_FILE, index=False)
                except Exception:
                    pass
                return df_gs
        except Exception:
            pass

    st.session_state["storage_mode"] = "Dysk serwera (Serwisowa posortowana.xlsx)"
    if not os.path.exists(LOCAL_EXCEL_FILE):
        empty_df = ensure_columns(pd.DataFrame(columns=REQUIRED_COLS))
        try:
            empty_df.to_excel(LOCAL_EXCEL_FILE, index=False)
        except Exception:
            pass
        return empty_df
    try:
        df = pd.read_excel(LOCAL_EXCEL_FILE)
        return ensure_columns(df)
    except Exception:
        return ensure_columns(pd.DataFrame(columns=REQUIRED_COLS))

def save_tickets(df):
    clean_df = ensure_columns(df)
    st.session_state["df_tickets"] = clean_df.copy()
    try:
        clean_df.to_excel(LOCAL_EXCEL_FILE, index=False)
    except Exception:
        pass

    ws = get_gsheet_worksheet()
    if ws is not None:
        try:
            export_df = clean_df.copy()
            for c in export_df.columns:
                export_df[c] = export_df[c].fillna("").astype(str)
            ws.clear()
            ws.update([export_df.columns.values.tolist()] + export_df.values.tolist())
            st.session_state["storage_mode"] = "Google Sheets (Chmura - Trwały Zapis)"
        except Exception:
            pass

def check_and_send_3day_reminders(df, email_cfg):
    today = datetime.date.today()
    reminders_sent = 0
    overdue_ids = []
    changed = False

    for idx, r in df.iterrows():
        dzial = str(r.get("Oddelegowano_Dzial", "")).strip()
        odbiorca = str(r.get("Oddelegowano_Email", "")).strip()
        otrzymano = str(r.get("Otrzymano_Oferte_Termin", "Nie")).strip()
        data_odd_str = str(r.get("Data_Oddelegowania", "")).strip()[:10]
        ost_przyp_str = str(r.get("Ostatnie_Przypomnienie", "")).strip()[:10]

        if dzial and odbiorca and otrzymano != "Tak" and data_odd_str:
            try:
                d_odd = datetime.date.fromisoformat(data_odd_str)
                days_diff = (today - d_odd).days
                if days_diff >= 3:
                    overdue_ids.append(r["ID_Zgloszenia"])
                    should_send = True
                    if ost_przyp_str:
                        try:
                            if (today - datetime.date.fromisoformat(ost_przyp_str)).days < 3:
                                should_send = False
                        except Exception:
                            pass
                    if should_send and email_cfg.get("sender_email") and email_cfg.get("sender_password"):
                        subj = f"🚨 PRZYPOMNIENIE (Minęło {days_diff} dni): Brak oferty/terminu - {r['ID_Zgloszenia']} ({r['Klient']})"
                        body = "\n".join([
                            "Dzień dobry,", "",
                            f"System PROMAG OZE przypomina o oddelegowanym zleceniu z dnia {data_odd_str} ({dzial}), dla którego nadal nie zaznaczono przygotowania oferty lub terminu.",
                            f"- Klient: {r['Klient']} (Tel: {r['Telefon']})",
                            f"- Adres: {r['Kod_Pocztowy']} {r['Miasto']}, {r['Adres']}",
                            f"- Opis: {r['Opis_Zgloszenia']}",
                            f"- Notatka dla działu: {r['Notatka_Dla_Dzialu']}"
                        ])
                        ok, _ = send_system_email(email_cfg, odbiorca, subj, body)
                        if ok:
                            df.at[idx, "Ostatnie_Przypomnienie"] = str(today)
                            reminders_sent += 1
                            changed = True
            except Exception:
                pass
    if changed:
        save_tickets(df)
    return overdue_ids, reminders_sent

# ==========================================
# 5. INICJALIZACJA SESJI + LOGOWANIE KODEM PIN + ONBOARDING
# ==========================================
if "profiles" not in st.session_state:
    p_l, l_l, e_l, c_l, s_l, d_l = load_config_and_profiles()
    st.session_state["profiles"] = p_l
    st.session_state["logo_b64"] = l_l
    st.session_state["email_cfg"] = e_l
    st.session_state["cennik_tab"] = c_l
    st.session_state["stawka_rbh"] = s_l
    st.session_state["dojazd_kw"] = d_l

if "df_tickets" not in st.session_state:
    st.session_state["df_tickets"] = load_tickets()
else:
    st.session_state["df_tickets"] = ensure_columns(st.session_state["df_tickets"])

if "df_montaze" not in st.session_state:
    st.session_state["df_montaze"] = load_montaze_database()

for k_s in ["logged_role", "logged_user", "active_day", "active_ticket_id", "last_created_sms", "last_planned_route_sms", "victory_banner"]:
    if k_s not in st.session_state:
        st.session_state[k_s] = None

if "tech_nav_tab" not in st.session_state:
    st.session_state["tech_nav_tab"] = "start"

if "visit_step" not in st.session_state:
    st.session_state["visit_step"] = 1

if "daily_quote" not in st.session_state:
    st.session_state["daily_quote"] = random.choice(MYSLI_NA_DZIS_QUOTES)

def resolve_pin_to_user(pin_code, profiles_dict):
    p_clean = str(pin_code).strip()
    if not p_clean:
        return None, None
    if p_clean == PIN_KOORDYNATOR:
        return "koordynator", "Koordynator"
    if p_clean == PIN_ZARZAD:
        return "zarzad", "Zarząd"
    for tech_name, t_data in profiles_dict.items():
        if str(t_data.get("pin", "")).strip() == p_clean:
            return "serwisant", tech_name
    return None, None

if st.session_state["logged_role"] is None:
    qp_token = st.query_params.get("token", None)
    if qp_token:
        r_found, u_found = resolve_pin_to_user(qp_token, st.session_state["profiles"])
        if r_found:
            st.session_state["logged_role"] = r_found
            st.session_state["logged_user"] = u_found

if st.session_state["logged_role"] is None:
    render_brand_logo(max_width=240, centered=True)

    col_login, col_onboard = st.columns([1, 1.35])
    with col_login:
        with st.container(border=True):
            st.subheader("🔐 Szybkie Logowanie PIN")
            st.caption("Wpisz swój unikalny 4-cyfrowy kod PIN – system sam rozpozna Twój profil i rolę.")
            pin_input = st.text_input("Twój 4-cyfrowy kod PIN:", type="password", max_chars=8, placeholder="••••")
            remember_me = st.checkbox("🔒 Zapamiętaj mnie na tym urządzeniu", value=True)

            if st.button("🌿 Wejdź do systemu", type="primary"):
                role_res, user_res = resolve_pin_to_user(pin_input, st.session_state["profiles"])
                if role_res:
                    st.session_state["logged_role"] = role_res
                    st.session_state["logged_user"] = user_res
                    st.session_state["tech_nav_tab"] = "start"
                    st.session_state["active_day"] = None
                    st.session_state["active_ticket_id"] = None
                    st.session_state["visit_step"] = 1
                    st.session_state["daily_quote"] = random.choice(MYSLI_NA_DZIS_QUOTES)
                    if remember_me:
                        st.query_params["token"] = str(pin_input).strip()
                    st.rerun()
                else:
                    st.error("❌ Nieznany kod PIN. Sprawdź kod i spróbuj ponownie.")

    with col_onboard:
        with st.container(border=True):
            st.subheader("🎓 Onboarding i Grywalizacja PROMAG Serwis")
            st.markdown("""
            <div class="onboard-card">
                <b>🏠 1. Start (Styl InPost + ⚡ Waty i 🔥 Seria Combo):</b><br>
                • Zbieraj <b>⚡ Waty</b> za zamknięte wizyty (+10 ⚡), pilne awarie (+50 ⚡) i Leady (+25 ⚡). Buduj <b>🔥 Serię napraw za 1. razem</b> i odblokowuj odznaki!
            </div>
            <div class="onboard-card">
                <b>🗺️ 2. Plan (Twoja trasa + 3-krokowa Karta Wizyty ze strzałkami):</b><br>
                • Rano sprawdzasz <b>Checklistę załadunku</b>. U klienta przechodzisz przez 3 kroki dużymi przyciskami <b>⬅ WSTECZ</b> / <b>DALEJ ➡</b>.
            </div>
            <div class="onboard-card">
                <b>📋 3. Przegląd & 🚐 4. Bus:</b><br>
                • W zakładce <b>Przegląd</b> wypełniasz interaktywne protokoły przeglądowe (+15 ⚡), a w zakładce <b>Bus</b> kontrolujesz stan części i narzędzi na pace.
            </div>
            """, unsafe_allow_html=True)
    st.stop()

# Nagłówek zalogowanego użytkownika
top_c1, top_c2, top_c3 = st.columns([2.1, 1.6, 0.8])
with top_c1:
    render_brand_logo(max_width=185, centered=False)
with top_c2:
    st.markdown(f"**Użytkownik:** `{st.session_state['logged_user']}` ({st.session_state['logged_role'].capitalize()})")
    st.caption(f"💾 Baza: {st.session_state.get('storage_mode', 'Lokalna')}")
with top_c3:
    if st.button("🚪 Wyloguj"):
        st.query_params.clear()
        st.session_state["logged_role"] = None
        st.session_state["logged_user"] = None
        st.rerun()

# ==========================================
# 6. NOWY PANEL SERWISANTA W STYLU INPOST + GRYWALIZACJA (SERIA COMBO, WATY, ODZNAKI)
# ==========================================
if st.session_state["logged_role"] == "serwisant":
    tech_name = st.session_state["logged_user"]
    df = st.session_state["df_tickets"]
    profiles = st.session_state["profiles"]
    my_prof = profiles[tech_name]
    metrics = calculate_tech_metrics(df, tech_name)
    gami = calculate_gamification_stats(df, tech_name, my_prof)

    # Efekt Zwycięstwa po zamknięciu wizyty (jeśli przed chwilą zamknięto zlecenie)
    if st.session_state.get("victory_banner"):
        vb = st.session_state["victory_banner"]
        st.balloons()
        st.success(vb)
        st.session_state["victory_banner"] = None

    today_dt = datetime.date.today()
    today_str = str(today_dt)
    dzien_tyg = DNI_TYGODNIA_PL.get(today_dt.weekday(), "")
    data_pl_str = f"{dzien_tyg}, {today_dt.strftime('%d.%m.%Y')}"

    my_open_tickets = df[(df["Serwisant"] == tech_name) & (df["Status"].isin(["Zaplanowane", "W trakcie"]))].copy()
    if not my_open_tickets.empty:
        my_open_tickets["Data_Grupa"] = my_open_tickets["Data_Planowana"].astype(str).str.strip().replace("", today_str)
    else:
        my_open_tickets["Data_Grupa"] = pd.Series(dtype=str)

    active_tab = st.session_state.get("tech_nav_tab", "start")

    # ----------------------------------------------------
    # ZAKŁADKA 1: START (EKRAN GŁÓWNY W STYLU INPOST + GRYWALIZACJA)
    # ----------------------------------------------------
    if active_tab == "start":
        # 1. Górny kafelek powitalny: Cześć [Imię], licznik Watów ⚡ w rogu (jak punkty InPost), data + Myśl na dziś
        st.markdown(f"""
        <div class="inpost-welcome-card">
            <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:8px;">
                <div>
                    <span style="font-size:24px; font-weight:900; color:#ffffff;">👋 Cześć, {tech_name}!</span>
                    <div style="font-size:13.5px; font-weight:700; color:#86efac; margin-top:2px;">📅 Dziś jest {data_pl_str}</div>
                </div>
                <div class="waty-pill" title="Zbieraj Waty za zamknięte wizyty (+10), pilne awarie (+50), protokoły (+15) i leady (+25)!">
                    ⚡ {gami['total_waty']} Watów
                </div>
            </div>
            <div style="margin-top:12px; font-size:12px; font-weight:800; text-transform:uppercase; color:#38bdf8; letter-spacing:0.6px;">
                💡 Myśl serwisowa na dziś:
            </div>
            <div style="margin-top:4px; font-size:15px; font-style:italic; color:#f8fafc; line-height:1.4;">
                {st.session_state['daily_quote']}
            </div>
        </div>
        """, unsafe_allow_html=True)

        if st.button("🎲 Wylosuj inną myśl na dziś"):
            st.session_state["daily_quote"] = random.choice(MYSLI_NA_DZIS_QUOTES)
            st.rerun()

        pts_total = len(my_open_tickets)
        cities_today = ", ".join([m for m in my_open_tickets["Miasto"].astype(str).unique() if m.strip()][:3]) or "Brak zaplanowanych punktów"
        val_today = float(my_open_tickets["Kwota_Przewidywana"].sum()) if pts_total > 0 else 0.0
        missing_to_next = metrics["missing_tier_1"] if metrics["monthly_revenue"] < COMMISSION_TIER_1 else metrics["missing_tier_2"]
        next_tier_lbl = "progu 3% (15 tys.)" if metrics["monthly_revenue"] < COMMISSION_TIER_1 else "progu 4% (30 tys.)"

        # 2. Siatka kafelków 2x2 (wzorowana na aplikacji InPost + Licznik Serii Combo 🔥)
        row1_c1, row1_c2 = st.columns(2)
        with row1_c1:
            st.markdown(f"""
            <div class="inpost-tile">
                <div class="inpost-tile-label">💰 TWOJA PREMIA ({metrics['active_rate_str']})</div>
                <div class="inpost-tile-val-green">{metrics['commission_pln']:,.0f} zł</div>
                <div class="inpost-tile-sub">Obrót w mies.: <b>{metrics['monthly_revenue']:,.0f} zł</b></div>
                <div style="font-size:12.5px; color:#64748b; margin-top:4px;">
                    Do {next_tier_lbl} brakuje: <b style="color:#0a2540;">{missing_to_next:,.0f} zł</b>
                </div>
            </div>
            """, unsafe_allow_html=True)
            st.progress(metrics["progress_pct"] / 100.0)

        with row1_c2:
            streak_txt = f"🔥 Seria: {gami['current_streak']} na strzała! (Rekord: {gami['best_streak']})"
            st.markdown(f"""
            <div class="inpost-tile">
                <div class="inpost-tile-label">🎯 SKUTECZNOŚĆ (1. PODEJŚCIE)</div>
                <div class="inpost-tile-val-blue">{metrics['ftfr_pct']}%</div>
                <div class="streak-pill">{streak_txt}</div>
                <div style="font-size:12.5px; color:#64748b; margin-top:6px;">
                    Zamknięte: <b style="color:#0a2540;">{metrics['total_closed']} szt.</b> | Leady: <b>{metrics['leads_count']}</b>
                </div>
            </div>
            """, unsafe_allow_html=True)

        row2_c1, row2_c2 = st.columns(2)
        with row2_c1:
            st.markdown(f"""
            <div class="inpost-tile">
                <div class="inpost-tile-label">🚚 PODSUMOWANIE TRASY</div>
                <div class="inpost-tile-val-dark">{pts_total} {'punkt' if pts_total == 1 else 'punkty' if 2 <= pts_total <= 4 else 'punktów'}</div>
                <div class="inpost-tile-sub" style="color:#0284c7;">📍 {cities_today}</div>
                <div style="font-size:13px; color:#16a34a; font-weight:800; margin-top:4px;">
                    Wartość zleceń: {val_today:,.0f} zł
                </div>
            </div>
            """, unsafe_allow_html=True)
            if st.button("🚀 Otwórz Plan Trasy ➔", key="tile_go_plan", type="primary"):
                st.session_state["tech_nav_tab"] = "plan"
                st.rerun()

        with row2_c2:
            low_stock = get_low_stock_items(my_prof)
            low_txt = f"⚠️ Braki w busie: {len(low_stock)} poz." if low_stock else "✅ Stan paki w normie"
            st.markdown(f"""
            <div class="inpost-tile">
                <div class="inpost-tile-label">📋 PROTOKOŁY & KASA</div>
                <div class="inpost-tile-val-amber">Gotówka: {metrics['cash_in_hand']:,.0f} zł</div>
                <div class="inpost-tile-sub">Do zdania w biurze</div>
                <div style="font-size:12.5px; color:#b45309; font-weight:700; margin-top:4px;">
                    {low_txt}
                </div>
            </div>
            """, unsafe_allow_html=True)
            if st.button("📋 Otwórz Protokół Przeglądu ➔", key="tile_go_przeglad"):
                st.session_state["tech_nav_tab"] = "przeglad"
                st.rerun()

        # 3. Sekcja Odznak i Osiągnięć (Grywalizacja)
        with st.expander(f"🏅 Twoje Odznaki i Osiągnięcia ({gami['unlocked_cnt']} z {len(gami['badges'])} odblokowanych) — Zbieraj Waty ⚡!", expanded=True):
            st.caption("Za każdą czysto zamkniętą wizytę dostajesz **+10 ⚡**, za protokół **+15 ⚡**, za Lead handlowy **+25 ⚡**, a za pilną awarię **+50 ⚡**!")
            b_cols = st.columns(3)
            for idx_b, b_item in enumerate(gami["badges"]):
                b_cls = "badge-card-unlocked" if b_item["unlocked"] else "badge-card-locked"
                status_tag = "✅ ODBLOKOWANA!" if b_item["unlocked"] else f"🔒 Postęp: {b_item['postep']}"
                tag_col = "#15803d" if b_item["unlocked"] else "#64748b"
                with b_cols[idx_b % 3]:
                    st.markdown(f"""
                    <div class="{b_cls}">
                        <div style="font-size:28px;">{b_item['ikona']}</div>
                        <div style="font-size:13.5px; font-weight:900; color:#0a2540; margin-top:2px;">{b_item['nazwa']}</div>
                        <div style="font-size:11.5px; color:#475569; margin-top:2px;">{b_item['opis']}</div>
                        <div style="font-size:11.5px; font-weight:800; color:{tag_col}; margin-top:5px;">{status_tag}</div>
                    </div>
                    """, unsafe_allow_html=True)

        # 4. Pasek wyszukiwarki pod kafelkami (w bazie wiedzy PDF i w internecie)
        st.markdown("#### 🔍 Wyszukiwarka w bazie wiedzy (PDF) i w internecie")
        with st.container(border=True):
            q_home = st.text_input(
                "Wpisz kod błędu lub objaw (np. F08, AL007, Deye wyłącza backup, naczynie przeponowe):",
                placeholder="Szukaj rozwiązania usterki...",
                key="inpost_home_search"
            )
            p_c1, p_c2, p_c3 = st.columns(3)
            with p_c1:
                show_pdf_codes = st.checkbox("📚 Kody z Bazy", value=False, key="p_kb_1")
            with p_c2:
                show_wifi_help = st.checkbox("📶 Ściąga Wi-Fi Deye", value=False, key="p_kb_2")
            with p_c3:
                if q_home.strip():
                    _, l_g_h, _, _ = przeszukaj_praktyke_i_internet(q_home)
                    st.link_button("🌐 Szukaj w Google", l_g_h)

            if q_home.strip():
                wyn_pdf = przeszukaj_lokalne_instrukcje(q_home)
                if wyn_pdf:
                    for w in wyn_pdf[:3]:
                        st.markdown(f"**📘 {w['zrodlo']} — {w['naglowek']}**")
                        st.info(w["fragment"])
                porady_h, l_goog, l_elek, l_yt = przeszukaj_praktyke_i_internet(q_home)
                if porady_h:
                    for p in porady_h:
                        st.success(f"💡 **{p['tytul']}**\n\n{p['opis']}")
                l_c1, l_c2, l_c3 = st.columns(3)
                with l_c1:
                    st.link_button("🌐 Wyniki Google", l_goog)
                with l_c2:
                    st.link_button("🛠️ Forum Elektroda", l_elek)
                with l_c3:
                    st.link_button("▶️ Wideo YouTube", l_yt)

            if show_pdf_codes:
                kb_cat = st.selectbox("Wybierz urządzenie:", list(DIAGNOSTYKA_BAZA.keys()))
                kb_err = st.selectbox("Kod błędu:", list(DIAGNOSTYKA_BAZA[kb_cat].keys()))
                st.success(DIAGNOSTYKA_BAZA[kb_cat][kb_err])

            if show_wifi_help:
                st.markdown(INSTRUKCJE_KONFIGURACJI_TEXT)

    # ----------------------------------------------------
    # ZAKŁADKA 2: PLAN (ZLECENIA NA DZIŚ + ZAŁADUNEK + 3-KARTIOWY KREATOR ZE STRZAŁKAMI)
    # ----------------------------------------------------
    elif active_tab == "plan":
        if st.session_state.get("active_ticket_id") is None:
            st.markdown(f"""
            <div class="promag-header">
                <span style="font-size:20px; font-weight:800;">🗺️ Twój Plan Trasy i Lista Załadunku</span><br>
                <span style="font-size:13px; color:#86efac;">Zleceń do realizacji: {len(my_open_tickets)} | 🔥 Twoja seria na strzała: {gami['current_streak']} | ⚡ {gami['total_waty']} Watów</span>
            </div>
            """, unsafe_allow_html=True)

            if my_open_tickets.empty:
                st.success("🌿 Wszystkie zaplanowane zlecenia zostały wykonane! Odpocznij lub sprawdź zakładkę Przegląd.")
            else:
                days_avail = sorted(my_open_tickets["Data_Grupa"].unique())
                sel_day_plan = st.selectbox("📅 Dzień trasy:", days_avail, index=0)
                day_tickets = my_open_tickets[my_open_tickets["Data_Grupa"] == sel_day_plan].copy()

                opisy_dnia = " ".join(day_tickets["Opis_Zgloszenia"].astype(str).tolist()).lower()
                potrzebne_mat = []
                for m in my_prof.get("materialy", []):
                    klucz = str(m.get("slowo_klucz", "")).lower()
                    if klucz and klucz in opisy_dnia:
                        potrzebne_mat.append(f"✅ **{m['nazwa']}** (masz w busie: `{m['ilosc']} szt.`)")

                with st.container(border=True):
                    st.markdown("#### 📦 WEŹ NA AUTO PRZED WYJAZDEM (Lista potrzebnych materiałów na dziś):")
                    if potrzebne_mat:
                        for pm_str in potrzebne_mat:
                            st.markdown(pm_str)
                    else:
                        st.caption("Standardowy zestaw serwisowy (glikol, anoda 3/4, czujniki NTC, miernik Riso, refraktometr).")

                for idx_pt, (t_idx, t_row) in enumerate(day_tickets.iterrows(), 1):
                    t_id = t_row["ID_Zgloszenia"]
                    is_urg = (str(t_row.get("Priorytet_Krytyczny", "Nie")) == "Tak")
                    box_cls = "day-card-urgent" if is_urg else "day-card"
                    urg_badge = "🚨 PILNE (+50 ⚡) | " if is_urg else ""

                    with st.container(border=True):
                        st.markdown(f"""
                        <div class="{box_cls}">
                            <div style="display:flex; justify-content:space-between;">
                                <span style="font-size:13px; font-weight:800; color:{'#b91c1c' if is_urg else '#0284c7'};">{urg_badge}Punkt #{idx_pt} z {len(day_tickets)} ({t_id})</span>
                                <span style="font-size:13px; font-weight:700; color:#16a34a;">Wycena: {float(t_row['Kwota_Przewidywana']):,.0f} zł</span>
                            </div>
                            <div style="font-size:20px; font-weight:900; color:#0a2540; margin-top:4px;">{t_row['Klient']}</div>
                            <div style="font-size:14.5px; color:#334155; margin-top:2px;">📍 <b>{t_row['Kod_Pocztowy']} {t_row['Miasto']}, {t_row['Adres']}</b></div>
                            <div style="font-size:14px; color:#0284c7; font-weight:700; margin-top:4px;">🔧 {t_row['Typ_Instalacji']} — {t_row['Opis_Zgloszenia']}</div>
                        </div>
                        """, unsafe_allow_html=True)

                        if str(t_row.get("Trudny_Klient", "Nie")) == "Tak":
                            st.warning("🚩 **UWAGA: Wymagana płatność z góry / trudny klient!**")

                        b_c1, b_c2, b_c3 = st.columns([1, 1, 1.3])
                        phones = extract_phone_numbers(str(t_row["Telefon"]), str(t_row["Klient"]))
                        with b_c1:
                            if phones:
                                st.link_button(f"📞 Zadzwoń ({phones[0]})", f"tel:+48{phones[0]}")
                        with b_c2:
                            maps_q = f"{t_row['Adres']} {t_row['Kod_Pocztowy']} {t_row['Miasto']}".replace(" ", "+")
                            st.link_button("🗺️ Nawiguj (Mapa)", f"https://www.google.com/maps/search/?api=1&query={maps_q}")
                        with b_c3:
                            if st.button(f"🛠️ OTWÓRZ WIZYTĘ ({t_id}) ➔", key=f"start_v_{t_id}", type="primary"):
                                df.loc[t_idx, "Status"] = "W trakcie"
                                save_tickets(df)
                                st.session_state["active_ticket_id"] = t_id
                                st.session_state["visit_step"] = 1
                                st.rerun()

        else:
            # WIDOK AKTYWNEGO ZLECENIA: 3 KARTY PRZEŁĄCZANE DUŻYMI PRZYCISKAMI ZE STRZAŁKĄ
            selected_id = st.session_state["active_ticket_id"]
            matching_rows = df[df["ID_Zgloszenia"] == selected_id]
            if matching_rows.empty:
                st.session_state["active_ticket_id"] = None
                st.rerun()
            t_idx = matching_rows.index[0]
            t_row = df.loc[t_idx]
            curr_step = int(st.session_state.get("visit_step", 1))

            top_v1, top_v2 = st.columns([2.4, 1])
            with top_v1:
                st.markdown(f"""
                <div class="promag-header" style="margin-bottom:8px;">
                    <span style="font-size:19px; font-weight:900;">🛠️ {t_row['Klient']} ({selected_id})</span><br>
                    <span style="font-size:13.5px; color:#86efac;">Karta {curr_step} z 3: {'1. Diagnoza i dane instalacji' if curr_step == 1 else '2. Zużyte części z busa i zdjęcia' if curr_step == 2 else '3. Rozliczenie finałowe i podpis'}</span>
                </div>
                """, unsafe_allow_html=True)
            with top_v2:
                if st.button("⬅️ Wróć do listy zleceń"):
                    st.session_state["active_ticket_id"] = None
                    st.rerun()

            s1_col = "#16a34a" if curr_step >= 1 else "#cbd5e1"
            s2_col = "#16a34a" if curr_step >= 2 else "#cbd5e1"
            s3_col = "#16a34a" if curr_step >= 3 else "#cbd5e1"
            st.markdown(f"""
            <div class="step-progress-bar">
                <div style="text-align:center; flex:1;">
                    <span style="background:{s1_col}; color:white; padding:6px 13px; border-radius:50%; font-weight:900;">1</span>
                    <div style="font-size:12px; font-weight:800; color:#0a2540; margin-top:4px;">1. Diagnoza</div>
                </div>
                <div style="flex:1; height:4px; background:{s2_col};"></div>
                <div style="text-align:center; flex:1;">
                    <span style="background:{s2_col}; color:white; padding:6px 13px; border-radius:50%; font-weight:900;">2</span>
                    <div style="font-size:12px; font-weight:800; color:#0a2540; margin-top:4px;">2. Części i Foto</div>
                </div>
                <div style="flex:1; height:4px; background:{s3_col};"></div>
                <div style="text-align:center; flex:1;">
                    <span style="background:{s3_col}; color:white; padding:6px 13px; border-radius:50%; font-weight:900;">3</span>
                    <div style="font-size:12px; font-weight:800; color:#0a2540; margin-top:4px;">3. Finał i Podpis</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            usage_key = f"used_qty_{selected_id}"
            if usage_key not in st.session_state:
                st.session_state[usage_key] = {i: 0 for i in range(len(my_prof.get("materialy", [])))}

            # --- KARTA 1: DIAGNOZA I DANE INSTALACJI ---
            if curr_step == 1:
                with st.container(border=True):
                    st.markdown(f"### 📍 Adres: `{t_row['Kod_Pocztowy']} {t_row['Miasto']}, {t_row['Adres']}`")
                    st.markdown(f"📞 **Telefon:** `{t_row['Telefon']}` | 🔧 **Urządzenie:** `{t_row['Typ_Instalacji']}`")
                    if str(t_row.get("Opis_Instalacji_Montaz", "")).strip():
                        st.markdown(f"""
                        <div class="install-badge">
                            📋 <b>Zamontowana instalacja u klienta (Montaż: {t_row.get('Data_Montazu', 'Brak')}):</b><br>
                            {t_row['Opis_Instalacji_Montaz']}
                        </div>
                        """, unsafe_allow_html=True)
                    st.info(f"📝 **Zadanie / Opis usterki od Koordynatora:**\n\n{t_row['Opis_Zgloszenia']}")
                    if float(t_row.get("Zgoda_Klienta_Kwota", 0.0)) > 0:
                        st.success(f"💰 **Klient zaakceptował koszt do: {float(t_row['Zgoda_Klienta_Kwota']):,.0f} zł**")

                    st.divider()
                    st.markdown("#### 🔍 Szybka pomoc diagnostyczna (PDF + Google)")
                    q_in = st.text_input("Wpisz kod błędu lub objaw na miejscu:", key=f"diag_q_{selected_id}")
                    if q_in.strip():
                        for rl in przeszukaj_lokalne_instrukcje(q_in)[:2]:
                            st.warning(f"**{rl['zrodlo']} ({rl['naglowek']}):**\n{rl['fragment']}")
                        por_i, lg_i, _, _ = przeszukaj_praktyke_i_internet(q_in)
                        for pi in por_i:
                            st.success(f"💡 **{pi['tytul']}:** {pi['opis']}")
                        st.link_button("🌐 Szukaj w Google", lg_i)

                nav1, nav2 = st.columns([1, 1.4])
                with nav1:
                    if st.button("📋 Otwórz Protokół Przeglądu", key=f"to_prot_{selected_id}"):
                        st.session_state["przeglad_ticket_id"] = selected_id
                        st.session_state["tech_nav_tab"] = "przeglad"
                        st.rerun()
                with nav2:
                    if st.button("DALEJ: CZĘŚCI I ZDJĘCIA ➡", key=f"next_to_2_{selected_id}", type="primary"):
                        st.session_state["visit_step"] = 2
                        st.rerun()

            # --- KARTA 2: ZUŻYTE CZĘŚCI Z BUSA (PRZYCISKI - / +) ORAZ ZDJĘCIA ---
            elif curr_step == 2:
                with st.container(border=True):
                    st.markdown("#### 🚐 Zużyte materiały z Twojego busa (dotknij `-` lub `+`):")
                    van_mats = my_prof.get("materialy", [])
                    for idx_m, m_item in enumerate(van_mats):
                        cur_u = st.session_state[usage_key].get(idx_m, 0)
                        mc1, mc2, mc3, mc4 = st.columns([2.4, 0.7, 0.6, 0.7])
                        with mc1:
                            st.markdown(f"**{m_item['nazwa']}**  \n<span style='font-size:12.5px; color:#64748b;'>W busie: {m_item['ilosc']} szt. | Cena: {m_item['cena_klienta']:.0f} zł</span>", unsafe_allow_html=True)
                        with mc2:
                            if st.button("➖", key=f"dec_m_{selected_id}_{idx_m}"):
                                st.session_state[usage_key][idx_m] = max(0, cur_u - 1)
                                st.rerun()
                        with mc3:
                            st.markdown(f"<div style='text-align:center; font-size:20px; font-weight:900; padding-top:8px; color:{'#16a34a' if cur_u > 0 else '#0a2540'};'>{cur_u}</div>", unsafe_allow_html=True)
                        with mc4:
                            if st.button("➕", key=f"inc_m_{selected_id}_{idx_m}"):
                                st.session_state[usage_key][idx_m] = min(int(m_item.get("ilosc", 20)), cur_u + 1)
                                st.rerun()

                    st.divider()
                    st.markdown("#### 📸 Zdjęcie tabliczki znamionowej i instalacji")
                    up_photos = st.file_uploader(
                        "Dotknij, aby zrobić zdjęcie aparatem lub wybrać z galerii:",
                        type=["png", "jpg", "jpeg"],
                        accept_multiple_files=True,
                        key=f"up_ph_{selected_id}"
                    )
                    if up_photos:
                        st.session_state[f"photos_cnt_{selected_id}"] = len(up_photos)
                    saved_ph_cnt = st.session_state.get(f"photos_cnt_{selected_id}", 0)
                    chk_ph = st.checkbox(
                        f"✅ Potwierdzam wykonanie zdjęć tabliczki i instalacji ({saved_ph_cnt} załączono)",
                        value=bool(saved_ph_cnt > 0 or st.session_state.get(f"chk_ph_val_{selected_id}", False)),
                        key=f"chk_ph_box_{selected_id}"
                    )
                    st.session_state[f"chk_ph_val_{selected_id}"] = chk_ph

                nav_b1, nav_b2 = st.columns([1, 1.3])
                with nav_b1:
                    if st.button("⬅ WSTECZ (DIAGNOZA)", key=f"back_to_1_{selected_id}"):
                        st.session_state["visit_step"] = 1
                        st.rerun()
                with nav_b2:
                    if st.button("DALEJ (FINAŁ I PODPIS) ➡", key=f"next_to_3_{selected_id}", type="primary"):
                        st.session_state["visit_step"] = 3
                        st.rerun()

            # --- KARTA 3: ROZLICZENIE (4 OPCJE WYNIKU WIZYTY + PODPIS) ---
            elif curr_step == 3:
                van_mats = my_prof.get("materialy", [])
                used_materials_price = 0.0
                used_materials_cost = 0.0
                used_items_desc = []
                for idx_m, m_item in enumerate(van_mats):
                    q_u = int(st.session_state[usage_key].get(idx_m, 0))
                    if q_u > 0:
                        used_materials_price += q_u * float(m_item.get("cena_klienta", 0.0))
                        used_materials_cost += q_u * float(m_item.get("koszt_zakupu", 0.0))
                        used_items_desc.append(f"{m_item['nazwa']} x{q_u}")

                with st.container(border=True):
                    if used_items_desc:
                        st.info(f"🚐 **Wybrane części z busa ({used_materials_price:,.0f} zł):** {', '.join(used_items_desc)}")

                    cennik_slownik = pobierz_slownik_cennika(st.session_state["cennik_tab"], st.session_state["stawka_rbh"])
                    wybrane_uslugi_tech = st.multiselect("Dodatkowe czynności z cennika (opcjonalnie):", list(cennik_slownik.keys()), key=f"srv_t_{selected_id}")
                    extra_srv_sum = sum(cennik_slownik[u]["cena_calkowita"] for u in wybrane_uslugi_tech)

                    rc1, rc2 = st.columns(2)
                    with rc1:
                        wynik_wizyty = st.radio(
                            "🏁 Wynik wizyty:",
                            [
                                "✅ Tak, udało się (Naprawione / Zakończone)",
                                "❌ Nie udało się naprawić (Podaj powód)",
                                "📦 Wymagane zamówienie części",
                                "💼 Potrzebna oferta handlowa / modernizacja"
                            ]
                        )
                    with rc2:
                        forma_plat = st.selectbox("💳 Forma płatności:", ["Gotówka pobrana", "Przelew (Faktura)", "BLIK / Terminal", "Gwarancja / Brak opłaty (0 zł)"])
                        typ_serwisu_real = st.selectbox("Kwalifikacja:", ["Płatny", "Gwarancyjny"], index=0 if t_row["Typ_Serwisu"] != "Gwarancyjny" else 1)
                        work_minutes = st.number_input("Czas pracy na miejscu (min):", 15, 600, 90, 15)

                    powod_niewykonania_txt = ""
                    potrzebne_czesci_txt = ""
                    lead_opis = ""

                    if wynik_wizyty.startswith("❌"):
                        powod_niewykonania_txt = st.text_area(
                            "⚠️ Dlaczego nie udało się wykonać zlecenia? (obowiązkowe):",
                            value=str(t_row.get("Powod_Niewykonania", "")),
                            placeholder="np. uszkodzona sprężarka / wymagana ekipa monterska / brak dostępu / odmowa klienta..."
                        )
                    elif wynik_wizyty.startswith("📦"):
                        potrzebne_czesci_txt = st.text_input("📦 Wpisz dokładny model części do zamówienia (obowiązkowe):", value=str(t_row.get("Potrzebne_Czesci", "")))
                    elif wynik_wizyty.startswith("💼"):
                        lead_opis = st.text_input("💼 Czego dotyczy potrzebna oferta? (obowiązkowe):", value=str(t_row.get("Lead_Handlowy", "")))

                    czy_dod_lead = st.checkbox("🔥 Klient pytał też o ofertę handlową (+25 ⚡ za Lead!)", value=bool(wynik_wizyty.startswith("💼")))
                    if czy_dod_lead and not wynik_wizyty.startswith("💼"):
                        lead_opis = st.text_input("Opis leada dla Handlowca (np. magazyn energii, wymiana PC):")

                    base_expected = float(t_row["Kwota_Przewidywana"])
                    suggested_total = 0.0 if (typ_serwisu_real == "Gwarancyjny" or forma_plat.startswith("Gwarancja") or wynik_wizyty.startswith("❌")) else (base_expected + extra_srv_sum + used_materials_price)
                    final_price = st.number_input("💰 Kwota końcowa do zapłaty przez klienta (zł):", 0.0, 50000.0, float(suggested_total), 50.0)
                    internal_note = st.text_area("Notatka serwisanta po wizycie:", value=str(t_row["Notatka_Serwisanta"]), height=70)

                    st.markdown('<div class="client-box">', unsafe_allow_html=True)
                    st.markdown(f"#### ✍️ Podpis Klienta — Do zapłaty: **{final_price:,.2f} zł** ({forma_plat})")
                    client_recommendations = st.text_input("Zalecenia dla klienta:", value=str(t_row["Zalecenia_Klienta"]))
                    st_canvas(
                        fill_color="rgba(255, 255, 255, 0)", stroke_width=3, stroke_color="#0a2540",
                        background_color="#ffffff", height=150, width=550, drawing_mode="freedraw", key=f"canv_{selected_id}"
                    )
                    st.markdown('</div>', unsafe_allow_html=True)

                nav_f1, nav_f2 = st.columns([1, 1.5])
                with nav_f1:
                    if st.button("⬅ WSTECZ (CZĘŚCI / FOTO)", key=f"back_to_2_{selected_id}"):
                        st.session_state["visit_step"] = 2
                        st.rerun()
                with nav_f2:
                    if st.button("✅ ZATWIERDŹ I ZAKOŃCZ WIZYTĘ", key=f"finish_v_{selected_id}", type="primary"):
                        if not st.session_state.get(f"chk_ph_val_{selected_id}", False):
                            st.error("Wróć o krok (⬅ WSTECZ) i potwierdź wykonanie zdjęcia tabliczki znamionowej!")
                        elif wynik_wizyty.startswith("❌") and not powod_niewykonania_txt.strip():
                            st.error("Podaj powód, dlaczego nie udało się wykonać zlecenia!")
                        elif wynik_wizyty.startswith("📦") and not potrzebne_czesci_txt.strip():
                            st.error("Wpisz dokładny model części do zamówienia!")
                        elif wynik_wizyty.startswith("💼") and not lead_opis.strip():
                            st.error("Wpisz, czego dotyczy oferta handlowa!")
                        else:
                            for idx_m, m_item in enumerate(van_mats):
                                q_u = int(st.session_state[usage_key].get(idx_m, 0))
                                if q_u > 0:
                                    my_prof["materialy"][idx_m]["ilosc"] = max(0, int(my_prof["materialy"][idx_m]["ilosc"]) - q_u)
                            save_config_and_profiles(profiles)

                            if wynik_wizyty.startswith("✅"):
                                new_stat, ftfr_val = "Zakończone", "Tak"
                            elif wynik_wizyty.startswith("❌"):
                                new_stat, ftfr_val = "Niewykonane", "Nie"
                            elif wynik_wizyty.startswith("📦"):
                                new_stat, ftfr_val = "Do zamówienia części", "Nie"
                            else:
                                new_stat, ftfr_val = "Oddelegowane", "Nie"

                            df.loc[t_idx, "Status"] = new_stat
                            df.loc[t_idx, "Naprawione_1_Wizyta"] = ftfr_val
                            df.loc[t_idx, "Powod_Niewykonania"] = powod_niewykonania_txt.strip()
                            df.loc[t_idx, "Potrzebne_Czesci"] = potrzebne_czesci_txt.strip()
                            df.loc[t_idx, "Typ_Serwisu"] = typ_serwisu_real
                            df.loc[t_idx, "Forma_Platnosci"] = forma_plat
                            df.loc[t_idx, "Gotowka_Rozliczona"] = "Nie" if (forma_plat == "Gotówka pobrana" and final_price > 0) else "Tak"
                            df.loc[t_idx, "Kwota_Koncowa"] = float(final_price)
                            df.loc[t_idx, "Koszt_Materialow"] = float(used_materials_cost)
                            df.loc[t_idx, "Czas_Pracy_Min"] = float(work_minutes)
                            prev_wyn = str(df.loc[t_idx, "Wyniki_Przegladu"]).strip()
                            cz_str = "; ".join(used_items_desc + wybrane_uslugi_tech)
                            df.loc[t_idx, "Wyniki_Przegladu"] = (prev_wyn + " | " + cz_str).strip(" |")
                            df.loc[t_idx, "Zdjecia_Potwierdzone"] = "Tak"
                            df.loc[t_idx, "Zalecenia_Klienta"] = str(client_recommendations)
                            note_c = str(internal_note).strip()
                            if powod_niewykonania_txt.strip():
                                note_c = f"[POWÓD NIEWYKONANIA: {powod_niewykonania_txt.strip()}] " + note_c
                            df.loc[t_idx, "Notatka_Serwisanta"] = note_c
                            if lead_opis.strip():
                                df.loc[t_idx, "Lead_Handlowy"] = lead_opis.strip()
                            df.loc[t_idx, "Data_Zamkniecia"] = today_str
                            df.loc[t_idx, "Podpis_Zlozony"] = "Tak"
                            save_tickets(df)

                            earned_waty = (10 if new_stat == "Zakończone" else 0) + (25 if lead_opis.strip() else 0) + (50 if str(t_row.get("Priorytet_Krytyczny")) == "Tak" else 0)
                            premia_visit = round(float(final_price) * 0.04, 2)
                            st.session_state["victory_banner"] = (
                                f"🔥 Bam! Zlecenie {selected_id} zamknięte! Doliczono +{final_price:,.0f} zł do obrotu "
                                f"(do premii wpada ok. +{premia_visit:,.1f} zł) oraz +{earned_waty} ⚡ Watów!"
                            )
                            st.session_state["active_ticket_id"] = None
                            st.session_state["visit_step"] = 1
                            st.rerun()

    # ----------------------------------------------------
    # ZAKŁADKA 3: PRZEGLĄD (INTERAKTYWNE PROTOKOŁY PRZEGLĄDÓW DLA URZĄDZEŃ)
    # ----------------------------------------------------
    elif active_tab == "przeglad":
        st.markdown("""
        <div class="promag-header" style="background: linear-gradient(135deg, #0284c7 0%, #0369a1 100%);">
            <span style="font-size:20px; font-weight:900;">📋 Interaktywny Protokół Przeglądu OZE (+15 ⚡)</span><br>
            <span style="font-size:13px; color:#e0f2fe;">Oddzielny protokół przeglądowy (dla zlecenia z listy lub wpisany z ręki po serwisie)</span>
        </div>
        """, unsafe_allow_html=True)

        with st.container(border=True):
            my_all_tickets = df[df["Serwisant"] == tech_name].copy()
            t_options = ["✍️ Wpisz dane klienta z ręki (Nowy protokół dodatkowy)"]
            t_map = {}
            default_opt_idx = 0
            pre_tid = st.session_state.get("przeglad_ticket_id", None)

            for _, r_t in my_all_tickets.iterrows():
                lbl_t = f"{r_t['ID_Zgloszenia']} | {r_t['Klient']} ({r_t['Miasto']}, {r_t['Adres']}) - {r_t['Typ_Instalacji']}"
                t_options.append(lbl_t)
                t_map[lbl_t] = r_t
                if pre_tid and r_t["ID_Zgloszenia"] == pre_tid:
                    default_opt_idx = len(t_options) - 1

            chosen_prot_target = st.selectbox("👤 Wybierz zlecenie klienta lub wpisz z ręki:", t_options, index=default_opt_idx)

            if chosen_prot_target in t_map:
                sel_r = t_map[chosen_prot_target]
                p_klient = sel_r["Klient"]
                p_adres = f"{sel_r['Kod_Pocztowy']} {sel_r['Miasto']}, {sel_r['Adres']}"
                p_tel = sel_r["Telefon"]
                p_tid = sel_r["ID_Zgloszenia"]
                def_dev = sel_r["Typ_Instalacji"]
            else:
                p_tid = None
                c_p1, c_p2, c_p3 = st.columns(3)
                with c_p1:
                    p_klient = st.text_input("Imię i nazwisko klienta:")
                with c_p2:
                    p_tel = st.text_input("Telefon klienta:")
                with c_p3:
                    p_adres = st.text_input("Adres instalacji (Miasto, ulica):")
                def_dev = "Pompa Ciepła ARGOM"

            proc_keys = list(PROCEDURY_PRZEGLADOW.keys())
            d_idx = 0
            for i_pk, pk in enumerate(proc_keys):
                if pk.lower() in str(def_dev).lower():
                    d_idx = i_pk
                    break

            wybrany_typ_urzadzenia = st.selectbox("🔧 Wybierz badane urządzenie OZE:", proc_keys, index=d_idx)
            proc_info = get_procedure_for_type(wybrany_typ_urzadzenia)

            st.divider()
            st.markdown(f"### {proc_info['nazwa']}")
            st.caption("Dotknij i odhacz punkty kontrolne wykonane podczas przeglądu:")

            checked_list = []
            for idx_c, ch_item in enumerate(proc_info["checklist"]):
                if st.checkbox(ch_item, value=True, key=f"prot_chk_{idx_c}"):
                    checked_list.append(ch_item)

            st.markdown("#### 📊 Wyniki pomiarów kontrolnych:")
            pm_cols = st.columns(2)
            measured_list = []
            for idx_m, (m_lbl, m_def) in enumerate(proc_info["pola_pomiarowe"]):
                with pm_cols[idx_m % 2]:
                    val_m = st.text_input(m_lbl, value=m_def, key=f"prot_meas_{idx_m}")
                    if val_m.strip():
                        measured_list.append(f"{m_lbl}: {val_m.strip()}")

            prot_uwagi = st.text_input("Zalecenia po przeglądzie / uwagi do protokołu:", value="Instalacja sprawna, dopuszczona do dalszej eksploatacji.")

            st.markdown("#### ✍️ Podpis klienta pod protokołem przeglądu:")
            st_canvas(
                fill_color="rgba(255, 255, 255, 0)", stroke_width=3, stroke_color="#0a2540",
                background_color="#ffffff", height=135, width=550, drawing_mode="freedraw", key="canv_przeglad_tab"
            )

            if st.button("✅ ZAPISZ PROTOKÓŁ PRZEGLĄDU W BAZIE (+15 ⚡)", type="primary"):
                podsumowanie_prot = (
                    f"[PROTOKÓŁ {today_str} - {wybrany_typ_urzadzenia}] "
                    f"Pomiary: {', '.join(measured_list)} | "
                    f"Skontrolowano: {len(checked_list)}/{len(proc_info['checklist'])} pkt | "
                    f"Uwagi: {prot_uwagi}"
                )
                if p_tid:
                    idx_t = df[df["ID_Zgloszenia"] == p_tid].index[0]
                    df.loc[idx_t, "Wyniki_Przegladu"] = podsumowanie_prot
                    df.loc[idx_t, "Wymagany_Przeglad"] = "Tak"
                    save_tickets(df)
                    st.balloons()
                    st.success(f"Zapisano protokół przeglądu dla zlecenia {p_tid} ({p_klient})! Zdobywasz +15 ⚡ Watów!")
                else:
                    if not p_klient.strip():
                        st.error("Podaj imię i nazwisko klienta!")
                    else:
                        new_pid = f"ZGL-{len(df) + 1001}"
                        new_prow = {
                            "ID_Zgloszenia": new_pid, "Klient": p_klient.strip(), "Telefon": p_tel.strip(),
                            "Kod_Pocztowy": extract_postal_code(p_adres), "Miasto": p_adres.strip(), "Adres": p_adres.strip(),
                            "Typ_Instalacji": wybrany_typ_urzadzenia, "Opis_Zgloszenia": f"Przegląd okresowy: {wybrany_typ_urzadzenia}",
                            "Status": "Zakończone", "Serwisant": tech_name, "Data_Planowana": today_str,
                            "Kwota_Przewidywana": 350.0, "Zgoda_Klienta_Kwota": 350.0, "Kwota_Koncowa": 350.0,
                            "Typ_Serwisu": "Płatny", "Wymagany_Przeglad": "Tak", "Naprawione_1_Wizyta": "Tak",
                            "Wyniki_Przegladu": podsumowanie_prot, "Zalecenia_Klienta": prot_uwagi,
                            "Data_Zamkniecia": today_str, "Podpis_Zlozony": "Tak", "Zdjecia_Potwierdzone": "Tak"
                        }
                        df = pd.concat([df, pd.DataFrame([new_prow])], ignore_index=True)
                        save_tickets(df)
                        st.balloons()
                        st.success(f"Utworzono i zapisano nowy protokół przeglądu ({new_pid}) dla klienta {p_klient}! Zdobywasz +25 ⚡ Watów!")

    # ----------------------------------------------------
    # ZAKŁADKA 4: BUS (2 LISTY: MATERIAŁY NA PACE I NARZĘDZIA)
    # ----------------------------------------------------
    elif active_tab == "bus":
        flota = my_prof.get("flota", {})
        st.markdown(f"""
        <div class="promag-header">
            <span style="font-size:20px; font-weight:900;">🚐 Twój Bus Serwisowy ({flota.get('nr_rej', 'WPR 00000')})</span><br>
            <span style="font-size:13px; color:#86efac;">Aktualny przebieg: {int(flota.get('przebieg_km', 120000)):,} km | Stan części i narzędzi na pace</span>
        </div>
        """, unsafe_allow_html=True)

        with st.container(border=True):
            st.subheader("📦 1. Materiały eksploatacyjne na pace")
            st.caption("Pozycje w kolorze czerwonym mają stan <= 1 szt. – uzupełnij je w magazynie przed wyjazdem!")
            for idx_b, m_b in enumerate(my_prof.get("materialy", [])):
                qty_b = int(m_b.get("ilosc", 0))
                is_low = (qty_b <= 1)
                bg_b = "#fef2f2" if is_low else "#f8fafc"
                brd_b = "1.5px solid #ef4444" if is_low else "1px solid #cbd5e1"
                col_b = "#b91c1c" if is_low else "#0a2540"
                alert_b = "⚠️ UZUPEŁNIJ W BAZIE! " if is_low else ""

                bc1, bc2, bc3 = st.columns([2.6, 0.7, 0.7])
                with bc1:
                    st.markdown(
                        f"""<div style="background:{bg_b}; border:{brd_b}; border-radius:10px; padding:10px 14px; margin-bottom:4px;">
                            <span style="font-size:15px; font-weight:800; color:{col_b};">{alert_b}{m_b['nazwa']}</span><br>
                            <span style="font-size:13px; color:#475569;">Stan na aucie: <b>{qty_b} szt.</b> | Cena dla klienta: {m_b['cena_klienta']:.0f} zł</span>
                        </div>""",
                        unsafe_allow_html=True
                    )
                with bc2:
                    if st.button("➖ 1 szt.", key=f"bus_dec_{idx_b}"):
                        my_prof["materialy"][idx_b]["ilosc"] = max(0, qty_b - 1)
                        save_config_and_profiles(profiles)
                        st.rerun()
                with bc3:
                    if st.button("➕ 1 szt.", key=f"bus_inc_{idx_b}"):
                        my_prof["materialy"][idx_b]["ilosc"] = qty_b + 1
                        save_config_and_profiles(profiles)
                        st.rerun()

        with st.container(border=True):
            st.subheader("🧰 2. Narzędzia przypisane do Twojego busa")
            for idx_t, t_item in enumerate(my_prof.get("narzedzia", [])):
                st.markdown(
                    f"""<div style="background:#f8fafc; border:1px solid #cbd5e1; border-left:5px solid #16a34a; border-radius:10px; padding:10px 14px; margin-bottom:6px; display:flex; justify-content:space-between;">
                        <span style="font-size:14.5px; font-weight:700; color:#0a2540;">🔧 {t_item['nazwa']}</span>
                        <span style="font-size:14px; font-weight:800; color:#15803d;">{t_item.get('ilosc', 1)} szt. [OK]</span>
                    </div>""",
                    unsafe_allow_html=True
                )

            st.divider()
            odo_c1, odo_c2 = st.columns([1.8, 1])
            with odo_c1:
                new_odo = st.number_input("Aktualny stan licznika kilometrów [km]:", min_value=0, step=100, value=int(flota.get("przebieg_km", 120000)))
            with odo_c2:
                st.write("")
                if st.button("💾 Zapisz przebieg busa", type="primary"):
                    my_prof["flota"]["przebieg_km"] = int(new_odo)
                    save_config_and_profiles(profiles)
                    st.success("Zaktualizowano przebieg!")
                    st.rerun()

    render_bottom_nav_bar(active_tab)

# ==========================================
# 7. PANEL KOORDYNATORA (TRYB PILNY, 1-KLIK DO HANDLOWCA/MONTAŻU, SMS-Y, NAPRAWIONY PLANER)
# ==========================================
elif st.session_state["logged_role"] == "koordynator":
    df = st.session_state["df_tickets"]
    df_montaze = st.session_state["df_montaze"]
    profiles = st.session_state["profiles"]
    email_cfg = st.session_state["email_cfg"]
    cennik_tab = st.session_state["cennik_tab"]
    stawka_rbh = st.session_state["stawka_rbh"]
    dojazd_kw = st.session_state["dojazd_kw"]

    overdue_delegated, sent_cnt = check_and_send_3day_reminders(df, email_cfg)
    if overdue_delegated:
        st.error(f"🚨 **PRZETERMINOWANE ODDELEGOWANIA (>3 DNI BEZ OFERTY/TERMINU):** `{', '.join(overdue_delegated)}`")

    failed_df = df[df["Status"] == "Niewykonane"]
    if not failed_df.empty:
        st.error(f"⚠️ **ZLECENIA NIEWYKONANE W TERENIE ({len(failed_df)} szt.) – Wymagana decyzja koordynatora!** Sprawdź powód w tabeli zgłoszeń poniżej.")

    for t_n, t_p in profiles.items():
        low_s = get_low_stock_items(t_p)
        if low_s:
            st.warning(f"🚐 **Uzupełnij auto serwisanta {t_n} (<=1 szt.):** {', '.join(low_s)}")

    now_time = datetime.datetime.now().time()
    today_iso = str(datetime.date.today())
    callbacks_df = df[df["Do_Oddzwonienia"] == "Tak"].copy()
    is_after_1530 = (now_time >= datetime.time(15, 30))

    if not callbacks_df.empty:
        header_cb = "🔔 KLIENCI DO ODDZWONIENIA (PO 15:30 - OBOWIĄZKOWA LISTA DO SKUTKU!)" if is_after_1530 else f"📞 Klienci do oddzwonienia ({len(callbacks_df)} poz. - główny blok o 15:30)"
        with st.expander(header_cb, expanded=(is_after_1530 or len(callbacks_df) > 0)):
            for cb_idx, cb_row in callbacks_df.iterrows():
                cb_id = cb_row["ID_Zgloszenia"]
                cb_phones = extract_phone_numbers(str(cb_row["Telefon"]), str(cb_row["Klient"]))
                cb_phone_main = cb_phones[0] if cb_phones else ""
                ost_wpis = cb_row["Notatka_Serwisanta"] or cb_row["Opis_Zgloszenia"]
                with st.container(border=True):
                    cc1, cc2, cc3 = st.columns([1.6, 1.2, 1.2])
                    with cc1:
                        st.markdown(f"🔴 **{cb_row['Klient']}** (`{cb_id}`) – 📞 `{cb_row['Telefon']}` – 📅 Data kontaktu: **{cb_row['Data_Oddzwonienia'] or today_iso}**")
                        st.markdown(f"📍 `{cb_row['Kod_Pocztowy']} {cb_row['Miasto']}, {cb_row['Adres']}`")
                        st.caption(f"💬 Ostatni wpis: {ost_wpis}")
                    with cc2:
                        if cb_phone_main:
                            st.link_button(f"📞 Zadzwoń: +48{cb_phone_main}", f"tel:+48{cb_phone_main}")
                            sms_m = "Dzien dobry, kontaktowalismy sie z Dzialu Serwisu PROMAG OZE. Prosimy o kontakt zwrotny."
                            st.link_button("📲 Wyślij SMS (Nie odbiera)", f"sms:+48{cb_phone_main}?body={urllib.parse.quote(sms_m)}")
                        nowa_notatka_cb = st.text_input("Ustalenia z rozmowy:", key=f"cb_n_{cb_id}")
                    with cc3:
                        if st.button("✅ Dodzwoniono (Zdejmij)", key=f"cb_ok_{cb_id}", type="primary"):
                            df.loc[cb_idx, "Do_Oddzwonienia"] = "Nie"
                            df.loc[cb_idx, "Potwierdzony_Tel"] = "Tak"
                            if nowa_notatka_cb.strip():
                                df.loc[cb_idx, "Notatka_Serwisanta"] = f"[{today_iso} Tel OK]: {nowa_notatka_cb.strip()} | " + str(df.loc[cb_idx, "Notatka_Serwisanta"])
                            save_tickets(df)
                            st.rerun()
                        if st.button("⏳ Nie odbiera (Zostaw do skutku)", key=f"cb_w_{cb_id}"):
                            stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
                            df.loc[cb_idx, "Notatka_Serwisanta"] = f"[Próba {stamp} - nie odbiera] " + str(df.loc[cb_idx, "Notatka_Serwisanta"])
                            save_tickets(df)
                            st.rerun()

    tab_klienci, tab_plan, tab_oddeleguj, tab_cennik_kadry, tab_ustawienia = st.tabs([
        "📋 1. Przyjęcie Zgłoszenia (1-Klik Handlowiec/Montaż/Pilne + SMS)",
        "🗺️ 2. Planer Korytarzowy + Generator SMS dla Trasy",
        "📤 3. Oddelegowania (Watchdog 3 dni), Części i Leady",
        "💰 4. Cennik (100 zł/h), Kadry, Narzędzia i Magazyn Busa",
        "⚙️ 5. Logo, Baza Montaży, Eksport Kontaktów Google i SMTP"
    ])

    with tab_klienci:
        if st.session_state.get("last_created_sms"):
            sms_info = st.session_state["last_created_sms"]
            st.markdown('<div class="sms-box">', unsafe_allow_html=True)
            st.markdown(f"#### 📲 Gotowy SMS z podsumowaniem dla klienta: **{sms_info['klient']}** (`{sms_info['id']}` — {sms_info['tryb_opis']})")
            st.code(sms_info["tekst"], language=None)
            s_c1, s_c2, s_c3 = st.columns([1.2, 1.2, 1])
            with s_c1:
                if sms_info.get("uri"):
                    st.link_button(f"📲 WYŚLIJ SMS NA TELEFON ({sms_info['tel']})", sms_info["uri"], type="primary")
                else:
                    st.warning("Brak numeru telefonu do wysyłki SMS.")
            with s_c2:
                if sms_info.get("email_status"):
                    st.info(sms_info["email_status"])
            with s_c3:
                if st.button("✖️ Zamknij powiadomienie SMS"):
                    st.session_state["last_created_sms"] = None
                    st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)

        st.subheader("➕ Szybkie Przyjęcie Zgłoszenia (Wyszukiwarka na żywo + Kalkulator + 1-Klik Przekazanie i SMS)")

        with st.container(border=True):
            st.markdown("#### Krok 1: Wyszukaj klienta na żywo (po Nazwisku, Numerze Telefonu lub Adresie)")

            @st.cache_data(show_spinner=False)
            def build_unified_search_options(montaze_df, tickets_df):
                opts = ["➕ Nowy klient (wpisz dane ręcznie poniżej)"]
                lookup = {}

                if tickets_df is not None and not tickets_df.empty:
                    for _, sr in tickets_df.iterrows():
                        kl = str(sr.get("Klient", "")).strip()
                        tel = str(sr.get("Telefon", "")).strip()
                        kod = str(sr.get("Kod_Pocztowy", "")).strip() or extract_postal_code(f"{sr.get('Adres', '')} {sr.get('Miasto', '')}")
                        adr = f"{kod} {sr.get('Miasto', '')}, {sr.get('Adres', '')}".strip(" ,")
                        if kl or tel or adr:
                            label = f"🛠️ {kl} | 📞 {tel} | 📍 {adr} [{sr.get('ID_Zgloszenia', '')}]"
                            if label not in lookup:
                                opts.append(label)
                                lookup[label] = {
                                    "Klient": kl, "Telefon": tel, "Kod_Pocztowy": kod,
                                    "Miasto": str(sr.get("Miasto", "")), "Adres": str(sr.get("Adres", "")),
                                    "Data_Montazu": str(sr.get("Data_Montazu", "")),
                                    "Opis_Instalacji": str(sr.get("Opis_Instalacji_Montaz", "") or sr.get("Typ_Instalacji", ""))
                                }

                if montaze_df is not None and not montaze_df.empty:
                    for _, mr in montaze_df.iterrows():
                        kl = str(mr.get("Klient", "")).strip()
                        tel = str(mr.get("Telefon", "")).strip()
                        kod = str(mr.get("Kod_Pocztowy", "")).strip() or extract_postal_code(f"{mr.get('Adres', '')} {mr.get('Miasto', '')}")
                        adr = f"{kod} {mr.get('Miasto', '')}, {mr.get('Adres', '')}".strip(" ,")
                        dm = str(mr.get("Data_Montazu", "")).strip()
                        label = f"🏠 {kl} | 📞 {tel} | 📍 {adr} | Montaż: {dm}"
                        if label not in lookup:
                            opts.append(label)
                            lookup[label] = {
                                "Klient": kl, "Telefon": tel, "Kod_Pocztowy": kod,
                                "Miasto": str(mr.get("Miasto", "")), "Adres": str(mr.get("Adres", "")),
                                "Data_Montazu": dm, "Opis_Instalacji": str(mr.get("Opis_Instalacji", ""))
                            }
                return opts, lookup

            all_options, options_lookup = build_unified_search_options(df_montaze, df)

            chosen_client_option = st.selectbox(
                "🔍 Wpisz w tym polu Nazwisko, Telefon lub Adres klienta (podpowiedzi filtrują się na żywo poniżej):",
                options=all_options,
                index=0,
                key="live_autocomplete_client"
            )

            selected_prefill = options_lookup.get(chosen_client_option, {
                "Klient": "", "Telefon": "", "Kod_Pocztowy": "", "Miasto": "",
                "Adres": "", "Data_Montazu": "", "Opis_Instalacji": ""
            })

            if selected_prefill["Klient"] or selected_prefill["Telefon"]:
                st.markdown(f"""
                <div class="install-badge">
                    👤 <b>Wybrany klient:</b> <b>{selected_prefill['Klient']}</b> | 📞 <b>Telefon:</b> <code>{selected_prefill['Telefon'] or 'Brak'}</code> | 📍 <b>Adres:</b> <code>{selected_prefill['Kod_Pocztowy']} {selected_prefill['Miasto']}, {selected_prefill['Adres']}</code><br>
                    📅 <b>Data zakończenia montażu (Miesiąc i Rok):</b> <code>{selected_prefill['Data_Montazu'] or 'Brak w bazie'}</code><br>
                    📋 <b>Opis instalacji z tabeli montaży:</b> {selected_prefill['Opis_Instalacji'] or 'Brak szczegółowego opisu w bazie montaży'}
                </div>
                """, unsafe_allow_html=True)

            st.divider()
            st.markdown("#### Krok 2: Wybierz czynności z cennika (Kalkulator na żywo)")

            cennik_slownik = pobierz_slownik_cennika(cennik_tab, stawka_rbh)
            opcje_cennika = list(cennik_slownik.keys())

            k2_c1, k2_c2 = st.columns([2.2, 1])
            with k2_c1:
                wybrane_czynnosci = st.multiselect(
                    "🛠️ Wybierz pozycje do wyceny:",
                    options=opcje_cennika,
                    key="live_services_multiselect"
                )
            with k2_c2:
                czy_doliczyc_dojazd = st.checkbox(f"🚗 Dolicz dojazd z cennika (+{dojazd_kw:.0f} zł)", value=True)
                uwagi_dodatkowe = st.text_input("Dodatkowe uwagi od klienta (dopiszą się za czynnościami):", placeholder="np. cieknie przy grupie pompowej...")

            koszyk = oblicz_koszyk_uslug(
                wybrane_czynnosci, tabela=cennik_tab, stawka_rbh=stawka_rbh,
                dolicz_dojazd=czy_doliczyc_dojazd, kwota_dojazdu=dojazd_kw
            )

            auto_opis_pelny = koszyk["auto_notatka"]
            if uwagi_dodatkowe.strip():
                auto_opis_pelny = (auto_opis_pelny + " | UWAGI KLIENTA: " + uwagi_dodatkowe.strip()).strip(" |")

            m_k1, m_k2, m_k3, m_k4 = st.columns(4)
            m_k1.metric("Materiały z cennika", f"{koszyk['suma_materialow']:,.0f} zł")
            m_k2.metric(f"Robocizna ({koszyk['suma_czasu_h']}h x {stawka_rbh:.0f} zł)", f"{koszyk['suma_robocizny']:,.0f} zł")
            m_k3.metric("Opłata za dojazd", f"{koszyk['dojazd']:,.0f} zł")
            m_k4.metric("💰 RAZEM DLA KLIENTA", f"{koszyk['razem_dla_klienta']:,.0f} zł")

            st.divider()
            st.markdown("#### Krok 3: Dane zgłoszenia, Tryb Pilny („Straż Pożarna”) i Jednoprzyciskowa Ścieżka Realizacji")

            f_c1, f_c2, f_c3 = st.columns(3)
            with f_c1:
                in_klient = st.text_input("Imię i nazwisko klienta:", value=selected_prefill["Klient"])
                in_tel = st.text_input("Telefon:", value=selected_prefill["Telefon"])
                in_kod = st.text_input("Kod pocztowy (XX-XXX):", value=selected_prefill["Kod_Pocztowy"])
            with f_c2:
                in_miasto = st.text_input("Miejscowość:", value=selected_prefill["Miasto"])
                in_adres = st.text_input("Ulica i nr domu:", value=selected_prefill["Adres"])
                in_data_montazu = st.text_input("Miesiąc i Rok montażu:", value=selected_prefill["Data_Montazu"])
            with f_c3:
                in_typ_inst = st.selectbox("Kategoria instalacji:", [
                    "Kolektory Słoneczne (Solar)", "Pompa Ciepła ARGOM", "Falownik Deye (PV / Hybryda)",
                    "Falownik GoodWe (PV / Hybryda)", "Magazyn Energii", "Pompa Ciepła + PV", "Inne"
                ])
                in_kryt = st.selectbox("🚨 Priorytet krytyczny (Pali się / Awaria)?", ["Nie", "Tak"])
                in_trudny = st.selectbox("🚩 Trudny klient / Przedpłata?", ["Nie", "Tak"])

            in_opis_instalacji = st.text_input("📋 Opis instalacji z bazy montaży:", value=selected_prefill["Opis_Instalacji"])
            in_notatka_konserwatora = st.text_area("📝 Notatka / Zakres prac / Prośba dla Handlowca lub Montażu:", value=auto_opis_pelny, height=80)

            b_k1, b_k2, b_k3, b_k4 = st.columns(4)
            with b_k1:
                in_kwota = st.number_input("Wycena zlecenia (zł):", 0.0, 50000.0, float(koszyk["razem_dla_klienta"]), 50.0)
            with b_k2:
                in_zgoda = st.number_input("✅ Zgoda klienta na kwotę (zł):", 0.0, 50000.0, float(koszyk["razem_dla_klienta"]), 50.0)
            with b_k3:
                in_oddzwonic = st.checkbox("📞 Do oddzwonienia (blok 15:30)?", value=False)
            with b_k4:
                in_data_oddzw = st.date_input("Data oddzwonienia:", value=datetime.date.today())

            st.markdown("##### ⚡ Wybierz docelową ścieżkę zlecenia (1-Klik zapisuje zlecenie, wysyła e-mail do działu i generuje odpowiedni SMS):")
            p_col1, p_col2 = st.columns([1.3, 1.5])
            with p_col1:
                sciezka_zlecenia = st.radio(
                    "Kto ma zrealizować to zgłoszenie?",
                    [
                        "🛠️ 1. Zwykłe zgłoszenie serwisowe (SMS: wyznaczenie terminu w 24h)",
                        "🚒 2. PILNE / Z Księżyca – Przypisz serwisanta i datę OD RAZU (SMS z datą wizyty)",
                        "👷 3. Przekaż do Ekipy Monterskiej (1-Klik E-mail Montaż + SMS: termin do 5 dni)",
                        "💼 4. Przekaż prośbę o Ofertę do Handlowca (1-Klik E-mail Handlowiec + SMS: oferta do 3 dni)"
                    ]
                )
            with p_col2:
                pilny_serwisant = ""
                pilna_data = str(datetime.date.today())
                wybrany_email_handlowca = ""
                if sciezka_zlecenia.startswith("🚒"):
                    st.info("🚒 **Tryb Pilny (Straż Pożarna):** Zlecenie omija kolejkę i od razu wskakuje do trasy wybranego konserwatora!")
                    pilny_serwisant = st.selectbox("Wybierz konserwatora do pilnego wyjazdu:", list(profiles.keys()), key="urgent_tech_sel")
                    pilna_data = str(st.date_input("Data pilnej wizyty:", value=datetime.date.today(), key="urgent_date_sel"))
                elif sciezka_zlecenia.startswith("👷"):
                    st.info(f"👷 Zlecenie zostanie od razu wysłane na adres Działu Montażu: `{email_cfg.get('email_montaz', 'montaz@promag-gaz.pl')}`")
                elif sciezka_zlecenia.startswith("💼"):
                    h_list = email_cfg.get("handlowcy", [])
                    h_opts = [f"{h['nazwa']} <{h['email']}>" for h in h_list]
                    sel_h_oneclick = st.selectbox("Wybierz Handlowca (1-Klik wyśle do niego prośbę o ofertę):", h_opts) if h_opts else ""
                    if "<" in sel_h_oneclick:
                        wybrany_email_handlowca = sel_h_oneclick.split("<")[1].split(">")[0]
                    st.info(f"💼 Zlecenie zapisze się i od razu wyśle e-mail z prośbą o ofertę na: `{wybrany_email_handlowca}` (Watchdog 3 dni).")
                else:
                    st.caption("Standardowe przyjęcie do puli nowych zgłoszeń – klient otrzyma SMS o wyznaczeniu terminu w ciągu 24 godzin.")

            if st.button("🚀 ZAPISZ ZGŁOSZENIE I GENERUJ SMS / WYŚLIJ DO DZIAŁU", type="primary"):
                if not in_klient.strip():
                    st.error("Podaj imię i nazwisko klienta!")
                else:
                    new_id = f"ZGL-{len(df) + 1001}"
                    kod_final = in_kod.strip() or extract_postal_code(f"{in_adres} {in_miasto}")
                    status_init = "Nowe"
                    serw_init = ""
                    data_plan_init = ""
                    odd_dzial = ""
                    odd_email = ""
                    odd_data = ""
                    email_status_msg = ""

                    if sciezka_zlecenia.startswith("🚒"):
                        status_init = "Zaplanowane"
                        serw_init = pilny_serwisant
                        data_plan_init = pilna_data
                        in_kryt = "Tak"
                        sms_txt = build_sms_content("zaplanowano_wizyte", new_id, in_klient, pilny_serwisant, pilna_data, in_kwota)
                        tryb_opis = f"PILNE -> {pilny_serwisant} ({pilna_data})"
                    elif sciezka_zlecenia.startswith("👷"):
                        status_init = "Oddelegowane"
                        odd_dzial = "Dział Montażu"
                        odd_email = email_cfg.get("email_montaz", "montaz@promag-gaz.pl")
                        odd_data = today_iso
                        sms_txt = build_sms_content("montaz_5dni", new_id, in_klient)
                        tryb_opis = "Ekipa Monterska (Termin do 5 dni)"
                        subj_m = f"[PROMAG SERWIS -> MONTAŻ] {new_id} - {in_klient}"
                        body_m = f"Przekazano zgłoszenie dla Ekipy Monterskiej ({new_id}):\nKlient: {in_klient} | Tel: {in_tel}\nAdres: {kod_final} {in_miasto}, {in_adres}\nOpis: {in_notatka_konserwatora}"
                        _, email_status_msg = send_system_email(email_cfg, odd_email, subj_m, body_m)
                    elif sciezka_zlecenia.startswith("💼"):
                        status_init = "Oddelegowane"
                        odd_dzial = "Dział Handlowy"
                        odd_email = wybrany_email_handlowca
                        odd_data = today_iso
                        sms_txt = build_sms_content("handlowiec_3dni", new_id, in_klient)
                        tryb_opis = f"Oferta Handlowca ({odd_email} - do 3 dni)"
                        subj_h = f"[PROMAG SERWIS -> PROŚBA O OFERTĘ] {new_id} - {in_klient}"
                        body_h = f"Prośba o przygotowanie oferty handlowej (w ciągu max 3 dni) dla klienta:\nKlient: {in_klient} | Tel: {in_tel}\nAdres: {kod_final} {in_miasto}, {in_adres}\nInstalacja: {in_opis_instalacji}\nSzczegóły zapytania: {in_notatka_konserwatora}"
                        _, email_status_msg = send_system_email(email_cfg, odd_email, subj_h, body_h)
                    else:
                        sms_txt = build_sms_content("zwykle_24h", new_id, in_klient)
                        tryb_opis = "Zwykłe zgłoszenie (Termin w 24h)"

                    new_row = {
                        "ID_Zgloszenia": new_id, "Klient": in_klient.strip(), "Telefon": in_tel.strip(),
                        "Kod_Pocztowy": kod_final, "Miasto": in_miasto.strip(), "Adres": in_adres.strip(),
                        "Typ_Instalacji": in_typ_inst, "Data_Montazu": in_data_montazu.strip(),
                        "Opis_Instalacji_Montaz": in_opis_instalacji.strip(),
                        "Wymagany_Przeglad": "Tak" if "Przegląd" in in_notatka_konserwatora else "Nie",
                        "Opis_Zgloszenia": in_notatka_konserwatora.strip(),
                        "Priorytet_Krytyczny": in_kryt,
                        "Typ_Serwisu": "Gwarancyjny" if float(in_kwota) == 0 else "Płatny",
                        "Serwisant": serw_init, "Data_Planowana": data_plan_init, "Status": status_init,
                        "Naprawione_1_Wizyta": "Tak",
                        "Kwota_Przewidywana": float(in_kwota), "Zgoda_Klienta_Kwota": float(in_zgoda),
                        "Kwota_Koncowa": 0.0, "Koszt_Materialow": float(koszyk["suma_materialow"]),
                        "Forma_Platnosci": "Przelew", "Gotowka_Rozliczona": "Tak", "Czas_Pracy_Min": float(koszyk["suma_czasu_h"] * 60),
                        "Wyniki_Przegladu": "", "Zdjecia_Potwierdzone": "Nie", "Zalecenia_Klienta": "",
                        "Notatka_Serwisanta": "", "Powod_Niewykonania": "", "Data_Zamkniecia": "", "Podpis_Zlozony": "Nie",
                        "Do_Oddzwonienia": "Tak" if in_oddzwonic else "Nie",
                        "Data_Oddzwonienia": str(in_data_oddzw) if in_oddzwonic else "",
                        "Potwierdzony_Tel": "Nie" if in_oddzwonic else "Tak",
                        "Trudny_Klient": in_trudny, "Potrzebne_Czesci": "",
                        "Lead_Handlowy": in_notatka_konserwatora.strip() if sciezka_zlecenia.startswith("💼") else "",
                        "Oddelegowano_Dzial": odd_dzial, "Oddelegowano_Email": odd_email, "Data_Oddelegowania": odd_data,
                        "Notatka_Dla_Dzialu": in_notatka_konserwatora.strip() if odd_dzial else "",
                        "Otrzymano_Oferte_Termin": "Nie", "Ostatnie_Przypomnienie": ""
                    }
                    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                    save_tickets(df)

                    st.session_state["last_created_sms"] = {
                        "id": new_id,
                        "klient": in_klient.strip(),
                        "tel": in_tel.strip(),
                        "tryb_opis": tryb_opis,
                        "tekst": sms_txt,
                        "uri": make_sms_uri(in_tel.strip(), sms_txt),
                        "email_status": email_status_msg
                    }
                    st.rerun()

        st.divider()
        st.subheader("📋 Lista Aktywnych Zgłoszeń + Szybki Generator SMS i Dysponowanie Awaryjne")

        flt1, flt2, flt3 = st.columns([1.4, 1, 1])
        with flt1:
            q_table_filter = st.text_input("🔍 Filtruj tabelę na żywo (nazwisko, telefon, kod pocztowy, miasto, treść):", placeholder="Szukaj w zgłoszeniach...")
        with flt2:
            stat_filter = st.multiselect("Pokaż statusy:", ["Nowe", "Zaplanowane", "W trakcie", "Niewykonane", "Do zamówienia części", "Oddelegowane", "Zakończone"], default=["Nowe", "Zaplanowane", "W trakcie", "Niewykonane", "Do zamówienia części", "Oddelegowane"])
        with flt3:
            tech_filter = st.multiselect("Serwisant:", [""] + list(profiles.keys()))

        view_tickets = df.copy()
        if stat_filter:
            view_tickets = view_tickets[view_tickets["Status"].isin(stat_filter)]
        if tech_filter:
            view_tickets = view_tickets[view_tickets["Serwisant"].isin(tech_filter)]
        if q_table_filter.strip():
            qf = q_table_filter.strip()
            m_f = (
                view_tickets["Klient"].astype(str).str.contains(qf, case=False, na=False) |
                view_tickets["Telefon"].astype(str).str.contains(qf, case=False, na=False) |
                view_tickets["Kod_Pocztowy"].astype(str).str.contains(qf, case=False, na=False) |
                view_tickets["Adres"].astype(str).str.contains(qf, case=False, na=False) |
                view_tickets["Miasto"].astype(str).str.contains(qf, case=False, na=False) |
                view_tickets["Opis_Zgloszenia"].astype(str).str.contains(qf, case=False, na=False) |
                view_tickets["ID_Zgloszenia"].astype(str).str.contains(qf, case=False, na=False)
            )
            view_tickets = view_tickets[m_f]

        priority_display_cols = [
            "ID_Zgloszenia", "Klient", "Telefon", "Kod_Pocztowy", "Miasto", "Adres",
            "Typ_Instalacji", "Opis_Zgloszenia", "Kwota_Przewidywana", "Status",
            "Serwisant", "Data_Planowana", "Powod_Niewykonania"
        ]
        avail_cols = [c for c in priority_display_cols if c in view_tickets.columns]
        st.dataframe(
            view_tickets[avail_cols],
            column_config={
                "ID_Zgloszenia": st.column_config.TextColumn("Nr Zgł."),
                "Klient": st.column_config.TextColumn("Nazwisko i Imię Klienta"),
                "Telefon": st.column_config.TextColumn("Telefon"),
                "Kod_Pocztowy": st.column_config.TextColumn("Kod Pocztowy"),
                "Miasto": st.column_config.TextColumn("Miasto"),
                "Adres": st.column_config.TextColumn("Ulica i nr"),
                "Typ_Instalacji": st.column_config.TextColumn("Instalacja"),
                "Opis_Zgloszenia": st.column_config.TextColumn("Treść Zgłoszenia / Opis Usterki", width="large"),
                "Kwota_Przewidywana": st.column_config.NumberColumn("Wycena (zł)"),
                "Status": st.column_config.TextColumn("Status"),
                "Serwisant": st.column_config.TextColumn("Serwisant"),
                "Data_Planowana": st.column_config.TextColumn("Data Serwisu"),
                "Powod_Niewykonania": st.column_config.TextColumn("Powód Niewykonania")
            },
            use_container_width=True,
            hide_index=True
        )

        if not view_tickets.empty:
            with st.expander("📲 Szybki SMS / 1-Klik Przekazanie do Handlowca lub Ekipy / Pilne Przypisanie dla wybranego zlecenia z tabeli", expanded=False):
                q_opts = [f"{r['ID_Zgloszenia']} | {r['Klient']} | Tel: {r['Telefon']} | {r['Kod_Pocztowy']} {r['Miasto']}" for _, r in view_tickets.iterrows()]
                sel_q_str = st.selectbox("Wybierz zlecenie z listy:", q_opts, key="quick_act_sel")
                sel_q_id = sel_q_str.split(" | ")[0]
                q_idx = df[df["ID_Zgloszenia"] == sel_q_id].index[0]
                q_row = df.loc[q_idx]

                qa1, qa2, qa3 = st.columns(3)
                with qa1:
                    st.markdown("**📲 Wyślij SMS z podsumowaniem:**")
                    sms_type_q = st.selectbox("Szablon SMS:", [
                        "Zwykłe zgłoszenie (Termin w 24h)",
                        "Ekipa Monterska (Termin do 5 dni)",
                        "Oferta Handlowca (Kontakt do 3 dni)",
                        "Potwierdzenie daty wizyty konserwatora"
                    ], key="q_sms_t")
                    if sms_type_q.startswith("Zwykłe"):
                        s_body_q = build_sms_content("zwykle_24h", sel_q_id, q_row["Klient"])
                    elif sms_type_q.startswith("Ekipa"):
                        s_body_q = build_sms_content("montaz_5dni", sel_q_id, q_row["Klient"])
                    elif sms_type_q.startswith("Oferta"):
                        s_body_q = build_sms_content("handlowiec_3dni", sel_q_id, q_row["Klient"])
                    else:
                        s_body_q = build_sms_content("zaplanowano_wizyte", sel_q_id, q_row["Klient"], q_row["Serwisant"] or "Serwisant", q_row["Data_Planowana"] or today_iso, q_row["Kwota_Przewidywana"])
                    st.code(s_body_q, language=None)
                    uri_q = make_sms_uri(q_row["Telefon"], s_body_q)
                    if uri_q:
                        st.link_button(f"📲 Otwórz SMS do {q_row['Klient']}", uri_q, type="primary")

                with qa2:
                    st.markdown("**🚒 Pilne przypisanie konserwatora (Na ostro):**")
                    u_t = st.selectbox("Serwisant:", list(profiles.keys()), key="q_urg_t")
                    u_d = st.date_input("Data wizyty:", value=datetime.date.today(), key="q_urg_d")
                    if st.button("🚀 Przypisz natychmiast i wygeneruj SMS z datą"):
                        df.loc[q_idx, "Serwisant"] = u_t
                        df.loc[q_idx, "Data_Planowana"] = str(u_d)
                        df.loc[q_idx, "Status"] = "Zaplanowane"
                        save_tickets(df)
                        s_plan = build_sms_content("zaplanowano_wizyte", sel_q_id, q_row["Klient"], u_t, str(u_d), q_row["Kwota_Przewidywana"])
                        st.session_state["last_created_sms"] = {
                            "id": sel_q_id, "klient": q_row["Klient"], "tel": q_row["Telefon"],
                            "tryb_opis": f"Zaplanowano wizytę: {u_t} ({u_d})", "tekst": s_plan,
                            "uri": make_sms_uri(q_row["Telefon"], s_plan), "email_status": ""
                        }
                        st.rerun()

                with qa3:
                    st.markdown("**💼 1-Klik Wyślij prośbę do Handlowca / Montażu:**")
                    h_list_q = email_cfg.get("handlowcy", [])
                    h_opts_q = [f"{h['nazwa']} <{h['email']}>" for h in h_list_q] + [f"Dział Montażu <{email_cfg.get('email_montaz', 'montaz@promag-gaz.pl')}>"]
                    sel_hq = st.selectbox("Odbiorca:", h_opts_q, key="q_del_h")
                    if st.button("📧 Wyślij 1-Klikiem i przygotuj SMS"):
                        em_target = sel_hq.split("<")[1].split(">")[0] if "<" in sel_hq else ""
                        is_mont = "Montaż" in sel_hq
                        dz_name = "Dział Montażu" if is_mont else "Dział Handlowy"
                        df.loc[q_idx, "Status"] = "Oddelegowane"
                        df.loc[q_idx, "Oddelegowano_Dzial"] = dz_name
                        df.loc[q_idx, "Oddelegowano_Email"] = em_target
                        df.loc[q_idx, "Data_Oddelegowania"] = today_iso
                        save_tickets(df)
                        ok_e, msg_e = send_system_email(
                            email_cfg, em_target,
                            f"[PROMAG SERWIS -> {dz_name.upper()}] {sel_q_id} - {q_row['Klient']}",
                            f"Klient: {q_row['Klient']} | Tel: {q_row['Telefon']}\nAdres: {q_row['Kod_Pocztowy']} {q_row['Miasto']}, {q_row['Adres']}\nOpis: {q_row['Opis_Zgloszenia']}"
                        )
                        s_del = build_sms_content("montaz_5dni" if is_mont else "handlowiec_3dni", sel_q_id, q_row["Klient"])
                        st.session_state["last_created_sms"] = {
                            "id": sel_q_id, "klient": q_row["Klient"], "tel": q_row["Telefon"],
                            "tryb_opis": dz_name, "tekst": s_del,
                            "uri": make_sms_uri(q_row["Telefon"], s_del), "email_status": msg_e
                        }
                        st.rerun()

    # ZAKŁADKA 2: PLANER KORYTARZOWY
    with tab_plan:
        st.subheader("🗺️ Planer Korytarzowy („Po Drodze”) + Automatyczny Generator SMS dla Klientów z Trasy")
        st.caption("💡 **Wskazówka na gorący okres:** Jeśli musisz wysłać konserwatora do pilnej awarii (pożaru), wpisz kod pocztowy tej awarii jako **Kod Docelowy (Cel)** – planer pokaże Ci zlecenia leżące dokładnie po drodze, żeby nie robić pustych przebiegów!")

        if st.session_state.get("last_planned_route_sms"):
            route_sms_pack = st.session_state["last_planned_route_sms"]
            st.markdown('<div class="sms-box">', unsafe_allow_html=True)
            st.markdown(f"#### 📲 Powiadomienia SMS dla zaplanowanej trasy ({route_sms_pack['serwisant']} — Dzień: **{route_sms_pack['data']}**)")
            st.caption("Kliknij przycisk przy kliencie, aby wysłać SMS z datą wizyty, lub skopiuj zbiorczą listę poniżej:")
            all_sms_lines = []
            for item_s in route_sms_pack["items"]:
                all_sms_lines.append(f"Do: {item_s['tel']} ({item_s['klient']})\n{item_s['tekst']}\n")
                r_c1, r_c2 = st.columns([2.5, 1])
                with r_c1:
                    st.markdown(f"👤 **{item_s['klient']}** (`{item_s['tel']}`): *{item_s['tekst']}*")
                with r_c2:
                    if item_s["uri"]:
                        st.link_button(f"📲 Wyślij SMS ({item_s['tel']})", item_s["uri"], type="primary")
            with st.expander("📋 Zbiorcza lista SMS do skopiowania"):
                st.code("\n".join(all_sms_lines), language=None)
            if st.button("✖️ Ukryj listę SMS trasy"):
                st.session_state["last_planned_route_sms"] = None
                st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)

        pc1, pc2, pc3, pc4, pc5 = st.columns(5)
        with pc1:
            sel_tech_plan = st.selectbox("Serwisant:", list(profiles.keys()), key="cor_tech")
        with pc2:
            start_postal = st.text_input("Baza Startowa (Kod):", value="05-552")
        with pc3:
            target_postal = st.text_input("Kod Docelowy (Cel / Pożar):", value="05-800")
        with pc4:
            corridor_width_km = st.slider("Szerokość korytarza [km]:", 5, 120, 25, 5)
        with pc5:
            plan_date = st.date_input("Data trasy:", value=datetime.date.today() + datetime.timedelta(days=1))

        start_lat, start_lon = get_coords_for_postal(start_postal)
        dest_lat, dest_lon = get_coords_for_postal(target_postal)
        axis_length_km = round(haversine_km(start_lat, start_lon, dest_lat, dest_lon) * ROAD_TORTUOSITY, 1)
        st.info(f"📏 **Długość osi trasy ({start_postal} ➔ {target_postal}):** ok. `{axis_length_km} km` drogowych | Szukam zgłoszeń w pasie `±{corridor_width_km} km` od tej linii.")

        pending_df = df[df["Status"].isin(["Nowe", "Niewykonane", "Do zaplanowania", ""])].copy()
        if not pending_df.empty:
            valid_rows = []
            missing_zip_rows = []

            for idx_p, r in pending_df.iterrows():
                rlat, rlon, resolved_zip = extract_coords_from_row(r)
                if resolved_zip and not str(r.get("Kod_Pocztowy", "")).strip():
                    pending_df.at[idx_p, "Kod_Pocztowy"] = resolved_zip
                    df.at[idx_p, "Kod_Pocztowy"] = resolved_zip

                if rlat is None or rlon is None:
                    missing_zip_rows.append(pending_df.loc[idx_p])
                else:
                    d_base = round(haversine_km(start_lat, start_lon, rlat, rlon) * ROAD_TORTUOSITY, 1)
                    d_dest = round(haversine_km(dest_lat, dest_lon, rlat, rlon) * ROAD_TORTUOSITY, 1)
                    d_cor = round(distance_from_segment_km(rlat, rlon, start_lat, start_lon, dest_lat, dest_lon), 1)
                    row_copy = pending_df.loc[idx_p].copy()
                    row_copy["Kod_Pocztowy"] = resolved_zip or str(r.get("Kod_Pocztowy", ""))
                    row_copy["Km_od_Osi_Trasy"] = d_cor
                    row_copy["Km_od_Bazy"] = d_base
                    row_copy["Km_od_Celu"] = d_dest
                    row_copy["_coords"] = (rlat, rlon)
                    valid_rows.append(row_copy)

            valid_plan_df = pd.DataFrame(valid_rows) if valid_rows else pd.DataFrame()

            if not valid_plan_df.empty:
                max_allowed_dist = max(axis_length_km + corridor_width_km * 1.5, corridor_width_km * 2.0)
                filtered_plan = valid_plan_df[
                    (valid_plan_df["Km_od_Osi_Trasy"] <= corridor_width_km) &
                    (valid_plan_df["Km_od_Bazy"] <= max_allowed_dist)
                ].sort_values(by=["Priorytet_Krytyczny", "Km_od_Bazy"], ascending=[False, True]).copy()

                if not filtered_plan.empty:
                    disp_cols = [
                        "ID_Zgloszenia", "Priorytet_Krytyczny", "Kod_Pocztowy", "Miasto", "Adres",
                        "Km_od_Osi_Trasy", "Km_od_Bazy", "Klient", "Telefon", "Kwota_Przewidywana", "Opis_Zgloszenia"
                    ]
                    ed_df = filtered_plan[disp_cols].copy()
                    ed_df.insert(0, "Wybierz", False)
                    edited_route = st.data_editor(
                        ed_df,
                        column_config={
                            "Wybierz": st.column_config.CheckboxColumn("Wybierz do trasy"),
                            "ID_Zgloszenia": st.column_config.TextColumn("Nr Zgł."),
                            "Priorytet_Krytyczny": st.column_config.TextColumn("Pilne?"),
                            "Kod_Pocztowy": st.column_config.TextColumn("Kod Pocztowy"),
                            "Miasto": st.column_config.TextColumn("Miejscowość"),
                            "Adres": st.column_config.TextColumn("Ulica i nr"),
                            "Km_od_Osi_Trasy": st.column_config.NumberColumn("Odchylenie od trasy (km)"),
                            "Km_od_Bazy": st.column_config.NumberColumn("Odległość od bazy (km)"),
                            "Kwota_Przewidywana": st.column_config.NumberColumn("Wycena (zł)")
                        },
                        use_container_width=True, hide_index=True, key="cor_ed"
                    )

                    chosen_ids = edited_route[edited_route["Wybierz"] == True]["ID_Zgloszenia"].tolist()
                    chosen_rows = filtered_plan[filtered_plan["ID_Zgloszenia"].isin(chosen_ids)]
                    gross_rev = float(chosen_rows["Kwota_Przewidywana"].sum())
                    parts_c = float(chosen_rows["Koszt_Materialow"].sum())

                    econ = estimate_route_economics(
                        chosen_rows["_coords"].tolist(), (start_lat, start_lon), gross_rev, parts_cost=parts_c,
                        hourly_rate=stawka_rbh, fuel_price=KOSZT_PALIWA_LITR,
                        burn_l_100km=DEFAULT_BURN_L_100KM, onsite_hours_per_client=SR_CZAS_WIZYTY_H
                    )

                    ek1, ek2, ek3, ek4, ek5 = st.columns(5)
                    ek1.metric("Liczba punktów (Cel: 4-5)", f"{len(chosen_ids)} pkt", "Optymalnie!" if len(chosen_ids) >= 4 else "Dobierz do 4-5 pkt")
                    ek2.metric("Przychód z trasy (Do prowizji)", f"{gross_rev:,.0f} zł", f"Przy 20 dniach: {gross_rev*20:,.0f} zł/mies.")
                    ek3.metric("Dystans i Paliwo (8 zł/l)", f"{econ['road_km']} km (-{econ['fuel_cost']:,.0f} zł)")
                    ek4.metric(f"Koszt czasu ({stawka_rbh:.0f} zł/h)", f"-{(econ['labor_drive_cost'] + econ['labor_onsite_cost']):,.0f} zł")
                    ek5.metric("💎 CZYSTY ZYSK NETTO", f"{econ['net_profit']:,.0f} zł", delta=f"{econ['net_profit'] - TARGET_NET_PROFIT:,.0f} zł vs 1200 zł")

                    if len(chosen_ids) > 0:
                        if st.button(f"🚀 PRZYPISZ TRASĘ ({len(chosen_ids)} zleceń) DO: {sel_tech_plan} NA DZIEŃ {plan_date} I WYGENERUJ SMS-Y DO KLIENTÓW", type="primary"):
                            sms_items_list = []
                            for cid in chosen_ids:
                                idx_m = df[df["ID_Zgloszenia"] == cid].index[0]
                                df.loc[idx_m, "Serwisant"] = sel_tech_plan
                                df.loc[idx_m, "Data_Planowana"] = str(plan_date)
                                df.loc[idx_m, "Status"] = "Zaplanowane"
                                r_cl = df.loc[idx_m]
                                s_txt = build_sms_content("zaplanowano_wizyte", cid, r_cl["Klient"], sel_tech_plan, str(plan_date), r_cl["Kwota_Przewidywana"])
                                sms_items_list.append({
                                    "id": cid, "klient": r_cl["Klient"], "tel": r_cl["Telefon"],
                                    "tekst": s_txt, "uri": make_sms_uri(r_cl["Telefon"], s_txt)
                                })
                            save_tickets(df)
                            st.session_state["last_planned_route_sms"] = {
                                "serwisant": sel_tech_plan, "data": str(plan_date), "items": sms_items_list
                            }
                            st.success("Trasa zaplanowana! Poniżej wygenerowano powiadomienia SMS dla wszystkich klientów z trasy.")
                            st.rerun()
                else:
                    st.info("Brak zgłoszeń mieszczących się w zadanym korytarzu trasy. Zwiększ szerokość korytarza suwakiem lub wpisz inny kod docelowy.")

            if missing_zip_rows:
                with st.expander(f"⚠️ Zgłoszenia bez rozpoznanego kodu pocztowego ({len(missing_zip_rows)} szt. – odfiltrowane, by nie przekłamywać kilometrów)", expanded=False):
                    st.caption("Uzupełnij kod pocztowy XX-XXX w poniższej tabeli, aby pojawiły się na mapie korytarza:")
                    miss_df = pd.DataFrame(missing_zip_rows)[["ID_Zgloszenia", "Klient", "Telefon", "Kod_Pocztowy", "Miasto", "Adres", "Opis_Zgloszenia"]]
                    ed_miss = st.data_editor(miss_df, use_container_width=True, hide_index=True, key="ed_missing_zips")
                    if st.button("💾 Zapisz uzupełnione kody pocztowe"):
                        for _, rm in ed_miss.iterrows():
                            im = df[df["ID_Zgloszenia"] == rm["ID_Zgloszenia"]].index[0]
                            df.loc[im, "Kod_Pocztowy"] = str(rm["Kod_Pocztowy"]).strip()
                            df.loc[im, "Miasto"] = str(rm["Miasto"]).strip()
                        save_tickets(df)
                        st.success("Zaktualizowano kody pocztowe!")
                        st.rerun()
        else:
            st.info("Brak nowych zgłoszeń oczekujących na zaplanowanie.")

    # ZAKŁADKA 3: ODDELEGOWANIA, CZĘŚCI I LEADY
    with tab_oddeleguj:
        st.subheader("📤 Status Zleceń Oddelegowanych, Części do Zamówienia i Leadów Handlowych")
        col_od1, col_od2 = st.columns(2)
        with col_od1:
            st.markdown("#### ⏳ Oddelegowane do Handlowców / Montażu (Watchdog 3 dni)")
            del_df = df[df["Oddelegowano_Dzial"].astype(str).str.strip() != ""]
            if del_df.empty:
                st.info("Brak oddelegowanych zleceń.")
            for di, dr in del_df.iterrows():
                did = dr["ID_Zgloszenia"]
                got_o = (dr["Otrzymano_Oferte_Termin"] == "Tak")
                with st.container(border=True):
                    st.markdown(f"**{did} – {dr['Klient']}** (`{dr['Oddelegowano_Dzial']}` ➔ `{dr['Oddelegowano_Email']}`)")
                    st.caption(f"Data przekazania: {dr['Data_Oddelegowania']} | Status: {'🟢 Otrzymano ofertę/termin' if got_o else '🔴 Czeka na reakcję działu'}")
                    st.markdown(f"📝 {dr['Notatka_Dla_Dzialu'] or dr['Opis_Zgloszenia']}")
                    if not got_o:
                        if st.button(f"✅ Oznacz: Otrzymano ofertę / termin ({did})", key=f"ok_del_{did}"):
                            df.loc[di, "Otrzymano_Oferte_Termin"] = "Tak"
                            save_tickets(df)
                            st.rerun()

        with col_od2:
            st.markdown("#### 📦 Zlecenia oczekujące na zamówienie części oraz 🔥 Leady z terenu")
            parts_df = df[df["Status"] == "Do zamówienia części"]
            if not parts_df.empty:
                st.warning(f"📦 **Wymagane zamówienie części ({len(parts_df)} szt.):**")
                st.dataframe(parts_df[["ID_Zgloszenia", "Klient", "Telefon", "Serwisant", "Potrzebne_Czesci", "Notatka_Serwisanta"]], use_container_width=True, hide_index=True)
            leads_df = df[df["Lead_Handlowy"].astype(str).str.strip() != ""]
            if not leads_df.empty:
                st.success(f"🔥 **Leady Handlowe przywiezione przez serwisantów ({len(leads_df)} szt.):**")
                st.dataframe(leads_df[["ID_Zgloszenia", "Klient", "Telefon", "Miasto", "Serwisant", "Lead_Handlowy"]], use_container_width=True, hide_index=True)

    # ZAKŁADKA 4: CENNIK I KADRY
    with tab_cennik_kadry:
        st.subheader("💰 Zestandaryzowany Cennik Usług (Nazwa usługi | Cena materiałów | Czas pracy x Stawka godzinowa)")
        ck1, ck2 = st.columns(2)
        with ck1:
            new_stawka_rbh = st.number_input("Stawka godzinowa pracy serwisanta (zł netto / h):", 50.0, 500.0, float(stawka_rbh), 10.0)
        with ck2:
            new_dojazd_kw = st.number_input("Stała kwota za dojazd doliczana do zlecenia (zł):", 0.0, 1000.0, float(dojazd_kw), 20.0)

        cennik_df_view = pd.DataFrame(cennik_tab)
        cennik_df_view["Wyliczona_Cena_Uslugi_ZL"] = cennik_df_view.apply(lambda r: wylicz_cene_uslugi(r, new_stawka_rbh), axis=1)
        ed_cennik_df = st.data_editor(
            cennik_df_view,
            column_config={
                "usluga": st.column_config.TextColumn("Nazwa usługi / czynności"),
                "cena_materialow": st.column_config.NumberColumn("Cena materiałów (zł netto)", min_value=0.0, step=10.0),
                "czas_pracy_h": st.column_config.NumberColumn("Czas pracy (h)", min_value=0.0, step=0.5),
                "Wyliczona_Cena_Uslugi_ZL": st.column_config.NumberColumn("Cena pozycji (Materiały + h * Stawka)", disabled=True)
            },
            num_rows="dynamic", use_container_width=True, key="ed_cennik_table"
        )
        if st.button("💾 Zapisz Cennik i Stawki", type="primary"):
            clean_c_list = ed_cennik_df[["usluga", "cena_materialow", "czas_pracy_h"]].to_dict(orient="records")
            st.session_state["cennik_tab"] = clean_c_list
            st.session_state["stawka_rbh"] = float(new_stawka_rbh)
            st.session_state["dojazd_kw"] = float(new_dojazd_kw)
            save_config_and_profiles(profiles, cennik_tab=clean_c_list, stawka_rbh=new_stawka_rbh, dojazd_kw=new_dojazd_kw)
            st.success("Zaktualizowano cennik!")
            st.rerun()

        st.divider()
        st.subheader("👷 Kadry, Unikalne Kody PIN, Narzędzia i Stan Materiałowy na Samochodach")
        sel_t_k = st.selectbox("Wybierz serwisanta:", list(profiles.keys()))
        t_prof = profiles[sel_t_k]
        t_m = calculate_tech_metrics(df, sel_t_k)

        pin_c1, pin_c2 = st.columns([1, 2])
        with pin_c1:
            new_tech_pin = st.text_input(f"🔑 Unikalny kod PIN dla {sel_t_k}:", value=str(t_prof.get("pin", "")))
            if st.button(f"💾 Zapisz PIN ({sel_t_k})"):
                profiles[sel_t_k]["pin"] = new_tech_pin.strip()
                save_config_and_profiles(profiles)
                st.success("Zapisano nowy PIN!")

        if t_m["cash_in_hand"] > 0:
            if st.button(f"💵 Rozlicz gotówkę zdaną w biurze przez {sel_t_k} ({t_m['cash_in_hand']:,.0f} zł)"):
                df.loc[(df["Serwisant"] == sel_t_k) & (df["Forma_Platnosci"] == "Gotówka pobrana"), "Gotowka_Rozliczona"] = "Tak"
                save_tickets(df)
                st.rerun()

        ic1, ic2 = st.columns(2)
        with ic1:
            st.markdown(f"**📦 Stan materiałowy na aucie ({sel_t_k}):**")
            ed_m_k = st.data_editor(pd.DataFrame(t_prof.get("materialy", [])), num_rows="dynamic", use_container_width=True, key=f"km_{sel_t_k}")
            if st.button(f"💾 Zapisz materiały ({sel_t_k})"):
                profiles[sel_t_k]["materialy"] = ed_m_k.to_dict(orient="records")
                save_config_and_profiles(profiles)
                st.success("Zapisano!")
        with ic2:
            st.markdown(f"**🧰 Narzędzia przypisane ({sel_t_k}):**")
            ed_t_k = st.data_editor(pd.DataFrame(t_prof.get("narzedzia", [])), num_rows="dynamic", use_container_width=True, key=f"kt_{sel_t_k}")
            if st.button(f"💾 Zapisz narzędzia ({sel_t_k})"):
                profiles[sel_t_k]["narzedzia"] = ed_t_k.to_dict(orient="records")
                save_config_and_profiles(profiles)
                st.success("Zapisano!")

    # ZAKŁADKA 5: USTAWIENIA, BACKUP ZGŁOSZEŃ, LOGO, BAZA MONTAŻY, SMTP
    with tab_ustawienia:
        st.subheader("💾 Trwałość Danych i Kopia Zapasowa Zgłoszeń (`Serwisowa posortowana.xlsx`)")
        st.info(f"**Aktualny tryb zapisu:** `{st.session_state.get('storage_mode', 'Lokalny')}`. Poniżej możesz pobrać lub wgrać pełny plik zgłoszeń Excela:")
        b_col1, b_col2 = st.columns(2)
        with b_col1:
            excel_buf = BytesIO()
            df.to_excel(excel_buf, index=False)
            st.download_button(
                label="📥 Pobierz aktualną bazę zgłoszeń (Serwisowa posortowana.xlsx)",
                data=excel_buf.getvalue(),
                file_name="Serwisowa posortowana.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
        with b_col2:
            up_tickets_xlsx = st.file_uploader("📤 Przywróć / Wgraj bazę zgłoszeń z pliku .xlsx:", type=["xlsx"])
            if up_tickets_xlsx is not None:
                if st.button("⚡ Załaduj wgrany plik zgłoszeń do systemu"):
                    restored_df = ensure_columns(pd.read_excel(up_tickets_xlsx))
                    save_tickets(restored_df)
                    st.success(f"Załadowano {len(restored_df)} zgłoszeń!")
                    st.rerun()

        st.divider()
        st.subheader("🎨 Uproszczone Logo Aplikacji i Ikona Paska Przeglądarki (Favicon)")
        up_logo = st.file_uploader("Wgraj uproszczone logo (PNG / JPG):", type=["png", "jpg", "jpeg"])
        if up_logo is not None:
            if st.button("💾 Zapisz uproszczone logo na pasku przeglądarki i w aplikacji", type="primary"):
                img_l = Image.open(up_logo)
                img_l.thumbnail((360, 120))
                img_l.save(LOGO_FILE, format="PNG")
                buf = BytesIO()
                img_l.save(buf, format="PNG")
                b64_str = base64.b64encode(buf.getvalue()).decode("utf-8")
                st.session_state["logo_b64"] = b64_str
                save_config_and_profiles(profiles, logo_b64=b64_str)
                st.success("Zapisano! Logo jest już widoczne na pasku przeglądarki oraz w nagłówku.")
                st.rerun()

        st.divider()
        st.subheader(f"🏠 Pełna Baza Montaży (`skonczone_clean.parquet` / `skonczone.csv`) — Załadowano rekordów: {len(df_montaze)}")
        up_montaze = st.file_uploader("📥 Wgraj plik `skonczone_clean.parquet`, `.csv` lub `.xlsx`:", type=["parquet", "csv", "xlsx"])
        if up_montaze is not None:
            if st.button("⚡ Przetwórz i połącz bazę wszystkich klientów", type="primary"):
                if up_montaze.name.endswith(".parquet"):
                    r_df = pd.read_parquet(up_montaze)
                elif up_montaze.name.endswith(".csv"):
                    try:
                        r_df = pd.read_csv(up_montaze, sep=None, engine="python", on_bad_lines="skip")
                    except Exception:
                        up_montaze.seek(0)
                        r_df = pd.read_csv(up_montaze, encoding="cp1250", sep=None, engine="python", on_bad_lines="skip")
                else:
                    r_df = pd.read_excel(up_montaze)
                c_df = clean_and_parse_montaze_df(r_df)
                st.session_state["df_montaze"] = c_df
                try:
                    c_df.to_parquet(MONTAZE_PARQUET_CACHE, index=False)
                except Exception:
                    pass
                st.cache_data.clear()
                st.success(f"Załadowano i oczyszczono {len(c_df)} klientów!")
                st.rerun()

        st.divider()
        st.subheader("📱 Eksport Kontaktów do Google Contacts (Paczki po 2 500 szt. z `skonczone_lekkie.csv`)")
        batch_sz = st.number_input("Wielkość pojedynczej paczki (zalecane 2500, limit Google to 3000):", min_value=500, max_value=3000, value=2500, step=250)
        if st.button("📇 Przygotuj paczki CSV do importu w Google Contacts", type="primary"):
            batches_list, zip_data, total_cnt = generate_google_contacts_batches(df_montaze, df, batch_size=int(batch_sz))
            st.session_state["gcontacts_batches"] = batches_list
            st.session_state["gcontacts_zip"] = zip_data
            st.session_state["gcontacts_total"] = total_cnt

        if st.session_state.get("gcontacts_batches"):
            st.success(f"✅ Wyodrębniono **{st.session_state['gcontacts_total']}** unikalnych klientów z numerami telefonów i podzielono na **{len(st.session_state['gcontacts_batches'])}** paczek!")
            st.download_button(
                label=f"📦 POBIERZ WSZYSTKIE PACZKI W JEDNYM PLIKU ZIP ({len(st.session_state['gcontacts_batches'])} plików CSV)",
                data=st.session_state["gcontacts_zip"],
                file_name="Kontakty_Google_PROMAG_Paczki.zip",
                mime="application/zip",
                type="primary"
            )

        st.divider()
        st.subheader("📧 Konfiguracja E-mail (SMTP) i Listy Handlowców")
        ec1, ec2 = st.columns(2)
        with ec1:
            s_host = st.text_input("Serwer SMTP:", value=email_cfg.get("smtp_server", "smtp.gmail.com"))
            s_port = st.number_input("Port SMTP:", value=int(email_cfg.get("smtp_port", 587)))
            s_user = st.text_input("E-mail nadawcy:", value=email_cfg.get("sender_email", ""))
            s_pass = st.text_input("Hasło aplikacyjne:", value=email_cfg.get("sender_password", ""), type="password")
            e_mont = st.text_input("E-mail Działu Montażu:", value=email_cfg.get("email_montaz", ""))
            e_tech = st.text_input("E-mail Specjalisty Technicznego:", value=email_cfg.get("email_technik", ""))
            if st.button("💾 Zapisz SMTP"):
                email_cfg.update({"smtp_server": s_host, "smtp_port": int(s_port), "sender_email": s_user, "sender_password": s_pass, "email_montaz": e_mont, "email_technik": e_tech})
                st.session_state["email_cfg"] = email_cfg
                save_config_and_profiles(profiles, email_cfg=email_cfg)
                st.success("Zapisano SMTP!")
        with ec2:
            st.markdown("**💼 Lista Handlowców:**")
            ed_h = st.data_editor(pd.DataFrame(email_cfg.get("handlowcy", [])), num_rows="dynamic", use_container_width=True)
            if st.button("💾 Zapisz Handlowców"):
                email_cfg["handlowcy"] = ed_h.to_dict(orient="records")
                st.session_state["email_cfg"] = email_cfg
                save_config_and_profiles(profiles, email_cfg=email_cfg)
                st.success("Zapisano listę handlowców!")

# ==========================================
# 8. PANEL ZARZĄDU (FINANSE + RANKING GRYWALIZACJI WATÓW I SERII)
# ==========================================
elif st.session_state["logged_role"] == "zarzad":
    df = st.session_state["df_tickets"]
    profiles = st.session_state["profiles"]
    st.subheader("📊 Panel Zarządu – Wyniki Finansowe, Prowizje, Seria Combo i Waty ⚡")
    rows_z = []
    for t_name, t_prof in profiles.items():
        tm = calculate_tech_metrics(df, t_name)
        gm = calculate_gamification_stats(df, t_name, t_prof)
        base_s = float(t_prof.get("wynagrodzenie_bazowe", 6000.0))
        rows_z.append({
            "Serwisant": t_name,
            "⚡ Waty (Pkt)": gm["total_waty"],
            "🔥 Seria Combo": f"{gm['current_streak']} (Rek: {gm['best_streak']})",
            "🏅 Odznaki": f"{gm['unlocked_cnt']}/{len(gm['badges'])}",
            "Zamknięte wizyty": tm["total_closed"],
            "Skuteczność 1. wizyty": f"{tm['ftfr_pct']}%",
            "Leady Handlowe": tm["leads_count"],
            "Gotówka do zdania (zł)": tm["cash_in_hand"],
            "Obrót miesięczny (zł)": tm["monthly_revenue"],
            "Podstawa (zł)": base_s,
            "Prowizja (zł)": tm["commission_pln"],
            "Do wypłaty Łącznie (zł)": round(base_s + tm["commission_pln"], 2)
        })
    st.dataframe(pd.DataFrame(rows_z), use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("📋 Pełna Tabela Zgłoszeń w Systemie")
    st.dataframe(df, use_container_width=True, hide_index=True)
