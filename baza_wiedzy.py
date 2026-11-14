
# -*- coding: utf-8 -*-
"""
Moduł Bazy Wiedzy, Cenników, Domyślnych Magazynów i Współrzędnych Kodów Pocztowych
dla systemu PROMAG OZE.
"""

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
    "Wymiana anody magnezowej / tytanowej": 450.0,
    "Serwis Gwarancyjny (Rozliczenie wewnętrzne)": 0.0
}

DOMYSLNE_NARZEDZIA = [
    {"nazwa": "Miernik cęgowy DC/AC + Megaomomierz Riso", "ilosc": 1, "stan": "Sprawny"},
    {"nazwa": "Refraktometr optyczny do glikolu + paski pH", "ilosc": 1, "stan": "Sprawny"},
    {"nazwa": "Pompa elektryczna do napełniania i płukania solarów", "ilosc": 1, "stan": "Sprawny"},
    {"nazwa": "Zestaw manometrów chłodniczych (R32 / R410A) + detektor gazu", "ilosc": 1, "stan": "Sprawny"},
    {"nazwa": "Klucz dynamometryczny (M3-M10) + zaciskarka MC4/RJ45", "ilosc": 1, "stan": "Sprawny"},
    {"nazwa": "Wiertarka udarowa + skrzynka narzędziowa instalatora", "ilosc": 1, "stan": "Sprawny"}
]

DOMYSLNE_MATERIALY = [
    {"kod": "MAT-01", "nazwa": "Glikol propylenowy solarny/PC (bańka 20L)", "ilosc": 4, "cena_zakupu": 180.0, "cena_sprzedazy": 320.0},
    {"kod": "MAT-02", "nazwa": "Anoda magnezowa 3/4 cala do zasobnika", "ilosc": 4, "cena_zakupu": 45.0, "cena_sprzedazy": 120.0},
    {"kod": "MAT-03", "nazwa": "Naczynie przeponowe solarne/CO 18L", "ilosc": 2, "cena_zakupu": 140.0, "cena_sprzedazy": 280.0},
    {"kod": "MAT-04", "nazwa": "Czujnik temperatury NTC 10k / PT1000", "ilosc": 5, "cena_zakupu": 20.0, "cena_sprzedazy": 80.0},
    {"kod": "MAT-05", "nazwa": "Odpowietrznik automatyczny solarny 150 st.C", "ilosc": 3, "cena_zakupu": 35.0, "cena_sprzedazy": 90.0},
    {"kod": "MAT-06", "nazwa": "Zawór bezpieczeństwa 3 bar / 6 bar", "ilosc": 3, "cena_zakupu": 30.0, "cena_sprzedazy": 85.0},
    {"kod": "MAT-07", "nazwa": "Wkładka ochronnika przepięć SPD DC", "ilosc": 4, "cena_zakupu": 55.0, "cena_sprzedazy": 140.0},
    {"kod": "MAT-08", "nazwa": "Komplet złączy MC4 oryginał (para + i -)", "ilosc": 10, "cena_zakupu": 8.0, "cena_sprzedazy": 25.0}
]

DIAGNOSTYKA_BAZA = {
    "Falownik Deye (Hybrydowy SUN-3K-12K-SG05LP3 / Stringowy)": {
        "F01 / F07 - Błąd polaryzacji DC / Awaria START DC": (
            "CO TO ZNACZY: Odwrotna polaryzacja na wejściu PV lub napięcie szyny BUS nie może wystartować z PV/baterii.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Wyłącz rozłącznik DC i zmierz multimetrem biegunowość stringów (+ do +, - do -).\n"
            "2. Zrestartuj inwerter (Bateria -> DC -> AC)."
        ),
        "F08 - GFDI_Relay_Failure (Błąd przekaźnika uziemienia / N-PE)": (
            "CO TO ZNACZY: Falownik wykrył problem z uziemieniem obudowy lub przebicie między N a PE.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Wyłącz falownik (AC, DC, Bateria) i odczekaj 5 minut.\n"
            "2. Zmierz miernikiem napięcie AC między szyną N a PE w rozdzielnicy (norma < 5-10V).\n"
            "3. W menu Deye (Advanced Function -> kod 7777) zweryfikuj opcję 'Signal ISLAND MODE' (powinna być włączona przy pracy z obwodem Backup/LOAD)."
        ),
        "F13 - Working_Mode_Change (Reset trybu pracy / Zmiana parametrów sieci)": (
            "CO TO ZNACZY: Zmiana parametrów sieci, przełączenie w tryb 'Brak aku' lub zmiana trybu pracy.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Odczekaj 3 minuty – falownik zazwyczaj sam odliczy czas i wystartuje.\n"
            "2. Jeśli wisi: wyłącz AC, przełącz hebel DC na OFF, wyłącz baterię. Odczekaj 2 minuty do wygaszenia LCD. Włącz: AC -> DC -> Bateria."
        ),
        "F15 / F18 / F34 - Tz_Ac_OverCurr_Fault (Zwarcie lub przetężenie po stronie AC / LOAD)": (
            "CO TO ZNACZY: Przeciążenie prądowe na wyjściu AC lub w obwodzie awaryjnym LOAD (Backup).\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Wyłącz bezpiecznik obwodu LOAD (wyjście awaryjne) i sprawdź, czy falownik wstaje bez obciążenia domu.\n"
            "2. Sprawdź obwody odbiorcze pod kątem zwarć L-N / L-PE oraz urządzeń o dużym prądzie rozruchowym (pompy, sprężarki)."
        ),
        "F20 / F21 - Tz_Dc_OverCurr_Fault (Przetężenie po stronie DC / BUS)": (
            "CO TO ZNACZY: Prąd z paneli PV lub akumulatora przekroczył limit sprzętowy (lub start w trybie Off-Grid ze zbyt dużym obciążeniem).\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Zmniejsz obciążenie na porcie LOAD i sprawdź nastawy Max A Charge / Discharge w Battery Setting.\n"
            "2. Rozepnij MC4 i zmierz napięcie Uoc (<800V) oraz prąd zwarciowy Isc stringów (limit roboczy 20A/26A na MPPT)."
        ),
        "F22 - Tz_EmergStop_Fault (Aktywowany wyłącznik PPOŻ / Zdalne wyłączenie RSD)": (
            "CO TO ZNACZY: Otwarty obwód wyłącznika awaryjnego RSD na płycie sterującej (kostka CN2).\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Odkręć pokrywę komory przyłączy falownika Deye.\n"
            "2. Sprawdź kostkę RSD_input (zaciski B, B, +, -). Jeśli brak zewnętrznego wyłącznika PPOŻ, upewnij się, że fabryczna zworka jest dokręcona."
        ),
        "F16 / F23 / F24 - Tz_GFCI_OC / DC_Insulation_Fault (Prąd upływu / Błąd izolacji DC)": (
            "CO TO ZNACZY: Przebicie prądu z paneli PV lub kabli solarnych przez wilgoć do uziemienia (Riso < 1 MOhm).\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Wyłącz AC i DC, wypnij wszystkie stringi z falownika.\n"
            "2. Zmierz megaomomierzem rezystancję izolacji Riso każdego bieguna PV do uziemienia PE (norma >= 1.0 MOhm).\n"
            "3. Wypnij wkładki ochronnika przepięć (SPD) DC w skrzynce – jeśli błąd zniknie, wymień zawilgoconą wkładkę SPD."
        ),
        "F26 - Bus_Unbalance_Fault (Niezrównoważenie szyny DC)": (
            "CO TO ZNACZY: Duża asymetria obciążenia 3 faz lub upływ DC.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Włącz w menu Advanced Function (kod 7777) opcję 'Asymmetric phase feeding'.\n"
            "2. Zrestartuj falownik 2-3 razy i sprawdź równowagę faz na porcie LOAD."
        ),
        "F29 / F41 - Parallel_Comm_Fault (Błąd komunikacji równoległej CAN)": (
            "CO TO ZNACZY: W menu włączono pracę kaskadową (Parallel), a pracuje pojedynczy falownik (lub 1 jednostka w kaskadzie się wyłączyła).\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Wejdź w: Advanced Function -> kod 7777 -> odznacz 'Parallel', 'Master' oraz 'Slave' przy pojedynczym falowniku."
        ),
        "F35 / F42 / F47 / F48 - Błędy napięcia lub częstotliwości sieci AC (GRID)": (
            "CO TO ZNACZY: Brak sieci na zaciskach GRID lub napięcie/częstotliwość poza zakresem normy.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Zmierz napięcia L1-N, L2-N, L3-N na kostce GRID (moment dokręcenia 1.2 Nm) oraz sprawdź kolejność wirowania faz.\n"
            "2. W menu Grid Setting sprawdź wybrany standard (EN50549) oraz aktywuj funkcję V(Q) / V(W)."
        ),
        "F55 / F56 / F58 - Błędy baterii (Za wysokie/niskie napięcie DC / Brak komunikacji BMS)": (
            "CO TO ZNACZY: Rozładowany akumulator (<40V) lub przerwana komunikacja CAN/RS485 z BMS (gdy włączone BMS_Err_Stop).\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Sprawdź śruby M10 akumulatora (moment 24.5 Nm) oraz napięcie DC. Wybudź baterię przyciskiem Black Start / RST.\n"
            "2. Sprawdź kabel RJ45 w porcie BMS 485/CAN (piny 4: CAN-H, 5: CAN-L) oraz pierścień magnetyczny (4 zwoje).\n"
            "3. W Battery Setting zaznacz 'Activate Battery' i w razie awarii kabla tymczasowo odznacz 'BMS_Err_Stop' w Advanced Function."
        ),
        "F64 - Heatsink_High_Temp_Fault (Przegrzanie radiatora)": (
            "CO TO ZNACZY: Temperatura radiatora przekroczyła próg bezpieczeństwa.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Sprawdź drożność wentylatorów i zachowanie odstępów montażowych (500 mm góra/dół/boki). Odczekaj 10 minut przy wyłączonym falowniku."
        )
    },
    "Falownik GoodWe (Seria ET Plus / BT / EH / ES / SDT G2 / XS)": {
        "Isolation Failure / Błąd ISO (Niska rezystancja izolacji DC)": (
            "CO TO ZNACZY: Rezystancja izolacji stringu PV do ziemi spadła poniżej normy (zwykle w deszczowe dni lub przez uszkodzony MC4/SPD).\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Odłącz wyłącznik DC, wypnij stringi i podłączaj pojedynczo, aby wskazać uszkodzony obwód.\n"
            "2. Zmierz rezystancję izolacji do PE oraz napięcie między N a PE na zaciskach AC (musi być < 10V)."
        ),
        "Ground I Failure (Kod 10 - Prąd upływowy doziemny zbyt wysoki)": (
            "CO TO ZNACZY: Nieprawidłowe połączenie N i PE po stronie AC lub upływ z paneli do ziemi.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Zmierz multimetrem napięcie między uziemieniem a ramą falownika (powinno być bliskie 0V).\n"
            "2. Sprawdź poprawność podłączenia żył L, N i PE we wtyczce AC oraz zewnętrzny przewód PE na obudowie (śruba M5, 1.2-2 Nm)."
        ),
        "Vac Failure / Fac Failure / Utility Loss (Kody 09, 17, 29 - Błędy sieci AC)": (
            "CO TO ZNACZY: Zanik sieci AC lub napięcie (>253V) / częstotliwość poza zakresem.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Rozkręć wtyczkę AC i sprawdź dokręcenie tulejek (luźne N powoduje skoki napięcia).\n"
            "2. Połącz się przez SolarGo (hasło: goodwe2010 lub 1111 na LCD), ustaw właściwy kod kraju i włącz regulację mocy biernej Q(U)."
        ),
        "Phase Sequence / Phase Angle Fault (Kody 14ETU, 27ETU - Błąd kolejności faz)": (
            "CO TO ZNACZY: Nieprawidłowe wirowanie faz na wejściu trójfazowym AC.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Wyłącz zabezpieczenie AC i zamień miejscami przewody fazowe L2 i L3 we wtyczce AC falownika."
        ),
        "Relay Check Failure (Kod 07 / 25 - Błąd testu przekaźników AC)": (
            "CO TO ZNACZY: Przekaźniki nie przeszły autotestu (często przez wysokie napięcie N-PE > 10V).\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Zmierz napięcie N-PE w sieci (<10V) oraz rezystancję N-PE na wyjściu Backup (<10 Ohm).\n"
            "2. Zrestartuj falownik (odłącz AC i DC na min. 5 minut)."
        ),
        "Back-Up Over Load / Awaria wentylatora (Kod 12 / 12ETU)": (
            "CO TO ZNACZY: Przeciążenie wyjścia rezerwowego Backup (dioda USTERKA miga pojedynczo) lub zablokowany wentylator.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Zmniejsz obciążenie podłączone pod obwód Backup (odłącz urządzenia o dużym prądzie rozruchowym) i sprawdź wentylatory."
        ),
        "BMS Loss / Konflikt Licznika Smart Meter (Dioda COM miga lub zgaszona)": (
            "CO TO ZNACZY: Brak komunikacji z akumulatorem lub inteligentnym licznikiem GM3000.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Uruchom test licznika w SolarGo / PV Master (pobór z sieci musi wynosić min. 150W, a przekładniki CT A/B/C muszą być na właściwych fazach L1/L2/L3 strzałką w stronę sieci).\n"
            "2. Jeśli BMS zakłóca licznik: odłącz wtyczkę BMS i upewnij się, że kabel BMS wykorzystuje tylko 2 wymagane piny komunikacyjne oraz ma rezystor terminujący 120 Ohm."
        )
    },
    "Pompa Ciepła ARGOM (Seria ARG-PC 6-22 kW, Automatyka CAREL) i Solar": {
        "AL007 / AL137 - Water flow switch / Flow switch alarm (Brak przepływu wody)": (
            "CO TO ZNACZY: Zbyt niski przepływ wody/glikolu przez wymiennik – najczęstsza usterka eksploatacyjna.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. MANOMETR: Dopuść ciśnienie w układzie do 1.5 - 1.8 bar.\n"
            "2. FILTRY: Zakręć zawory i wyczyść filtr siatkowy (skośny) oraz separator magnetyczny z czarnego szlamu.\n"
            "3. ODPOWIETRZENIE: Odpowietrz układ oraz pompę obiegową, upewnij się że min. 2 pętle grzewcze są otwarte na stałe."
        ),
        "AL138 / AL139 / AL140 - High/Low temp. alarm / Temp. delta alarm (Delta T > 8 st.C)": (
            "CO TO ZNACZY: Woda wylotowa >62 st.C (AL138), <5 st.C (AL139 - antyzamarzanie) lub zbyt duża różnica Zasilanie-Powrót >8 st.C (AL140).\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Wyczyść filtr siatkowy i separator magnetyczny (zdławiony przepływ powoduje skok Delty T).\n"
            "2. Zwiększ bieg pompy obiegowej i sprawdź stężenie glikolu refraktometrem (zbyt gęsty glikol <-35 st.C dławi przepływ)."
        ),
        "AL003..AL006 / AL130..AL136 - Błędy czujników temperatury i przetworników ciśnienia": (
            "CO TO ZNACZY: AL003 (powrót wody), AL004 (zasilanie wody), AL005 (otoczenie), AL006 (wężownica), AL130 (tłoczenie), AL131 (ssanie), AL132 (ciśn. HP), AL133 (ciśn. LP), AL134 (zasobnik CWU), AL135/136 (układ EVI).\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Sprawdź wtyczki sondy na płycie sterownika CAREL.\n"
            "2. Zmierz rezystancję czujnika NTC omomierzem. W razie przerwy lub zwarcia wymień czujnik z magazynu w busie."
        ),
        "AL001 / AL002 / AL008 / AL152 - Zasilanie, kolejność faz i pamięć EEPROM CAREL": (
            "CO TO ZNACZY: AL001/AL002 (błąd zapisu pamięci EEPROM), AL008 (błąd kolejności lub zaniku faz w modelach 3F), AL152 (niestabilne napięcie zasilania).\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Przy AL001/AL002: odłącz zasilanie pompy bezpiecznikiem na 3-5 minut i uruchom ponownie.\n"
            "2. Przy AL008: zmierz napięcia międzyfazowe 400V w rozdzielnicy; jeśli wszystkie 3 fazy są obecne, zamień miejscami dwie fazy zasilające (L1 z L2)."
        ),
        "AL009 / AL010 / AL011 / AL012 - Ostrzeżenia o czasie pracy podzespołów (Przegląd)": (
            "CO TO ZNACZY: Przekroczony licznik godzin pracy jednostki (AL009), pompy wody (AL010), sprężarki (AL011) lub wentylatora (AL012).\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Wykonaj pełny przegląd konserwacyjny urządzenia i zresetuj licznik przeglądowy w menu serwisowym sterownika CAREL."
        ),
        "AL013..AL151 / AL035..AL129 / AL153..AL174 - Układ chłodniczy R32, EEV/EVI, Inwerter Power+": (
            "CO TO ZNACZY: Praca sprężarki Panasonic poza obwiednią Envelope (AL035-AL129), błąd zaworu rozprężnego EEV/EVI (AL013-AL151), alarm modułu inwertera Power+ (AL051-AL114) lub wentylatora (AL153-AL174).\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Odłącz zasilanie pompy na pełne 5 minut w celu zresetowania falownika Power+ i pozycjonowania krokowego zaworu EEV.\n"
            "2. Sprawdź czystość parownika (brak oblodzenia) oraz poprawność ilości czynnika R32."
        )
    }
}

POSTAL_PREFIX_COORDS = {
    "00": (52.23, 21.01), "01": (52.25, 20.95), "02": (52.18, 21.00), "03": (52.27, 21.08),
    "04": (52.22, 21.15), "05": (52.12, 20.88), "06": (52.88, 20.61), "07": (52.60, 21.80),
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
