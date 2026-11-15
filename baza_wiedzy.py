# -*- coding: utf-8 -*-
"""
Moduł Bazy Wiedzy, Przeszukiwarki PDF/Zdjęć (OCR) oraz Praktycznych Rozwiązań (baza_wiedzy.py)
"""
import os
import re
import urllib.parse

try:
    import pypdf
except ImportError:
    pypdf = None

DIAGNOSTYKA_BAZA = {
    "Falownik Deye (SUN-3K..12K-SG05LP3-EU-SM2 / String)": {
        "F01 / F07 - DC Reverse Polarity / DC START Fault": (
            "CO TO ZNACZY: Odwrotna polaryzacja wejścia DC (PV) lub brak możliwości zbudowania napięcia szyny DC BUS.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Wyłącz rozłącznik DC i zmierz polaryzację każdego stringu PV miernikiem (+ do +, - do -).\n"
            "2. Zrestartuj falownik (wyłącz Baterię -> DC -> AC, odczekaj 2 minuty, włącz ponownie)."
        ),
        "F08 - GFDI_Relay_Failure (Błąd przekaźnika uziemienia / N-PE)": (
            "CO TO ZNACZY: Problem z uziemieniem obudowy lub nieprawidłowy potencjał między N a PE.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Zmierz napięcie AC między szyną N a PE w rozdzielnicy (norma < 5-10V).\n"
            "2. W menu Deye (Advanced Function -> kod 7777) sprawdź opcję 'Signal Island Mode' (przekaźnik zwierający N-PE w trybie wyspowym na porcie LOAD)."
        ),
        "F13 - Working_Mode_Change (Reset trybu pracy / Zmiana sieci)": (
            "CO TO ZNACZY: Zmiana parametrów sieci, przełączenie w tryb Brak Aku lub zmiana System Work Mode.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Odczekaj 3 minuty – zazwyczaj znika automatycznie po autoteście.\n"
            "2. Jeśli wisi: wyłącz DC, Baterię i AC na 2 minuty i uruchom ponownie."
        ),
        "F15 / F18 / F34 - AC Overcurrent Fault (Przetężenie / Zwarcie po stronie AC lub LOAD)": (
            "CO TO ZNACZY: Przekroczenie prądu na wyjściu AC lub obwodzie awaryjnym LOAD (Backup).\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Wyłącz zabezpieczenie obwodu rezerwowego LOAD i sprawdź, czy falownik wstaje bez obciążenia domu.\n"
            "2. Sprawdź odbiorniki pod kątem zwarć L-N / L-PE oraz prądów rozruchowych."
        ),
        "F20 / F21 / F55 - DC Overcurrent / HV Bus Overvoltage": (
            "CO TO ZNACZY: Zbyt wysoki prąd DC (PV/Bateria) lub przekroczone napięcie szyny DC BUS (Voc > 800V lub napięcie baterii za wysokie).\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Zmierz napięcie rozwarcia stringów PV (Voc nie może przekraczać 800V!).\n"
            "2. W trybie Off-Grid zmniejsz duże obciążenie startowe i sprawdź limity prądu baterii w Battery Setting."
        ),
        "F22 - Tz_EmergStop_Fault (Aktywowany wyłącznik PPOŻ / Brak zwory RSD)": (
            "CO TO ZNACZY: Otwarty obwód zatrzymania awaryjnego RSD / zdalnego wyłączenia.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Sprawdź złącze RSD_input wewnątrz komory komunikacyjnej.\n"
            "2. Jeśli budynek nie ma zewnętrznego przycisku PPOŻ, upewnij się, że fabryczna zworka jest dokręcona."
        ),
        "F16 / F23 / F24 - GFCI / DC Insulation Fault (Prąd upływu / Błąd izolacji DC)": (
            "CO TO ZNACZY: Przebicie prądu z paneli PV lub kabli do uziemienia (wilgoć, uszkodzony MC4 lub SPD DC).\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Wypnij stringi i zmierz rezystancję izolacji Riso megaomomierzem (norma >= 1 MOhm) lub napięcie biegunów względem PE.\n"
            "2. Sprawdź wkładki ochronnika przepięć SPD DC w rozdzielnicy oraz ciągłość przewodu PE obudowy."
        ),
        "F26 - DC Busbar Unbalanced (Niezrównoważona szyna DC)": (
            "CO TO ZNACZY: Bardzo duża asymetria obciążenia między 3 fazami lub upływ DC.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Włącz w menu Advanced Function (kod 7777) opcję 'Asymmetric phase feeding'.\n"
            "2. Zrestartuj falownik 2-3 razy."
        ),
        "F29 / F41 - Parallel CAN Bus Fault / Parallel System Stop": (
            "CO TO ZNACZY: Błąd komunikacji równoległej (kaskady) lub wyłączenie jednego z falowników w kaskadzie.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Przy pojedynczym falowniku: wejdź w Advanced Function (kod 7777) i odznacz 'Parallel' (Master/Slave).\n"
            "2. W kaskadzie: sprawdź kable RJ45 w portach Parallel A/B oraz unikalne adresy Modbus SN."
        ),
        "F35 / F42 / F47 / F48 - Brak sieci AC / Napięcie lub Częstotliwość poza zakresem": (
            "CO TO ZNACZY: Brak napięcia na kostce GRID lub parametry sieci poza normą.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Zmierz napięcia L1-N, L2-N, L3-N oraz międzyfazowe (400V) na listwie GRID.\n"
            "2. Sprawdź poprawność wirowania faz oraz kod sieci w Grid Setting (EN50549)."
        ),
        "F56 / F58 - DC_VoltLow / BMS_Comm_Fault (Rozładowana bateria / Brak komunikacji BMS)": (
            "CO TO ZNACZY: Napięcie baterii poniżej progu odcięcia (<40-44V) lub przerwana komunikacja CAN/RS485 z BMS.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Wybudź baterię (Black Start - przytrzymaj przycisk SW/RST na magazynie przez 5s).\n"
            "2. W menu Battery Setting włącz 'Activate Battery' oraz wymuś doładowanie z sieci ('Grid Charge').\n"
            "3. Sprawdź kabel RJ45 w porcie BMS 485/CAN (piny 4-CAN_H, 5-CAN_L) i kod protokołu Lithium Mode (np. 00)."
        ),
        "F63 / F64 - ARC Fault / High Heatsink Temp (Łuk DC / Przegrzanie radiatora)": (
            "CO TO ZNACZY: Wykrycie łuku elektrycznego DC (F63) lub przegrzanie radiatora >85C (F64).\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Dla F63: sprawdź złącza MC4 pod kątem upalenia i wyczyść błąd w Advanced Function ('Clear Arc Fault').\n"
            "2. Dla F64: oczyść radiator i wentylatory, odczekaj 10 minut po wyłączeniu."
        )
    },
    "Falownik GoodWe (ET / ET PLUS+ / BT / EH / ES / SDT G2 / XS)": {
        "Isolation Failure / Błąd ISO (Rezystancja izolacji DC < 30-100 kOhm)": (
            "CO TO ZNACZY: Zwarcie stringu PV do ziemi lub zawilgocenie złączy MC4.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Odłącz stringi i podłączaj pojedynczo, aby wskazać uszkodzony obwód.\n"
            "2. Zmierz rezystancję izolacji DC względem PE oraz sprawdź napięcie między N a PE na AC (musi być < 10V)."
        ),
        "Ground I Failure (Kod 10) / Relay Check Failure (Kod 25)": (
            "CO TO ZNACZY: Zbyt wysoki prąd upływu doziemnego lub błąd testu przekaźników AC.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Zmierz multimetrem napięcie między uziemieniem (ramą) a przewodem N – musi być poniżej 10V.\n"
            "2. Sprawdź poprawność podłączenia N i PE w złączu AC oraz na wyjściu Backup."
        ),
        "Vac Failure / Fac Failure / Utility Loss (Kody 09, 17, 29)": (
            "CO TO ZNACZY: Brak sieci AC lub napięcie (>253V) / częstotliwość poza zakresem.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Zmierz napięcie na wtyczce AC i dokręć zaciski N oraz fazowe.\n"
            "2. W aplikacji SolarGo / PV Master (hasło: goodwe2010) sprawdź Safety Country (Polska EN50549) i włącz regulację mocy biernej Q(U)."
        ),
        "Błąd kolejności faz (Kod 14ETU) / Błąd kąta fazowego (Kod 27ETU)": (
            "CO TO ZNACZY: Nieprawidłowe wirowanie faz na wejściu AC falownika trójfazowego.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Wyłącz zabezpieczenie AC i zamień miejscami przewody L2 i L3 we wtyczce AC falownika."
        ),
        "BMS Loss / Błąd komunikacji z akumulatorem / Błąd testu licznika Smart Meter": (
            "CO TO ZNACZY: Brak komunikacji z magazynem energii lub konflikt na magistrali licznika.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Sprawdź w SolarGo/PV Master wybór modelu akumulatora oraz rezystor terminujący 120 Ohm.\n"
            "2. Jeśli licznik nie przechodzi testu: upewnij się, że pobór z sieci > 150W, kierunek przekładników CT (strzałka House->Grid) jest właściwy."
        )
    },
    "Pompa Ciepła ARGOM (seria ARG-PC, automatyka CAREL, kody AL001 - AL174)": {
        "AL007 / AL137 - Water flow switch / Flow switch alarm (Brak przepływu wody)": (
            "CO TO ZNACZY: Zadziałał czujnik przepływu – zbyt mały przepływ wody/glikolu przez wymiennik.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Sprawdź ciśnienie na manometrze (dopuść do 1.5 - 1.8 bar) i odpowietrz instalację.\n"
            "2. Wyczyść filtr siatkowy (skośny) oraz separator magnetyczny ze szlamu!\n"
            "3. Sprawdź pracę pompy obiegowej i upewnij się, że zawory/pętle CO są otwarte."
        ),
        "AL138 / AL139 / AL140 - High temp (>62C) / Low temp (<5C) / Temp.delta alarm (Delta T >8C)": (
            "CO TO ZNACZY: Przekroczenie temperatur granicznych wody wylotowej lub zbyt duża różnica Zasilanie-Powrót.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Wyczyść filtr wody i odpowietrz układ – zdławiony przepływ powoduje skok Delty T powyżej 8 st.C.\n"
            "2. Zwiększ bieg pompy obiegowej i sprawdź stężenie glikolu."
        ),
        "AL003 / AL004 / AL005 / AL006 / AL134 - Błędy czujników temp. wody, otoczenia, wymiennika i CWU": (
            "CO TO ZNACZY: AL003=powrót (Inlet), AL004=zasilanie (Outlet), AL005=zewnętrzny (Ambient), AL006=wężownica (Condenser), AL134=zasobnik CWU.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Sprawdź podłączenie wtyczek czujników na płycie sterownika CAREL.\n"
            "2. Zmierz rezystancję sondy omomierzem (przerwa lub zwarcie oznacza konieczność wymiany czujnika NTC)."
        ),
        "AL008 / AL152 - Phase sequ.prot.alarm / Supply power error (Błąd faz lub zasilania)": (
            "CO TO ZNACZY: Błąd kolejności/zaniku faz (w modelach 3-fazowych) lub niestabilne napięcie sieci.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Zmierz napięcia fazowe w rozdzielnicy (sprawdź bezpieczniki).\n"
            "2. W modelach 3-fazowych zamień miejscami dwie fazy na listwie zasilającej."
        )
    },
    "Kolektory Słoneczne (Solar) / Sterowniki Tech / Geco / Hewalex": {
        "Zapowietrzenie układu / Duża Delta T (T1 > 110C, T2 < 35C)": (
            "CO TO ZNACZY: Pompa solarna pracuje, ale glikol nie krąży (poduszka powietrzna na dachu lub zużyty glikol).\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Podłącz pompę serwisową i przepłucz układ pod ciśnieniem do całkowitego usunięcia powietrza.\n"
            "2. Sprawdź ciśnienie wstępne naczynia przeponowego (zwykle 1.5 - 2.0 bar) oraz drożność grupy pompowej."
        ),
        "Uszkodzony czujnik temperatury kolektora PT1000 / NTC (Błąd czujnika T1/T2)": (
            "CO TO ZNACZY: Przerwana żyła przewodu na dachu (np. przez ptaki/kuny) lub uszkodzona sonda po przegrzaniu.\n"
            "PROCEDURA KROK PO KROKU:\n"
            "1. Odepnij czujnik ze sterownika i zmierz rezystancję (PT1000 ma ~1000 Ohm w 0C i ~1078 Ohm w 20C).\n"
            "2. Wymień czujnik na wysokotemperaturowy (silikonowy) i zabezpiecz przewód w peszlu UV."
        )
    }
}

PRAKTYCZNE_WSKAZOWKI_FORUM = [
    {
        "tagi": ["f08", "deye", "gfdi", "uziemienie", "n-pe", "load", "wyspa"],
        "tytul": "Praktyka z terenu: Błąd F08 w Deye przy przełączaniu na Backup",
        "opis": "Najczęstsza przyczyna na polskich instalacjach: brak zaznaczonej opcji 'Signal Island Mode' w menu Advanced Function (kod 7777) lub luźny przewód PE obudowy. Gdy sieć zanika, falownik rozłącza przekaźnik sieciowy i traci uziemienie punktu N – włączenie Signal Island Mode zwiera wewnętrznie N z PE w trybie wyspowym."
    },
    {
        "tagi": ["al007", "al137", "argom", "przepływ", "flow", "e01", "kolejność faz", "delta"],
        "tytul": "Praktyka z terenu: Argom AL007 / AL137 mimo czystego filtra",
        "opis": "Jeśli filtr siatkowy i magnetyczny są czyste, a nadal wybija AL007: 1) Sprawdź czy zawór 3-drogowy CWU/CO nie zablokował się w pozycji pośredniej, 2) Sprawdź odpowietrznik na górze wymiennika płytowego wewnątrz jednostki, 3) Przy świeżym montażu 3-fazowym sprawdź kolejność faz (E01/AL008)."
    },
    {
        "tagi": ["goodwe", "smart meter", "licznik", "ct", "test", "bms", "faza"],
        "tytul": "Praktyka z terenu: GoodWe ET/BT nie przechodzi testu licznika (Meter Test Fail)",
        "opis": "Trzy klasyczne pułapki: 1) W domu podczas testu musi być włączony odbiornik o mocy min. 150-200W na każdej fazie (np. czajnik/farelka), 2) Przekładniki CT muszą być założone na tych samych fazach L1-L1, L2-L2, L3-L3 co zasilanie falownika i samego licznika, 3) Strzałka na przekładniku CT musi wskazywać kierunek KUPOWANIA energii (w stronę sieci/licznika operatora)."
    },
    {
        "tagi": ["solar", "glikol", "zapowietrzenie", "t1", "naczynie", "pompka", "cieknie"],
        "tytul": "Praktyka z terenu: Solary zapowietrzają się ponownie po tygodniu od serwisu",
        "opis": "Jeśli po odpowietrzeniu układ znów staje po kilku dniach: 1) Dziurawa membrana w naczyniu przeponowym (po naciśnięciu wentyla samochodowego na dole naczynia leci glikol zamiast powietrza), 2) Nieszczelny automatyczny odpowietrznik na dachu zasysa powietrze gdy układ stygnie w nocy (warto zakręcić zaworek kulowy pod odpowietrznikiem po odpowietrzeniu!)."
    },
    {
        "tagi": ["f56", "f58", "bateria", "magazyn", "rozładowany", "black start", "deye"],
        "tytul": "Praktyka z terenu: Martwy magazyn energii (błąd F56 / F58 w Deye)",
        "opis": "Gdy bateria zjedzie poniżej 42V i BMS odetnie wyjście: odłącz obciążenie LOAD, wejdź w Battery Setting -> zaznacz 'Activate Battery' oraz w Time of Use zaznacz 'Grid Charge' na 100%. Przytrzymaj przycisk RST/SW na baterii przez 5 sekund aż zamrugają diody RUN – falownik od razu złapie napięcie i zacznie ładować z sieci."
    }
]

INSTRUKCJE_KONFIGURACJI_TEXT = """
### 📶 1. Podłączenie do Internetu (Wi-Fi 2.4 GHz) i Aplikacje Mobilne
* **Falowniki Deye (Deye Cloud / SOLARMAN Smart):**
  1. Upewnij się, że router klienta nadaje w pasmie **2.4 GHz**.
  2. Połącz telefon z siecią Wi-Fi modułu loggera: **`AP_XXXXXXXXX`** (gdzie X to S/N loggera). Domyślne hasło: **`12345678`**.
  3. Otwórz przeglądarkę i wejdź na **`http://10.10.100.254`** (login: `admin`, hasło: `admin`).
  4. W zakładce **Wizard / Wi-Fi STA** wybierz domową sieć Wi-Fi, wpisz hasło i zrestartuj moduł.
* **Falowniki GoodWe (SolarGo / PV Master / SEMS Portal):**
  1. Naciśnij krótko przycisk Wi-Fi Reload na falowniku (lub użyj menu LCD `W/L Reset`).
  2. Połącz telefon z siecią **`Solar-WiFiXXXXXXXX`**. Hasło domyślne: **`12345678`**.
  3. Otwórz aplikację **SolarGo** lub **PV Master** (hasło instalatora: **`goodwe2010`**) albo wejdź na **`http://10.10.100.253`**.

### 🛠️ 2. Rutynowa Konserwacja i Bezpieczeństwo (Standard Serwisowy)
* **Zasada 5 minut:** Przed pracami wyłącz AC, DC i baterię oraz **odczekaj min. 5 minut**.
* **Momenty dokręcania:** Deye śruby aku M10: **24,5 Nm**, zaciski AC: **1,2 Nm**.
* **Hasła serwisowe Deye (Ekran LCD):** Ustawienia zaawansowane: **`7777`** | Reset fabryczny: **`9999`**.
"""

POSTAL_PREFIX_COORDS = {
    "00": (52.23, 21.01), "01": (52.25, 20.95), "02": (52.18, 21.00), "03": (52.27, 21.08),
    "04": (52.22, 21.15), "05": (52.08, 20.85), "06": (52.88, 20.61), "07": (52.60, 21.80),
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

def przeszukaj_lokalne_instrukcje(zapytanie, folder_sciezka="."):
    """Przeszukuje wszystkie pliki PDF i TXT w katalogu oraz w podfolderze instrukcje/."""
    q = str(zapytanie).strip().lower()
    if len(q) < 2:
        return []
    wyniki = []
    for urzadzenie, bledy in DIAGNOSTYKA_BAZA.items():
        for kod, opis in bledy.items():
            if q in kod.lower() or q in opis.lower() or q in urzadzenie.lower():
                wyniki.append({
                    "zrodlo": f"📘 Baza Wewnętrzna: {urzadzenie}",
                    "naglowek": kod,
                    "fragment": opis
                })

    search_dirs = [folder_sciezka]
    inst_sub = os.path.join(folder_sciezka, "instrukcje")
    if os.path.isdir(inst_sub):
        search_dirs.append(inst_sub)

    if pypdf is not None:
        for s_dir in search_dirs:
            try:
                for fname in os.listdir(s_dir):
                    if fname.lower().endswith(".pdf"):
                        fpath = os.path.join(s_dir, fname)
                        try:
                            reader = pypdf.PdfReader(fpath)
                            for p_idx, page in enumerate(reader.pages):
                                txt = page.extract_text() or ""
                                if q in txt.lower():
                                    pos = txt.lower().find(q)
                                    start = max(0, pos - 140)
                                    end = min(len(txt), pos + 260)
                                    snippet = re.sub(r'\s+', ' ', txt[start:end]).strip()
                                    wyniki.append({
                                        "zrodlo": f"📄 Plik PDF: {fname}",
                                        "naglowek": f"Strona {p_idx + 1}",
                                        "fragment": f"...{snippet}..."
                                    })
                                    if len(wyniki) >= 15:
                                        return wyniki
                        except Exception:
                            pass
            except Exception:
                pass
    return wyniki

def przeszukaj_praktyke_i_internet(zapytanie):
    """Zwraca praktyczne porady z terenu oraz bezpośrednie linki do wyszukiwania w Google / Forach OZE."""
    q = str(zapytanie).strip().lower()
    trafienia = []
    for p in PRAKTYCZNE_WSKAZOWKI_FORUM:
        if any(t in q for t in p["tagi"]) or q in p["tytul"].lower() or q in p["opis"].lower():
            trafienia.append(p)
    encoded_q = urllib.parse.quote_plus(f"{zapytanie} serwis naprawa forum OZE")
    encoded_yt = urllib.parse.quote_plus(f"{zapytanie} fault repair tutorial")
    link_google = f"https://www.google.com/search?q={encoded_q}"
    link_elektroda = f"https://www.google.com/search?q=site:elektroda.pl+{encoded_q}"
    link_yt = f"https://www.youtube.com/results?search_query={encoded_yt}"
    return trafienia, link_google, link_elektroda, link_yt
