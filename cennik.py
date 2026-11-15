# -*- coding: utf-8 -*-
"""
Moduł Cennika Usług i Materiałów PROMAG OZE (cennik.py)
Stawka godzinowa serwisanta dla wycen: 100 zł netto / h
Wzór pozycji: Cena_Materialow + (Czas_Pracy_h * STAWKA_RBH_NETTO)
Do całego zlecenia (gdy wybrano min. 1 czynność płatną) doliczana jest stała kwota za dojazd (RYCZALT_DOJAZD).
"""

STAWKA_RBH_NETTO = 100.0   # 100 zł netto za godzinę pracy serwisanta (do kalkulacji usług i trasy)
KOSZT_PALIWA_LITR = 8.0    # 8 zł / litr paliwa
SR_CZAS_WIZYTY_H = 1.5     # 1.5h średni czas u klienta
RYCZALT_DOJAZD = 200.0     # Stała opłata za dojazd doliczana automatycznie do wyceny klienta

# Zestandaryzowana tabela cennika: Nazwa usługi | Cena materiałów (zł netto) | Czas pracy (h)
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
    mat = float(wiersz.get("cena_materialow", 0.0))
    czas = float(wiersz.get("czas_pracy_h", 0.0))
    return round(mat + (czas * stawka_rbh), 2)

def pobierz_slownik_cennika(tabela=None, stawka_rbh=STAWKA_RBH_NETTO):
    if tabela is None:
        tabela = TABELA_CENNIKA
    out = {}
    for w in tabela:
        nazwa = w["usluga"]
        out[nazwa] = {
            "cena_materialow": float(w.get("cena_materialow", 0.0)),
            "czas_pracy_h": float(w.get("czas_pracy_h", 0.0)),
            "cena_calkowita": wylicz_cene_uslugi(w, stawka_rbh)
        }
    return out

def oblicz_koszyk_uslug(wybrane_uslugi, tabela=None, stawka_rbh=STAWKA_RBH_NETTO, dolicz_dojazd=True, kwota_dojazdu=RYCZALT_DOJAZD):
    """Zwraca słownik z podsumowaniem wyceny na żywo oraz automatycznym tekstem notatki dla serwisanta."""
    slownik = pobierz_slownik_cennika(tabela, stawka_rbh)
    suma_mat = 0.0
    suma_czas = 0.0
    suma_uslug = 0.0
    czynnosci_opis = []

    for u in wybrane_uslugi:
        if u in slownik:
            info = slownik[u]
            suma_mat += info["cena_materialow"]
            suma_czas += info["czas_pracy_h"]
            suma_uslug += info["cena_calkowita"]
            czynnosci_opis.append(f"{u} ({info['cena_calkowita']:.0f} zł)")

    platne_wybrane = [u for u in wybrane_uslugi if "Gwarancyjny" not in u]
    dojazd_naliczony = float(kwota_dojazdu) if (dolicz_dojazd and len(platne_wybrane) > 0) else 0.0
    razem_netto = round(suma_uslug + dojazd_naliczony, 2)

    if czynnosci_opis:
        auto_notatka = "ZAKRES PRAC Z KALKULATORA: " + ", ".join(czynnosci_opis)
        if dojazd_naliczony > 0:
            auto_notatka += f" + Dojazd ({dojazd_naliczony:.0f} zł)."
    else:
        auto_notatka = ""

    return {
        "suma_materialow": round(suma_mat, 2),
        "suma_czasu_h": round(suma_czas, 2),
        "suma_robocizny": round(suma_czas * stawka_rbh, 2),
        "dojazd": round(dojazd_naliczony, 2),
        "razem_dla_klienta": razem_netto,
        "auto_notatka": auto_notatka
          }
  
