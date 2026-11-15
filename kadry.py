# -*- coding: utf-8 -*-
"""
Moduł Kadrowy, Flotowy i Magazynów Samochodowych PROMAG OZE (kadry.py)
Zarządza profilami serwisantów, narzędziami, stanami materiałowymi na busach,
rozliczeniami gotówki oraz progami prowizyjnymi.
"""
import datetime

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

def get_default_tech_profile(name, pin, role_desc, umowa, base_sal):
    return {
        "pin": pin,
        "pelne_dane": f"{name} - {role_desc}",
        "telefon": "",
        "umowa": umowa,
        "wynagrodzenie_bazowe": base_sal,
        "bhp_data": "2026-01-15",
        "bhp_protokol": "Ważne",
        "medycyna_data": "2026-01-10",
        "medycyna_zakres": "Praca pow. 3m, SEP E+D, kat. B",
        "urlopy_l4": [],
        "notatki_poufne": "",
        "narzedzia": [dict(x) for x in DOMYSLNE_NARZEDZIA],
        "materialy": [dict(x) for x in DOMYSLNE_MATERIALY],
        "flota": {
            "nr_rej": "WPR 00000",
            "przebieg_km": 120000,
            "limit_olej_km": 130000,
            "data_oc_przeglad": "2027-05-01"
        }
    }

DEFAULT_TECHS = {
    "Leszek": get_default_tech_profile("Leszek", "4829", "Serwisant OZE", "Umowa o pracę", 6500.0),
    "Mikołaj": get_default_tech_profile("Mikołaj", "7315", "Serwisant OZE", "B2B", 7000.0),
    "Wojtek": get_default_tech_profile("Wojtek", "5942", "Serwisant OZE", "Umowa o pracę", 6000.0),
    "Piotr": get_default_tech_profile("Piotr", "3190", "Serwisant OZE", "Umowa o pracę", 6000.0),
    "Stanisław": get_default_tech_profile("Stanisław", "8164", "Serwisant OZE", "Umowa zlecenie", 6000.0)
}

def ensure_tech_profile_fields(profiles):
    for t_name, p in profiles.items():
        if "narzedzia" not in p or not isinstance(p["narzedzia"], list):
            p["narzedzia"] = [dict(x) for x in DOMYSLNE_NARZEDZIA]
        if "materialy" not in p or not isinstance(p["materialy"], list):
            p["materialy"] = [dict(x) for x in DOMYSLNE_MATERIALY]
        if "flota" not in p or not isinstance(p["flota"], dict):
            p["flota"] = {
                "nr_rej": "WPR 00000",
                "przebieg_km": 120000,
                "limit_olej_km": 130000,
                "data_oc_przeglad": "2027-05-01"
            }
    return profiles

def get_low_stock_items(tech_profile):
    """Zwraca listę materiałów, których stan na samochodzie wynosi <= 1 szt."""
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
        "total_closed": total_closed,
        "ftfr_pct": ftfr_pct,
        "paid_count": paid_count,
        "warranty_count": warranty_count,
        "leads_count": leads_count,
        "cash_in_hand": cash_in_hand,
        "monthly_revenue": monthly_revenue,
        "monthly_parts_cost": monthly_parts_cost,
        "missing_tier_1": missing_tier_1,
        "missing_tier_2": missing_tier_2,
        "commission_pln": round(commission_pln, 2),
        "active_rate_str": active_rate_str,
        "progress_pct": progress_pct
    }
