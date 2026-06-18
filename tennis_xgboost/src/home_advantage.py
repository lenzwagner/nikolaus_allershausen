"""Phase 6b: Home advantage feature.

A player is considered "at home" when their nationality matches the host country
of the current tournament. Effect is real in tennis: Nadal at Roland Garros,
Djokovic at Australian Open, local wildcards reaching late rounds, etc.

Feature added: home_advantage_diff = is_home_a − is_home_b
  +1  → player A is playing in their home country
   0  → neither (or both, very rare)
  −1  → player B is playing in their home country
"""
import logging
from pathlib import Path

import pandas as pd
from rapidfuzz import process as fz_process

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

# ── Player nationality (IOC / ISO-3 codes) ───────────────────────────────────
PLAYER_NATIONALITY: dict[str, str] = {
    # ATP
    "Novak Djokovic": "SRB",
    "Carlos Alcaraz": "ESP",
    "Jannik Sinner": "ITA",
    "Alexander Zverev": "GER",
    "Daniil Medvedev": "RUS",
    "Andrey Rublev": "RUS",
    "Stefanos Tsitsipas": "GRE",
    "Casper Ruud": "NOR",
    "Holger Rune": "DEN",
    "Taylor Fritz": "USA",
    "Ben Shelton": "USA",
    "Hubert Hurkacz": "POL",
    "Alex de Minaur": "AUS",
    "Grigor Dimitrov": "BUL",
    "Sebastian Korda": "USA",
    "Tommy Paul": "USA",
    "Francisco Cerundolo": "ARG",
    "Lorenzo Musetti": "ITA",
    "Flavio Cobolli": "ITA",
    "Felix Auger-Aliassime": "CAN",
    "Rafael Nadal": "ESP",
    "Roger Federer": "SUI",
    "Andy Murray": "GBR",
    "Stan Wawrinka": "SUI",
    "Dominic Thiem": "AUT",
    "Gael Monfils": "FRA",
    "Jo-Wilfried Tsonga": "FRA",
    "Richard Gasquet": "FRA",
    "Lucas Pouille": "FRA",
    "Adrian Mannarino": "FRA",
    "Ugo Humbert": "FRA",
    "Arthur Fils": "FRA",
    "Giovanni Mpetshi Perricard": "FRA",
    "Cameron Norrie": "GBR",
    "Dan Evans": "GBR",
    "Jack Draper": "GBR",
    "Nick Kyrgios": "AUS",
    "John Millman": "AUS",
    "Thanasi Kokkinakis": "AUS",
    "Chris Eubanks": "USA",
    "Frances Tiafoe": "USA",
    "Reilly Opelka": "USA",
    "Jenson Brooksby": "USA",
    "Denis Shapovalov": "CAN",
    "Milos Raonic": "CAN",
    "Vasek Pospisil": "CAN",
    "Pablo Carreno Busta": "ESP",
    "Roberto Bautista Agut": "ESP",
    "David Ferrer": "ESP",
    "Fernando Verdasco": "ESP",
    "Pedro Martinez": "ESP",
    "Alejandro Davidovich Fokina": "ESP",
    "Albert Ramos-Vinolas": "ESP",
    "Karen Khachanov": "RUS",
    "Aslan Karatsev": "RUS",
    "Emil Ruusuvuori": "FIN",
    "Laslo Djere": "SRB",
    "Borna Coric": "CRO",
    "Marin Cilic": "CRO",
    "Botic van de Zandschulp": "NED",
    "Tallon Griekspoor": "NED",
    "Robin Haase": "NED",
    "Juan Martin del Potro": "ARG",
    "Diego Schwartzman": "ARG",
    "Guido Pella": "ARG",
    "Nishioka Yoshihito": "JPN",
    "Kei Nishikori": "JPN",
    "Taro Daniel": "JPN",
    "Mackenzie McDonald": "USA",
    "Jiri Lehecka": "CZE",
    "Tomas Machac": "CZE",
    # WTA
    "Iga Swiatek": "POL",
    "Aryna Sabalenka": "BLR",
    "Coco Gauff": "USA",
    "Elena Rybakina": "KAZ",
    "Jessica Pegula": "USA",
    "Barbora Krejcikova": "CZE",
    "Marketa Vondrousova": "CZE",
    "Ons Jabeur": "TUN",
    "Karolina Muchova": "CZE",
    "Mirra Andreeva": "RUS",
    "Daria Kasatkina": "RUS",
    "Madison Keys": "USA",
    "Jasmine Paolini": "ITA",
    "Beatriz Haddad Maia": "BRA",
    "Caroline Garcia": "FRA",
    "Simona Halep": "ROU",
    "Petra Kvitova": "CZE",
    "Karolina Pliskova": "CZE",
    "Garbine Muguruza": "ESP",
    "Sloane Stephens": "USA",
    "Sofia Kenin": "USA",
    "Jennifer Brady": "USA",
    "Danielle Collins": "USA",
    "Emma Raducanu": "GBR",
    "Katie Boulter": "GBR",
    "Johanna Konta": "GBR",
    "Harriet Dart": "GBR",
    "Ash Barty": "AUS",
    "Ajla Tomljanovic": "AUS",
    "Storm Hunter": "AUS",
    "Alycia Parks": "USA",
    "Clara Tauson": "DEN",
    "Amanda Anisimova": "USA",
    "Anett Kontaveit": "EST",
    "Kaja Juvan": "SLO",
    "Tamara Zidansek": "SLO",
    "Elina Svitolina": "UKR",
    "Marta Kostyuk": "UKR",
    "Victoria Azarenka": "BLR",
    "Jelena Ostapenko": "LAT",
    "Anastasia Sevastova": "LAT",
    "Sara Sorribes Tormo": "ESP",
    "Nuria Parrizas Diaz": "ESP",
    "Paula Badosa": "ESP",
    "Maria Sakkari": "GRE",
    "Caroline Wozniacki": "DEN",
    "Naomi Osaka": "JPN",
    "Bianca Andreescu": "CAN",
    "Leylah Fernandez": "CAN",
    "Belinda Bencic": "SUI",
    "Viktorija Golubic": "SUI",
    "Sorana Cirstea": "ROU",
    "Irina-Camelia Begu": "ROU",
    "Ana Bogdan": "ROU",
}

# ── Tournament → host country ────────────────────────────────────────────────
TOURNAMENT_COUNTRY: dict[str, str] = {
    # Grand Slams
    "Australian Open": "AUS",
    "Roland Garros": "FRA",
    "French Open": "FRA",
    "Wimbledon": "GBR",
    "US Open": "USA",
    # ATP Masters / WTA Premier Mandatory
    "Indian Wells Masters": "USA", "BNP Paribas Open": "USA",
    "Miami Open": "USA",
    "Monte-Carlo Masters": "MON",
    "Madrid Open": "ESP",
    "Rome Masters": "ITA", "Internazionali BNL d'Italia": "ITA",
    "Canada Masters": "CAN", "National Bank Open": "CAN",
    "Cincinnati Masters": "USA", "Western & Southern Open": "USA",
    "Shanghai Masters": "CHN",
    "Paris Masters": "FRA", "Rolex Paris Masters": "FRA",
    "Vienna Open": "AUT",
    # ATP 500
    "Queens Club": "GBR", "Cinch Championships": "GBR",
    "Halle Open": "GER", "Terra Wortmann Open": "GER",
    "Barcelona Open": "ESP",
    "Basel Open": "SUI", "Swiss Indoors": "SUI",
    "Tokyo": "JPN", "Rakuten Japan Open": "JPN",
    "Beijing": "CHN",
    "Hamburg Open": "GER",
    "Washington": "USA", "Mubadala Citi DC Open": "USA",
    "Rotterdam Open": "NED",
    "Rio Open": "BRA",
    "Dubai Tennis Championships": "UAE",
    "Acapulco": "MEX",
    # ATP 250
    "Eastbourne": "GBR", "Rothesay International Eastbourne": "GBR",
    "Nottingham Open": "GBR",
    "Bastad Open": "SWE", "SkiStar Swedish Open": "SWE",
    "Gstaad Open": "SUI",
    "Kitzbuhel": "AUT",
    "Umag": "CRO", "Croatia Open": "CRO",
    "Geneva": "SUI",
    "Lyon": "FRA",
    "Marrakech": "MAR",
    "Buenos Aires": "ARG",
    "Santiago": "CHI",
    "Estoril": "POR",
    "Munich": "GER",
    "Stuttgart": "GER",
    "Cologne": "GER",
    "Metz": "FRA",
    "Antwerp": "BEL",
    "Stockholm": "SWE",
    "Moscow": "RUS",
    "Sofia": "BUL",
    "Nur-Sultan": "KAZ",
    "Astana": "KAZ",
    "Adelaide": "AUS",
    "Auckland": "NZL",
    "Sydney": "AUS", "United Cup Sydney": "AUS",
    "Brisbane": "AUS",
    "Doha": "QAT",
    "Dallas": "USA",
    "Delray Beach": "USA",
    "Houston": "USA",
    "Atlanta": "USA",
    "Newport": "USA",
    "Los Cabos": "MEX",
    "Winston-Salem": "USA",
    "San Diego": "USA",
    "Tel Aviv": "ISR",
    "Gijon": "ESP",
    "Malaga": "ESP",
    "Florence": "ITA",
    "Napoli": "ITA",
    "Parma": "ITA",
    "Belgrade": "SRB",
    "Bucharest": "ROU",
    "Munich Indoor": "GER",
}


def _get_nationality(player: str) -> str | None:
    """Look up player nationality with fuzzy fallback."""
    # Direct lookup
    for name, nat in PLAYER_NATIONALITY.items():
        if player.lower() == name.lower():
            return nat
    # Fuzzy match
    names = list(PLAYER_NATIONALITY.keys())
    result = fz_process.extractOne(player, names, score_cutoff=75)
    if result:
        return PLAYER_NATIONALITY[result[0]]
    return None


def _get_tournament_country(tournament: str) -> str | None:
    """Look up tournament host country with fuzzy fallback."""
    for name, country in TOURNAMENT_COUNTRY.items():
        if tournament.lower() == name.lower():
            return country
    names = list(TOURNAMENT_COUNTRY.keys())
    result = fz_process.extractOne(tournament, names, score_cutoff=70)
    if result:
        return TOURNAMENT_COUNTRY[result[0]]
    return None


def add_home_advantage(df: pd.DataFrame) -> pd.DataFrame:
    """Add home_advantage_diff = is_home_a - is_home_b to the dataframe."""
    log.info("Adding home_advantage_diff feature")
    df = df.copy()

    is_home_a = []
    is_home_b = []

    for _, row in df.iterrows():
        tournament = str(row.get("tournament", ""))
        pa = str(row.get("player_a", ""))
        pb = str(row.get("player_b", ""))

        tourn_country = _get_tournament_country(tournament)
        nat_a = _get_nationality(pa)
        nat_b = _get_nationality(pb)

        is_home_a.append(1 if (tourn_country and nat_a and tourn_country == nat_a) else 0)
        is_home_b.append(1 if (tourn_country and nat_b and tourn_country == nat_b) else 0)

    df["is_home_a"] = is_home_a
    df["is_home_b"] = is_home_b
    df["home_advantage_diff"] = df["is_home_a"] - df["is_home_b"]

    n_home = df["home_advantage_diff"].abs().sum()
    log.info(f"  home_advantage_diff: {n_home} matches with a home player "
             f"({n_home/len(df)*100:.1f}%)")
    return df
