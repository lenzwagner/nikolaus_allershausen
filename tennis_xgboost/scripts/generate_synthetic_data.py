"""
Generate realistic synthetic tennis data for offline testing.
Uses real player names, plausible ELO spreads, and correct schema.
"""
import random
import sys
from pathlib import Path

import numpy as np
import pandas as pd

ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT / "src"))
random.seed(42)
np.random.seed(42)

ATP_PLAYERS = [
    ("Novak Djokovic",     2410, 2380, 2450, 2260),
    ("Carlos Alcaraz",     2390, 2360, 2380, 2430),
    ("Jannik Sinner",      2370, 2410, 2340, 2200),
    ("Alexander Zverev",   2280, 2290, 2270, 2180),
    ("Daniil Medvedev",    2320, 2390, 2180, 2150),
    ("Andrey Rublev",      2200, 2210, 2190, 2080),
    ("Stefanos Tsitsipas", 2230, 2180, 2300, 2100),
    ("Casper Ruud",        2150, 2110, 2280, 1990),
    ("Holger Rune",        2130, 2100, 2150, 2050),
    ("Taylor Fritz",       2120, 2200, 2060, 2090),
    ("Ben Shelton",        2060, 2130, 2010, 1980),
    ("Hubert Hurkacz",     2110, 2070, 2090, 2140),
    ("Alex de Minaur",     2080, 2100, 2050, 2060),
    ("Grigor Dimitrov",    2050, 2070, 2030, 2020),
    ("Sebastian Korda",    2000, 2020, 1970, 1950),
    ("Tommy Paul",         2010, 2050, 1970, 1940),
    ("Francisco Cerundolo",1960, 1910, 2040, 1870),
    ("Lorenzo Musetti",    1980, 1900, 2020, 1970),
    ("Flavio Cobolli",     1890, 1870, 1880, 1840),
    ("Felix Auger-Aliassime",2040,2060,2000,2010),
]

WTA_PLAYERS = [
    ("Iga Swiatek",        2400, 2310, 2500, 2100),
    ("Aryna Sabalenka",    2360, 2390, 2300, 2200),
    ("Coco Gauff",         2250, 2270, 2200, 2190),
    ("Elena Rybakina",     2210, 2240, 2130, 2200),
    ("Jessica Pegula",     2150, 2180, 2090, 2080),
    ("Barbora Krejcikova", 2100, 2050, 2120, 2150),
    ("Marketa Vondrousova",2070, 2030, 2060, 2160),
    ("Ons Jabeur",         2090, 2050, 2080, 2110),
    ("Karolina Muchova",   2060, 2030, 2080, 2030),
    ("Mirra Andreeva",     1980, 1990, 1970, 1960),
    ("Daria Kasatkina",    2010, 1980, 2060, 1940),
    ("Madison Keys",       2030, 2050, 1980, 1990),
    ("Jasmine Paolini",    2000, 1960, 2050, 1980),
    ("Beatriz Haddad Maia",1970, 1940, 2030, 1880),
    ("Caroline Garcia",    2000, 1990, 1960, 2010),
]

SURFACES = ["Hard", "Clay", "Grass"]
SURFACE_WEIGHTS = [0.55, 0.30, 0.15]

TOURNAMENTS = {
    "Hard":  [("Australian Open", "Grand Slam"), ("US Open", "Grand Slam"),
              ("Indian Wells Masters", "Masters"), ("Miami Open", "Masters"),
              ("Canada Masters", "Masters"), ("Cincinnati Masters", "Masters"),
              ("Vienna Open", "ATP500"), ("Basel Open", "ATP500")],
    "Clay":  [("Roland Garros", "Grand Slam"), ("Monte-Carlo Masters", "Masters"),
              ("Madrid Open", "Masters"), ("Rome Masters", "Masters"),
              ("Barcelona Open", "ATP500"), ("Hamburg Open", "ATP250"),
              ("Bastad Open", "ATP250"), ("Gstaad Open", "ATP250")],
    "Grass": [("Wimbledon", "Grand Slam"), ("Queens Club", "ATP500"),
              ("Halle Open", "ATP500"), ("Eastbourne", "ATP250"),
              ("Nottingham Open", "ATP250")],
}

ROUNDS = ["R128", "R64", "R32", "R16", "QF", "SF", "F"]
LEVEL_ROUND_FILTER = {
    "Grand Slam": ROUNDS,
    "Masters":    ROUNDS[2:],   # R32 onwards
    "ATP500":     ROUNDS[3:],   # R16 onwards
    "ATP250":     ROUNDS[4:],   # QF onwards
}


def win_prob(elo_a, elo_b) -> float:
    return 1.0 / (1.0 + 10 ** ((elo_b - elo_a) / 400.0))


def serve_stats(elo: float) -> dict:
    """Generate plausible serve stats correlated with ELO."""
    strength = (elo - 1800) / 600.0  # 0..1 range
    return {
        "first_in":   int(np.clip(np.random.normal(0.64 + 0.05*strength, 0.04), 0.50, 0.78) * 100),
        "first_won":  int(np.clip(np.random.normal(0.71 + 0.06*strength, 0.04), 0.58, 0.85) * 100),
        "second_won": int(np.clip(np.random.normal(0.51 + 0.06*strength, 0.04), 0.38, 0.68) * 100),
        "svpt":       int(np.random.normal(80, 10)),
        "aces":       max(0, int(np.random.normal(4 + 3*strength, 2))),
        "df":         max(0, int(np.random.normal(3 - 1*strength, 1.5))),
        "bp_faced":   max(0, int(np.random.normal(4 - 1.5*strength, 2))),
        "bp_saved":   max(0, int(np.random.normal(2.5 - 0.5*strength, 1.5))),
    }


def generate_matches(players, tour: str, n_matches: int) -> pd.DataFrame:
    rows = []
    start = pd.Timestamp("2019-01-01")
    end = pd.Timestamp.today() - pd.Timedelta(days=7)
    names = [p[0] for p in players]
    elos_ov = {p[0]: p[1] for p in players}
    elos_hard = {p[0]: p[2] for p in players}
    elos_clay = {p[0]: p[3] for p in players}
    elos_grass = {p[0]: p[4] for p in players}

    for _ in range(n_matches):
        surface = random.choices(SURFACES, SURFACE_WEIGHTS)[0]
        tourn_name, level = random.choice(TOURNAMENTS[surface])
        rounds = LEVEL_ROUND_FILTER.get(level, ROUNDS[3:])
        rnd = random.choice(rounds)

        pa, pb = random.sample(names, 2)
        elo_map = {"Hard": elos_hard, "Clay": elos_clay, "Grass": elos_grass}
        elo_a = elo_map[surface][pa]
        elo_b = elo_map[surface][pb]
        p_a_wins = win_prob(elo_a, elo_b)
        won_a = random.random() < p_a_wins

        # Random date
        days = random.randint(0, (end - start).days)
        match_date = start + pd.Timedelta(days=days)

        s_a = serve_stats(elo_a)
        s_b = serve_stats(elo_b)

        rows.append({
            "match_date": match_date,
            "tournament": tourn_name,
            "surface": surface,
            "level": level,
            "round": rnd,
            "tour": tour,
            "player_a": pa,
            "player_b": pb,
            "rank_a": names.index(pa) + 1,
            "rank_b": names.index(pb) + 1,
            "winner": "A" if won_a else "B",
            "aces_a": s_a["aces"],     "aces_b": s_b["aces"],
            "df_a": s_a["df"],         "df_b": s_b["df"],
            "first_in_a": s_a["first_in"],  "first_in_b": s_b["first_in"],
            "first_won_a": s_a["first_won"],"first_won_b": s_b["first_won"],
            "second_won_a": s_a["second_won"],"second_won_b": s_b["second_won"],
            "svpt_a": s_a["svpt"],     "svpt_b": s_b["svpt"],
            "bp_faced_a": s_a["bp_faced"],"bp_faced_b": s_b["bp_faced"],
            "bp_saved_a": s_a["bp_saved"],"bp_saved_b": s_b["bp_saved"],
            "retirement": None, "walkover": None,
        })

    return pd.DataFrame(rows).sort_values("match_date").reset_index(drop=True)


def generate_and_save(n_atp=8000, n_wta=5000):
    out_dir = ROOT / "data" / "processed"
    out_dir.mkdir(parents=True, exist_ok=True)

    print(f"Generating {n_atp} ATP + {n_wta} WTA synthetic matches …")
    atp = generate_matches(ATP_PLAYERS, "atp", n_atp)
    wta = generate_matches(WTA_PLAYERS, "wta", n_wta)
    df = pd.concat([atp, wta], ignore_index=True).sort_values("match_date").reset_index(drop=True)

    # Save as both sackmann and api placeholder
    df.to_parquet(out_dir / "matches_sackmann.parquet", index=False)
    pd.DataFrame(columns=df.columns).to_parquet(out_dir / "matches_api.parquet", index=False)
    print(f"  Saved {len(df)} synthetic matches to data/processed/")
    return df


if __name__ == "__main__":
    generate_and_save()
    print("Done. Run  python scripts/demo_predictions.py  next.")
