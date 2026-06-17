"""Phase 1: Fetch Jeff Sackmann ATP/WTA CSVs from GitHub for ELO warmup."""
import logging
from pathlib import Path

import pandas as pd
import yaml

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

ROOT = Path(__file__).parent.parent
OUTPUT = ROOT / "data" / "processed" / "matches_sackmann.parquet"

SACKMANN_ATP_URL = "https://raw.githubusercontent.com/JeffSackmann/tennis_atp/master/atp_matches_{year}.csv"
SACKMANN_WTA_URL = "https://raw.githubusercontent.com/JeffSackmann/tennis_wta/master/wta_matches_{year}.csv"

START_YEAR = 2010

COLUMN_MAP = {
    "tourney_date": "match_date",
    "tourney_name": "tournament",
    "surface": "surface",
    "tourney_level": "level",
    "round": "round",
    "winner_name": "player_a",
    "loser_name": "player_b",
    "winner_rank": "rank_a",
    "loser_rank": "rank_b",
    "w_ace": "aces_a",
    "l_ace": "aces_b",
    "w_df": "df_a",
    "l_df": "df_b",
    "w_1stWon": "first_won_a",
    "l_1stWon": "first_won_b",
    "w_1stIn": "first_in_a",
    "l_1stIn": "first_in_b",
    "w_2ndWon": "second_won_a",
    "l_2ndWon": "second_won_b",
    "w_bpFaced": "bp_faced_a",
    "l_bpFaced": "bp_faced_b",
    "w_bpSaved": "bp_saved_a",
    "l_bpSaved": "bp_saved_b",
    "w_svpt": "svpt_a",
    "l_svpt": "svpt_b",
}


def load_tour(url_template: str, tour: str, start_year: int) -> pd.DataFrame:
    import datetime
    end_year = datetime.date.today().year
    frames = []
    for year in range(start_year, end_year + 1):
        url = url_template.format(year=year)
        try:
            df = pd.read_csv(url, low_memory=False)
            df["tour"] = tour
            frames.append(df)
            log.info(f"Loaded {tour} {year}: {len(df)} rows")
        except Exception as e:
            log.warning(f"Failed to load {tour} {year}: {e}")
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def normalize(df: pd.DataFrame) -> pd.DataFrame:
    df = df.rename(columns={k: v for k, v in COLUMN_MAP.items() if k in df.columns})
    if "match_date" in df.columns:
        df["match_date"] = pd.to_datetime(df["match_date"].astype(str), format="%Y%m%d", errors="coerce")
    if "surface" in df.columns:
        df["surface"] = df["surface"].str.capitalize().str.strip()
    df["winner"] = "A"  # In Sackmann, winner_name is always player_a
    return df


def fetch_all(force: bool = False) -> pd.DataFrame:
    if OUTPUT.exists() and not force:
        log.info(f"Skipping fetch_sackmann – output already exists: {OUTPUT}")
        return pd.read_parquet(OUTPUT)

    log.info("Phase 1b START: Fetching Sackmann CSVs")
    atp = normalize(load_tour(SACKMANN_ATP_URL, "atp", START_YEAR))
    wta = normalize(load_tour(SACKMANN_WTA_URL, "wta", START_YEAR))
    df = pd.concat([atp, wta], ignore_index=True)
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    df.to_parquet(OUTPUT, index=False)
    log.info(f"Phase 1b END: {len(df)} rows → {OUTPUT}")
    return df


if __name__ == "__main__":
    fetch_all()
