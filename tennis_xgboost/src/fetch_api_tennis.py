"""Phase 1: Fetch match data from api-tennis.com with incremental JSON cache."""
import json
import logging
import os
import time
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
import requests
import yaml
from dotenv import load_dotenv

load_dotenv()
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

ROOT = Path(__file__).parent.parent
with open(ROOT / "config.yaml") as f:
    CFG = yaml.safe_load(f)

API_KEY = os.getenv("API_TENNIS_KEY", "")
BASE_URL = CFG["api_tennis"]["base_url"].rstrip("/")
SLEEP = CFG["api_tennis"]["sleep_between_calls"]
LOOKBACK_YEARS = CFG["api_tennis"]["lookback_years"]
CACHE_DIR = ROOT / "data" / "raw" / "api_tennis"
CACHE_DIR.mkdir(parents=True, exist_ok=True)

HEADERS = {"X-Auth-Token": API_KEY}


def fetch_day(day: date) -> list[dict]:
    """Fetch matches for a single day, using JSON cache."""
    cache_file = CACHE_DIR / f"{day.isoformat()}.json"
    if cache_file.exists():
        with open(cache_file) as f:
            return json.load(f)

    url = f"{BASE_URL}/matches"
    params = {"dateFrom": day.isoformat(), "dateTo": day.isoformat()}
    try:
        resp = requests.get(url, headers=HEADERS, params=params, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        matches = data if isinstance(data, list) else data.get("matches", data.get("data", []))
    except Exception as e:
        log.warning(f"Failed to fetch {day}: {e}")
        matches = []

    with open(cache_file, "w") as f:
        json.dump(matches, f)
    time.sleep(SLEEP)
    return matches


def fetch_all(force: bool = False) -> pd.DataFrame:
    """Fetch all matches from lookback_years ago until today."""
    output_path = ROOT / "data" / "processed" / "matches_api.parquet"
    if output_path.exists() and not force:
        log.info(f"Skipping fetch_api_tennis – output already exists: {output_path}")
        return pd.read_parquet(output_path)

    log.info("Phase 1 START: Fetching api-tennis match data")
    start_date = date.today() - timedelta(days=LOOKBACK_YEARS * 365)
    end_date = date.today()

    all_matches = []
    current = start_date
    while current <= end_date:
        matches = fetch_day(current)
        all_matches.extend(matches)
        current += timedelta(days=1)

    if not all_matches:
        log.warning("No matches fetched from api-tennis")
        return pd.DataFrame()

    df = pd.json_normalize(all_matches)
    df = _normalize_columns(df)
    (ROOT / "data" / "processed").mkdir(parents=True, exist_ok=True)
    df.to_parquet(output_path, index=False)
    log.info(f"Phase 1 END: {len(df)} rows → {output_path}")
    return df


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize column names to a common schema."""
    rename = {
        "date": "match_date",
        "tournament.name": "tournament",
        "tournament.surface": "surface",
        "tournament.level": "level",
        "round": "round",
        "player1.name": "player_a",
        "player2.name": "player_b",
        "player1.rank": "rank_a",
        "player2.rank": "rank_b",
        "result.winner": "winner",
        "score": "score",
        "stats.player1.aces": "aces_a",
        "stats.player2.aces": "aces_b",
        "stats.player1.doubleFaults": "df_a",
        "stats.player2.doubleFaults": "df_b",
        "stats.player1.firstServePointsWon": "first_won_a",
        "stats.player2.firstServePointsWon": "first_won_b",
        "stats.player1.firstServeIn": "first_in_a",
        "stats.player2.firstServeIn": "first_in_b",
        "stats.player1.secondServePointsWon": "second_won_a",
        "stats.player2.secondServePointsWon": "second_won_b",
        "stats.player1.breakPointsFaced": "bp_faced_a",
        "stats.player2.breakPointsFaced": "bp_faced_b",
        "stats.player1.breakPointsSaved": "bp_saved_a",
        "stats.player2.breakPointsSaved": "bp_saved_b",
        "stats.player1.servicePointsPlayed": "svpt_a",
        "stats.player2.servicePointsPlayed": "svpt_b",
        "tour": "tour",
        "retirement": "retirement",
        "walkover": "walkover",
    }
    df = df.rename(columns={k: v for k, v in rename.items() if k in df.columns})
    if "match_date" in df.columns:
        df["match_date"] = pd.to_datetime(df["match_date"], errors="coerce")
    if "surface" in df.columns:
        df["surface"] = df["surface"].str.capitalize().str.strip()
    return df


if __name__ == "__main__":
    fetch_all()
