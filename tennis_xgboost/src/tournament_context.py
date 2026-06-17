"""Phase 6: Tournament context features (Grand Slam proximity, level, round encoding)."""
import logging
from pathlib import Path

import numpy as np
import pandas as pd

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

ROOT = Path(__file__).parent.parent

GRAND_SLAMS = [
    {"name": "Australian Open", "month": 1, "day": 14},
    {"name": "Roland Garros",   "month": 5, "day": 26},
    {"name": "Wimbledon",       "month": 6, "day": 30},
    {"name": "US Open",         "month": 8, "day": 26},
]

LEVEL_MAP = {
    "grand slam": 0, "grandslam": 0, "g": 0,
    "masters": 1, "masters 1000": 1, "premier mandatory": 1, "pm": 1, "m": 1,
    "atp500": 2, "wta500": 2, "500": 2, "a": 2,
    "atp250": 3, "wta250": 3, "250": 3, "international": 3, "d": 3, "f": 3,
}

ROUND_MAP = {
    "r128": 0, "r64": 1, "r32": 2, "r16": 3,
    "qf": 4, "sf": 5, "f": 6, "rr": 2,
}


def gs_nearest_days(match_date: pd.Timestamp) -> int:
    """Return days to nearest Grand Slam start (positive = before, negative = after)."""
    if pd.isna(match_date):
        return 999
    year = match_date.year
    min_days = 999
    for gs in GRAND_SLAMS:
        for y in [year - 1, year, year + 1]:
            try:
                gs_date = pd.Timestamp(year=y, month=gs["month"], day=gs["day"])
                diff = (gs_date - match_date).days
                if abs(diff) < abs(min_days):
                    min_days = diff
            except Exception:
                pass
    return min_days


def add_tournament_context(df: pd.DataFrame) -> pd.DataFrame:
    log.info("Phase 6: Adding tournament context features")
    if "match_date" in df.columns:
        df["gs_days"] = df["match_date"].apply(gs_nearest_days)
        df["is_pre_grand_slam_tournament"] = (df["gs_days"].between(0, 21)).astype(int)
    else:
        df["gs_days"] = 0
        df["is_pre_grand_slam_tournament"] = 0

    if "level" in df.columns:
        df["level_encoded"] = df["level"].str.lower().str.strip().map(LEVEL_MAP).fillna(3).astype(int)
    else:
        df["level_encoded"] = 3

    if "round" in df.columns:
        df["round_encoded"] = df["round"].str.lower().str.strip().map(ROUND_MAP).fillna(2).astype(int)
    else:
        df["round_encoded"] = 2

    if "surface" in df.columns:
        df["surface_encoded"] = df["surface"].str.capitalize().map({"Hard": 0, "Clay": 1, "Grass": 2}).fillna(0).astype(int)
    else:
        df["surface_encoded"] = 0

    return df
