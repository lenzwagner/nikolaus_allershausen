"""Phase 8: Sample weight computation for XGBoost training."""
import logging
from pathlib import Path

import numpy as np
import pandas as pd
import yaml

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

ROOT = Path(__file__).parent.parent
with open(ROOT / "config.yaml") as f:
    CFG = yaml.safe_load(f)

W = CFG["weights"]
SURFACE_BOOST = W["surface_boost"]
GS_RADIUS_WEEKS = W["grand_slam_weeks_radius"]
GS_BOOST = W["grand_slam_boost"]
RECENCY_LAMBDA = W["recency_lambda"]
INJURY_WEIGHT = W["injury_match_weight"]

# (month, day) of typical Grand Slam start dates
GRAND_SLAM_DATES = [
    (1, 14),   # Australian Open
    (5, 26),   # Roland Garros
    (6, 30),   # Wimbledon
    (8, 26),   # US Open
]


def _is_near_grand_slam(match_date: pd.Timestamp, radius_days: int) -> bool:
    if pd.isna(match_date):
        return False
    year = match_date.year
    for month, day in GRAND_SLAM_DATES:
        for y in [year - 1, year, year + 1]:
            try:
                gs_date = pd.Timestamp(year=y, month=month, day=day)
                if abs((match_date - gs_date).days) <= radius_days:
                    return True
            except Exception:
                pass
    return False


def compute_weights(df: pd.DataFrame, target_surface: str) -> np.ndarray:
    """Compute multiplicative sample weights for a given target surface.

    weight = w_surface * w_grand_slam * w_recency * w_injury
    """
    log.info(f"Computing weights for surface={target_surface}")
    today = pd.Timestamp.today()
    radius_days = int(GS_RADIUS_WEEKS * 7)
    weights = np.ones(len(df))

    for i, (_, row) in enumerate(df.iterrows()):
        surface = str(row.get("surface", ""))
        raw_date = row.get("match_date")
        match_date = pd.Timestamp(raw_date) if pd.notna(raw_date) else None

        w_surface = SURFACE_BOOST if surface.lower() == target_surface.lower() else 1.0
        w_gs = GS_BOOST if (match_date and _is_near_grand_slam(match_date, radius_days)) else 1.0
        w_recency = np.exp(-RECENCY_LAMBDA * (today - match_date).days / 365.0) if match_date else 0.5
        inj_a = int(row.get("is_returning_from_injury_a", 0) or 0)
        inj_b = int(row.get("is_returning_from_injury_b", 0) or 0)
        w_injury = INJURY_WEIGHT if (inj_a or inj_b) else 1.0

        weights[i] = w_surface * w_gs * w_recency * w_injury

    return weights
