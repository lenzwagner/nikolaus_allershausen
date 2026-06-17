"""Phase 4: Head-to-head features, chronologically correct (no leakage)."""
import logging
from collections import defaultdict
from pathlib import Path

import numpy as np
import pandas as pd

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

ROOT = Path(__file__).parent.parent
OUTPUT = ROOT / "data" / "processed" / "matches_with_h2h.parquet"


def h2h_key(pa: str, pb: str) -> tuple:
    return tuple(sorted([pa, pb]))


def compute_h2h(df: pd.DataFrame, force: bool = False) -> pd.DataFrame:
    if OUTPUT.exists() and not force:
        log.info(f"Skipping h2h – output already exists: {OUTPUT}")
        return pd.read_parquet(OUTPUT)

    log.info("Phase 4 START: Computing H2H features")
    df = df.sort_values("match_date").reset_index(drop=True)

    h2h_total: dict[tuple, dict] = defaultdict(lambda: {"total": 0, "wins_a": 0})
    h2h_surface: dict[tuple, dict] = defaultdict(lambda: {"total": 0, "wins_a": 0})

    rows = []
    for _, row in df.iterrows():
        pa = str(row.get("player_a", ""))
        pb = str(row.get("player_b", ""))
        surface = str(row.get("surface", "Hard"))
        winner = str(row.get("winner", "A"))
        key = h2h_key(pa, pb)
        is_a_first = pa <= pb

        rec_total = h2h_total[key]
        rec_surf = h2h_surface[(key[0], key[1], surface)]

        h2h_count = rec_total["total"]
        h2h_win_rate_a = rec_total["wins_a"] / h2h_count if h2h_count > 0 else 0.5
        if not is_a_first:
            h2h_win_rate_a = 1.0 - h2h_win_rate_a

        h2h_surf_count = rec_surf["total"]
        h2h_surf_win_rate_a = rec_surf["wins_a"] / h2h_surf_count if h2h_surf_count > 0 else 0.5
        if not is_a_first:
            h2h_surf_win_rate_a = 1.0 - h2h_surf_win_rate_a

        rows.append({
            **row.to_dict(),
            "h2h_count": h2h_count,
            "h2h_win_rate_a": h2h_win_rate_a,
            "h2h_surface_count": h2h_surf_count,
            "h2h_surface_win_rate_a": h2h_surf_win_rate_a,
        })

        # Update counts AFTER reading (prevent leakage)
        won_canonical = (winner == "A" and is_a_first) or (winner == "B" and not is_a_first)
        rec_total["total"] += 1
        rec_total["wins_a"] += 1 if won_canonical else 0
        rec_surf["total"] += 1
        rec_surf["wins_a"] += 1 if won_canonical else 0

    result = pd.DataFrame(rows)
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    result.to_parquet(OUTPUT, index=False)
    log.info(f"Phase 4 END: {len(result)} rows → {OUTPUT}")
    return result


if __name__ == "__main__":
    df = pd.read_parquet(ROOT / "data" / "processed" / "matches_with_rolling.parquet")
    compute_h2h(df)
