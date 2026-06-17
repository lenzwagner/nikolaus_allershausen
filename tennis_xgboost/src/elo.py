"""Phase 2: ELO rating calculation (Overall + surface-specific) with Sackmann warmup."""
import logging
from collections import defaultdict
from pathlib import Path

import pandas as pd
import yaml

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

ROOT = Path(__file__).parent.parent
with open(ROOT / "config.yaml") as f:
    CFG = yaml.safe_load(f)

K = CFG["elo"]["k_factor"]
DEFAULT_ELO = 1500.0
SURFACES = ["Hard", "Clay", "Grass"]
OUTPUT = ROOT / "data" / "processed" / "matches_with_elo.parquet"


def expected_score(ra: float, rb: float) -> float:
    return 1.0 / (1.0 + 10 ** ((rb - ra) / 400.0))


def update_elo(ra: float, rb: float, won: bool, k: float = K):
    ea = expected_score(ra, rb)
    sa = 1.0 if won else 0.0
    new_ra = ra + k * (sa - ea)
    new_rb = rb + k * ((1 - sa) - (1 - ea))
    return new_ra, new_rb


def compute_elo(df: pd.DataFrame, force: bool = False) -> pd.DataFrame:
    if OUTPUT.exists() and not force:
        log.info(f"Skipping elo – output already exists: {OUTPUT}")
        return pd.read_parquet(OUTPUT)

    log.info("Phase 2 START: Computing ELO ratings")
    df = df.sort_values("match_date").reset_index(drop=True)

    elo_overall: dict[str, float] = defaultdict(lambda: DEFAULT_ELO)
    elo_surface: dict[tuple, float] = defaultdict(lambda: DEFAULT_ELO)

    rows = []
    for _, row in df.iterrows():
        pa, pb = str(row.get("player_a", "")), str(row.get("player_b", ""))
        surface = str(row.get("surface", "Hard"))
        winner = str(row.get("winner", "A"))

        ra_ov = elo_overall[pa]
        rb_ov = elo_overall[pb]
        ra_sf = elo_surface[(pa, surface)]
        rb_sf = elo_surface[(pb, surface)]

        rows.append({
            **row.to_dict(),
            "elo_a": ra_ov,
            "elo_b": rb_ov,
            f"elo_{surface.lower()}_a": ra_sf,
            f"elo_{surface.lower()}_b": rb_sf,
            "elo_diff": ra_ov - rb_ov,
            f"elo_{surface.lower()}_diff": ra_sf - rb_sf,
        })

        won_a = winner == "A"
        new_ra_ov, new_rb_ov = update_elo(ra_ov, rb_ov, won_a)
        new_ra_sf, new_rb_sf = update_elo(ra_sf, rb_sf, won_a)
        elo_overall[pa] = new_ra_ov
        elo_overall[pb] = new_rb_ov
        elo_surface[(pa, surface)] = new_ra_sf
        elo_surface[(pb, surface)] = new_rb_sf

    result = pd.DataFrame(rows)
    # Fill missing surface ELO columns
    for s in SURFACES:
        col_a = f"elo_{s.lower()}_a"
        col_b = f"elo_{s.lower()}_b"
        col_diff = f"elo_{s.lower()}_diff"
        if col_a not in result.columns:
            result[col_a] = DEFAULT_ELO
        if col_b not in result.columns:
            result[col_b] = DEFAULT_ELO
        if col_diff not in result.columns:
            result[col_diff] = 0.0

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    result.to_parquet(OUTPUT, index=False)
    log.info(f"Phase 2 END: {len(result)} rows → {OUTPUT}")
    return result


if __name__ == "__main__":
    sackmann = pd.read_parquet(ROOT / "data" / "processed" / "matches_sackmann.parquet")
    api = pd.read_parquet(ROOT / "data" / "processed" / "matches_api.parquet")
    combined = pd.concat([sackmann, api], ignore_index=True).drop_duplicates()
    compute_elo(combined)
