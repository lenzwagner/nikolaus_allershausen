"""Master pipeline runner — executes all phases sequentially with skip logic.

Usage:
    python run_pipeline.py           # skip phases whose output already exists
    python run_pipeline.py --force   # recompute everything from scratch
"""
import argparse
import logging
import sys
from pathlib import Path

import pandas as pd
import yaml

ROOT = Path(__file__).parent
sys.path.insert(0, str(ROOT / "src"))

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)


def run_pipeline(force: bool = False):
    log.info("=== Tennis XGBoost Pipeline START ===")

    # Phase 1a – api-tennis data
    from fetch_api_tennis import fetch_all as fetch_api
    api_df = fetch_api(force=force)

    # Phase 1b – Sackmann warmup data
    from fetch_sackmann import fetch_all as fetch_sackmann
    sackmann_df = fetch_sackmann(force=force)

    # Merge & deduplicate for ELO warmup
    elo_input_path = ROOT / "data" / "processed" / "elo_input.parquet"
    if not elo_input_path.exists() or force:
        combined = pd.concat([sackmann_df, api_df], ignore_index=True)
        if "match_date" in combined.columns and "player_a" in combined.columns:
            combined = combined.drop_duplicates(
                subset=["match_date", "player_a", "player_b"], keep="last"
            )
        combined = combined.sort_values("match_date").reset_index(drop=True)
        elo_input_path.parent.mkdir(parents=True, exist_ok=True)
        combined.to_parquet(elo_input_path, index=False)
        log.info(f"Combined dataset: {len(combined)} rows")
    else:
        combined = pd.read_parquet(elo_input_path)

    # Phase 2 – ELO
    from elo import compute_elo
    elo_df = compute_elo(combined, force=force)

    # Phase 3 – Rolling serve stats & form
    from features import compute_rolling
    rolling_df = compute_rolling(elo_df, force=force)

    # Phase 4 – Head-to-head features
    from h2h import compute_h2h
    h2h_df = compute_h2h(rolling_df, force=force)

    # Phase 5 – Scraping (rankings, ELO, injuries)
    from scrape_rankings import scrape_all as scrape_rankings
    from scrape_elo import scrape_all as scrape_elo
    from scrape_injuries import scrape_injuries
    scrape_rankings(force=force)
    scrape_elo(force=force)
    scrape_injuries(force=force)

    # Phase 6 – Tournament context
    from tournament_context import add_tournament_context
    h2h_df = add_tournament_context(h2h_df)

    # Phase 7 – Injury features
    from injury_features import add_injury_features
    full_df = add_injury_features(h2h_df, force=force)

    # Phase 9 – Train/test split (time-based)
    with open(ROOT / "config.yaml") as f:
        cfg = yaml.safe_load(f)
    test_weeks = cfg["split"]["test_weeks"]
    cutoff = pd.Timestamp.today() - pd.Timedelta(weeks=test_weeks)

    train_path = ROOT / "data" / "processed" / "train.parquet"
    test_path = ROOT / "data" / "processed" / "test.parquet"
    if not train_path.exists() or force:
        train_df = full_df[full_df["match_date"] < cutoff]
        test_df = full_df[full_df["match_date"] >= cutoff]
        train_df.to_parquet(train_path, index=False)
        test_df.to_parquet(test_path, index=False)
        log.info(f"Split: train={len(train_df)} rows, test={len(test_df)} rows")

    # Phase 10 – XGBoost training (3 surface models)
    from train import train_all
    train_all(force=force)

    # Phase 11 – Evaluation
    from evaluate import evaluate_all
    evaluate_all(force=force)

    # Phase 12 – Calibration
    from calibrate import calibrate_all
    calibrate_all(force=force)

    log.info("=== Tennis XGBoost Pipeline END ===")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Run the full Tennis XGBoost pipeline")
    parser.add_argument("--force", action="store_true",
                        help="Force recompute all phases (ignore cached outputs)")
    args = parser.parse_args()
    run_pipeline(force=args.force)
