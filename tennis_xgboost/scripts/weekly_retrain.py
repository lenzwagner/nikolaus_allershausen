"""Phase 13: Weekly retraining script — incremental update + retrain all models."""
import logging
import subprocess
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT / "src"))

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(ROOT / "reports" / f"retrain_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"),
    ],
)
log = logging.getLogger(__name__)


def run_step(description: str, fn, **kwargs):
    log.info(f"START: {description}")
    try:
        result = fn(**kwargs)
        log.info(f"END: {description}")
        return result
    except Exception as e:
        log.error(f"FAILED: {description} — {e}")
        raise


def weekly_retrain():
    log.info("=== Weekly Retraining START ===")
    ts = datetime.now().strftime("%Y%m%d")

    from fetch_api_tennis import fetch_all
    from fetch_sackmann import fetch_all as fetch_sackmann
    from scrape_rankings import scrape_all as scrape_rankings
    from scrape_elo import scrape_all as scrape_elo
    from scrape_injuries import scrape_injuries
    from elo import compute_elo
    from features import compute_rolling
    from h2h import compute_h2h
    from tournament_context import add_tournament_context
    from injury_features import add_injury_features
    from train import train_surface_model
    from evaluate import evaluate_all
    import pandas as pd
    import yaml
    import shutil

    # Step 1: Fetch new data (force=False → incremental)
    run_step("Fetch api-tennis data", fetch_all, force=False)
    run_step("Fetch Sackmann data", fetch_sackmann, force=False)

    # Step 2–3: Scraping
    run_step("Scrape rankings", scrape_rankings, force=True)
    run_step("Scrape ELO", scrape_elo, force=True)
    run_step("Scrape injuries", scrape_injuries, force=True)

    # Steps 4–8: Rebuild features
    api_df = pd.read_parquet(ROOT / "data" / "processed" / "matches_api.parquet")
    sackmann_df = pd.read_parquet(ROOT / "data" / "processed" / "matches_sackmann.parquet")
    combined = pd.concat([sackmann_df, api_df], ignore_index=True).drop_duplicates(
        subset=["match_date", "player_a", "player_b"], keep="last"
    ).sort_values("match_date").reset_index(drop=True)
    combined.to_parquet(ROOT / "data" / "processed" / "elo_input.parquet", index=False)

    elo_df = run_step("Compute ELO", compute_elo, df=combined, force=True)
    rolling_df = run_step("Compute rolling features", compute_rolling, df=elo_df, force=True)
    h2h_df = run_step("Compute H2H", compute_h2h, df=rolling_df, force=True)
    h2h_df = add_tournament_context(h2h_df)
    full_df = run_step("Compute injury features", add_injury_features, df=h2h_df, force=True)

    # Step 5: Update train/test split
    with open(ROOT / "config.yaml") as f:
        cfg = yaml.safe_load(f)
    test_weeks = cfg["split"]["test_weeks"]
    cutoff = pd.Timestamp.today() - pd.Timedelta(weeks=test_weeks)
    train_df = full_df[full_df["match_date"] < cutoff]
    test_df = full_df[full_df["match_date"] >= cutoff]
    train_df.to_parquet(ROOT / "data" / "processed" / "train.parquet", index=False)
    test_df.to_parquet(ROOT / "data" / "processed" / "test.parquet", index=False)
    log.info(f"Split: train={len(train_df)}, test={len(test_df)}")

    # Step 6: Retrain all models with timestamp
    models_dir = ROOT / "models"
    for surface in ["Hard", "Clay", "Grass"]:
        model = run_step(f"Train {surface} model", train_surface_model, train_df=train_df, surface=surface, force=True)
        src = models_dir / f"xgb_{surface.lower()}.json"
        dst = models_dir / f"xgb_{surface.lower()}_{ts}.json"
        if src.exists():
            shutil.copy2(src, dst)
            log.info(f"Timestamped model saved: {dst}")

    # Step 7: Evaluate
    run_step("Evaluate models", evaluate_all, force=True)

    log.info("=== Weekly Retraining END ===")


if __name__ == "__main__":
    (ROOT / "reports").mkdir(exist_ok=True)
    weekly_retrain()
