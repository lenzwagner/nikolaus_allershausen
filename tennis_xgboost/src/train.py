"""Phase 10: Train surface-specific XGBoost models with sample weights."""
import logging
import sys
from pathlib import Path

import numpy as np
import pandas as pd
import xgboost as xgb
import yaml

ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT / "src"))

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

with open(ROOT / "config.yaml") as f:
    CFG = yaml.safe_load(f)

MODELS_DIR = ROOT / "models"
MODELS_DIR.mkdir(exist_ok=True)

FEATURE_COLS = [
    "elo_diff", "elo_hard_diff", "elo_clay_diff", "elo_grass_diff",
    "rank_a", "rank_b",
    "h2h_count", "h2h_win_rate_a", "h2h_surface_count", "h2h_surface_win_rate_a",
    "form_a", "form_b", "form_surface_a", "form_surface_b",
    "days_since_last_a", "days_since_last_b",
    "first_serve_pct_a", "first_serve_pct_b",
    "second_serve_pct_a", "second_serve_pct_b",
    "bp_saved_pct_a", "bp_saved_pct_b",
    "aces_per_match_a", "aces_per_match_b",
    "df_per_match_a", "df_per_match_b",
    "surf_first_serve_pct_a", "surf_first_serve_pct_b",
    "surf_second_serve_pct_a", "surf_second_serve_pct_b",
    "surf_bp_saved_pct_a", "surf_bp_saved_pct_b",
    "days_since_last_injury_a", "days_since_last_injury_b",
    "is_returning_from_injury_a", "is_returning_from_injury_b",
    "retirement_rate_12m_a", "retirement_rate_12m_b",
    "days_since_last_injury_diff",
    "level_encoded", "round_encoded", "surface_encoded",
    "is_pre_grand_slam_tournament", "gs_days",
]

XGB_PARAMS = {
    "n_estimators": 500,
    "max_depth": 6,
    "learning_rate": 0.05,
    "subsample": 0.8,
    "colsample_bytree": 0.8,
    "eval_metric": "logloss",
    "random_state": 42,
    "tree_method": "hist",
}


def get_available_features(df: pd.DataFrame) -> list[str]:
    return [c for c in FEATURE_COLS if c in df.columns]


def train_surface_model(train_df: pd.DataFrame, surface: str, force: bool = False):
    from weights import compute_weights

    output = MODELS_DIR / f"xgb_{surface.lower()}.json"
    if output.exists() and not force:
        log.info(f"Skipping training {surface} – model already exists: {output}")
        model = xgb.XGBClassifier(**XGB_PARAMS)
        model.load_model(str(output))
        return model

    log.info(f"Training {surface} model on {len(train_df)} samples")
    available = get_available_features(train_df)
    X = train_df[available].copy()
    y = (train_df["winner"] == "A").astype(int)
    sample_weights = compute_weights(train_df, surface)

    # Median imputation per column
    medians = X.median(numeric_only=True)
    X = X.fillna(medians)

    model = xgb.XGBClassifier(**XGB_PARAMS)
    model.fit(X, y, sample_weight=sample_weights)
    model.save_model(str(output))
    log.info(f"Saved {surface} model → {output}")

    feat_imp = pd.Series(model.feature_importances_, index=available).sort_values(ascending=False)
    log.info(f"Top 10 features ({surface}):\n{feat_imp.head(10).to_string()}")
    imp_path = ROOT / "reports" / f"feature_importance_{surface.lower()}.csv"
    imp_path.parent.mkdir(exist_ok=True)
    feat_imp.to_csv(imp_path)

    return model


def train_all(force: bool = False):
    log.info("Phase 10 START: Training all surface models")
    train_path = ROOT / "data" / "processed" / "train.parquet"
    if not train_path.exists():
        log.error("train.parquet not found — run run_pipeline.py first (through Phase 9)")
        return None

    train_df = pd.read_parquet(train_path)
    models = {}
    for surface in ["Hard", "Clay", "Grass"]:
        models[surface] = train_surface_model(train_df, surface, force=force)

    log.info("Phase 10 END: All models trained")
    return models


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--force", action="store_true")
    args = parser.parse_args()
    train_all(force=args.force)
