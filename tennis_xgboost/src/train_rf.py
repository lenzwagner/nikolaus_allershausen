"""Random Forest surface models — same feature set and sample weights as XGBoost."""
import logging
import pickle
import sys
from pathlib import Path

import numpy as np
import pandas as pd
import yaml
from sklearn.ensemble import RandomForestClassifier

ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT / "src"))

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

with open(ROOT / "config.yaml") as f:
    CFG = yaml.safe_load(f)

MODELS_DIR = ROOT / "models"
MODELS_DIR.mkdir(exist_ok=True)

from feature_config import FEATURE_COLS

RF_PARAMS = {
    "n_estimators": 500,
    "max_depth": 12,
    "min_samples_leaf": 5,
    "max_features": "sqrt",
    "n_jobs": -1,
    "random_state": 42,
    "class_weight": None,   # we pass sample_weight explicitly
}


def train_rf_surface(train_df: pd.DataFrame, surface: str, force: bool = False):
    from weights import compute_weights

    output = MODELS_DIR / f"rf_{surface.lower()}.pkl"
    if output.exists() and not force:
        log.info(f"Skipping RF training {surface} — model exists")
        with open(output, "rb") as f:
            return pickle.load(f)

    log.info(f"Training RF {surface} model on {len(train_df)} samples")
    available = [c for c in FEATURE_COLS if c in train_df.columns]
    X = train_df[available].copy().fillna(train_df[available].median(numeric_only=True))
    y = (train_df["winner"] == "A").astype(int)
    sample_weights = compute_weights(train_df, surface)

    model = RandomForestClassifier(**RF_PARAMS)
    model.fit(X, y, sample_weight=sample_weights)

    with open(output, "wb") as f:
        pickle.dump(model, f)
    log.info(f"Saved RF {surface} model → {output}")

    feat_imp = pd.Series(model.feature_importances_, index=available).sort_values(ascending=False)
    log.info(f"Top 10 RF features ({surface}):\n{feat_imp.head(10).to_string()}")
    imp_path = ROOT / "reports" / f"rf_feature_importance_{surface.lower()}.csv"
    imp_path.parent.mkdir(exist_ok=True)
    feat_imp.to_csv(imp_path)

    return model


def train_rf_all(force: bool = False):
    log.info("Phase 10b START: Training all RF surface models")
    train_path = ROOT / "data" / "processed" / "train.parquet"
    if not train_path.exists():
        log.error("train.parquet not found — run run_pipeline.py first")
        return None
    train_df = pd.read_parquet(train_path)
    models = {}
    for surface in ["Hard", "Clay", "Grass"]:
        models[surface] = train_rf_surface(train_df, surface, force=force)
    log.info("Phase 10b END: All RF models trained")
    return models


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--force", action="store_true")
    args = parser.parse_args()
    train_rf_all(force=args.force)
