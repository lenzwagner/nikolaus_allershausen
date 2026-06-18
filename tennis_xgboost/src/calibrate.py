"""Phase 12: Isotonic regression calibration of XGBoost surface models."""
import logging
import pickle
from pathlib import Path

import pandas as pd
import xgboost as xgb
from sklearn.calibration import CalibratedClassifierCV

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

ROOT = Path(__file__).parent.parent
MODELS_DIR = ROOT / "models"

import sys
sys.path.insert(0, str(ROOT / "src"))
from feature_config import FEATURE_COLS

XGB_PARAMS = {
    "n_estimators": 500, "max_depth": 6, "learning_rate": 0.05,
    "subsample": 0.8, "colsample_bytree": 0.8,
    "eval_metric": "logloss", "random_state": 42, "tree_method": "hist",
}


def calibrate_surface(surface: str, val_df: pd.DataFrame, force: bool = False):
    output = MODELS_DIR / f"xgb_{surface.lower()}_calibrated.pkl"
    if output.exists() and not force:
        log.info(f"Skipping calibration {surface} — already exists")
        with open(output, "rb") as f:
            return pickle.load(f)

    model_path = MODELS_DIR / f"xgb_{surface.lower()}.json"
    if not model_path.exists():
        log.warning(f"Base model not found: {model_path}")
        return None

    base = xgb.XGBClassifier(**XGB_PARAMS)
    base.load_model(str(model_path))

    available = [c for c in FEATURE_COLS if c in val_df.columns]
    X = val_df[available].fillna(val_df[available].median(numeric_only=True))
    y = (val_df["winner"] == "A").astype(int)

    calibrated = CalibratedClassifierCV(base, cv="prefit", method="isotonic")
    calibrated.fit(X, y)

    with open(output, "wb") as f:
        pickle.dump(calibrated, f)
    log.info(f"Saved calibrated model → {output}")
    return calibrated


def calibrate_all(force: bool = False):
    log.info("Phase 12 START: Calibrating all surface models")
    test_path = ROOT / "data" / "processed" / "test.parquet"
    if not test_path.exists():
        log.error("test.parquet not found — run run_pipeline.py first")
        return

    val_df = pd.read_parquet(test_path)
    for surface in ["Hard", "Clay", "Grass"]:
        calibrate_surface(surface, val_df, force=force)
    log.info("Phase 12 END: Calibration complete")


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--force", action="store_true")
    args = parser.parse_args()
    calibrate_all(force=args.force)
