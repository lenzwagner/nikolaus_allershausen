"""Phase 11: Evaluate surface models — accuracy, log-loss, Brier, ROC-AUC, calibration."""
import logging
from pathlib import Path

import numpy as np
import pandas as pd
import xgboost as xgb
from sklearn.metrics import (accuracy_score, brier_score_loss, log_loss,
                             roc_auc_score)

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

ROOT = Path(__file__).parent.parent
REPORTS_DIR = ROOT / "reports"
REPORTS_DIR.mkdir(exist_ok=True)

import sys
sys.path.insert(0, str(ROOT / "src"))
from feature_config import FEATURE_COLS

XGB_PARAMS = {
    "n_estimators": 500, "max_depth": 6, "learning_rate": 0.05,
    "subsample": 0.8, "colsample_bytree": 0.8,
    "eval_metric": "logloss", "random_state": 42, "tree_method": "hist",
}


def baseline_accuracy(df: pd.DataFrame) -> float:
    if "rank_a" not in df.columns or "rank_b" not in df.columns:
        return np.nan
    y_true = (df["winner"] == "A").astype(int)
    pred = (df["rank_a"].fillna(999) < df["rank_b"].fillna(999)).astype(int)
    return float(accuracy_score(y_true, pred))


def evaluate_surface(model: xgb.XGBClassifier, test_df: pd.DataFrame, surface: str) -> dict:
    available = [c for c in FEATURE_COLS if c in test_df.columns]
    X = test_df[available].fillna(test_df[available].median(numeric_only=True))
    y = (test_df["winner"] == "A").astype(int)

    proba = model.predict_proba(X)[:, 1]
    pred = (proba >= 0.5).astype(int)

    n_classes = len(y.unique())
    return {
        "surface": surface,
        "n_matches": len(test_df),
        "accuracy": float(accuracy_score(y, pred)),
        "log_loss": float(log_loss(y, proba)),
        "brier_score": float(brier_score_loss(y, proba)),
        "roc_auc": float(roc_auc_score(y, proba)) if n_classes > 1 else np.nan,
        "baseline_accuracy": baseline_accuracy(test_df),
    }


def evaluate_all(force: bool = False):
    report_path = REPORTS_DIR / "evaluation.md"
    if report_path.exists() and not force:
        log.info("Skipping evaluation — report already exists")
        return

    log.info("Phase 11 START: Evaluating surface models")
    test_path = ROOT / "data" / "processed" / "test.parquet"
    if not test_path.exists():
        log.error("test.parquet not found — run run_pipeline.py first")
        return

    test_df = pd.read_parquet(test_path)
    lines = ["# Model Evaluation Report\n", f"Generated on test set: {len(test_df)} total matches\n"]

    for surface in ["Hard", "Clay", "Grass"]:
        model_path = ROOT / "models" / f"xgb_{surface.lower()}.json"
        if not model_path.exists():
            lines.append(f"## {surface}\n\nModel not found.\n")
            continue

        model = xgb.XGBClassifier(**XGB_PARAMS)
        model.load_model(str(model_path))

        surf_df = test_df[test_df.get("surface", pd.Series(dtype=str)).str.capitalize() == surface]
        if surf_df.empty:
            lines.append(f"## {surface}\n\nNo test data available.\n")
            continue

        m = evaluate_surface(model, surf_df, surface)
        lines += [
            f"## {surface}\n",
            f"- Matches: {m['n_matches']}",
            f"- Accuracy: {m['accuracy']:.4f}",
            f"- Log-Loss: {m['log_loss']:.4f}",
            f"- Brier Score: {m['brier_score']:.4f}",
            f"- ROC-AUC: {m['roc_auc']:.4f}",
            f"- Baseline (higher rank wins): {m['baseline_accuracy']:.4f}\n",
        ]
        log.info(
            f"{surface}: acc={m['accuracy']:.4f} logloss={m['log_loss']:.4f} "
            f"auc={m['roc_auc']:.4f} baseline={m['baseline_accuracy']:.4f}"
        )

    report_path.write_text("\n".join(lines))
    log.info(f"Phase 11 END: Report → {report_path}")


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--force", action="store_true")
    args = parser.parse_args()
    evaluate_all(force=args.force)
