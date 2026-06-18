"""
Backtest: compare model predictions against actual match outcomes.

Loads the test parquet (matches the model never saw during training),
predicts each match, and reports:
  - Overall accuracy vs baseline (higher-ranked player wins)
  - Breakdown by surface, ELO-gap bucket, and confidence tier
  - Per-match table for the most recent N matches
  - Comparison: XGBoost vs Random Forest

Usage:  python scripts/backtest.py [--n_recent 30] [--surface hard|clay|grass]
"""
import argparse
import pickle
import sys
from pathlib import Path

import numpy as np
import pandas as pd
import xgboost as xgb

ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT / "src"))

from feature_config import FEATURE_COLS

XGB_PARAMS = {
    "n_estimators": 500, "max_depth": 6, "learning_rate": 0.05,
    "subsample": 0.8, "colsample_bytree": 0.8,
    "eval_metric": "logloss", "random_state": 42, "tree_method": "hist",
}


def load_model(surface: str, model_type: str):
    if model_type == "rf":
        p = ROOT / "models" / f"rf_{surface.lower()}.pkl"
        if not p.exists():
            return None
        with open(p, "rb") as f:
            return pickle.load(f)
    # XGBoost
    cal = ROOT / "models" / f"xgb_{surface.lower()}_calibrated.pkl"
    base = ROOT / "models" / f"xgb_{surface.lower()}.json"
    if cal.exists():
        with open(cal, "rb") as f:
            return pickle.load(f)
    if base.exists():
        m = xgb.XGBClassifier(**XGB_PARAMS)
        m.load_model(str(base))
        return m
    return None


def predict_batch(model, df: pd.DataFrame) -> np.ndarray:
    available = [c for c in FEATURE_COLS if c in df.columns]
    X = df[available].fillna(0)
    return model.predict_proba(X)[:, 1]


def accuracy_report(label: str, y_true: pd.Series, proba: np.ndarray) -> dict:
    pred = (proba >= 0.5).astype(int)
    correct = (pred == y_true.values).sum()
    acc = correct / len(y_true)
    # Confidence buckets: how often is the model right when it's confident?
    confident_mask = (proba >= 0.65) | (proba <= 0.35)
    if confident_mask.sum() > 0:
        conf_acc = (pred[confident_mask] == y_true.values[confident_mask]).mean()
    else:
        conf_acc = np.nan
    return {"label": label, "n": len(y_true), "accuracy": acc,
            "n_confident": int(confident_mask.sum()), "confident_accuracy": conf_acc}


def baseline_pred(df: pd.DataFrame) -> np.ndarray:
    """Baseline: higher-ranked player (lower rank number) wins."""
    ra = df.get("rank_a", pd.Series([999]*len(df))).fillna(999)
    rb = df.get("rank_b", pd.Series([999]*len(df))).fillna(999)
    # Return 1 if rank_a < rank_b (player A has better rank = wins)
    return (ra < rb).astype(float).values


def print_match_table(df: pd.DataFrame, xgb_proba: np.ndarray, rf_proba: np.ndarray | None,
                      n: int = 30):
    """Print a per-match comparison table."""
    df = df.copy().reset_index(drop=True)
    actual_winner = df["winner"].map({"A": df["player_a"], "B": df["player_b"]})

    rows = []
    for i in range(min(n, len(df))):
        row = df.iloc[i]
        pa = str(row.get("player_a", "?"))[:20]
        pb = str(row.get("player_b", "?"))[:20]
        surf = str(row.get("surface", "?"))[:5]
        actual = "A" if row.get("winner") == "A" else "B"
        p_a_xgb = xgb_proba[i]
        pred_xgb = "A" if p_a_xgb >= 0.5 else "B"
        correct_xgb = "✓" if pred_xgb == actual else "✗"

        rf_str = ""
        if rf_proba is not None:
            p_a_rf = rf_proba[i]
            pred_rf = "A" if p_a_rf >= 0.5 else "B"
            correct_rf = "✓" if pred_rf == actual else "✗"
            rf_str = f"  RF {p_a_rf*100:4.1f}% {correct_rf}"

        date = str(row.get("match_date", ""))[:10]
        home = ""
        if "home_advantage_diff" in row.index:
            hd = int(row["home_advantage_diff"])
            if hd == 1:
                home = " 🏠A"
            elif hd == -1:
                home = " 🏠B"

        rows.append(
            f"  {date}  {surf:<5}  {pa:<20} vs {pb:<20}  "
            f"Actual:{actual}  XGB {p_a_xgb*100:4.1f}% {correct_xgb}{rf_str}{home}"
        )

    print("\n".join(rows))


def run_backtest(surface_filter: str | None = None, n_recent: int = 40):
    test_path = ROOT / "data" / "processed" / "test.parquet"
    if not test_path.exists():
        print("test.parquet not found. Run demo_predictions.py first.")
        return

    test = pd.read_parquet(test_path)
    test = test.sort_values("match_date").reset_index(drop=True)
    y_true = (test["winner"] == "A").astype(int)

    surfaces = [surface_filter.capitalize()] if surface_filter else ["Hard", "Clay", "Grass"]
    surf_col = test.get("surface", pd.Series(["Hard"] * len(test)))

    print("\n" + "="*72)
    print("  BACKTEST — Vorhersage vs. echter Ausgang (Test-Set)")
    print("="*72)
    print(f"  Test-Matches gesamt: {len(test)}")
    print(f"  Zeitraum: {test['match_date'].min().date()} → {test['match_date'].max().date()}")
    print()

    # ── Baseline ─────────────────────────────────────────────────────────────
    base_proba = baseline_pred(test)
    base_correct = ((base_proba >= 0.5) == y_true.values).sum()
    print(f"  Baseline (höher Gerankte gewinnt): {base_correct/len(test)*100:.1f}%  "
          f"({base_correct}/{len(test)})\n")

    all_results = []

    for surface in surfaces:
        mask = surf_col.str.capitalize() == surface
        surf_test = test[mask].reset_index(drop=True)
        if surf_test.empty:
            print(f"  [{surface}] keine Daten im Test-Set\n")
            continue

        y_surf = (surf_test["winner"] == "A").astype(int)
        base_surf = baseline_pred(surf_test)
        base_surf_acc = ((base_surf >= 0.5) == y_surf.values).mean()

        print(f"{'─'*72}")
        print(f"  Surface: {surface}  ({len(surf_test)} Matches)")
        print(f"  Baseline: {base_surf_acc*100:.1f}%")

        for model_type, label in [("xgb", "XGBoost"), ("rf", "Random Forest")]:
            model = load_model(surface, model_type)
            if model is None:
                print(f"  [{label}] Kein Modell gefunden")
                continue
            proba = predict_batch(model, surf_test)
            r = accuracy_report(label, y_surf, proba)
            all_results.append({**r, "surface": surface, "model": model_type})

            conf_str = (f"  | bei ≥65% Konfidenz: {r['confident_accuracy']*100:.1f}% "
                        f"({r['n_confident']} Matches)")
            print(f"  {label:<15}: {r['accuracy']*100:.1f}%  "
                  f"({int(r['accuracy']*r['n'])}/{r['n']}){conf_str}")

        # ELO-gap buckets (how accurate for easy vs hard matches?)
        if "elo_diff" in surf_test.columns:
            xgb_model = load_model(surface, "xgb")
            if xgb_model:
                xgb_proba_surf = predict_batch(xgb_model, surf_test)
                print(f"\n  XGBoost Genauigkeit nach ELO-Abstand:")
                for lo, hi, tag in [(-999, -50, "B klar besser"), (-50, 50, "Ausgeglichen"),
                                     (50, 200, "A leicht besser"), (200, 9999, "A klar besser")]:
                    em = (surf_test["elo_diff"] >= lo) & (surf_test["elo_diff"] < hi)
                    if em.sum() == 0:
                        continue
                    bucket_acc = ((xgb_proba_surf[em] >= 0.5) == y_surf.values[em]).mean()
                    print(f"    ELO-Diff {lo:>+5} .. {hi:>+5}  ({tag:<18}): "
                          f"{bucket_acc*100:.1f}%  n={em.sum()}")
        print()

    # ── Recent matches table ──────────────────────────────────────────────────
    print(f"{'─'*72}")
    print(f"  Die letzten {n_recent} Matches im Test-Set (chronologisch)")
    print(f"  Format: Datum  Surface  Spieler_A vs Spieler_B  "
          f"Actual  XGB-P(A) ✓/✗  RF-P(A) ✓/✗  🏠=Heimvorteil")
    print(f"{'─'*72}")

    recent = test.tail(n_recent).reset_index(drop=True)
    y_recent = (recent["winner"] == "A").astype(int)

    # Get XGB and RF probabilities for each surface in recent matches
    xgb_proba_recent = np.full(len(recent), 0.5)
    rf_proba_recent  = np.full(len(recent), 0.5)
    has_rf = False
    for surface in ["Hard", "Clay", "Grass"]:
        sm = recent.get("surface", pd.Series(["Hard"]*len(recent))).str.capitalize() == surface
        if sm.sum() == 0:
            continue
        xgb_m = load_model(surface, "xgb")
        rf_m  = load_model(surface, "rf")
        if xgb_m:
            surf_slice = recent[sm].reset_index(drop=True)
            xgb_proba_recent[sm.values] = predict_batch(xgb_m, surf_slice)
        if rf_m:
            surf_slice = recent[sm].reset_index(drop=True)
            rf_proba_recent[sm.values] = predict_batch(rf_m, surf_slice)
            has_rf = True

    print_match_table(recent, xgb_proba_recent, rf_proba_recent if has_rf else None, n_recent)

    # Summary
    xgb_acc = ((xgb_proba_recent >= 0.5) == y_recent.values).mean()
    rf_acc  = ((rf_proba_recent  >= 0.5) == y_recent.values).mean() if has_rf else None
    base_acc = ((baseline_pred(recent) >= 0.5) == y_recent.values).mean()
    print(f"\n  In diesen {n_recent} Matches: XGB {xgb_acc*100:.1f}%"
          + (f"  |  RF {rf_acc*100:.1f}%" if rf_acc else "")
          + f"  |  Baseline {base_acc*100:.1f}%")
    print("="*72 + "\n")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Backtest: Prediction vs. actual outcomes")
    parser.add_argument("--surface", default=None, choices=["hard", "clay", "grass"],
                        help="Filter to one surface (default: all)")
    parser.add_argument("--n_recent", type=int, default=40,
                        help="Number of recent matches to show in detail table")
    args = parser.parse_args()
    run_backtest(surface_filter=args.surface, n_recent=args.n_recent)
