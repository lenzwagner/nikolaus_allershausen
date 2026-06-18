"""
Demo script: fetch Sackmann data (or fall back to synthetic data), train both XGBoost and
Random Forest models, then make predictions for today's typical grass-court fixtures.

Usage:  python scripts/demo_predictions.py
"""
import logging
import sys
from pathlib import Path

ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT / "src"))

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

# ── Matches to predict (typical grass-court June fixtures) ────────────────────
DEMO_MATCHES = [
    # (player_a, player_b, surface, level, tour, description)
    ("Carlos Alcaraz",   "Novak Djokovic",   "grass", "atp500",   "atp",
     "Queen's Club QF (Grass)"),
    ("Jannik Sinner",    "Taylor Fritz",     "grass", "atp500",   "atp",
     "Halle QF (Grass)"),
    ("Iga Swiatek",      "Aryna Sabalenka",  "clay",  "grandslam","wta",
     "Roland Garros Final (Clay)"),
    ("Alexander Zverev", "Flavio Cobolli",   "clay",  "atp250",   "atp",
     "Bastad QF (Clay)"),
    ("Coco Gauff",       "Elena Rybakina",   "grass", "grandslam","wta",
     "Wimbledon SF (Grass)"),
]
# ─────────────────────────────────────────────────────────────────────────────

def run_demo():
    import pandas as pd

    # ── Step 1: Load data — try Sackmann GitHub, fall back to synthetic ────────
    log.info("Step 1: Loading match data")
    sack_path = ROOT / "data" / "processed" / "matches_sackmann.parquet"
    api_path  = ROOT / "data" / "processed" / "matches_api.parquet"

    if sack_path.exists() and pd.read_parquet(sack_path).shape[0] > 500:
        df = pd.read_parquet(sack_path)
        log.info(f"  Loaded {len(df)} cached matches from Sackmann parquet")
    else:
        log.info("  Trying Sackmann GitHub CSVs …")
        try:
            from fetch_sackmann import load_tour, normalize, SACKMANN_ATP_URL, SACKMANN_WTA_URL
            atp = normalize(load_tour(SACKMANN_ATP_URL, "atp", 2021))
            wta = normalize(load_tour(SACKMANN_WTA_URL, "wta", 2021))
            df  = pd.concat([atp, wta], ignore_index=True)
            if len(df) < 100:
                raise ValueError("Too few rows from GitHub — falling back to synthetic")
            sack_path.parent.mkdir(parents=True, exist_ok=True)
            df.to_parquet(sack_path, index=False)
            log.info(f"  Fetched {len(df)} matches from GitHub")
        except Exception as e:
            log.warning(f"  GitHub fetch failed ({e}). Generating synthetic data …")
            from generate_synthetic_data import generate_and_save
            df = generate_and_save()

    if not api_path.exists():
        pd.DataFrame(columns=df.columns).to_parquet(api_path, index=False)
    log.info(f"  Using {len(df)} training matches")

    # ── Step 2: ELO ──────────────────────────────────────────────────────────
    log.info("Step 2: Computing ELO ratings")
    from elo import compute_elo
    elo_path = ROOT / "data" / "processed" / "matches_with_elo.parquet"
    elo_df = compute_elo(df, force=not elo_path.exists())

    # ── Step 3: Rolling stats ─────────────────────────────────────────────────
    log.info("Step 3: Rolling serve stats")
    from features import compute_rolling
    roll_path = ROOT / "data" / "processed" / "matches_with_rolling.parquet"
    rolling_df = compute_rolling(elo_df, force=not roll_path.exists())

    # ── Step 4: H2H ──────────────────────────────────────────────────────────
    log.info("Step 4: Head-to-head features")
    from h2h import compute_h2h
    h2h_path = ROOT / "data" / "processed" / "matches_with_h2h.parquet"
    h2h_df = compute_h2h(rolling_df, force=not h2h_path.exists())

    # ── Step 5-7: Context + injuries (skip scraping, use empty files) ────────
    log.info("Step 5-7: Tournament context + injury stubs")
    from tournament_context import add_tournament_context
    h2h_df = add_tournament_context(h2h_df)

    inj_path = ROOT / "data" / "raw" / "injuries.csv"
    if not inj_path.exists():
        import csv
        inj_path.parent.mkdir(parents=True, exist_ok=True)
        with open(inj_path, "w") as f:
            f.write("player,tour,injury_start,injury_end,injury_type,source\n")

    from injury_features import add_injury_features
    full_df = add_injury_features(h2h_df)

    # ── Step 8: Train/test split ──────────────────────────────────────────────
    log.info("Step 8: Train/test split")
    import yaml
    with open(ROOT / "config.yaml") as f:
        cfg = yaml.safe_load(f)
    cutoff = pd.Timestamp.today() - pd.Timedelta(weeks=cfg["split"]["test_weeks"])
    train_df = full_df[full_df["match_date"] < cutoff]
    test_df  = full_df[full_df["match_date"] >= cutoff]
    train_df.to_parquet(ROOT / "data" / "processed" / "train.parquet", index=False)
    test_df.to_parquet(ROOT / "data" / "processed"  / "test.parquet",  index=False)
    log.info(f"  train={len(train_df)}, test={len(test_df)}")

    if len(train_df) < 100:
        log.warning("Very little training data — predictions will be weak. "
                    "Set START_YEAR=2010 in fetch_sackmann for more data.")

    # ── Step 9: Train XGBoost models ─────────────────────────────────────────
    log.info("Step 9: Training XGBoost models")
    from train import train_surface_model
    for surf in ["Hard", "Clay", "Grass"]:
        train_surface_model(train_df, surf, force=True)

    # ── Step 10: Train Random Forest models ──────────────────────────────────
    log.info("Step 10: Training Random Forest models")
    from train_rf import train_rf_surface
    for surf in ["Hard", "Clay", "Grass"]:
        train_rf_surface(train_df, surf, force=True)

    # ── Step 11: Evaluate on test set ────────────────────────────────────────
    log.info("Step 11: Evaluating models")
    from evaluate import evaluate_all
    evaluate_all(force=True)

    # ── Step 12: Predictions ─────────────────────────────────────────────────
    print("\n" + "="*65)
    print("   TENNIS MATCH PREDICTIONS — Today's Demo Fixtures")
    print("="*65)

    import pickle
    import numpy as np
    import xgboost as xgb
    sys.path.insert(0, str(ROOT))
    from predict import build_feature_vector, FEATURE_COLS, XGB_PARAMS

    for pa, pb, surf, level, tour, desc in DEMO_MATCHES:
        print(f"\n{'─'*65}")
        print(f"  {desc}")
        print(f"{'─'*65}")
        X = build_feature_vector(pa, pb, surf, level, tour, h2h_df)
        avail = [c for c in FEATURE_COLS if c in X.columns]
        X_m = X[avail].fillna(0)

        for mtype, label in [("xgb", "XGBoost"), ("rf", "Random Forest")]:
            if mtype == "xgb":
                mp = ROOT / "models" / f"xgb_{surf}.json"
                if not mp.exists():
                    continue
                m = xgb.XGBClassifier(**XGB_PARAMS)
                m.load_model(str(mp))
            else:
                mp = ROOT / "models" / f"rf_{surf}.pkl"
                if not mp.exists():
                    continue
                with open(mp, "rb") as f:
                    m = pickle.load(f)

            p_a = float(m.predict_proba(X_m)[0][1])
            p_b = 1.0 - p_a
            winner = pa if p_a >= p_b else pb
            bar_a = "█" * int(p_a * 20)
            bar_b = "█" * int(p_b * 20)
            print(f"  [{label}]")
            print(f"    {pa:<25} {bar_a:<20} {p_a*100:5.1f}%")
            print(f"    {pb:<25} {bar_b:<20} {p_b*100:5.1f}%")
            print(f"    → Predicted winner: {winner}")

    print("\n" + "="*65)
    print("  Done. See reports/evaluation.md for model metrics.")
    print("="*65 + "\n")


if __name__ == "__main__":
    run_demo()
