"""CLI interface: predict win probabilities for a tennis match.

Usage:
    python predict.py --player_a "Alexander Zverev" --player_b "Flavio Cobolli" \
                      --surface grass --level atp250 --tour atp
"""
import argparse
import logging
import pickle
import sys
from pathlib import Path

import numpy as np
import pandas as pd
import xgboost as xgb
import yaml
from rapidfuzz import process as fz_process

ROOT = Path(__file__).parent
sys.path.insert(0, str(ROOT / "src"))

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

with open(ROOT / "config.yaml") as f:
    CFG = yaml.safe_load(f)

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
    "n_estimators": 500, "max_depth": 6, "learning_rate": 0.05,
    "subsample": 0.8, "colsample_bytree": 0.8,
    "eval_metric": "logloss", "random_state": 42, "tree_method": "hist",
}

LEVEL_MAP = {
    "grandslam": 0, "masters": 1, "atp500": 2, "wta500": 2,
    "atp250": 3, "wta250": 3, "international": 3,
}
SURFACE_MAP = {"hard": 0, "clay": 1, "grass": 2}


def _fuzzy_match(name: str, candidates: list[str], threshold: int = 65) -> str | None:
    result = fz_process.extractOne(name.lower(), [c.lower() for c in candidates],
                                   score_cutoff=threshold)
    if result:
        return candidates[[c.lower() for c in candidates].index(result[0])]
    return None


def _load_elo(player: str, tour: str) -> dict:
    path = ROOT / "data" / "raw" / f"elo_{tour}.csv"
    defaults = {"elo": 1500.0, "elo_hard": 1500.0, "elo_clay": 1500.0, "elo_grass": 1500.0}
    if not path.exists():
        return defaults
    df = pd.read_csv(path)
    if "player" not in df.columns:
        return defaults
    names = df["player"].tolist()
    match = _fuzzy_match(player, names)
    if not match:
        return defaults
    row = df[df["player"] == match].iloc[0]
    for k in defaults:
        if k in df.columns:
            try:
                defaults[k] = float(row[k])
            except Exception:
                pass
    return defaults


def _load_rank(player: str, tour: str) -> int:
    path = ROOT / "data" / "raw" / f"rankings_{tour}.csv"
    if not path.exists():
        return 999
    df = pd.read_csv(path)
    if "player" not in df.columns:
        return 999
    match = _fuzzy_match(player, df["player"].tolist())
    if not match:
        return 999
    row = df[df["player"] == match].iloc[0]
    try:
        return int(str(row.get("rank", "999")).replace(".", "").split()[0])
    except Exception:
        return 999


def _get_rolling_stats(player: str, h2h_df: pd.DataFrame) -> dict:
    mask = (
        h2h_df["player_a"].str.lower().str.contains(player.lower(), na=False) |
        h2h_df["player_b"].str.lower().str.contains(player.lower(), na=False)
    )
    hist = h2h_df[mask]
    if hist.empty:
        return {}
    row = hist.iloc[-1]
    # Prefer player_a columns if player appears as player_a
    if player.lower() in str(row.get("player_a", "")).lower():
        suf = "a"
    else:
        suf = "b"
    stat_keys = [
        "form", "form_surface", "days_since_last",
        "first_serve_pct", "second_serve_pct", "bp_saved_pct",
        "aces_per_match", "df_per_match",
        "surf_first_serve_pct", "surf_second_serve_pct", "surf_bp_saved_pct",
    ]
    return {k: row.get(f"{k}_{suf}", np.nan) for k in stat_keys}


def _h2h_stats(player_a: str, player_b: str, h2h_df: pd.DataFrame) -> dict:
    mask = (
        (h2h_df["player_a"].str.lower().str.contains(player_a.lower(), na=False) &
         h2h_df["player_b"].str.lower().str.contains(player_b.lower(), na=False)) |
        (h2h_df["player_a"].str.lower().str.contains(player_b.lower(), na=False) &
         h2h_df["player_b"].str.lower().str.contains(player_a.lower(), na=False))
    )
    hist = h2h_df[mask]
    if hist.empty:
        return {"h2h_count": 0, "h2h_win_rate_a": 0.5,
                "h2h_surface_count": 0, "h2h_surface_win_rate_a": 0.5}
    last = hist.iloc[-1]
    return {
        "h2h_count": int(last.get("h2h_count", 0)),
        "h2h_win_rate_a": float(last.get("h2h_win_rate_a", 0.5)),
        "h2h_surface_count": int(last.get("h2h_surface_count", 0)),
        "h2h_surface_win_rate_a": float(last.get("h2h_surface_win_rate_a", 0.5)),
    }


def build_feature_vector(player_a: str, player_b: str, surface: str,
                         level: str, tour: str, h2h_df: pd.DataFrame) -> pd.DataFrame:
    from tournament_context import gs_nearest_days, LEVEL_MAP as TC_LEVEL_MAP, ROUND_MAP as TC_ROUND_MAP
    from injury_features import load_injuries, compute_injury_features

    elo_a = _load_elo(player_a, tour)
    elo_b = _load_elo(player_b, tour)
    rank_a = _load_rank(player_a, tour)
    rank_b = _load_rank(player_b, tour)
    h2h = _h2h_stats(player_a, player_b, h2h_df)
    stats_a = _get_rolling_stats(player_a, h2h_df)
    stats_b = _get_rolling_stats(player_b, h2h_df)

    injuries = load_injuries()
    api_path = ROOT / "data" / "processed" / "matches_api.parquet"
    api_df = pd.read_parquet(api_path) if api_path.exists() else None
    today = pd.Timestamp.today()
    inj_a = compute_injury_features(today, player_a, injuries, api_df)
    inj_b = compute_injury_features(today, player_b, injuries, api_df)

    gs_days_val = gs_nearest_days(today)
    level_enc = TC_LEVEL_MAP.get(level.lower().replace(" ", ""), 3)
    surface_enc = SURFACE_MAP.get(surface.lower(), 0)

    days_inj_a = inj_a["days_since_last_injury"]
    days_inj_b = inj_b["days_since_last_injury"]
    inj_diff = (days_inj_a - days_inj_b) if not (np.isnan(days_inj_a) or np.isnan(days_inj_b)) else np.nan

    feat = {
        "elo_diff": elo_a["elo"] - elo_b["elo"],
        "elo_hard_diff": elo_a["elo_hard"] - elo_b["elo_hard"],
        "elo_clay_diff": elo_a["elo_clay"] - elo_b["elo_clay"],
        "elo_grass_diff": elo_a["elo_grass"] - elo_b["elo_grass"],
        "rank_a": rank_a, "rank_b": rank_b,
        **h2h,
        "form_a": stats_a.get("form", 0.5),
        "form_b": stats_b.get("form", 0.5),
        "form_surface_a": stats_a.get("form_surface", 0.5),
        "form_surface_b": stats_b.get("form_surface", 0.5),
        "days_since_last_a": stats_a.get("days_since_last", np.nan),
        "days_since_last_b": stats_b.get("days_since_last", np.nan),
        "first_serve_pct_a": stats_a.get("first_serve_pct", np.nan),
        "first_serve_pct_b": stats_b.get("first_serve_pct", np.nan),
        "second_serve_pct_a": stats_a.get("second_serve_pct", np.nan),
        "second_serve_pct_b": stats_b.get("second_serve_pct", np.nan),
        "bp_saved_pct_a": stats_a.get("bp_saved_pct", np.nan),
        "bp_saved_pct_b": stats_b.get("bp_saved_pct", np.nan),
        "aces_per_match_a": stats_a.get("aces_per_match", np.nan),
        "aces_per_match_b": stats_b.get("aces_per_match", np.nan),
        "df_per_match_a": stats_a.get("df_per_match", np.nan),
        "df_per_match_b": stats_b.get("df_per_match", np.nan),
        "surf_first_serve_pct_a": stats_a.get("surf_first_serve_pct", np.nan),
        "surf_first_serve_pct_b": stats_b.get("surf_first_serve_pct", np.nan),
        "surf_second_serve_pct_a": stats_a.get("surf_second_serve_pct", np.nan),
        "surf_second_serve_pct_b": stats_b.get("surf_second_serve_pct", np.nan),
        "surf_bp_saved_pct_a": stats_a.get("surf_bp_saved_pct", np.nan),
        "surf_bp_saved_pct_b": stats_b.get("surf_bp_saved_pct", np.nan),
        "days_since_last_injury_a": days_inj_a,
        "days_since_last_injury_b": days_inj_b,
        "is_returning_from_injury_a": inj_a["is_returning_from_injury"],
        "is_returning_from_injury_b": inj_b["is_returning_from_injury"],
        "retirement_rate_12m_a": inj_a["retirement_rate_12m"],
        "retirement_rate_12m_b": inj_b["retirement_rate_12m"],
        "days_since_last_injury_diff": inj_diff,
        "level_encoded": level_enc,
        "round_encoded": 4,  # default: QF
        "surface_encoded": surface_enc,
        "is_pre_grand_slam_tournament": int(0 <= gs_days_val <= 21),
        "gs_days": gs_days_val,
    }
    return pd.DataFrame([feat])


def predict(player_a: str, player_b: str, surface: str, level: str, tour: str):
    surface_cap = surface.capitalize()
    calibrated_path = ROOT / "models" / f"xgb_{surface.lower()}_calibrated.pkl"
    model_path = ROOT / "models" / f"xgb_{surface.lower()}.json"

    if not model_path.exists():
        print(f"\nModel not found: {model_path}")
        print("Run  python run_pipeline.py  first to train the models.")
        return

    if calibrated_path.exists():
        with open(calibrated_path, "rb") as f:
            model = pickle.load(f)
        model_label = "calibrated"
    else:
        model = xgb.XGBClassifier(**XGB_PARAMS)
        model.load_model(str(model_path))
        model_label = "base"

    h2h_path = ROOT / "data" / "processed" / "matches_with_h2h.parquet"
    if not h2h_path.exists():
        print("H2H data not found. Run run_pipeline.py first.")
        return
    h2h_df = pd.read_parquet(h2h_path)

    X = build_feature_vector(player_a, player_b, surface, level, tour, h2h_df)
    available = [c for c in FEATURE_COLS if c in X.columns]
    X_model = X[available].fillna(0)

    proba_a = float(model.predict_proba(X_model)[0][1])
    proba_b = 1.0 - proba_a

    print(f"\nSurface-Modell: {surface_cap} ({model_label})")
    print(f"{player_a}:  {proba_a*100:.1f}%")
    print(f"{player_b}: {proba_b*100:.1f}%")

    # Feature importance / top diffs
    base = getattr(model, "estimator", model)
    if hasattr(base, "feature_importances_"):
        imp = pd.Series(base.feature_importances_, index=available)
        top = imp.nlargest(5)
        print("\nTop Features (Δ):")
        for feat_name, _ in top.items():
            val = X[feat_name].iloc[0] if feat_name in X.columns else "N/A"
            if isinstance(val, float):
                print(f"  {feat_name:<40} {val:+.3f}")
            else:
                print(f"  {feat_name:<40} {val}")
    print()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Tennis match win-probability predictor")
    parser.add_argument("--player_a", required=True, help="Name of player A")
    parser.add_argument("--player_b", required=True, help="Name of player B")
    parser.add_argument("--surface", required=True, choices=["hard", "clay", "grass"])
    parser.add_argument("--level", default="atp250",
                        help="Tournament level: grandslam | masters | atp500 | atp250")
    parser.add_argument("--tour", default="atp", choices=["atp", "wta"])
    args = parser.parse_args()
    predict(args.player_a, args.player_b, args.surface, args.level, args.tour)
