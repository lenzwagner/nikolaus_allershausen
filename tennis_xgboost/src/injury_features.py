"""Phase 7: Build injury-related features from injuries.csv and match data."""
import logging
from pathlib import Path

import numpy as np
import pandas as pd
from rapidfuzz import process as fz_process

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

ROOT = Path(__file__).parent.parent
INJURY_CSV = ROOT / "data" / "raw" / "injuries.csv"


def normalize_name(name: str) -> str:
    return str(name).lower().strip()


def load_injuries() -> pd.DataFrame:
    if not INJURY_CSV.exists():
        log.warning("injuries.csv not found; injury features will be NaN")
        return pd.DataFrame(columns=["player", "tour", "injury_start", "injury_end", "injury_type"])
    df = pd.read_csv(INJURY_CSV, parse_dates=["injury_start", "injury_end"])
    df["player_norm"] = df["player"].apply(normalize_name)
    return df


def match_player(name: str, known_names: list[str], threshold: int = 80) -> str | None:
    if not known_names:
        return None
    result = fz_process.extractOne(normalize_name(name), known_names, score_cutoff=threshold)
    return result[0] if result else None


def compute_injury_features(match_date: pd.Timestamp, player: str, injuries: pd.DataFrame, api_df: pd.DataFrame | None = None) -> dict:
    """Compute injury features for a player before a given match date."""
    player_norm = normalize_name(player)
    known = injuries["player_norm"].tolist() if not injuries.empty else []
    matched = match_player(player, known)

    days_since_last_injury = np.nan
    is_returning_from_injury = 0
    retirement_rate_12m = 0.0

    if matched and not injuries.empty:
        player_inj = injuries[injuries["player_norm"] == matched]
        past_inj = player_inj[player_inj["injury_end"] < match_date].dropna(subset=["injury_end"])
        if not past_inj.empty:
            latest_end = past_inj["injury_end"].max()
            days_since_last_injury = (match_date - latest_end).days
            is_returning_from_injury = int(days_since_last_injury <= 30)

    if api_df is not None and not api_df.empty:
        cutoff = match_date - pd.Timedelta(days=365)
        recent = api_df[
            (api_df["match_date"] >= cutoff) & (api_df["match_date"] < match_date)
        ]
        p_matches = recent[
            (recent["player_a"].apply(normalize_name) == player_norm) |
            (recent["player_b"].apply(normalize_name) == player_norm)
        ]
        if len(p_matches) > 0:
            ret_cols = ["retirement", "walkover"]
            ret_count = 0
            for col in ret_cols:
                if col in p_matches.columns:
                    ret_count += p_matches[col].astype(str).str.lower().isin(["true", "1", "yes", "a", "b"]).sum()
            retirement_rate_12m = ret_count / len(p_matches)

    return {
        "days_since_last_injury": days_since_last_injury,
        "is_returning_from_injury": is_returning_from_injury,
        "retirement_rate_12m": retirement_rate_12m,
    }


def add_injury_features(df: pd.DataFrame, force: bool = False) -> pd.DataFrame:
    log.info("Phase 7 START: Adding injury features")
    injuries = load_injuries()
    api_path = ROOT / "data" / "processed" / "matches_api.parquet"
    api_df = pd.read_parquet(api_path) if api_path.exists() else None

    rows = []
    for _, row in df.iterrows():
        md = pd.Timestamp(row.get("match_date")) if pd.notna(row.get("match_date")) else pd.Timestamp.today()
        feats_a = compute_injury_features(md, str(row.get("player_a", "")), injuries, api_df)
        feats_b = compute_injury_features(md, str(row.get("player_b", "")), injuries, api_df)
        new_feats = {
            "days_since_last_injury_a": feats_a["days_since_last_injury"],
            "days_since_last_injury_b": feats_b["days_since_last_injury"],
            "is_returning_from_injury_a": feats_a["is_returning_from_injury"],
            "is_returning_from_injury_b": feats_b["is_returning_from_injury"],
            "retirement_rate_12m_a": feats_a["retirement_rate_12m"],
            "retirement_rate_12m_b": feats_b["retirement_rate_12m"],
            "days_since_last_injury_diff": (feats_a["days_since_last_injury"] - feats_b["days_since_last_injury"])
                if not (np.isnan(feats_a["days_since_last_injury"]) or np.isnan(feats_b["days_since_last_injury"])) else np.nan,
        }
        rows.append({**row.to_dict(), **new_feats})

    result = pd.DataFrame(rows)
    log.info(f"Phase 7 END: {len(result)} rows with injury features")
    return result
