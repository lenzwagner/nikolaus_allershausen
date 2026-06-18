"""Phase 3: Rolling serve stats and form features."""
import logging
from pathlib import Path

import numpy as np
import pandas as pd
import yaml

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

ROOT = Path(__file__).parent.parent
with open(ROOT / "config.yaml") as f:
    CFG = yaml.safe_load(f)

FORM_W = CFG["rolling"]["form_window"]
SERVE_W = CFG["rolling"]["serve_stats_window"]
SURF_FORM_W = CFG["rolling"]["surface_form_window"]
OUTPUT = ROOT / "data" / "processed" / "matches_with_rolling.parquet"


def safe_div(a, b, default=np.nan):
    try:
        if np.isnan(a) or np.isnan(b) or b == 0:
            return default
        return float(a) / float(b)
    except Exception:
        return default


def compute_rolling(df: pd.DataFrame, force: bool = False) -> pd.DataFrame:
    if OUTPUT.exists() and not force:
        log.info(f"Skipping features – output already exists: {OUTPUT}")
        return pd.read_parquet(OUTPUT)

    log.info("Phase 3 START: Computing rolling features")
    df = df.sort_values("match_date").reset_index(drop=True)

    # Track per-player rolling stats
    player_matches: dict[str, list] = {}

    result_rows = []
    for idx, row in df.iterrows():
        pa, pb = str(row.get("player_a", "")), str(row.get("player_b", ""))
        surface = str(row.get("surface", "Hard"))

        feats = {}
        for player, prefix in [(pa, "a"), (pb, "b")]:
            hist = player_matches.get(player, [])
            feats.update(_rolling_features(hist, prefix, surface, FORM_W, SERVE_W, SURF_FORM_W))

        result_rows.append({**row.to_dict(), **feats})

        # Update history for both players
        match_record = {
            "surface": surface,
            "won": True,  # placeholder; updated below
            "aces": row.get("aces_a", np.nan),
            "df": row.get("df_a", np.nan),
            "first_won": row.get("first_won_a", np.nan),
            "first_in": row.get("first_in_a", np.nan),
            "second_won": row.get("second_won_a", np.nan),
            "svpt": row.get("svpt_a", np.nan),
            "bp_faced": row.get("bp_faced_a", np.nan),
            "bp_saved": row.get("bp_saved_a", np.nan),
            "match_date": row.get("match_date"),
        }
        winner = str(row.get("winner", "A"))
        rec_a = {**match_record, "won": winner == "A",
                 "aces": row.get("aces_a", np.nan), "df": row.get("df_a", np.nan),
                 "first_won": row.get("first_won_a", np.nan), "first_in": row.get("first_in_a", np.nan),
                 "second_won": row.get("second_won_a", np.nan), "svpt": row.get("svpt_a", np.nan),
                 "bp_faced": row.get("bp_faced_a", np.nan), "bp_saved": row.get("bp_saved_a", np.nan)}
        rec_b = {**match_record, "won": winner == "B",
                 "aces": row.get("aces_b", np.nan), "df": row.get("df_b", np.nan),
                 "first_won": row.get("first_won_b", np.nan), "first_in": row.get("first_in_b", np.nan),
                 "second_won": row.get("second_won_b", np.nan), "svpt": row.get("svpt_b", np.nan),
                 "bp_faced": row.get("bp_faced_b", np.nan), "bp_saved": row.get("bp_saved_b", np.nan)}

        player_matches.setdefault(pa, []).append(rec_a)
        player_matches.setdefault(pb, []).append(rec_b)

    result = pd.DataFrame(result_rows)
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    result.to_parquet(OUTPUT, index=False)
    log.info(f"Phase 3 END: {len(result)} rows → {OUTPUT}")
    return result


def _rolling_features(hist: list, prefix: str, surface: str, form_w: int, serve_w: int, surf_form_w: int) -> dict:
    if not hist:
        return {
            f"form_{prefix}": np.nan,
            f"form_surface_{prefix}": np.nan,
            f"days_since_last_{prefix}": np.nan,
            f"first_serve_pct_{prefix}": np.nan,
            f"second_serve_pct_{prefix}": np.nan,
            f"bp_saved_pct_{prefix}": np.nan,
            f"aces_per_match_{prefix}": np.nan,
            f"df_per_match_{prefix}": np.nan,
            f"surf_first_serve_pct_{prefix}": np.nan,
            f"surf_second_serve_pct_{prefix}": np.nan,
            f"surf_bp_saved_pct_{prefix}": np.nan,
        }

    last_date = hist[-1].get("match_date")
    recent = hist[-form_w:]
    surf_recent = [m for m in hist if m.get("surface") == surface][-surf_form_w:]

    form = np.mean([m["won"] for m in recent]) if recent else np.nan
    form_surface = np.mean([m["won"] for m in surf_recent]) if surf_recent else np.nan

    import datetime
    days_since = np.nan
    if last_date is not None:
        try:
            ref = pd.Timestamp.today()
            days_since = (ref - pd.Timestamp(last_date)).days
        except Exception:
            pass

    serve_hist = hist[-serve_w:]

    def avg_stat(records, key):
        vals = [m.get(key, np.nan) for m in records if not np.isnan(m.get(key, np.nan))]
        return np.mean(vals) if vals else np.nan

    first_won = avg_stat(serve_hist, "first_won")
    first_in = avg_stat(serve_hist, "first_in")
    second_won = avg_stat(serve_hist, "second_won")
    svpt = avg_stat(serve_hist, "svpt")
    bp_faced = avg_stat(serve_hist, "bp_faced")
    bp_saved = avg_stat(serve_hist, "bp_saved")
    aces = avg_stat(serve_hist, "aces")
    df_val = avg_stat(serve_hist, "df")

    first_serve_pct = safe_div(first_won, first_in)
    second_serve_pct = safe_div(second_won, svpt - first_in if not np.isnan(svpt) and not np.isnan(first_in) else np.nan)
    bp_saved_pct = safe_div(bp_saved, bp_faced)

    # Surface-specific
    surf_first_won = avg_stat(surf_recent, "first_won")
    surf_first_in = avg_stat(surf_recent, "first_in")
    surf_second_won = avg_stat(surf_recent, "second_won")
    surf_svpt = avg_stat(surf_recent, "svpt")
    surf_bp_faced = avg_stat(surf_recent, "bp_faced")
    surf_bp_saved = avg_stat(surf_recent, "bp_saved")

    surf_first_serve_pct = safe_div(surf_first_won, surf_first_in)
    surf_second_serve_pct = safe_div(surf_second_won, surf_svpt - surf_first_in if not np.isnan(surf_svpt) and not np.isnan(surf_first_in) else np.nan)
    surf_bp_saved_pct = safe_div(surf_bp_saved, surf_bp_faced)

    return {
        f"form_{prefix}": form,
        f"form_surface_{prefix}": form_surface,
        f"days_since_last_{prefix}": days_since,
        f"first_serve_pct_{prefix}": first_serve_pct,
        f"second_serve_pct_{prefix}": second_serve_pct,
        f"bp_saved_pct_{prefix}": bp_saved_pct,
        f"aces_per_match_{prefix}": aces,
        f"df_per_match_{prefix}": df_val,
        f"surf_first_serve_pct_{prefix}": surf_first_serve_pct,
        f"surf_second_serve_pct_{prefix}": surf_second_serve_pct,
        f"surf_bp_saved_pct_{prefix}": surf_bp_saved_pct,
    }


if __name__ == "__main__":
    df = pd.read_parquet(ROOT / "data" / "processed" / "matches_with_elo.parquet")
    compute_rolling(df)
