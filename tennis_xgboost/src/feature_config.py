"""Single source of truth for the model feature list.

Import FEATURE_COLS from here instead of redefining it in every module.
"""

FEATURE_COLS = [
    # ELO
    "elo_diff", "elo_hard_diff", "elo_clay_diff", "elo_grass_diff",
    # Ranking
    "rank_a", "rank_b",
    # H2H
    "h2h_count", "h2h_win_rate_a", "h2h_surface_count", "h2h_surface_win_rate_a",
    # Form
    "form_a", "form_b", "form_surface_a", "form_surface_b",
    "days_since_last_a", "days_since_last_b",
    # Serve stats
    "first_serve_pct_a", "first_serve_pct_b",
    "second_serve_pct_a", "second_serve_pct_b",
    "bp_saved_pct_a", "bp_saved_pct_b",
    "aces_per_match_a", "aces_per_match_b",
    "df_per_match_a", "df_per_match_b",
    # Surface-specific serve stats
    "surf_first_serve_pct_a", "surf_first_serve_pct_b",
    "surf_second_serve_pct_a", "surf_second_serve_pct_b",
    "surf_bp_saved_pct_a", "surf_bp_saved_pct_b",
    # Injury
    "days_since_last_injury_a", "days_since_last_injury_b",
    "is_returning_from_injury_a", "is_returning_from_injury_b",
    "retirement_rate_12m_a", "retirement_rate_12m_b",
    "days_since_last_injury_diff",
    # Tournament context
    "level_encoded", "round_encoded", "surface_encoded",
    "is_pre_grand_slam_tournament", "gs_days",
    # Home advantage
    "home_advantage_diff",
]
