"""Phase 5: Scrape ELO ratings from tennisabstract.com."""
import logging
from pathlib import Path

import pandas as pd
import requests
from bs4 import BeautifulSoup

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

ROOT = Path(__file__).parent.parent
HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; TennisBot/1.0)"}

URLS = {
    "atp": "https://tennisabstract.com/reports/atp_elo_ratings.html",
    "wta": "https://tennisabstract.com/reports/wta_elo_ratings.html",
}
OUTPUTS = {
    "atp": ROOT / "data" / "raw" / "elo_atp.csv",
    "wta": ROOT / "data" / "raw" / "elo_wta.csv",
}


def scrape_elo(tour: str, force: bool = False) -> pd.DataFrame:
    output = OUTPUTS[tour]
    if output.exists() and not force:
        log.info(f"Skipping elo_{tour} – already exists")
        return pd.read_csv(output)

    log.info(f"Scraping {tour.upper()} ELO ratings")
    try:
        resp = requests.get(URLS[tour], headers=HEADERS, timeout=30)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "lxml")
    except Exception as e:
        log.warning(f"Failed to scrape ELO for {tour}: {e}")
        return pd.DataFrame(columns=["player", "elo", "elo_hard", "elo_clay", "elo_grass", "tour"])

    rows = []
    table = soup.find("table", {"id": "reportable"}) or soup.find("table")
    if table:
        headers_row = table.find("tr")
        col_names = [th.get_text(strip=True).lower() for th in headers_row.find_all(["th", "td"])] if headers_row else []
        for tr in table.find_all("tr")[1:]:
            cells = [td.get_text(strip=True) for td in tr.find_all("td")]
            if cells:
                row = dict(zip(col_names, cells)) if col_names else {"raw": cells}
                row["tour"] = tour
                rows.append(row)

    df = pd.DataFrame(rows)
    # Normalize expected columns
    col_map = {}
    for col in df.columns:
        cl = col.lower()
        if "player" in cl or "name" in cl:
            col_map[col] = "player"
        elif "hard" in cl:
            col_map[col] = "elo_hard"
        elif "clay" in cl:
            col_map[col] = "elo_clay"
        elif "grass" in cl:
            col_map[col] = "elo_grass"
        elif "elo" in cl and col not in col_map:
            col_map[col] = "elo"
    df = df.rename(columns=col_map)

    output.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(output, index=False)
    log.info(f"Saved {len(df)} ELO rows → {output}")
    return df


def scrape_all(force: bool = False):
    for tour in ["atp", "wta"]:
        scrape_elo(tour, force=force)


if __name__ == "__main__":
    scrape_all()
