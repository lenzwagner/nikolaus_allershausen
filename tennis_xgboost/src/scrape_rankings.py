"""Phase 5: Scrape ATP/WTA rankings from live-tennis.eu."""
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
    "atp": "https://live-tennis.eu/de/atp-weltrangliste-live",
    "wta": "https://live-tennis.eu/de/wta-weltrangliste-live",
}
OUTPUTS = {
    "atp": ROOT / "data" / "raw" / "rankings_atp.csv",
    "wta": ROOT / "data" / "raw" / "rankings_wta.csv",
}


def scrape_rankings(tour: str, force: bool = False) -> pd.DataFrame:
    output = OUTPUTS[tour]
    if output.exists() and not force:
        log.info(f"Skipping rankings_{tour} – already exists")
        return pd.read_csv(output)

    log.info(f"Scraping {tour.upper()} rankings")
    try:
        resp = requests.get(URLS[tour], headers=HEADERS, timeout=30)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "lxml")
    except Exception as e:
        log.warning(f"Failed to scrape rankings for {tour}: {e}")
        return pd.DataFrame(columns=["rank", "player", "points", "tour"])

    rows = []
    table = soup.find("table")
    if table:
        for tr in table.find_all("tr")[1:]:
            cells = [td.get_text(strip=True) for td in tr.find_all(["td", "th"])]
            if len(cells) >= 2:
                rows.append({"rank": cells[0], "player": cells[1],
                             "points": cells[2] if len(cells) > 2 else None, "tour": tour})

    df = pd.DataFrame(rows)
    output.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(output, index=False)
    log.info(f"Saved {len(df)} rankings → {output}")
    return df


def scrape_all(force: bool = False):
    for tour in ["atp", "wta"]:
        scrape_rankings(tour, force=force)


if __name__ == "__main__":
    scrape_all()
