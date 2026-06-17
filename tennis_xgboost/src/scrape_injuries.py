"""Phase 5: Scrape injury data from multiple sources."""
import logging
import re
from pathlib import Path

import pandas as pd
import requests
from bs4 import BeautifulSoup

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

ROOT = Path(__file__).parent.parent
OUTPUT = ROOT / "data" / "raw" / "injuries.csv"
HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; TennisBot/1.0)"}


def scrape_uts_injuries() -> list[dict]:
    """Scrape ultimatetennisstatistics.com injuries page."""
    url = "https://www.ultimatetennisstatistics.com/injuries"
    rows = []
    try:
        resp = requests.get(url, headers=HEADERS, timeout=30)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "lxml")
        table = soup.find("table")
        if table:
            headers_row = table.find("tr")
            col_names = [th.get_text(strip=True).lower() for th in headers_row.find_all(["th", "td"])] if headers_row else []
            for tr in table.find_all("tr")[1:]:
                cells = [td.get_text(strip=True) for td in tr.find_all("td")]
                if cells:
                    row = dict(zip(col_names, cells)) if col_names else {}
                    rows.append(row)
        log.info(f"UTS injuries: {len(rows)} rows")
    except Exception as e:
        log.warning(f"UTS scrape failed: {e}")
    return rows


def extract_retirement_proxy(api_parquet: Path) -> list[dict]:
    """Use api-tennis retirement/walkover data as injury proxy."""
    rows = []
    if not api_parquet.exists():
        return rows
    df = pd.read_parquet(api_parquet)
    for col in ["retirement", "walkover"]:
        if col in df.columns:
            mask = df[col].notna() & (df[col].astype(str).str.lower().isin(["true", "1", "yes", "a", "b"]))
            for _, row in df[mask].iterrows():
                for player_col in ["player_a", "player_b"]:
                    rows.append({
                        "player": row.get(player_col, ""),
                        "tour": row.get("tour", ""),
                        "injury_start": row.get("match_date", ""),
                        "injury_end": row.get("match_date", ""),
                        "injury_type": col,
                        "source": "api_tennis_proxy",
                    })
    log.info(f"Retirement proxy injuries: {len(rows)} rows")
    return rows


def scrape_injuries(force: bool = False) -> pd.DataFrame:
    if OUTPUT.exists() and not force:
        log.info(f"Skipping scrape_injuries – already exists")
        return pd.read_csv(OUTPUT)

    log.info("Phase 5c START: Scraping injury data")
    all_rows = []

    # Source 1: UTS
    uts_rows = scrape_uts_injuries()
    for r in uts_rows:
        all_rows.append({
            "player": r.get("player", r.get("name", "")),
            "tour": r.get("tour", ""),
            "injury_start": r.get("from", r.get("start", r.get("date", ""))),
            "injury_end": r.get("to", r.get("end", "")),
            "injury_type": r.get("injury", r.get("type", "")),
            "source": "uts",
        })

    # Source 3: api-tennis retirement proxy
    api_parquet = ROOT / "data" / "processed" / "matches_api.parquet"
    all_rows.extend(extract_retirement_proxy(api_parquet))

    df = pd.DataFrame(all_rows, columns=["player", "tour", "injury_start", "injury_end", "injury_type", "source"])
    for col in ["injury_start", "injury_end"]:
        df[col] = pd.to_datetime(df[col], errors="coerce")

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(OUTPUT, index=False)
    log.info(f"Phase 5c END: {len(df)} injury records → {OUTPUT}")
    return df


if __name__ == "__main__":
    scrape_injuries()
