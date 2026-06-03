import os
import json
import requests
from datetime import datetime, timezone

FRED_API_KEY = os.environ.get("FRED_API_KEY", "")
SERIES_ID    = "TSFR1M"   # 1-Month Term SOFR (CME Group, via FRED)
OUTPUT_PATH  = "Loan Calculator/sofr.json"

def fetch_from_fred(api_key):
    url = (
        f"https://api.stlouisfed.org/fred/series/observations"
        f"?series_id={SERIES_ID}"
        f"&api_key={api_key}"
        f"&file_type=json"
        f"&sort_order=desc"
        f"&limit=5"          # fetch last 5 to skip any missing/null entries
    )
    resp = requests.get(url, timeout=15)
    resp.raise_for_status()
    data = resp.json()
    # Use the most recent observation with a valid numeric value
    for obs in data["observations"]:
        try:
            value = float(obs["value"])
            return value, obs["date"]
        except (ValueError, TypeError):
            continue
    raise ValueError("No valid observations returned from FRED")

def load_existing():
    try:
        with open(OUTPUT_PATH) as f:
            return json.load(f)
    except Exception:
        return {"sofr": None, "date": None, "updated": None}

def main():
    existing = load_existing()

    if not FRED_API_KEY:
        print("ERROR: FRED_API_KEY environment variable not set.")
        raise SystemExit(1)

    try:
        sofr, date = fetch_from_fred(FRED_API_KEY)
        print(f"Fetched 1-Month Term SOFR: {sofr}% (as of {date})")
    except Exception as e:
        print(f"ERROR fetching from FRED: {e}")
        if existing.get("sofr") is not None:
            print(f"Keeping last known value: {existing['sofr']}%")
            raise SystemExit(0)   # don't fail the action, keep old value
        raise SystemExit(1)

    output = {
        "sofr":    sofr,
        "date":    date,
        "updated": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    }

    with open(OUTPUT_PATH, "w") as f:
        json.dump(output, f, indent=2)

    print(f"Written to {OUTPUT_PATH}: {json.dumps(output)}")

if __name__ == "__main__":
    main()
