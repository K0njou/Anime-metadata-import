import requests
import pandas as pd
import time
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

API_URL = "https://graphql.anilist.co"
INPUT_FILE = "bajki.xlsx"
OUTPUT_FILE = "bajki-done.xlsx"
MAX_REQUESTS = 30
SLEEP_SHORT = 2
SLEEP_LONG = 5

GRAPHQL_QUERY = """
query($name: String) {
  Media(search: $name, type: ANIME, format_in: [TV, TV_SHORT, MOVIE, SPECIAL, ONA, OVA]) {
    season
    seasonYear
    siteUrl
    studios {
      edges {
        isMain
        node { name }
      }
    }
    genres
    tags { name }
  }
}
"""

def fetch_anime_data(name):
    try:
        response = requests.post(API_URL, json={"query": GRAPHQL_QUERY, "variables": {"name": name}})
        response.raise_for_status()
        return response.json()
    except requests.RequestException as e:
        print(f"❌ Error fetching '{name}': {e}")
        return None

def parse_anime_data(data):
    media = data.get("data", {}).get("Media", {})
    if not media:
        return ["Unknown"] * 5

    season = media.get("season")
    year = media.get("seasonYear")
    season_year = f"{season} {year}" if season and year else "Unknown"

    studios = media.get("studios", {}).get("edges", [])
    studio = next((s["node"]["name"] for s in studios if s.get("isMain")), None)
    if not studio and studios:
        studio = studios[0]["node"]["name"]
    studio = studio or "Unknown"

    genres = ", ".join(media.get("genres", [])) or "Unknown"
    tags = ", ".join(tag["name"] for tag in media.get("tags", [])[:5]) or "Unknown"
    link = media.get("siteUrl", "")

    return season_year, studio, genres, tags, link

def ensure_columns(df, columns):
    for col in columns:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].astype(str)

def update_dataframe(df):
    request_counter = 0
    for idx, row in df.iterrows():
        name = row["Bajka"]
        print(f"🔍 Fetching: {name}")
        data = fetch_anime_data(name)
        details = parse_anime_data(data) if data else ["Unknown"] * 5
        df.loc[idx, ["Sezon", "Studio", "Gatunki", "Tagi", "Link"]] = details

        request_counter += 1
        time.sleep(SLEEP_LONG if request_counter >= MAX_REQUESTS else SLEEP_SHORT)
        if request_counter >= MAX_REQUESTS:
            request_counter = 0

def highlight_unknowns(file_path):
    wb = load_workbook(file_path)
    ws = wb.active
    fill = PatternFill(start_color="FFFF9999", end_color="FFFF9999", fill_type="solid")

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        if any(cell.value == "Unknown" for cell in row):
            for cell in row:
                cell.fill = fill

    wb.save(file_path)

def main():
    df = pd.read_excel(INPUT_FILE)
    ensure_columns(df, ["Sezon", "Studio", "Gatunki", "Tagi", "Link"])
    update_dataframe(df)
    df.to_excel(OUTPUT_FILE, index=False, engine="openpyxl")
    highlight_unknowns(OUTPUT_FILE)
    print(f"✅ Zapisano dane do: {OUTPUT_FILE}")
    os.startfile(OUTPUT_FILE)

if __name__ == "__main__":
    main()
