from __future__ import annotations

import argparse
import os
import time
from collections import defaultdict, deque
from typing import Dict, Iterable, List, Optional, Tuple

import gspread
from dotenv import load_dotenv
from oauth2client.service_account import ServiceAccountCredentials


load_dotenv(override=True)

GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID", "").strip()
SCHEDULE_SHEET_ID = os.getenv("SCHEDULE_SHEET_ID", "").strip()
GOOGLE_CREDENTIALS_FILE = os.getenv("GOOGLE_CREDENTIALS_FILE", "google_credentials.json")

SCOPE = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
]

REPORT_TAB = "Schedule Check"
REPORT_HEADER = [
    "Season",
    "Date",
    "Scheduled Time",
    "Away",
    "Home",
    "Game Code",
    "Status",
    "Match ID",
    "Played Time",
    "Actual Team 1",
    "Score",
    "Actual Team 2",
    "Lagout",
    "Notes",
]


def authorize() -> gspread.Client:
    creds = ServiceAccountCredentials.from_json_keyfile_name(GOOGLE_CREDENTIALS_FILE, SCOPE)
    return gspread.authorize(creds)


def index_map(header: List[str]) -> Dict[str, int]:
    return {name.strip(): i for i, name in enumerate(header)}


def cell(row: List[str], idx: Dict[str, int], name: str) -> str:
    pos = idx.get(name)
    if pos is None or pos >= len(row):
        return ""
    return str(row[pos]).strip()


def normalize_name(value: str) -> str:
    text = " ".join(str(value or "").lower().replace("-", " ").split())
    for prefix in ("cwnhl ", "cwhl ", "wca "):
        if text.startswith(prefix):
            text = text[len(prefix) :]
    aliases = {
        "columbus": "columbus blue jackets",
        "golden knights": "vegas golden knights",
        "vegas knights": "vegas golden knights",
        "la kings": "la kings",
        "los angeles kings": "la kings",
    }
    return aliases.get(text, text)


def ensure_report_tab(api_book: gspread.Spreadsheet) -> gspread.Worksheet:
    try:
        ws = api_book.worksheet(REPORT_TAB)
    except gspread.exceptions.WorksheetNotFound:
        ws = api_book.add_worksheet(title=REPORT_TAB, rows="500", cols=str(len(REPORT_HEADER)))
    return ws


def read_schedule_teams(schedule_book: gspread.Spreadsheet) -> Tuple[Dict[str, str], Dict[str, str]]:
    rows = schedule_book.worksheet("Teams").get_all_values()
    if not rows:
        return {}, {}

    idx = index_map(rows[0])
    name_to_abbr = {}
    abbr_to_name = {}
    for row in rows[1:]:
        name = cell(row, idx, "Team Name")
        abbr = cell(row, idx, "Abbreviation").upper()
        if not name or not abbr:
            continue
        name_to_abbr[normalize_name(name)] = abbr
        abbr_to_name[abbr] = name
    return name_to_abbr, abbr_to_name


def read_api_team_map(
    api_book: gspread.Spreadsheet,
    schedule_name_to_abbr: Dict[str, str],
) -> Dict[str, str]:
    rows = api_book.worksheet("Team List").get_all_values()
    if not rows:
        return {}

    idx = index_map(rows[0])
    club_id_to_abbr = {}
    for row in rows[1:]:
        club_id = cell(row, idx, "Club ID")
        abbr = cell(row, idx, "Abbreviation").upper()
        team_name = normalize_name(cell(row, idx, "Team Name"))
        if not abbr and team_name in schedule_name_to_abbr:
            abbr = schedule_name_to_abbr[team_name]
        if club_id and abbr:
            club_id_to_abbr[club_id] = abbr
    return club_id_to_abbr


def build_played_games(
    api_book: gspread.Spreadsheet,
    club_id_to_abbr: Dict[str, str],
    abbr_to_name: Dict[str, str],
) -> Tuple[Dict[Tuple[str, Tuple[str, str]], deque], List[dict]]:
    rows = api_book.worksheet("Game Log").get_all_values()
    if not rows:
        return defaultdict(deque), []

    idx = index_map(rows[0])
    played_by_key = defaultdict(deque)
    all_played = []

    for row in rows[1:]:
        date_time = cell(row, idx, "Date")
        date = date_time[:10]
        if not date:
            continue

        id1 = cell(row, idx, "Team 1 ID")
        id2 = cell(row, idx, "Team 2 ID")
        abbr1 = club_id_to_abbr.get(id1)
        abbr2 = club_id_to_abbr.get(id2)
        if not abbr1 or not abbr2:
            continue

        game = {
            "date": date,
            "played_time": date_time[11:16],
            "match_id": cell(row, idx, "Match ID"),
            "team1": abbr_to_name.get(abbr1, cell(row, idx, "Team 1")),
            "team2": abbr_to_name.get(abbr2, cell(row, idx, "Team 2")),
            "score1": cell(row, idx, "Team 1 Score"),
            "score2": cell(row, idx, "Team 2 Score"),
            "lagout": cell(row, idx, "Lagout"),
            "stitched_ids": cell(row, idx, "Stitched Match IDs"),
            "key": (date, tuple(sorted([abbr1, abbr2]))),
            "used": False,
        }
        played_by_key[game["key"]].append(game)
        all_played.append(game)

    return played_by_key, all_played


def schedule_rows(
    schedule_book: gspread.Spreadsheet,
    schedule_name_to_abbr: Dict[str, str],
    tab_name: str,
) -> Iterable[dict]:
    try:
        rows = schedule_book.worksheet(tab_name).get_all_values()
    except gspread.exceptions.WorksheetNotFound:
        return []
    if not rows:
        return []

    idx = index_map(rows[0])
    season = "Playoffs" if "Playoff" in tab_name else "Regular Season"
    output = []
    for row in rows[1:]:
        date = cell(row, idx, "Date")
        away = cell(row, idx, "Away")
        home = cell(row, idx, "Home")
        away_abbr = schedule_name_to_abbr.get(normalize_name(away))
        home_abbr = schedule_name_to_abbr.get(normalize_name(home))
        if not date or not away or not home:
            continue

        output.append(
            {
                "season": season,
                "date": date,
                "time": cell(row, idx, "Time"),
                "away": away,
                "home": home,
                "game_code": cell(row, idx, "Game Code"),
                "key": (date, tuple(sorted([away_abbr or away, home_abbr or home]))),
            }
        )
    return output


def build_report_rows(api_book: gspread.Spreadsheet, schedule_book: gspread.Spreadsheet) -> List[List[str]]:
    schedule_name_to_abbr, abbr_to_name = read_schedule_teams(schedule_book)
    club_id_to_abbr = read_api_team_map(api_book, schedule_name_to_abbr)
    played_by_key, all_played = build_played_games(api_book, club_id_to_abbr, abbr_to_name)

    report = [REPORT_HEADER]
    scheduled_count = matched_count = 0
    scheduled_dates = set()
    scheduled_keys = set()

    for tab in ("Regular Season Schedule", "Playoff Schedule"):
        for scheduled in schedule_rows(schedule_book, schedule_name_to_abbr, tab):
            scheduled_count += 1
            scheduled_dates.add(scheduled["date"])
            scheduled_keys.add(scheduled["key"])
            played = played_by_key.get(scheduled["key"], deque())
            if played:
                game = played.popleft()
                game["used"] = True
                matched_count += 1
                if game["lagout"].lower() == "yes" and not game["stitched_ids"]:
                    status = "Unstitched Lagout"
                    notes = "Lagout was pulled, but no stitched continuation is attached"
                else:
                    status = "Played"
                    notes = (
                        f"Matched by date and teams; stitched with {game['stitched_ids']}"
                        if game["stitched_ids"]
                        else "Matched by date and teams"
                    )
                match_id = game["match_id"]
                played_time = game["played_time"]
                actual_team1 = game["team1"]
                actual_team2 = game["team2"]
                score = f"{game['score1']}-{game['score2']}"
                lagout = game["lagout"]
            else:
                status = "Missing"
                notes = "Scheduled game not found in Game Log"
                match_id = played_time = actual_team1 = actual_team2 = score = lagout = ""

            report.append(
                [
                    scheduled["season"],
                    scheduled["date"],
                    scheduled["time"],
                    scheduled["away"],
                    scheduled["home"],
                    scheduled["game_code"],
                    status,
                    match_id,
                    played_time,
                    actual_team1,
                    score,
                    actual_team2,
                    lagout,
                    notes,
                ]
            )

    extra_count = 0
    for game in all_played:
        if game["used"] or game["date"] not in scheduled_dates:
            continue
        extra_count += 1
        has_scheduled_matchup = game["key"] in scheduled_keys
        status = "Possible Lagout Extra" if has_scheduled_matchup and game["lagout"].lower() == "yes" else "Extra Played"
        notes = (
            "Same teams already matched a scheduled row; review if this should be stitched"
            if status == "Possible Lagout Extra"
            else "Played game did not match a scheduled row"
        )
        report.append(
            [
                "",
                game["date"],
                "",
                "",
                "",
                "",
                status,
                game["match_id"],
                game["played_time"],
                game["team1"],
                f"{game['score1']}-{game['score2']}",
                game["team2"],
                game["lagout"],
                notes,
            ]
        )

    report.append([])
    report.append(["Summary", "Scheduled", scheduled_count, "Matched", matched_count, "Extra Played", extra_count])
    return report


def update_schedule_check() -> None:
    if not GOOGLE_SHEET_ID:
        raise RuntimeError("GOOGLE_SHEET_ID is missing from .env")
    if not SCHEDULE_SHEET_ID:
        raise RuntimeError("SCHEDULE_SHEET_ID is missing from .env")

    gc = authorize()
    api_book = gc.open_by_key(GOOGLE_SHEET_ID)
    schedule_book = gc.open_by_key(SCHEDULE_SHEET_ID)
    report = build_report_rows(api_book, schedule_book)

    ws = ensure_report_tab(api_book)
    last_error = None
    for attempt in range(1, 4):
        try:
            ws.clear()
            ws.update(values=report, range_name="A1", value_input_option="USER_ENTERED")
            ws.freeze(rows=1)
            ws.format("1:1", {"textFormat": {"bold": True}})
            last_error = None
            break
        except gspread.exceptions.APIError as exc:
            last_error = exc
            if attempt == 3:
                raise
            time.sleep(attempt * 2)

    summary = report[-1] if len(report) >= 2 else []
    print(f"Schedule Check updated: {summary}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Cross-reference scraped games against the CWHL schedule.")
    parser.parse_args()
    update_schedule_check()


if __name__ == "__main__":
    main()
