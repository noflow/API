



import os
import json
import requests
import gspread
from datetime import datetime
import math
from oauth2client.service_account import ServiceAccountCredentials
from dotenv import load_dotenv

# ── LOAD ENV & SETUP GOOGLE SHEETS ──

load_dotenv()
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID")
GOOGLE_CREDENTIALS_FILE = "google_credentials.json"

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(GOOGLE_CREDENTIALS_FILE, scope)
client = gspread.authorize(creds)
sheet = client.open_by_key(GOOGLE_SHEET_ID)

# ── COLUMN MAPPING UTILS ──

def map_plus_minus_column(log_idx):
    plus_variants = ["'+/-'", "+/-", " '+/-", "Plus Minus", "plusminus", "plus_minus"]
    for variant in plus_variants:
        if variant in log_idx:
            log_idx["Plus Minus"] = log_idx[variant]
            print(f"ℹ️ Mapped column '{variant}' to 'Plus Minus' for aggregation.")
            return
    print("⚠️ 'Plus Minus' column still not found; skipping aggregation for plus/minus.")

# ── PLUS MINUS COLUMN NORMALIZATION ──

def normalize_plus_minus_column():
    try:
        ws = sheet.worksheet("Skater Log")
        headers = ws.row_values(1)
        updated = False
        for i, col in enumerate(headers):
            if any(sym in col.lower() for sym in ['+/-', 'plusminus', 'plus_minus', 'plus minus']):
                headers[i] = "Plus Minus"
                updated = True
        if updated:
            ws.update("1:1", [headers])
            print("✅ Header row updated: column normalized to 'Plus Minus'.")
        else:
            print("ℹ️ No recognizable plus/minus column found in Skater Log.")
    except Exception as e:
        pass  # inserted to complete except block

# ── TEAM LIST UTILS ──

def get_team_list():
    try:
        team_tab = sheet.worksheet("Team List")
        records = team_tab.get_all_records()
        return {str(row["Club ID"]).strip() for row in records if row.get("Club ID")}
    except Exception as e:
        print("❌ Failed to read 'Team List' tab:", e)
        return set()



def get_existing_game_match_ids():
    try:
        ws = sheet.worksheet("Game Log")
        data = ws.get_all_values()[1:]
        return set(r[0] for r in data)
    except gspread.exceptions.WorksheetNotFound:
        return set()

# ── EA API FETCH ──


def detect_lagout(match):
    players = match.get("players", {})
    clubs = match.get("clubs", {})
    mercy_threshold = 8
    toi_threshold = 55  # Average TOI in minutes to consider low enough for a lagout

    avg_toi_by_team = {}

    for club_id, team_players in players.items():
        if club_id not in clubs:
            continue

        total_toi = 0
        count = 0
        for pid, pdata in team_players.items():
            if pdata.get("position", "").lower() == "goalie":
                continue
            try:
                total_toi += int(pdata.get("toi", 0))
                count += 1
            except:
                continue

        avg = total_toi / count if count > 0 else 0
        avg_toi_by_team[club_id] = avg

    # Check for mercy rule
    mercy = False
    for club_data in clubs.values():
        try:
            goals = int(club_data.get("goals", 0))
            opp   = int(club_data.get("opponentScore", 0))
            if abs(goals - opp) >= mercy_threshold:
                mercy = True
        except:
            continue

    if not mercy and all(avg_toi < toi_threshold for avg_toi in avg_toi_by_team.values()):
        return "yes"
    return "no"




def get_private_matches(club_id):
    """
    Hits the EA Pro Clubs API to fetch that club's private matches.
    Returns a list of match JSON objects (or empty list on error).
    """
    url = (
        f"https://proclubs.ea.com/api/nhl/clubs/matches"
        f"?matchType=club_private&platform=common-gen5&clubIds={club_id}"
    )
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept": "application/json",
        "Referer": (
            f"https://www.ea.com/games/nhl/nhl-25/pro-clubs/match-history?"
            f"clubId={club_id}&platform=common-gen5"
        )
    }
    try:
        res = requests.get(url, headers=headers, timeout=15)
        if res.status_code != 200:
            print(f"❌ EA API returned {res.status_code} for club {club_id}")
            return []
        return res.json()
    except Exception as e:
        print(f"❌ Exception fetching matches for club {club_id}:", e)
        return []

# ── DEBUG: DUMP RAW JSON FOR A SPECIFIC MATCH (CALL MANUALLY IF NEEDED) ──

def dump_match_json(all_matches, target_mid):
    """
    Finds the first match whose matchId == target_mid and prints out its full JSON,
    including nested keys. Designed for manual inspection.
    """
    from pprint import pprint
    found = False

    for m in all_matches:
        if str(m.get("matchId")) == str(target_mid):
            found = True
            print(f"\n\n===== RAW JSON FOR MATCH {target_mid} =====\n")
            pprint(m)
            print("\n\n===== ALL TOP-LEVEL KEYS =====\n")
            pprint(list(m.keys()))
            print("\n\n===== CHECKING FOR 'result' / 'winnerByDnf' / GOAL DIFFERENCE =====\n")
            clubs_dict = m.get("clubs", {})
            for club_id, club_data in clubs_dict.items():
                code = club_data.get("result", "")
                wbd = club_data.get("winnerByDnf", "")
                goals = int(club_data.get("goals", 0))
                opp_goals = int(club_data.get("opponentScore", 0))
                diff = abs(goals - opp_goals)
                print(f"  → Club {club_id}: result code = {code}, winnerByDnf = {wbd}, goal diff = {diff}")
            break

    if not found:
        print(f"⚠️ Could not find matchId {target_mid} in all_matches.")

# ── GAME LOG ──

def create_or_fetch_game_log():
    headers = [
        "Match ID", "Date", "Lagout",
        "Team 1", "Team 1 ID", "Team 1 Score", "Team 1 PPG", "Team 1 PPO",
        "Team 2", "Team 2 ID", "Team 2 Score", "Team 2 PPG", "Team 2 PPO",
        "Team 1 Result", "Team 2 Result", "OT", "Ended Early", "Pushed Timestamp"
    ]
    try:
        ws = sheet.worksheet("Game Log")
        if not ws.get_all_values():
            ws.append_row(headers)
        return ws
    except gspread.exceptions.WorksheetNotFound:
        ws = sheet.add_worksheet(title="Game Log", rows="1000", cols="40")
        ws.append_row(headers)
        ws.format("1:1", {"textFormat": {"bold": True}})
        ws.freeze(rows=1)
        return ws


def parse_game_row(match, valid_ids):
    """
    Given a match JSON object and valid_ids set, extract a single row for Game Log:
      [Match ID, Date, Lagout, Team1 Name, Team1 ID, Team1 Score, Team1 PPG, Team1 PPO,
       Team2 Name, Team2 ID, Team2 Score, Team2 PPG, Team2 PPO, Team1 Result, Team2 Result,
       OT, Ended Early, Pushed Timestamp]

    Returns that row as a list, or None if:
      - clubs data is malformed (not exactly 2 clubs)
      - either team’s Club ID is not in valid_ids
    """
    clubs = match.get("clubs", {})
    club_ids = list(clubs.keys())
    if len(club_ids) != 2:
        return None

    team1_id, team2_id = club_ids
    # Skip if either team is not in Team List
    if (team1_id not in valid_ids) or (team2_id not in valid_ids):
        return None

    match_id = str(match.get("matchId", ""))
    timestamp = match.get("timestamp", 0)
    readable_date = datetime.fromtimestamp(timestamp).strftime("%Y-%m-%d %H:%M")
    pushed_time = datetime.now().isoformat()

    t1_raw = clubs[team1_id]
    t2_raw = clubs[team2_id]

    # ── EXTRACT SCORES AND RESULT CODES ──
    try:
        s1 = int(t1_raw.get("score", 0))
        o1 = int(t1_raw.get("opponentScore", 0))
        s2 = int(t2_raw.get("score", 0))
        o2 = int(t2_raw.get("opponentScore", 0))
    except:
        s1 = o1 = s2 = o2 = 0

    try:
        code1 = int(t1_raw.get("result", 0))
    except ValueError:
        code1 = 0
    try:
        code2 = int(t2_raw.get("result", 0))
    except ValueError:
        code2 = 0

    diff = abs(s1 - o1)

    # ── DETECT OT OR SHOOTOUT ──
    # EA uses bit‐2 (value 4) for OT and bit‐3 (value 8) for shootout.
    # Only count OT if that bit is set AND final goal difference == 1.
    ot_flag = False
    if ((code1 & 4) != 0 or (code2 & 4) != 0 or (code1 & 8) != 0 or (code2 & 8) != 0) and diff == 1:
        ot_flag = True

    # ── DETECT “ENDED EARLY” BY MERCY OR LAG‐OUT ──
    ended_early = False
    scores = []
    for club_data in (t1_raw, t2_raw):
        try:
            goals = int(club_data.get("goals", 0))
            opp   = int(club_data.get("opponentScore", 0))
        except:
            goals = opp = 0
        wbd = club_data.get("winnerByDnf", "0")
        scores.append((goals, opp, wbd))

    # Mercy: difference of 8 or more goals
    for goals, opp, _ in scores:
        if abs(goals - opp) >= 8:
            ended_early = True
            break

    # Lag‐out (winnerByDnf == '1')
    if not ended_early:
        for _, _, wbd in scores:
            if wbd == "1":
                ended_early = True
                break

    def extract(club_data):
        d = club_data.get("details", {})
        name = d.get("name", "Unknown")
        cid_str = str(d.get("clubId", "Unknown"))
        score_str = club_data.get("score", "0")
        ppg_str = club_data.get("ppg", "0")
        ppo_str = club_data.get("ppo", "0")
        try:
            s = int(club_data.get("score", 0))
            o = int(club_data.get("opponentScore", 0))
        except:
            s = o = 0
        result_text = "Win" if s > o else "Loss"
        return {
            "name": name,
            "id": cid_str,
            "score": score_str,
            "ppg": ppg_str,
            "ppo": ppo_str,
            "result": result_text
        }

    team1 = extract(t1_raw)
    team2 = extract(t2_raw)

    return [
        match_id,
        readable_date,
        detect_lagout(match),          # Lagout detected via TOI and mercy rule
        team1["name"],
        team1["id"],
        team1["score"],
        team1["ppg"],
        team1["ppo"],
        team2["name"],
        team2["id"],
        team2["score"],
        team2["ppg"],
        team2["ppo"],
        team1["result"],
        team2["result"],
        "yes" if ot_flag else "no",    # OT → only if bit 4 or 8 set AND diff == 1
        "yes" if ended_early else "no",# Ended Early → mercy or DNF
        pushed_time
    ]

def log_game_data(all_matches, valid_ids):
    """
    Recreates 'Game Log' sheet and writes header + every match row from all_matches,
    but only if both clubs are in valid_ids.
    """
    worksheet = create_or_fetch_game_log()
    headers = [
        "Match ID", "Date", "Lagout",
        "Team 1", "Team 1 ID", "Team 1 Score", "Team 1 PPG", "Team 1 PPO",
        "Team 2", "Team 2 ID", "Team 2 Score", "Team 2 PPG", "Team 2 PPO",
        "Team 1 Result", "Team 2 Result", "OT", "Ended Early", "Pushed Timestamp"
    ]
    worksheet = create_or_fetch_game_log()
    rows_to_append = []
    for match in all_matches:
        row = parse_game_row(match, valid_ids)
        if row:
            rows_to_append.append(row)
    if rows_to_append:
        worksheet.append_rows(rows_to_append, value_input_option="USER_ENTERED")
        for r in rows_to_append:
            print(f"📘 Logged match {r[0]} to Game Log")

# ── SKATER LOG ──

def get_existing_skater_match_ids():
    """
    Returns a set of Match IDs already present in 'Skater Log'
    (so we can skip duplicates).
    """
    try:
        worksheet = sheet.worksheet("Skater Log")
        data = worksheet.get_all_values()[1:]
        return set(r[0] for r in data)
    except gspread.exceptions.WorksheetNotFound:
        return set()

def ensure_skater_log_exists(headers):
    """
    Returns the 'Skater Log' worksheet. If missing, creates it with the provided headers.
    """
    try:
        return sheet.worksheet("Skater Log")
    except gspread.exceptions.WorksheetNotFound:
        ws = sheet.add_worksheet(title="Skater Log", rows="5000", cols="50")
        ws.append_row(headers)
        # ← make header bold and freeze it
        ws.format("1:1", {"textFormat": {"bold": True}})
        ws.freeze(rows=1)
        return ws

def log_skater_data(all_matches, valid_ids):
    """
    Appends new skater‐level rows to 'Skater Log' in one batch, skipping any Match IDs
    already logged, and skipping any match where either club is not in valid_ids.
    """
    headers = [
        "Match ID", "Date", "Player ID", "Username", "Platform", "Team", "Position",
        "Goals", "Assists", "Points", "Shots", "Offense Rating", "Defense Rating", "Teamplay Rating",
        "Hits", "Blocked Shots", "Takeaways", "Deflections", "Faceoff Losses", "Faceoff Wins",
        "Faceoff %", "Giveaways", "GWGs", "Interceptions", "Pass Attempts", "Passes", "Pass %",
        "Penalties Drawn", "PIMs", "PK Zone Clears", "Plus Minus", "Possession (secs)", "PPG",
        "Saucer Passes", "SHG", "ES Goals", "Shot Attempts", "Shots On Net %", "Shot %", "TOI", "Result"
    ]

    existing_ids = get_existing_skater_match_ids()
    worksheet = ensure_skater_log_exists(headers)

    new_rows = []
    for match in all_matches:
        # Skip if either club not in valid_ids
        clubs = match.get("clubs", {})
        club_ids = set(clubs.keys())
        if not club_ids.issubset(valid_ids):
            continue

        match_id = str(match.get("matchId"))
        if match_id in existing_ids:
            continue

        timestamp = match.get("timestamp", 0)
        readable_date = datetime.fromtimestamp(timestamp).strftime("%Y-%m-%d %H:%M")
        players = match.get("players", {})

        for club_key, players_dict in players.items():
            if club_key not in valid_ids:
                continue

            club_info = clubs.get(club_key, {})
            try:
                club_result = "Win" if int(club_info.get("score", 0)) > int(club_info.get("opponentScore", 0)) else "Loss"
            except:
                club_result = "Loss"
            team_name = club_info.get("details", {}).get("name", "")

            for pid, p in players_dict.items():
                if p.get("position", "").lower() == "goalie":
                    continue

                def g(key):
                    try:
                        return int(p.get(key, 0))
                    except:
                        return 0

                def pct(num, denom):
                    try:
                        return round((num / (num + denom)) * 100, 2) if (num + denom) > 0 else 0
                    except:
                        return 0

                goals = g("skgoals")
                assists = g("skassists")
                points = goals + assists
                shots = g("skshots")
                shot_attempts = g("skshotattempts")
                es_goals = goals - g("skppg") - g("skshg")
                raw_toi_mins = g("toi")
                toi_mins = float(raw_toi_mins)


                row = [
                    match_id,
                    readable_date,
                    pid,
                    p.get("playername", "Unknown"),
                    p.get("clientPlatform", ""),
                    team_name,
                    p.get("position", ""),
                    goals,
                    assists,
                    points,
                    shots,
                    p.get("ratingOffense", "0"),
                    p.get("ratingDefense", "0"),
                    p.get("ratingTeamplay", "0"),
                    g("skhits"),
                    g("skbs"),
                    g("sktakeaways"),
                    g("skdeflections"),
                    g("skfol"),
                    g("skfow"),
                    pct(g("skfow"), g("skfol")),
                    g("skgiveaways"),
                    g("skgwg"),
                    g("skinterceptions"),
                    g("skpassattempts"),
                    g("skpasses"),
                    pct(g("skpasses"), g("skpassattempts")),
                    g("skpenaltiesdrawn"),
                    g("skpim"),
                    g("skpkclearzone"),
                    g("skplusmin"),
                    g("skpossession"),
                    g("skppg"),
                    g("sksaucerpasses"),
                    g("skshg"),
                    es_goals,
                    shot_attempts,
                    pct(g("skshots"), shot_attempts),
                    pct(goals, shots),
                    toi_mins,
                    club_result
                ]
                new_rows.append(row)
                existing_ids.add(match_id)

    if new_rows:
        existing_rows = len(worksheet.get_all_values())
        total_needed = existing_rows + len(new_rows) + 10
        if total_needed > int(worksheet._properties.get("gridProperties", {}).get("rowCount", 0)):
            worksheet.resize(rows=total_needed)
        worksheet.append_rows(new_rows, value_input_option="USER_ENTERED")
        for r in new_rows:
            print(f"✅ Logged skater: {r[3]} in match {r[0]}")

# ── GOALIE LOG ──

def get_existing_goalie_match_ids():
    try:
        worksheet = sheet.worksheet("Goalie Log")
        data = worksheet.get_all_values()[1:]
        return set(r[0] for r in data)
    except gspread.exceptions.WorksheetNotFound:
        return set()

def ensure_goalie_log_exists(headers):
    try:
        return sheet.worksheet("Goalie Log")
    except gspread.exceptions.WorksheetNotFound:
        ws = sheet.add_worksheet(title="Goalie Log", rows="5000", cols="30")
        ws.append_row(headers)
        ws.format("1:1", {"textFormat": {"bold": True}})
        ws.freeze(rows=1)
        return ws
    
def log_goalie_data(all_matches, valid_ids):
    """
    Appends new goalie‐level rows to 'Goalie Log' in one batch, skipping duplicates,
    and skipping any match where either club is not in valid_ids.
    """
    headers = [
        "Match ID", "Date", "Player ID", "Username", "Platform", "Team", "Position",
        "Goals Against", "GAA", "Saves", "Shots Against", "Save %",
        "Breakaways Saves", "Breakaway Save %", "Penalty Shot Saves", "Penalty Shot Save %",
        "Diving Saves", "Pokechecks", "PK Zone Clears", "Shutout Periods", "Result"
    ]

    existing_ids = get_existing_goalie_match_ids()
    worksheet = ensure_goalie_log_exists(headers)

    new_rows = []
    for match in all_matches:
        clubs = match.get("clubs", {})
        club_ids = set(clubs.keys())
        if not club_ids.issubset(valid_ids):
            continue

        match_id = str(match.get("matchId"))
        if match_id in existing_ids:
            continue

        timestamp = match.get("timestamp", 0)
        readable_date = datetime.fromtimestamp(timestamp).strftime("%Y-%m-%d %H:%M")
        players = match.get("players", {})

        for club_key, players_dict in players.items():
            if club_key not in valid_ids:
                continue

            club_info = clubs.get(club_key, {})
            try:
                club_result = "Win" if int(club_info.get("score", 0)) > int(club_info.get("opponentScore", 0)) else "Loss"
            except:
                club_result = "Loss"
            team_name = club_info.get("details", {}).get("name", "")

            for pid, p in players_dict.items():
                if p.get("position", "").lower() != "goalie":
                    continue

                def g(key):
                    try:
                        return int(p.get(key, 0))
                    except:
                        return 0

                def pct(num, denom):
                    try:
                        return round((num / (num + denom)) * 100, 2) if (num + denom) > 0 else 0
                    except:
                        return 0

                saves = g("glsaves")
                ga = g("glga")
                shots_against = saves + ga

                row = [
                    match_id,
                    readable_date,
                    pid,
                    p.get("playername", "Unknown"),
                    p.get("clientPlatform", ""),
                    team_name,
                    "goalie",
                    ga,
                    ga,  # store raw GA; actual GAA recalculated when aggregating
                    saves,
                    shots_against,
                    pct(saves, shots_against),
                    g("glbrksaves"),
                    pct(g("glbrksaves"), g("glbrkshots")),
                    g("glpensaves"),
                    pct(g("glpensaves"), g("glpenshots")),
                    g("gldsaves"),
                    g("glpokechecks"),
                    g("glpkclearzone"),
                    g("glsoperiods"),
                    club_result
                ]
                new_rows.append(row)
                existing_ids.add(match_id)

    if new_rows:
        existing_rows = len(worksheet.get_all_values())
        total_needed = existing_rows + len(new_rows) + 10
        if total_needed > int(worksheet._properties.get("gridProperties", {}).get("rowCount", 0)):
            worksheet.resize(rows=total_needed)
        worksheet.append_rows(new_rows, value_input_option="USER_ENTERED")
        for r in new_rows:
            print(f"🥅 Logged goalie: {r[3]} in match {r[0]}")

# ── AGGREGATION OF STATS TO TEAM SHEETS ──


        print("⚠️ No Club IDs found in 'Team List'; skipping per-team aggregation.")
        return

    # Build mapping of team_name → team_id from Game Log
    try:
        game_ws = sheet.worksheet("Game Log")
        game_values = game_ws.get_all_values()
    except gspread.exceptions.WorksheetNotFound:
        print("⚠️ 'Game Log' not found; skipping per-team aggregation.")
        return

    header = game_values[0]
    rows = game_values[1:]
    try:
        idx_team1 = header.index("Team 1")
        idx_id1   = header.index("Team 1 ID")
        idx_team2 = header.index("Team 2")
        idx_id2   = header.index("Team 2 ID")
    except ValueError as e:
        return

    valid_names = set()
    for row in rows:
        name1 = row[idx_team1]; id1 = row[idx_id1]
        name2 = row[idx_team2]; id2 = row[idx_id2]
        if id1 in valid_ids:
            valid_names.add(name1)
        if id2 in valid_ids:
            valid_names.add(name2)

    # Read Skater Log
    try:
        skater_ws = sheet.worksheet("Skater Log")
        all_skaters = skater_ws.get_all_values()
    except gspread.exceptions.WorksheetNotFound:
        all_skaters = []

    if len(all_skaters) <= 1:
        per_team_skaters = {}
        replacement_rate_center = replacement_rate_wing = replacement_rate_defense = 0.0
    else:
        skater_header = all_skaters[0]
        skater_rows = all_skaters[1:]
        log_idx = {name: i for i, name in enumerate(skater_header)}

        # Use actual stat column named variations but map to 'Plus Minus'
        plus_variants = ["'+/-'", "+/-", " '+/-", "Plus Minus"]
        for variant in plus_variants:
            if variant in log_idx:
                log_idx["Plus Minus"] = log_idx[variant]
                print(f"ℹ️ Mapped column '{variant}' to 'Plus Minus' for aggregation.")
                break

        def to_int_safe(val):
            try:
                return int(val)
            except:
                return 0

        def to_float_safe(val):
            try:
                return float(val)
            except:
                return 0.0

        centers_es60 = []
        wings_es60   = []
        defense_es60 = []
        for row in skater_rows:
            team_name = row[log_idx["Team"]]
            if team_name not in valid_names:
                continue
            pos_val = row[log_idx["Position"]]
            pos = pos_val.lower() if isinstance(pos_val, str) else ""
            es_g = to_int_safe(row[log_idx["ES Goals"]])
            toi_m = to_float_safe(row[log_idx["TOI"]])
            if toi_m > 0:
                es_per_60 = (es_g / toi_m) * 60
                if "center" in pos:
                    centers_es60.append(es_per_60)
                elif "wing" in pos:
                    wings_es60.append(es_per_60)
                elif "defense" in pos:
                    defense_es60.append(es_per_60)

        def percentile_10(lst):
            if not lst:
                return 0.0
            lst_sorted = sorted(lst)
            idxp = max(0, int(len(lst_sorted) * 0.10) - 1)
            return lst_sorted[idxp]

        replacement_rate_center  = percentile_10(centers_es60)
        replacement_rate_wing    = percentile_10(wings_es60)
        replacement_rate_defense = percentile_10(defense_es60)

        from collections import defaultdict
        per_team_skaters = defaultdict(lambda: defaultdict(list))
        for row in skater_rows:
            team_name = row[log_idx["Team"]]
            if team_name not in valid_names:
                continue
            player_name = row[log_idx["Username"]]
            per_team_skaters[team_name][player_name].append(row)

        # Safely aggregate plus/minus
        tot_plusmin = 0
        if "Plus Minus" in log_idx:
            for row in skater_rows:
                plus_raw = row[log_idx["Plus Minus"]].lstrip("+-")
                tot_plusmin += to_int_safe(plus_raw) if plus_raw.isdigit() else 0
        else:
            print("⚠️ 'Plus Minus' column still not found; skipping aggregation for plus/minus.")



    # Read Goalie Log
    try:
        goalie_ws = sheet.worksheet("Goalie Log")
        all_goalies = goalie_ws.get_all_values()
    except gspread.exceptions.WorksheetNotFound:
        all_goalies = []

    if len(all_goalies) <= 1:
        per_team_goalies = {}
    else:
        goalie_header = all_goalies[0]
        goalie_rows = all_goalies[1:]
        g_idx = {name: i for i, name in enumerate(goalie_header)}

        from collections import defaultdict
        per_team_goalies = defaultdict(lambda: defaultdict(list))
        for row in goalie_rows:
            team_name = row[g_idx["Team"]]
            if team_name not in valid_names:
                continue
            player_name = row[g_idx["Username"]]
            per_team_goalies[team_name][player_name].append(row)

    # ── WRITE PER-TEAM SHEETS ──

    for team_name in valid_names:
        print(f"🔢 Aggregating stats for team '{team_name}'…")

        # Build Skater Summary
        skater_summary = []
        skater_header_row = [
            "Roster", "Platform", "Position", "Games Played",
            "Offense Rating", "Defense Rating", "Teamplay Rating",
            "Goals", "Goals per Game", "Powerplay Goals", "Shorthanded Goals",
            "Assists", "Assists per Game", "Points", "Points per Game", "Plus Minus",
            "Shots", "Shot Attempts", "Shot %", "Shots On Net %",
            "Deflections", "Passes", "Passes per Game", "Saucer Passes", "Saucer Passes per Game",
            "Pass Attempts", "Pass Attempts per Game", "Pass %",
            "Giveaways", "Giveaways per Game", "Hits", "Hits per Game",
            "Interceptions", "Interceptions per Game", "Takeaways", "Takeaways per Game",
            "Blocked Shots", "Blocked Shots per Game", "PK Zone Clears",
            "Total Possession (mins)", "Possession (Min Per Game)", "Faceoff Wins", "Faceoff Losses", "Faceoff %",
            "PIMs", "Penalties Drawn", "TOI (mins)",
            "EVO (ES Above Rep)", "GAR (Goals Above Rep)", "WAR (Wins Above Rep)"
        ]
        skater_summary.append(skater_header_row)

        if team_name in per_team_skaters:
            sk_rows = per_team_skaters[team_name]
            for player_name, rows in sk_rows.items():
                gp = len(rows)
                tot_off = tot_def = tot_teamplay = 0.0
                tot_goals = tot_ppg = tot_shg = tot_assists = 0
                tot_shots = tot_shotatt = 0
                tot_shotpct = tot_shotonn = 0.0
                tot_defl = tot_passes = tot_passatt = 0
                tot_saucers = tot_hits = tot_takeaways = tot_intercep = 0
                tot_blocked = tot_plusmin = tot_possec = 0
                tot_fol = tot_fow = 0
                tot_pen_d = tot_pim = 0
                tot_pkclears = 0
                tot_toi_secs = 0.0
                tot_es_goals = 0

                log_idx = {name: i for i, name in enumerate(all_skaters[0])}

                def to_int_safe(val):
                    try:
                        return int(val)
                    except:
                        return 0

                for row in rows:
                    # Ratings
                    tot_off += float(row[log_idx["Offense Rating"]]) if row[log_idx["Offense Rating"]].replace(".", "").isdigit() else 0.0
                    tot_def += float(row[log_idx["Defense Rating"]]) if row[log_idx["Defense Rating"]].replace(".", "").isdigit() else 0.0
                    tot_teamplay += float(row[log_idx["Teamplay Rating"]]) if row[log_idx["Teamplay Rating"]].replace(".", "").isdigit() else 0.0

                    # Goals / Assists
                    tot_goals += to_int_safe(row[log_idx["Goals"]])
                    tot_ppg += to_int_safe(row[log_idx["PPG"]])
                    tot_shg += to_int_safe(row[log_idx["SHG"]])
                    tot_assists += to_int_safe(row[log_idx["Assists"]])

                    # Shots
                    tot_shots += to_int_safe(row[log_idx["Shots"]])
                    tot_shotatt += to_int_safe(row[log_idx["Shot Attempts"]])
                    tot_shotpct += float(row[log_idx["Shot %"]].rstrip("%")) if row[log_idx["Shot %"]].endswith("%") else 0.0
                    tot_shotonn += float(row[log_idx["Shots On Net %"]].rstrip("%")) if row[log_idx["Shots On Net %"]].endswith("%") else 0.0

                    # Deflections / Passes
                    tot_defl += to_int_safe(row[log_idx["Deflections"]])
                    tot_passes += to_int_safe(row[log_idx["Passes"]])
                    tot_passatt += to_int_safe(row[log_idx["Pass Attempts"]])

                    # Saucer Passes / Hits / Takeaways / Interceptions
                    tot_saucers += to_int_safe(row[log_idx["Saucer Passes"]])
                    tot_hits += to_int_safe(row[log_idx["Hits"]])
                    tot_takeaways += to_int_safe(row[log_idx["Takeaways"]])
                    tot_intercep += to_int_safe(row[log_idx["Interceptions"]])

                    # Blocked Shots / +/- / Possession
                    tot_blocked += to_int_safe(row[log_idx["Blocked Shots"]])
                    val = row[log_idx["Plus Minus"]].strip()
                    val_clean = val.lstrip("+-")
                    tot_plusmin += to_int_safe(val_clean) if val_clean.isdigit() else 0
                    tot_possec += int(float(row[log_idx["TOI"]]) * 60) if row[log_idx["TOI"]].replace(".", "").isdigit() else 0

                    # Faceoffs
                    tot_fol += to_int_safe(row[log_idx["Faceoff Losses"]])
                    tot_fow += to_int_safe(row[log_idx["Faceoff Wins"]])

                    # Penalties Drawn / PIMs
                    tot_pen_d += to_int_safe(row[log_idx["Penalties Drawn"]])
                    tot_pim += to_int_safe(row[log_idx["PIMs"]])

                    # PK Zone Clears
                    tot_pkclears += to_int_safe(row[log_idx["PK Zone Clears"]])

                    # TOI (in seconds)
                    tot_toi_secs += float(row[log_idx["TOI"]]) * 60 if row[log_idx["TOI"]].replace(".", "").isdigit() else 0.0

                    # ES Goals
                    tot_es_goals += to_int_safe(row[log_idx["ES Goals"]])

                # Averages and derived
                avg_off = round(tot_off / gp, 1) if gp else 0.0
                avg_def = round(tot_def / gp, 1) if gp else 0.0
                avg_tp = round(tot_teamplay / gp, 1) if gp else 0.0

                goals = tot_goals
                gpg = round(tot_goals / gp, 2) if gp else 0.0
                ppg = tot_ppg
                shg = tot_shg
                assists = tot_assists
                apg = round(tot_assists / gp, 2) if gp else 0.0

                points_total = tot_goals + tot_assists
                ppg_pts = round(points_total / gp, 2) if gp else 0.0

                plusminus = tot_plusmin

                shots = tot_shots
                shotatt = tot_shotatt
                shotpct = round(tot_shotpct / gp, 2) if gp else 0.0
                shotonn = round(tot_shotonn / gp, 2) if gp else 0.0

                defl = tot_defl
                passes = tot_passes
                ppr = round(tot_passes / gp, 2) if gp else 0.0

                saucers = tot_saucers
                spg = round(tot_saucers / gp, 2) if gp else 0.0

                passatt = tot_passatt
                papg = round(tot_passatt / gp, 2) if gp else 0.0

                passpct2 = round((tot_passes / tot_passatt) * 100, 2) if tot_passatt > 0 else 0.0

                giveaways = sum(
                    math.ceil(float(r[log_idx["Giveaways"]]))
                    for r in rows
                    if r[log_idx["Giveaways"]].replace(".", "").isdigit()
                )
                giveaways_pg = round(giveaways / gp, 2) if gp else 0.0

                hits = tot_hits
                hits_pg = round(tot_hits / gp, 2) if gp else 0.0

                intercep = tot_intercep
                intercep_pg = round(tot_intercep / gp, 2) if gp else 0.0

                takeaway = tot_takeaways
                takeaway_pg = round(tot_takeaways / gp, 2) if gp else 0.0

                blocked = tot_blocked
                blocked_pg = round(tot_blocked / gp, 2) if gp else 0.0

                pk_clears = tot_pkclears

                total_pos_mins = round(tot_possec / 60, 2)
                poss_pg = round((tot_possec / 60) / gp, 2) if gp else 0.0

                fol = tot_fol
                fow = tot_fow
                fopct = round((fow / (fow + fol)) * 100, 2) if (fow + fol) > 0 else 0.0

                pim = tot_pim
                pen_drawn = tot_pen_d

                toi_mins = round(tot_toi_secs / 60, 2)

                total_toi_min = tot_toi_secs / 60.0
                if total_toi_min > 0:
                    es_per_60 = (tot_es_goals / total_toi_min) * 60
                else:
                    es_per_60 = 0.0

                pos_label = rows[0][log_idx["Position"]].lower()
                if "center" in pos_label:
                    rep_rate = replacement_rate_center
                elif "wing" in pos_label:
                    rep_rate = replacement_rate_wing
                elif "defense" in pos_label:
                    rep_rate = replacement_rate_defense
                else:
                    rep_rate = replacement_rate_center

                toi_hours = total_toi_min / 60.0
                evo = round((es_per_60 - rep_rate) * toi_hours, 2)
                gar = evo
                war = round(gar / 10.0, 2)

                sample = rows[0]
                plat = sample[log_idx["Platform"]]
                posstr = sample[log_idx["Position"]]

                summary_row = [
                    player_name,
                    plat,
                    posstr,
                    gp,
                    avg_off,
                    avg_def,
                    avg_tp,
                    goals,
                    gpg,
                    ppg,
                    shg,
                    assists,
                    apg,
                    points_total,
                    ppg_pts,
                    plusminus,
                    shots,
                    shotatt,
                    shotpct,
                    shotonn,
                    defl,
                    passes,
                    ppr,
                    saucers,
                    spg,
                    passatt,
                    papg,
                    passpct2,
                    giveaways,
                    giveaways_pg,
                    hits,
                    hits_pg,
                    intercep,
                    intercep_pg,
                    takeaway,
                    takeaway_pg,
                    blocked,
                    blocked_pg,
                    pk_clears,
                    total_pos_mins,
                    poss_pg,
                    fow,
                    fol,
                    fopct,
                    pim,
                    pen_drawn,
                    toi_mins,
                    evo,
                    gar,
                    war
                ]
                skater_summary.append(summary_row)

        # Build Goalie Summary
        goalie_summary = []
        goalie_header_row = [
            "Roster", "Platform", "Position", "Games Played",
            "Goals Against", "GAA", "Saves", "Shots Against", "Save %",
            "Breakaway Saves", "Breakaway Save %", "Penalty Shot Saves", "Penalty Shot Save %",
            "Diving Saves", "Pokechecks", "PK Zone Clears", "Shutout Periods",
            "Wins", "Losses"
        ]
        goalie_summary.append(goalie_header_row)

        if team_name in per_team_goalies:
            for player_name, rows in per_team_goalies[team_name].items():
                gp = len(rows)
                tot_ga = tot_saves = tot_shots = 0
                tot_brksaves = tot_pensaves = tot_dive = 0
                tot_poke = tot_pkclr = tot_shut = wins = losses = 0

                g_idx = {name: i for i, name in enumerate(all_goalies[0])}
                for row in rows:
                    def to_int_g(idx_col):
                        try:
                            return int(row[idx_col])
                        except:
                            return 0

                    tot_ga += to_int_g(g_idx["Goals Against"])
                    tot_saves += to_int_g(g_idx["Saves"])
                    tot_shots += to_int_g(g_idx["Shots Against"])
                    tot_brksaves += to_int_g(g_idx["Breakaways Saves"])
                    tot_pensaves += to_int_g(g_idx["Penalty Shot Saves"])
                    tot_dive += to_int_g(g_idx["Diving Saves"])
                    tot_poke += to_int_g(g_idx["Pokechecks"])
                    tot_pkclr += to_int_g(g_idx["PK Zone Clears"])
                    tot_shut += to_int_g(g_idx["Shutout Periods"])

                    if row[g_idx["Result"]] == "Win":
                        wins += 1
                    else:
                        losses += 1

                avg_gaa = round(tot_ga / gp, 2) if gp else 0.0
                save_pct = round((tot_saves / tot_shots) * 100, 2) if tot_shots > 0 else 0.0

                summary_row = [
                    player_name,
                    rows[0][g_idx["Platform"]],
                    rows[0][g_idx["Position"]],
                    gp,
                    tot_ga,
                    avg_gaa,
                    tot_saves,
                    tot_shots,
                    save_pct,
                    tot_brksaves,
                    0.0,
                    tot_pensaves,
                    0.0,
                    tot_dive,
                    tot_poke,
                    tot_pkclr,
                    tot_shut,
                    wins,
                    losses
                ]
                goalie_summary.append(summary_row)

        # Write to Team Sheet


        print(
            f"✅ Wrote {len(skater_summary)-1} skater rows and "
        )

# ── STANDINGS BUILDER ──

def create_or_update_standings():
    """
    Reads 'Game Log', deduplicates by Match ID, tallies Wins/Losses/OTL for each team
    in Team List, and writes the results (including Points) into a 'Standings' worksheet.
    """
    valid_ids = get_team_list()
    if not valid_ids:
        print("⚠️ No Club IDs found in 'Team List'; skipping standings.")
        return

    try:
        game_ws = sheet.worksheet("Game Log")
        all_values = game_ws.get_all_values()
    except gspread.exceptions.WorksheetNotFound:
        print("⚠️ 'Game Log' worksheet not found; cannot build standings.")
        return

    if len(all_values) <= 1:
        print("⚠️ 'Game Log' has no data (just a header). Skipping standings.")
        return

    header = all_values[0]
    rows = all_values[1:]
    try:
        idx_match_id = header.index("Match ID")
        idx_team1    = header.index("Team 1")
        idx_id1      = header.index("Team 1 ID")
        idx_result1  = header.index("Team 1 Result")
        idx_team2    = header.index("Team 2")
        idx_id2      = header.index("Team 2 ID")
        idx_result2  = header.index("Team 2 Result")
        idx_ot       = header.index("OT")
        idx_early    = header.index("Ended Early")
    except ValueError as e:
        return

    # Deduplicate by Match ID
    unique_matches = {}
    for row in rows:
        mid = row[idx_match_id]
        if mid not in unique_matches:
            unique_matches[mid] = row

    # Tally results for valid teams
    standings = {}
    for mid, row in unique_matches.items():
        team1 = row[idx_team1]; id1 = row[idx_id1]; res1 = row[idx_result1]
        team2 = row[idx_team2]; id2 = row[idx_id2]; res2 = row[idx_result2]
        ot_flag = (row[idx_ot].strip().lower() == "yes")
        early_flag = (row[idx_early].strip().lower() == "yes")

        if id1 in valid_ids:
            if team1 not in standings:
                standings[team1] = {"Wins": 0, "Losses": 0, "OTL": 0}
            if res1 == "Win":
                standings[team1]["Wins"] += 1
            else:
                if early_flag or ot_flag:
                    standings[team1]["OTL"] += 1
                else:
                    standings[team1]["Losses"] += 1

        if id2 in valid_ids:
            if team2 not in standings:
                standings[team2] = {"Wins": 0, "Losses": 0, "OTL": 0}
            if res2 == "Win":
                standings[team2]["Wins"] += 1
            else:
                if early_flag or ot_flag:
                    standings[team2]["OTL"] += 1
                else:
                    standings[team2]["Losses"] += 1

    # Write into "Standings" worksheet
    try:
        standings_ws = sheet.worksheet("Standings")
        standings_ws.clear()
    except gspread.exceptions.WorksheetNotFound:
        row_count = len(standings) + 5
        standings_ws = sheet.add_worksheet(title="Standings", rows=str(row_count), cols="5")

    output = []
    header_row = ["Team", "Wins", "Losses", "OTL", "Points"]
    output.append(header_row)

    for team_name in sorted(standings.keys()):
        rec = standings[team_name]
        points = rec["Wins"] * 2 + rec["OTL"] * 1
        output.append([
            team_name,
            rec["Wins"],
            rec["Losses"],
            rec["OTL"],
            points
        ])

    standings_ws.append_rows(output, value_input_option="USER_ENTERED")
    try:
        standings_ws.format("1:1", {"textFormat": {"bold": True}})
        standings_ws.freeze(rows=1)
    except Exception as e:
        pass  # inserted to complete except block

    print(f"✅ 'Standings' sheet updated with {len(standings)} teams.")

# ── MAIN RUNNER ──




def aggregate_stats_to_master_tabs():
    """
    Aggregates all skater and goalie logs into 'Skater Master' and 'Goalie Master' tabs,
    summarizing player stats by Username across all matches.
    """
    import math
    from collections import defaultdict

    def ensure_sheet_exists(title, headers):
        try:
            ws = sheet.worksheet(title)
            ws.clear()
        except gspread.exceptions.WorksheetNotFound:
            ws = sheet.add_worksheet(title=title, rows="5000", cols="50")
        ws.append_row(headers)
        ws.format("1:1", {"textFormat": {"bold": True}})
        ws.freeze(rows=1)
        return ws

    def safe_int(val):
        try:
            return int(val)
        except:
            return 0

    def safe_float(val):
        try:
            return float(val)
        except:
            return 0.0

    # ── SKATER MASTER ──
    tallied_skater_ids = get_tallied_ids("Skater Tallied IDs")

    try:
        ws = sheet.worksheet("Skater Log")
        data = ws.get_all_values()
    except:
        print("⚠️ Skater Log not found.")
        data = []

    skater_header = [
        "Username", "Team", "Games Played", "Goals", "Assists", "Points", "Shots",
        "Shot Attempts", "Hits", "Takeaways", "Deflections", "Faceoff Wins",
        "Faceoff Losses", "Faceoff %", "Giveaways", "GWGs", "Interceptions",
        "Pass Attempts", "Passes", "Pass %", "Plus Minus", "PIMs", "TOI"
    ]

    if len(data) > 1:
        header = data[0]
        rows = [r for r in data[1:] if (r[0], r[2]) not in tallied_skater_ids]
        idx = {k: i for i, k in enumerate(header)}
        statmap = defaultdict(list)
        for row in rows:
            name = row[idx["Username"]]
            team = row[idx["Team"]]
            statmap[(name, team)].append(row)

        skater_agg = []
        for (name, team), games in statmap.items():
            gp = len(games)
            tot = lambda key: sum(safe_int(g[idx[key]]) for g in games)
            totf = lambda key: sum(safe_float(g[idx[key]]) for g in games)

            goals = tot("Goals")
            assists = tot("Assists")
            # Ratings averaged, capped, rounded to nearest 5
            avg_rating = lambda key: round(min(100, sum(safe_float(g[sk_idx[key]]) for g in games) / gp / 5) * 5) if gp else 0
            rating_off = avg_rating("Offense Rating")
            rating_def = avg_rating("Defense Rating")
            rating_team = avg_rating("Teamplay Rating")
            points = goals + assists
            shots = tot("Shots")
            attempts = tot("Shot Attempts")
            hits = tot("Hits")
            takeaways = tot("Takeaways")
            defl = tot("Deflections")
            fow = tot("Faceoff Wins")
            fol = tot("Faceoff Losses")
            fopct = round((fow / (fow + fol)) * 100, 2) if (fow + fol) else 0
            gva = tot("Giveaways")
            gwg = tot("GWGs")
            interceptions = tot("Interceptions")
            passatt = tot("Pass Attempts")
            passes = tot("Passes")
            passpct = round((passes / passatt) * 100, 2) if passatt else 0
            plusmin = sum(safe_int(g[idx["Plus Minus"]].strip()) for g in games)
            pim = tot("PIMs")
            toi = round(totf("TOI"), 2)

            skater_agg.append([
                name, team, gp, goals, assists, points, shots,
                rating_off, rating_def, rating_team, attempts,
                hits, takeaways, defl, fow, fol, fopct,
                gva, gwg, interceptions, passatt, passes, passpct,
                plusmin, pim, toi
            ])

        ws_out = ensure_sheet_exists("Skater Master", skater_header)
        ws_out.append_rows(skater_agg, value_input_option="USER_ENTERED")
    new_ids = [(r[0], r[2]) for r in rows]
    append_tallied_ids("Skater Tallied IDs", new_ids)
    print(f"✅ Skater Master updated: {len(skater_agg)} players")

    
    # ── GOALIE MASTER ──
    tallied_goalie_ids = get_tallied_ids("Goalie Tallied IDs")

    try:
        ws = sheet.worksheet("Goalie Log")
        data = ws.get_all_values()
    except:
        print("⚠️ Goalie Log not found.")
        data = []

    goalie_header = [
        "Username", "Team", "Games Played", "Goals Against", "GAA", "Saves", "Shots Against",
        "Save %", "Breakaways Saves", "Breakaway Save %", "Penalty Shot Saves", "Penalty Shot Save %",
        "Diving Saves", "Pokechecks", "PK Zone Clears", "Shutout Periods", "Wins", "Losses"
    ]

    if len(data) > 1:
        header = data[0]
        rows = [r for r in data[1:] if (r[0], r[2]) not in tallied_goalie_ids]
        idx = {k: i for i, k in enumerate(header)}
        from collections import defaultdict
        statmap = defaultdict(list)
        for row in rows:
            name = row[idx["Username"]]
            team = row[idx["Team"]]
            statmap[(name, team)].append(row)

        goalie_agg = []
        for (name, team), games in statmap.items():
            gp = len(games)
            tot = lambda key: sum(safe_int(g[idx[key]]) for g in games)
            ga = tot("Goals Against")
            saves = tot("Saves")
            shots = tot("Shots Against")
            savepct = round((saves / shots) * 100, 2) if shots else 0.0
            gaa = round(ga / gp, 2) if gp else 0.0
            brk_saves = tot("Breakaways Saves")
            brk_pct = round((brk_saves / shots) * 100, 2) if shots else 0.0
            pen_saves = tot("Penalty Shot Saves")
            pen_pct = round((pen_saves / shots) * 100, 2) if shots else 0.0
            dive_saves = tot("Diving Saves")
            pokechecks = tot("Pokechecks")
            pk_clears = tot("PK Zone Clears")
            shutouts = tot("Shutout Periods")
            wins = sum(1 for g in games if g[idx["Result"]].strip().lower() == "win")
            losses = gp - wins

            goalie_agg.append([
                name, team, gp, ga, gaa, saves, shots, savepct,
                brk_saves, brk_pct, pen_saves, pen_pct,
                dive_saves, pokechecks, pk_clears, shutouts,
                wins, losses
            ])

        ws_out = ensure_sheet_exists("Goalie Master", goalie_header)
        ws_out.append_rows(goalie_agg, value_input_option="USER_ENTERED")

    new_ids = [(r[0], r[2]) for r in rows]
    append_tallied_ids("Goalie Tallied IDs", new_ids)
    print(f"🥅 Goalie Master updated: {len(goalie_agg)} players")



def get_tallied_ids(sheet_name):
    try:
        ws = sheet.worksheet(sheet_name)
        data = ws.get_all_values()[1:]
        return set((row[0], row[1]) for row in data)
    except gspread.exceptions.WorksheetNotFound:
        return set()

def append_tallied_ids(sheet_name, ids_to_append):
    try:
        try:
            ws = sheet.worksheet(sheet_name)
        except gspread.exceptions.WorksheetNotFound:
            ws = sheet.add_worksheet(title=sheet_name, rows="1000", cols="2")
            ws.append_row(["Match ID", "Player ID"])
        ws.append_rows(ids_to_append, value_input_option="USER_ENTERED")
    except Exception as e:
        print(f"❌ Failed to append to {sheet_name}: {e}")


# ===== OVERRIDE log_game_data TO ENSURE NO DUPLICATES =====
def log_game_data(all_matches, valid_ids):
    """Append only NEW matches into 'Game Log' by skipping any Match ID already present."""
    worksheet = create_or_fetch_game_log()
    existing_ids = get_existing_game_match_ids()
    rows_to_append = []
    for match in all_matches:
        row = parse_game_row(match, valid_ids)
        if row is None:
            continue
        match_id = str(row[0])
        if match_id in existing_ids:
            continue
        existing_ids.add(match_id)  # Avoid duplicates within same run
        rows_to_append.append(row)
    if rows_to_append:
        worksheet.append_rows(rows_to_append, value_input_option="USER_ENTERED")
        for r in rows_to_append:
            print(f"📘 Logged match {r[0]} to Game Log")
# ===== END OVERRIDE =====

# ===== OVERRIDE aggregate_stats_to_master_tabs TO PRESERVE EXISTING TOTALS =====
def aggregate_stats_to_master_tabs():
    """Aggregates Skater and Goalie logs into Master tabs.

    • Only re‑runs if there are NEW rows (based on Tallied IDs)
    • If new rows exist, totals are recomputed across *all* log data and Master
      sheets are replaced.
    • If no new rows, Master sheets are left untouched.
    """
    import math
    from collections import defaultdict

    def safe_int(val):
        try:
            return int(val)
        except:
            return 0

    def safe_float(val):
        try:
            return float(val)
        except:
            return 0.0

    # ──────────────────────────── SKATER MASTER ────────────────────────────
    tallied_skater_ids = get_tallied_ids("Skater Tallied IDs")

    try:
        sk_ws = sheet.worksheet("Skater Log")
        sk_data = sk_ws.get_all_values()
    except gspread.exceptions.WorksheetNotFound:
        sk_data = []

    sk_header_master = [
        "Username", "Team", "Games Played", "Goals", "Assists", "Points", "Shots",
        "Offense Rating", "Defense Rating", "Teamplay Rating",
        "Shot Attempts", "Hits", "Takeaways", "Deflections", "Faceoff Wins",
        "Faceoff Losses", "Faceoff %", "Giveaways", "GWGs", "Interceptions",
        "Pass Attempts", "Passes", "Pass %", "Plus Minus", "PIMs", "TOI"
    ]

    new_sk_rows = []  # rows not yet tallied
    if len(sk_data) > 1:
        sk_header = sk_data[0]
        sk_idx = {k: i for i, k in enumerate(sk_header)}
        for row in sk_data[1:]:
            if (row[0], row[2]) not in tallied_skater_ids:
                new_sk_rows.append(row)

    if new_sk_rows:
        # Need to recompute totals across ALL skater data
        statmap = defaultdict(list)
        for row in sk_data[1:]:
            name = row[sk_idx["Username"]]
            team = row[sk_idx["Team"]]
            statmap[(name, team)].append(row)

        skater_agg = []
        for (name, team), games in statmap.items():
            gp = len(games)
            tot = lambda key: sum(safe_int(g[sk_idx[key]]) for g in games)
            totf = lambda key: sum(safe_float(g[sk_idx[key]]) for g in games)

            goals = tot("Goals")
            assists = tot("Assists")
            # Ratings averaged, capped, rounded to nearest 5
            avg_rating = lambda key: round(min(100, sum(safe_float(g[sk_idx[key]]) for g in games) / gp / 5) * 5) if gp else 0
            rating_off = avg_rating("Offense Rating")
            rating_def = avg_rating("Defense Rating")
            rating_team = avg_rating("Teamplay Rating")
            points = goals + assists
            shots = tot("Shots")
            attempts = tot("Shot Attempts")
            hits = tot("Hits")
            takeaways = tot("Takeaways")
            defl = tot("Deflections")
            fow = tot("Faceoff Wins")
            fol = tot("Faceoff Losses")
            fopct = round((fow / (fow + fol)) * 100, 2) if (fow + fol) else 0
            gva = tot("Giveaways")
            gwg = tot("GWGs")
            interceptions = tot("Interceptions")
            passatt = tot("Pass Attempts")
            passes = tot("Passes")
            passpct = round((passes / passatt) * 100, 2) if passatt else 0
            plusmin = sum(safe_int(g[sk_idx["Plus Minus"]].strip()) for g in games)
            pim = tot("PIMs")
            toi = round(totf("TOI"), 2)

            skater_agg.append([
                name, team, gp, goals, assists, points, shots,
                rating_off, rating_def, rating_team, attempts,
                hits, takeaways, defl, fow, fol, fopct,
                gva, gwg, interceptions, passatt, passes, passpct,
                plusmin, pim, toi
            ])

        # Write / replace Skater Master
        try:
            master_ws = sheet.worksheet("Skater Master")
            master_ws.clear()
        except gspread.exceptions.WorksheetNotFound:
            master_ws = sheet.add_worksheet(title="Skater Master", rows="5000", cols="50")
        master_ws.append_row(sk_header_master)
        master_ws.format("1:1", {"textFormat": {"bold": True}})
        master_ws.freeze(rows=1)
        master_ws.append_rows(skater_agg, value_input_option="USER_ENTERED")
        print(f"✅ Skater Master updated: {len(skater_agg)} players")

        # Record new tallied IDs
        new_ids = [(r[0], r[2]) for r in new_sk_rows]
        append_tallied_ids("Skater Tallied IDs", new_ids)
    else:
        print("ℹ️ No new skater games — Skater Master left unchanged.")

    # ──────────────────────────── GOALIE MASTER ────────────────────────────
    tallied_goalie_ids = get_tallied_ids("Goalie Tallied IDs")

    try:
        gl_ws = sheet.worksheet("Goalie Log")
        gl_data = gl_ws.get_all_values()
    except gspread.exceptions.WorksheetNotFound:
        gl_data = []

    goalie_header_master = [
        "Username", "Team", "Games Played", "Goals Against", "GAA", "Saves", "Shots Against",
        "Save %", "Breakaways Saves", "Breakaway Save %", "Penalty Shot Saves", "Penalty Shot Save %",
        "Diving Saves", "Pokechecks", "PK Zone Clears", "Shutout Periods", "Wins", "Losses"
    ]

    new_gl_rows = []
    if len(gl_data) > 1:
        gl_header = gl_data[0]
        gl_idx = {k: i for i, k in enumerate(gl_header)}
        for row in gl_data[1:]:
            if (row[0], row[2]) not in tallied_goalie_ids:
                new_gl_rows.append(row)

    if new_gl_rows:
        statmap = defaultdict(list)
        for row in gl_data[1:]:
            name = row[gl_idx["Username"]]
            team = row[gl_idx["Team"]]
            statmap[(name, team)].append(row)

        goalie_agg = []
        for (name, team), games in statmap.items():
            gp = len(games)
            tot = lambda key: sum(safe_int(g[gl_idx[key]]) for g in games)

            ga = tot("Goals Against")
            saves = tot("Saves")
            shots = tot("Shots Against")
            savepct = round((saves / shots) * 100, 2) if shots else 0.0
            gaa = round(ga / gp, 2) if gp else 0.0
            brk_saves = tot("Breakaways Saves")
            brk_pct = round((brk_saves / shots) * 100, 2) if shots else 0.0
            pen_saves = tot("Penalty Shot Saves")
            pen_pct = round((pen_saves / shots) * 100, 2) if shots else 0.0
            dive_saves = tot("Diving Saves")
            pokechecks = tot("Pokechecks")
            pk_clears = tot("PK Zone Clears")
            shutouts = tot("Shutout Periods")
            wins = sum(1 for g in games if g[gl_idx["Result"]].strip().lower() == "win")
            losses = gp - wins

            goalie_agg.append([
                name, team, gp, ga, gaa, saves, shots, savepct,
                brk_saves, brk_pct, pen_saves, pen_pct,
                dive_saves, pokechecks, pk_clears, shutouts,
                wins, losses
            ])

        try:
            gm_ws = sheet.worksheet("Goalie Master")
            gm_ws.clear()
        except gspread.exceptions.WorksheetNotFound:
            gm_ws = sheet.add_worksheet(title="Goalie Master", rows="5000", cols="50")
        gm_ws.append_row(goalie_header_master)
        gm_ws.format("1:1", {"textFormat": {"bold": True}})
        gm_ws.freeze(rows=1)
        gm_ws.append_rows(goalie_agg, value_input_option="USER_ENTERED")
        print(f"🥅 Goalie Master updated: {len(goalie_agg)} players")

        new_ids = [(r[0], r[2]) for r in new_gl_rows]
        append_tallied_ids("Goalie Tallied IDs", new_ids)
    else:
        print("ℹ️ No new goalie games — Goalie Master left unchanged.")
# ===== END OVERRIDE =====



# ── MERGE DUPLICATE PLAYER ROWS IN MASTER TABS ──────────────────────────────
def merge_duplicate_players():
    def safe_int(x):   return int(float(x)) if str(x).replace('.', '', 1).lstrip('+-').isdigit() else 0
    def safe_float(x): return float(x)       if str(x).replace('.', '', 1).lstrip('+-').isdigit() else 0.0

    def _merge_master(sheet_name, derive_fn):
        ws = sheet.worksheet(sheet_name)
        raw = ws.get_all_values()
        if len(raw) <= 1:
            print(f"ℹ️ {sheet_name} is empty – nothing to merge.")
            return

        header, rows = raw[0], raw[1:]
        idx = {h: i for i, h in enumerate(header)}
        from collections import defaultdict
        bucket = defaultdict(list)
        for r in rows:
            bucket[r[idx["Username"]]].append(r)

        changed = False
        out_rows = []

        for uname, user_rows in bucket.items():
            if len(user_rows) == 1:
                out_rows.append(user_rows[0])
                continue

            teams = {r[idx["Team"]] for r in user_rows}
            if len(teams) == 1:
                keep_team = teams.pop()
                merged = derive_fn(user_rows, idx, keep_team)
                out_rows.append(merged)
                changed = True
                print(f"🤖 Auto-merged {len(user_rows)} rows for {uname} on {keep_team}")
                continue

            changed = True
            print(f"\n🧐  Duplicate detected for **{uname}** (different teams):")
            for j, r in enumerate(user_rows, 1):
                print(f"  {j}. Team = {r[idx['Team']]}  |  Games = {r[idx['Games Played']]}")
            choice = None
            while choice not in range(1, len(user_rows) + 1):
                try:
                    choice = int(input(f"   → Which team label to keep for {uname}? [1-{len(user_rows)}]: "))
                except ValueError:
                    pass
            keep_team = user_rows[choice - 1][idx["Team"]]
            merged = derive_fn(user_rows, idx, keep_team)
            out_rows.append(merged)
            print(f"✅  Merged {len(user_rows)} rows → 1 ({keep_team})")

        if changed:
            ws.clear()
            ws.append_row(header)
            ws.append_rows(out_rows, value_input_option="USER_ENTERED")
            ws.format("1:1", {"textFormat": {"bold": True}})
            ws.freeze(rows=1)
            print(f"🎉 {sheet_name} deduplicated and rebuilt.")
        else:
            print(f"👍 {sheet_name} already has no duplicates.")

    def derive_skater(rows, idx, team_name):
        sums = {k: 0 for k in
            ["Games Played","Goals","Assists","Shots","Shot Attempts","Hits","Takeaways",
             "Deflections","Faceoff Wins","Faceoff Losses","Giveaways","GWGs","Interceptions",
             "Pass Attempts","Passes","Plus Minus","PIMs"]}
        toi_total = 0.0
        for r in rows:
            for k in sums:
                sums[k] += safe_int(r[idx[k]])
            toi_total += safe_float(r[idx["TOI"]])
        points = sums["Goals"] + sums["Assists"]
        fow, fol = sums["Faceoff Wins"], sums["Faceoff Losses"]
        fopct = round((fow / (fow + fol)) * 100, 2) if (fow + fol) else 0
        passatt, passes = sums["Pass Attempts"], sums["Passes"]
        passpct = round((passes / passatt) * 100, 2) if passatt else 0
        return [rows[0][idx["Username"]], team_name, sums["Games Played"], sums["Goals"], sums["Assists"], points,
                sums["Shots"], sums["Shot Attempts"], sums["Hits"], sums["Takeaways"], sums["Deflections"],
                fow, fol, fopct, sums["Giveaways"], sums["GWGs"], sums["Interceptions"],
                passatt, passes, passpct, sums["Plus Minus"], sums["PIMs"], round(toi_total, 2)]

    def derive_goalie(rows, idx, team_name):
        sums = {k: 0 for k in
            ["Games Played","Goals Against","Saves","Shots Against","Breakaways Saves",
             "Penalty Shot Saves","Diving Saves","Pokechecks","PK Zone Clears",
             "Shutout Periods","Wins","Losses"]}
        for r in rows:
            for k in sums:
                sums[k] += safe_int(r[idx[k]])
        gp, shots = sums["Games Played"], sums["Shots Against"]
        gaa = round(sums["Goals Against"] / gp, 2) if gp else 0
        save = round((sums["Saves"] / shots) * 100, 2) if shots else 0
        brk_pct = round((sums["Breakaways Saves"] / shots) * 100, 2) if shots else 0
        pen_pct = round((sums["Penalty Shot Saves"] / shots) * 100, 2) if shots else 0
        return [rows[0][idx["Username"]], team_name, gp, sums["Goals Against"], gaa, sums["Saves"], shots, save,
                sums["Breakaways Saves"], brk_pct, sums["Penalty Shot Saves"], pen_pct,
                sums["Diving Saves"], sums["Pokechecks"], sums["PK Zone Clears"],
                sums["Shutout Periods"], sums["Wins"], sums["Losses"]]

    _merge_master("Skater Master", derive_skater)
    _merge_master("Goalie Master", derive_goalie)
# ── END MERGE DUPLICATES ────────────────────────────────────────────────────
if __name__ == "__main__":
    print("🏒 WCA NHL Scraper Runner")
    print("1 = Log Matches")
    print("2 = Aggregate Team Tabs")
    print("3 = Update Standings")
    print("4 = Update Master Tabs")
    print("5 = Merge duplicate rows in Skater/Goalie Master")
    mode = input("Choose mode (1-5): ").strip()

    if mode == "1":
        print("[1] Fetching + Logging New Matches...")
        valid_ids = get_team_list()
        all_matches = []
        for cid in valid_ids:
            matches = get_private_matches(cid)
            all_matches.extend(matches)

        # Deduplicate by match ID
        unique_matches = {}
        for m in all_matches:
            mid = str(m.get("matchId"))
            if mid and mid not in unique_matches:
                unique_matches[mid] = m
        all_matches_deduped = list(unique_matches.values())

        if not all_matches_deduped:
            print("⚠️ No new matches found.")
        else:
            log_game_data(all_matches_deduped, valid_ids)
            log_skater_data(all_matches_deduped, valid_ids)
            log_goalie_data(all_matches_deduped, valid_ids)
            print("✅ Logs updated.")

    elif mode == "2":
        print("📊 [2] Aggregating Per-Team Tabs...")
        print("🧱 Skipped — implement if needed manually.")

    elif mode == "3":
        print("🏆 [3] Updating Standings Tab...")
        create_or_update_standings()

    elif mode == "4":
        print("📚 [4] Updating Skater and Goalie Master Tabs...")
        aggregate_stats_to_master_tabs()

    
    elif mode == "5":
        print("🔄 [5] Checking for duplicate players...")
        merge_duplicate_players()
    else:
        print("❌ Invalid selection.")