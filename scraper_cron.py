# scraper_cron.py
from scraper_working import get_team_list, get_private_matches, log_game_data, log_skater_data, log_goalie_data

print("🏒 Running Scraper Automatically (Log Matches)")
valid_ids = get_team_list()
all_matches = []
for cid in valid_ids:
    matches = get_private_matches(cid)
    all_matches.extend(matches)

# Deduplicate by Match ID
unique_matches = {str(m["matchId"]): m for m in all_matches}
all_matches_deduped = list(unique_matches.values())

if all_matches_deduped:
    log_game_data(all_matches_deduped, valid_ids)
    log_skater_data(all_matches_deduped, valid_ids)
    log_goalie_data(all_matches_deduped, valid_ids)
    print("✅ Logs updated.")
else:
    print("⚠️ No new matches found.")
