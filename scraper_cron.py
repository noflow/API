# scraper_cron.py
from scraper_working import (
    EA_GAME_SLUG,
    EA_PLATFORM,
    combine_lagout_matches,
    filter_matches_to_league_window,
    set_active_log_phase,
    split_matches_by_season_phase,
    get_team_list,
    get_private_matches,
    log_game_data,
    log_skater_data,
    log_goalie_data,
)
from schedule_checker import update_schedule_check

print(f"Running Scraper Automatically (Log Matches) for {EA_GAME_SLUG} on {EA_PLATFORM}")
valid_ids = get_team_list()
all_matches = []
for cid in valid_ids:
    matches = get_private_matches(cid)
    all_matches.extend(matches)

# Deduplicate by Match ID
unique_matches = {str(m["matchId"]): m for m in all_matches}
all_matches_deduped = list(unique_matches.values())
all_matches_deduped = combine_lagout_matches(all_matches_deduped)
all_matches_deduped = filter_matches_to_league_window(all_matches_deduped)
matches_by_phase = split_matches_by_season_phase(all_matches_deduped)

if any(matches_by_phase.values()):
    for phase in ("regular", "playoffs"):
        phase_matches = matches_by_phase[phase]
        if not phase_matches:
            continue
        set_active_log_phase(phase)
        log_game_data(phase_matches, valid_ids)
        log_skater_data(phase_matches, valid_ids)
        log_goalie_data(phase_matches, valid_ids)
    set_active_log_phase("regular")
    print("Logs updated.")
else:
    print("No new matches found.")

try:
    update_schedule_check()
except Exception as exc:
    print(f"Schedule Check could not be updated: {exc}")
