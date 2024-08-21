#%%
import pandas as pd
import numpy as np

from pathlib import Path
from collections import defaultdict

def roster_in_week(week):

    df_week = pd.read_csv(week, index_col='Unnamed: 0')
    df_week = df_week.rename(columns=dict_full)
    df_week = df_week.dropna()
    
    return df_week

def team_total_roster(team, rosters):

    roster_check = []
    for week in weeks:
        roster_check.append(rosters[int(week)][0][team])

    unique_players = set(roster_check[0]).intersection(*roster_check[1:])
    return unique_players

teammapping_file_full = Path(r"C:\Users\Joseph\Google Drive\RBB\nfl-fantasy-football2022\teams\team_mapping_full.txt")
rosters_directory = Path(r"C:\Users\Joseph\Google Drive\RBB\nfl-fantasy-football2022\Data Analysis\weekly_rosters")

rosters_weeks = [roster for roster in rosters_directory.glob("*.csv")]

with open(teammapping_file_full, 'r') as f:
    dict_full = dict(eval(f.read()))

teams = [name for name in dict_full.values()]

weeks = [str(w) for w in range(1, len(rosters_weeks) - 2)]

# %%
total_rosters = defaultdict(list)
for week in rosters_weeks:
    week_id = int(week.name.split('wk_')[1].split('_')[0])
    total_rosters[week_id].append(roster_in_week(week))

total_rosters_p_team = defaultdict(list)
for team in teams:
    total_rosters_p_team[team].append(team_total_roster(team, total_rosters))

# %%
for team in teams:
    print(team)
    for player in total_rosters_p_team[team][0]:
        print(player)
    print("-----------")


# %%
