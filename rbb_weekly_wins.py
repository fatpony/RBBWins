#%%
import pandas as pd
import numpy as np
import gspread

from oauth2client.service_account import ServiceAccountCredentials
from pathlib import Path
from gspread_dataframe import set_with_dataframe

CURRENT_WEEK = int(input("Week: "))

teammapping_file_full = Path(r"H:\My Drive\nfl-fantasy-football2022\teams\team_mapping_full.txt")
scores_directory = Path(r"H:\My Drive\nfl-fantasy-football2022\Data Analysis\weekly_scores")
standings_file = Path(r"H:\My Drive\nfl-fantasy-football2022\standings.json")

scores = [score for score in scores_directory.glob("*.csv")]

with open(teammapping_file_full, 'r') as f:
    dict_full = dict(eval(f.read()))

teams = [name for name in dict_full.values()]

weeks = [str(w) for w in range(1, 17)]


df_total_scores = pd.DataFrame()
df_total_scores = pd.DataFrame(columns=weeks, index = teams)
for week in scores:
    df_scores = pd.read_csv(week, index_col='Unnamed: 0')
    df_scores = df_scores.rename(columns=dict_full)
    
    week_id = week.name.split('wk_')[1].split('_')[0]
    positions = ['QB', 'WR1', 'WR2', 'RB1', 'RB2', 'TE', 'W/R/T', 'K', 'DEF']
    for manager in df_scores:
        df_total_scores.loc[manager, week_id] = df_scores.loc[positions, manager].sum()

df_total_scores.replace(0, np.nan, inplace=True)

df_wins = pd.DataFrame()
df_wins = pd.DataFrame(columns=['Matchup Wins'], index = teams)

standings = pd.read_json(standings_file, orient = 'index')['league'][0][1]['standings'][0]['teams']

for team in standings:
    if team != 'count':
        df_wins.loc[teams[int(standings[team]['team'][0][19]['managers'][0]['manager']['manager_id'])-1]] = int(standings[team]['team'][2]['team_standings']['outcome_totals']['wins'])

df_wins["Points Wins"] = 0
for week in range(1, CURRENT_WEEK + 1):
    for team, _ in df_total_scores[str(week)].nlargest(6).iteritems():
        df_wins.loc[team, 'Points Wins'] = int(1 + df_wins.loc[team]['Points Wins'])

df_wins.index = df_wins.index.rename('Team')
df_wins['Total Wins'] = df_wins['Matchup Wins'] + df_wins['Points Wins']
df_wins["Position"] = df_wins["Total Wins"].rank(
    method="dense", ascending=False)
df_wins = df_wins.sort_values(by=["Position"])

file_name = 'Standings_Week' + str(CURRENT_WEEK) + '.xlsx'
df_wins.to_excel(file_name)
# %%
gc = gspread.service_account()
spreadsheet = gc.open('RBB Standings')
#spreadsheet.share('fatpony@gmail.com', perm_type='user', role='writer')
#%%
#spreadsheet.del_worksheet(placeholder)
#placeholder = spreadsheet.add_worksheet(title='placeholder', rows=1, cols=1)
#spreadsheet.del_worksheet(worksheet)
#worksheet = spreadsheet.add_worksheet(title='Standings', rows=1, cols=1)
worksheet = spreadsheet.get_worksheet(0)
#print(spreadsheet.worksheets())
set_with_dataframe(worksheet, df_wins, include_index=True, resize=True)

worksheet.format("A1:E1", {
    "backgroundColor": {
      "red": 0.0,
      "green": 0.0,
      "blue": 0.0
    },
    "horizontalAlignment": "CENTER",
    "textFormat": {
      "foregroundColor": {
        "red": 1.0,
        "green": 1.0,
        "blue": 1.0
      },
      "fontSize": 12,
      "bold": True
    }
})
worksheet.format("A2:A13", {
    "backgroundColor": {
      "red": 0.2,
      "green": 0.2,
      "blue": 0.2
    },
    "horizontalAlignment": "CENTER",
    "textFormat": {
      "foregroundColor": {
        "red": 1.0,
        "green": 1.0,
        "blue": 1.0
      },
      "fontSize": 12,
      "bold": True
    }
})
worksheet.format("B2:E13", {
    "backgroundColor": {
      "red": 1.0,
      "green": 1.0,
      "blue": 1.0
    },
    "horizontalAlignment": "CENTER",
    "textFormat": {
      "foregroundColor": {
        "red": 0.0,
        "green": 0.0,
        "blue": 0.0
      },
      "fontSize": 12,
      "bold": False
    }
})
#spreadsheet.del_worksheet(placeholder)
# %%
