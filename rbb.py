import pandas as pd
import matplotlib.pylab as pl
from pathlib import Path


def import_data(data):
    temp_frame = pd.read_csv(
        data, sep="\t", header=None, index_col=0, usecols=[0, 2, 3]
    )
    temp_frame[2] = temp_frame[2].str.strip()
    temp_frame[3] = temp_frame[3].str.strip()

    team = data.name.split(".")[0]
    wins = len(temp_frame[temp_frame[2] == "Win"])
    points = []
    for row in temp_frame.itertuples():
        if row[1] == "Win":
            points.append(max([float(x.strip()) for x in row[2].split("-")]))
        elif row[1] == "Loss":
            points.append(min([float(x.strip()) for x in row[2].split("-")]))

    return pd.DataFrame([points], index=[team]), [team, wins]


file_directory = Path(r"C:\Users\Joseph\Desktop\RBB 2019")
file_list = [f for f in file_directory.glob("*.csv")]

total_points = []
total_wins = {}
sums = []
for team in file_list:
    results_points, results_wins = import_data(team)
    total_points.append(results_points)
    sums.append(sum(results_points.sum(axis=1)))
    total_wins[results_wins[0]] = results_wins[1]

standings = pd.DataFrame.from_dict(
    total_wins, orient="index", columns=["Wins Standard"]
)
standings.insert(0, "Total Points", sums)
scores = pd.concat(total_points, axis=0)

for week in scores:
    for team in scores[week].nlargest(6).iteritems():
        total_wins[team[0]] += 1

standings["Wins New"] = total_wins.values()

standings["Rank Standard"] = standings["Wins Standard"].rank(
    method="dense", ascending=False
)

standings["Rank New"] = standings["Wins New"].rank(method="dense", ascending=False)

standings["Rank Change"] = standings["Rank New"] - standings["Rank Standard"]

standings["Playoffs Standard"] = [
    "No",
    "No",
    "No",
    "No",
    "Yes",
    "No",
    "No",
    "Yes",
    "No",
    "Yes",
    "Yes",
    "No",
]
standings["Playoffs New"] = [
    "No",
    "Yes",
    "No",
    "No",
    "Yes",
    "No",
    "No",
    "Yes",
    "No",
    "No",
    "Yes",
    "No",
]

standings.to_excel("Standings.xlsx")

