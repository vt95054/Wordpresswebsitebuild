import openpyxl

prediction_workbook = openpyxl.load_workbook('logreg_seed_starter.xlsx')
predictions = prediction_workbook.get_sheet_by_name('logreg_seed_starter')

team_workbook = openpyxl.load_workbook('DataFiles/Teams.xlsx')
teams = team_workbook.get_sheet_by_name('Teams')

def get_year_t1_t2(event_id):
    """Return a tuple with ints `year`, `team1` and `team2`."""
    return (int(x) for x in event_id.split('_'))

def get_team_by_id(team_id):
	for i in range(2, teams.max_row + 1):
		if team_id == teams.cell(row=i, column=1).value:
			return teams.cell(row=i, column=2).value

for i in range(2, predictions.max_row + 1):
	event_id = predictions.cell(row=i, column=1).value
	estimate = predictions.cell(row=i, column=2).value
	year, team_1_id, team_2_id = get_year_t1_t2(event_id)
	team_1_name = get_team_by_id(team_1_id)
	team_2_name = get_team_by_id(team_2_id)

	if estimate >= 0.5:
		winner = team_1_name
	else:
		winner = team_2_name

	predictions.cell(row=i, column=3).value = team_1_id
	predictions.cell(row=i, column=4).value = team_1_name
	predictions.cell(row=i, column=5).value = team_2_id
	predictions.cell(row=i, column=6).value = team_2_name
	predictions.cell(row=i, column=7).value = winner

prediction_workbook.save('predictions.xlsx')