import pandas as pd
import xlsxwriter

def print_all(team_elo):
    workbook = xlsxwriter.Workbook("output.xlsx")
    worksheet = workbook.add_worksheet('ELO')
    row, col = 1, 0
    BOLD = workbook.add_format({'bold': True})
    worksheet.write(0,0, 'Team', BOLD)
    worksheet.write(0,1, 'ELO', BOLD)
    for team in team_elo:
        worksheet.write(row, col, team)
        worksheet.write(row, col+1, team_elo[team])
        row+=1
    workbook.close()

if __name__ == '__main__':
    sheet = pd.read_excel('ipl_data.xlsx')
    team_elo = {}

    current_year=2008
    k = 20
    base_elo = 1200
    carry = 2/3
    home_adv = 100

    for i, row in sheet.iterrows():

        t1 = row['Team 1']
        t2 = row['Team 2']
        if t1 not in team_elo:
            team_elo[t1] = base_elo
        if t2 not in team_elo:
            team_elo[t2] = base_elo
        p1 = 1 / (1 + 10 ** ((team_elo[t1] - team_elo[t2]) / 400))
        p2 = 1 / (1 + 10 ** ((team_elo[t2] - team_elo[t1]) / 400))
        if row['Winner'] == t1:
            team_elo[t1] += k * (1 - p1)
            team_elo[t2] += k * (0 - p2)
        elif row['Winner'] == t2:
            team_elo[t2] += k * (1 - p2)
            team_elo[t1] += k * (0 - p1)
        # keep 2/3 of previous year's ELO
        if row['Year']>current_year:
            current_year=row['Year']
            for team in team_elo:
                team_elo[team] = carry*team_elo[team] + (1-carry)*base_elo

    print_all(team_elo)
    for team in team_elo:
        print(team, team_elo[team])
