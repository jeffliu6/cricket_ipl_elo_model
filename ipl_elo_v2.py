import pandas as pd
import xlsxwriter
from scipy.optimize import minimize

def print_all(team_elo):
    workbook = xlsxwriter.Workbook('elo.xlsx')
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

#if __name__ == '__main__':
def do_all(myList):
    sheet = pd.read_excel('ipl_data.xlsx')
    team_elo = {}

    columns = ['Team 1', 'Team 2', 'Team 1 Prob', 'Team 2 Prob', 'Winner Prediction', 'Actual Winner', 'Correct']
    tests = pd.DataFrame(columns=columns)
    total_correct = 0
    count = 0

    current_year = sheet['Year'][0]
    k = myList[0]
    base_elo = 1500
    carry = myList[1]
    home_adv = 0
    bad=0

    for i, row in sheet.iterrows():
        year = int(sheet['Date'][0][-4:])
        t1 = row['Team 1']
        t2 = row['Team 2']
        if t1 not in team_elo:
            team_elo[t1] = base_elo
        if t2 not in team_elo:
            team_elo[t2] = base_elo
        p2 = 1 / (1 + 10 ** (((team_elo[t1] + home_adv) - team_elo[t2]) / 400))
        p1 = 1 / (1 + 10 ** ((team_elo[t2] - (team_elo[t1] + home_adv)) / 400))
        # p2 = 1 / (1 + 10 ** ((team_elo[t1] - team_elo[t2]) / 400))
        # p1 = 1 / (1 + 10 ** ((team_elo[t2] - team_elo[t1]) / 400))

        # Predict
        winner = t1 if p1 > p2 else t2
        actual = row['Winner']
        correct = int(winner == actual)
        if(row['Year']>=2017):
            # win_factor = 1 if t1==actual else -1
            # if not correct:
            #     bad += (p2-p1) * win_factor
            # else:
            #     bad += (p2-p1) * (win_factor * 0.5)
            total_correct += correct
            count += 1
            #print(t1 + ", " + t2 + ", (" + str(p1) + ", " + str(p2) + "), " + str(correct))
        test_row = [t1, t2, p1, p2, winner, actual, correct]
        tests = tests.append({columns[i]: test_row[i] for i in range(len(columns))}, ignore_index=True)

        if row['Winner'] == t1:
            team_elo[t1] += k * (1 - p1)
            team_elo[t2] += k * (0 - p2)
        elif row['Winner'] == t2:
            team_elo[t2] += k * (1 - p2)
            team_elo[t1] += k * (0 - p1)
        # keep 2/3 of previous year's ELO
        if row['Year'] > current_year:
            #print("End of " + str(current_year))
            #for team in team_elo:
            #    print(team, team_elo[team])
            current_year=row['Year']
            for team in team_elo:
                team_elo[team] = carry*team_elo[team] + (1-carry)*base_elo

    #print("End of " + str(current_year))
    print_all(team_elo)
    for team in team_elo:
        print(team, team_elo[team])

    accuracy = total_correct / count
    print('Accuracy: {}'.format(accuracy))
    #print('Bad : ' + str(bad))
    #return bad
    return 1-accuracy #(1 - accuracy)*bad
    #writer = pd.ExcelWriter('tests.xlsx', engine='xlsxwriter')
    #tests.to_excel(writer, sheet_name='Predictions', index=False, columns=columns)

if __name__ == '__main__':
    #do_all()
    #do_all( [20, 1500, 2/3, 100])
    #temp = minimize(do_all, [45, 2/3], method='L-BFGS-B',options=dict(maxiter=20), bounds = [(0, 70), (0, 1)])
    #print(temp)
    do_all([40, 2/3])
