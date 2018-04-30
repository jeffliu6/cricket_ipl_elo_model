import pandas as pd
import xlsxwriter

def print_all(team_elo):
    workbook = xlsxwriter.Workbook('predictions.xlsx')
    worksheet = workbook.add_worksheet('Predictions')
    row, col = 1, 0
    BOLD = workbook.add_format({'bold': True})
    worksheet.write(0,0, 'Team 1', BOLD)
    worksheet.write(0,1, 'Team 2', BOLD)
    worksheet.write(0,2, 'P1', BOLD)
    worksheet.write(0,3, 'P2', BOLD)
    worksheet.write(0,4, 'Winner', BOLD)
    for t1 in team_elo:
        for t2 in team_elo:
            if t1 == t2:
                continue
            p1 = 1 / (1 + 10 ** ((team_elo[t2] - team_elo[t1]) / 400))
            p2 = 1 / (1 + 10 ** ((team_elo[t1] - team_elo[t2]) / 400))
            winner = t1 if p1 > p2 else t2
            worksheet.write(row, col, t1)
            worksheet.write(row, col + 1, t2)
            worksheet.write(row, col + 2, p1)
            worksheet.write(row, col + 3, p2)
            worksheet.write(row, col + 4, winner)
            row+=1
    workbook.close()

def main(k_carry):
    teams = {'Sunrisers Hyderabad', 'Chennai Super Kings', 'Kings XI Punjab', 'Kolkata Knight Riders', 'Rajasthan Royals', 'Mumbai Indians', 'Royal Challengers Bangalore', 'Delhi Daredevils'}
    sheet = pd.read_excel('ipl_data.xlsx')
    team_elo = {}

    columns = ['Team 1', 'Team 2', 'Team 1 Prob', 'Team 2 Prob', 'Winner Prediction', 'Actual Winner', 'Correct']
    tests = pd.DataFrame(columns=columns)
    total_correct = 0
    count = 0

    current_year = sheet['Year'][0]
    k = k_carry[0]
    carry = k_carry[1]
    base_elo = 1200
    home_adv = 0
    bad = 0
    for i, row in sheet.iterrows():
        t1 = row['Team 1']
        t2 = row['Team 2']
        if t1.strip() in teams and t2.strip() in teams:
            if t1 not in team_elo:
                team_elo[t1] = base_elo
            if t2 not in team_elo:
                team_elo[t2] = base_elo
            p1 = 1 / (1 + 10 ** ((team_elo[t2] - team_elo[t1]) / 400))
            p2 = 1 / (1 + 10 ** ((team_elo[t1] - team_elo[t2]) / 400))

            winner = t1 if p1 > p2 else t2
            actual = row['Winner']
            win_factor = 1 if actual == t1 else -1
            correct = int(winner == actual)
            if row['Year'] == 2017:
                bad += (p2 - p1) * win_factor
                total_correct += correct
                count += 1
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
            current_year = row['Year']
            for team in team_elo:
                team_elo[team] = carry*team_elo[team] + (1-carry)*base_elo

    print_all(team_elo)
    for team in team_elo:
        print(team, team_elo[team])

    accuracy = total_correct / count
    print('Accuracy: {}'.format(accuracy))
    print(bad)
    #writer = pd.ExcelWriter('tests.xlsx', engine='xlsxwriter')
    #tests.to_excel(writer, sheet_name='Predictions', index=False, columns=columns)
    return bad


if __name__ == '__main__':
    main([40, 0.6])
    #from scipy.optimize import minimize
    #res = minimize(main, [10.82017935,  0.50900893], bounds=[(0, None), (0, 1)], method='TNC', options=dict(maxiter=20))
    #print(res)
