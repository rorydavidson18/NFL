from __future__ import division
import xlrd
import operator
from collections import OrderedDict

workbook = xlrd.open_workbook('nfl_money_lines.xlsx')
sheet = workbook.sheet_by_index(0)
number_rows = sheet.nrows
number_columns = sheet.ncols
games_list = []
teams_dict = {}
spread_teams_dict = {}


def calculate_moneyline_payout(moneyline, bet):
    if moneyline==100 or moneyline==-100:
        return 0
    elif moneyline<0:
        payout_per_dollar = 100.00/(moneyline*(-1.00))
        return payout_per_dollar*bet
    elif moneyline>0:
        payout_per_dollar = (moneyline*1.00)/100.00
        return payout_per_dollar*bet

def spread_result(team1, team1_score, team2, team2_score, team1_spread, team2_spread):
    if team1_spread<0 and team1_score-team2_score>team1_spread:
        try:
            spread_teams_dict[team1][1] += 1
        except KeyError:
            spread_teams_dict[team1] = [0, 1]
        try:
            spread_teams_dict[team2][0] += 1
        except KeyError:
            spread_teams_dict[team2] = [1, 0]
    else:
        try:
            spread_teams_dict[team2][1] += 1
        except KeyError:
            spread_teams_dict[team2] = [0, 1]
        try:
            spread_teams_dict[team1][0] += 1
        except KeyError:
            spread_teams_dict[team1] = [1, 0]


if __name__ == "__main__":
    for row in range(number_rows):
        game = []
        for col in range(number_columns):
            value = sheet.cell(row, col).value
            game.append(value)
        games_list.append(game)
    for game in games_list:
        spread_result(game[0], game[4], game[2], game[5], game[6], game[7])
        if game[4]==game[5]:
            pass
        elif game[4]>game[5]:
            try:
                teams_dict[game[0]] = teams_dict[game[0]] + calculate_moneyline_payout(game[1], 100)
            except KeyError:
                teams_dict[game[0]] = calculate_moneyline_payout(game[1], 100)

            try:
                teams_dict[game[2]] = teams_dict[game[2]] - 100
            except KeyError:
                teams_dict[game[2]] = -100
        else:
            try:
                teams_dict[game[2]] = teams_dict[game[2]] + calculate_moneyline_payout(game[3], 100)
            except KeyError:
                teams_dict[game[2]] = calculate_moneyline_payout(game[3], 100)
            try:
                teams_dict[game[0]] = teams_dict[game[0]] - 100
            except KeyError:
                teams_dict[game[0]] = -100
    teams_dict = OrderedDict(sorted(teams_dict.items(), key=operator.itemgetter(1)))
    for team in teams_dict:
        teams_dict[team] = round(teams_dict[team], 2)
        print str(team) + ': ' + str(teams_dict[team])
        print str(team) + ': ' + str(spread_teams_dict[team][0]) + '-' + str(spread_teams_dict[team][1])
    # for foo in teams_dict:
    #     teams_dict[foo] = round(teams_dict[foo], 2)
    #     print str(foo) + ': ' + str(teams_dict[foo])
        # payout_away = calculate_payout(game[1], 100)
        # payout_home = calculate_payout(game[3], 100)

