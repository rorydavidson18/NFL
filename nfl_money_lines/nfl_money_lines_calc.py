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


def calculate_payout(odds, bet):
    if odds==100 or odds==-100:
        return bet
    elif odds<0:
        payout_per_dollar = 100.00/(odds*(-1.00))
        return payout_per_dollar*bet
    elif odds>0:
        payout_per_dollar = (odds*1.00)/100.00
        return payout_per_dollar*bet


if __name__ == "__main__":
    for row in range(number_rows):
        game = []
        for col in range(number_columns):
            value = sheet.cell(row, col).value
            game.append(value)
        games_list.append(game)
    for game in games_list:
        if game[4]==game[5]:
            pass
        elif game[4]>game[5]:
            try:
                teams_dict[game[0]] = teams_dict[game[0]] + calculate_payout(game[1], 100)
            except KeyError:
                teams_dict[game[0]] = calculate_payout(game[1], 100)

            try:
                teams_dict[game[2]] = teams_dict[game[2]] - 100
            except KeyError:
                teams_dict[game[2]] = -100
        else:
            try:
                teams_dict[game[2]] = teams_dict[game[2]] + calculate_payout(game[3], 100)
            except KeyError:
                teams_dict[game[2]] = calculate_payout(game[3], 100)
            try:
                teams_dict[game[0]] = teams_dict[game[0]] - 100
            except KeyError:
                teams_dict[game[0]] = -100
    teams_dict = OrderedDict(sorted(teams_dict.items(), key=operator.itemgetter(1)))
    for team in teams_dict:
        teams_dict[team] = round(teams_dict[team], 2)
        print str(team) + ': ' + str(teams_dict[team])
    # for foo in teams_dict:
    #     teams_dict[foo] = round(teams_dict[foo], 2)
    #     print str(foo) + ': ' + str(teams_dict[foo])
        # payout_away = calculate_payout(game[1], 100)
        # payout_home = calculate_payout(game[3], 100)

