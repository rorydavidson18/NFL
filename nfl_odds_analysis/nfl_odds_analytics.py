from __future__ import division
import xlrd
import operator
import sys, os
from collections import OrderedDict

workbook = xlrd.open_workbook('nfl_odds.xlsx')
sheet = workbook.sheet_by_index(0)
number_rows = sheet.nrows
number_columns = sheet.ncols
f = open('results.txt', 'wr')
games_list = []
teams_dict = {}
spread_teams_dict = {}  #the values are lists within a list, home record, away record, as favorite record, as underdog


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
    team1_is_favorite = False
    try:
        spread_teams_dict[team1]
    except KeyError:
        spread_teams_dict[team1] = [[0, 0, 0], [0, 0, 0], [0, 0 ,0], [0, 0, 0]]
    try:
        spread_teams_dict[team2]
    except KeyError:
        spread_teams_dict[team2] = [[0, 0, 0], [0, 0, 0], [0, 0 ,0], [0, 0, 0]]
    if team1_spread<0:
        team1_is_favorite = True
    if team1_score-team2_score==team1_spread or team1_score-team2_score==team2_spread:
        spread_teams_dict[team1][0][2] += 1
        spread_teams_dict[team2][1][2] += 1
        if team1_is_favorite:
            spread_teams_dict[team1][2][2] += 1
            spread_teams_dict[team2][3][2] += 1
        else:
            spread_teams_dict[team2][2][2] += 1
            spread_teams_dict[team1][3][2] += 1
    elif (team1_spread<0 and team1_score-team2_score<((-1.0)*team1_spread)) or (team2_spread<0 and team1_score-team2_score<team2_spread):
        spread_teams_dict[team1][0][1] += 1
        spread_teams_dict[team2][1][0] += 1
        if team1_is_favorite:
            spread_teams_dict[team1][2][1] += 1
            spread_teams_dict[team2][3][0] += 1
        else:
            spread_teams_dict[team1][3][1] += 1
            spread_teams_dict[team2][2][0] += 1
    else:
        spread_teams_dict[team2][1][1] += 1
        spread_teams_dict[team1][0][0] += 1
        if team1_is_favorite:
            spread_teams_dict[team1][2][0] += 1
            spread_teams_dict[team2][3][1] += 1
        else:
            spread_teams_dict[team1][3][0] += 1
            spread_teams_dict[team2][2][1] += 1


def calc_team_odds_and_ats():
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
    temp_teams_dict = OrderedDict(sorted(teams_dict.items(), key=operator.itemgetter(1)))
    return temp_teams_dict


if __name__ == "__main__":
    for row in range(number_rows):
        game = []
        for col in range(number_columns):
            value = sheet.cell(row, col).value
            game.append(value)
        games_list.append(game)
    teams_dict = calc_team_odds_and_ats()
    for team in teams_dict:
        teams_dict[team] = round(teams_dict[team], 2)
        f.write(str(team) + ':\n' + 'Moneyline: ' + str(teams_dict[team]) + '\n' + 'ATS:\n' +
                '\tHome: ' + str(spread_teams_dict[team][0][0]) + '-' + str(spread_teams_dict[team][0][1]) + '-' + str(spread_teams_dict[team][0][2]) +
                '\tAway: ' + str(spread_teams_dict[team][1][0]) + '-' + str(spread_teams_dict[team][1][1]) + '-' + str(spread_teams_dict[team][1][2]) +
                '\n\tAs Favorite: ' + str(spread_teams_dict[team][2][0]) + '-' + str(spread_teams_dict[team][2][1]) + '-' + str(spread_teams_dict[team][2][2]) +
                '\tAs Underdog: ' + str(spread_teams_dict[team][3][0]) + '-' + str(spread_teams_dict[team][3][1]) + '-' + str(spread_teams_dict[team][3][2]) + '\n\n')

