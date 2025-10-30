##########################################################################
# the imports
##########################################################################
import requests
import json
import re 
import xlsxwriter
import datetime 

##########################################################################
# the objects
##########################################################################
class team:
    def __init__(self, team_id, name, player_data):
        self.team_id = team_id 
        self.name = name
        self.player_data = player_data
        self.past_5_games = []


##########################################################################
# the variables 
##########################################################################
BASE_URL = 'https://data.sihf.ch/Statistic/api/cms/cache'
# results
RESULT_CACHE = '600?'
RESULT_PAGE = 'alias=results'
NEXT_GAME = '&size=nextday'
GAME_QUERY = '&searchQuery=1,2,4,10//1/////nextDayDelayed'
ALL_GAMES_QUERY = '&searchQuery=1,10//1'
GAME_FILTER = '&filterQuery=/1&filterBy=league,deferredState'
GAME_ORDER = '&orderBy=league,deferredState&orderByDescending=false'
FILTER_QUERY_PAST_GAMES = '&filterQuery='
SEASON_START_TIL_TODAY = ('/09.09.2025-' + datetime.datetime.now().strftime("%d.%m.%Y") + '//all/al')
FILTER_BY_PAST_GAMES = '&filterBy=season,phase,date,deferredState,team1,team2'
ORDER_BY_PAST_GAMES = '&orderBy=date&orderByDescending=true'
#kind of stat 
STAT_CACHE = '300?'
SUMMARY_STAT = 'alias=player'
GOAL_STATS = 'alias=playerGoalAssist'
SHOT_STAT = 'alias=playerShotDetail'
FACEOFF_STAT = 'alias=playerFaceoff'
GOALIE_SUMMARY_STAT = 'alias=goalkeeper'
GOALIE_SHOT_STAT = 'alias=goalkeeperShotDetail'
# strandard stuff for language, sorting and query definition 
STANDARD_SORTING = '&orderBy=player&orderByDescending=false'
STANDARD_INFO = '&skip=-1&language=de'
QUARY_BASIC = '&searchQuery=1//1&filterQuery=' 
##########################################################################
# FILTERS                                                                #
##########################################################################
# season 
CURRENT_SEASON = '2026'
# phase
PHASE_REGULAR = '/4940'
# Teams 
ALL_TEAMS = '/all'
BIEL = '/102128'
EHC_KLOTEN = '/101149'
EV_ZUG = '/101144'
FRIBOURG_GOTTERON = '103138'
GENF_SERVETTE = '/103140'
AJOIE = '/103144'
AMBRI_PIOTTA = '/101152'
DAVOS = '/101151'
LUGANO = '/101150'
LAUSANNE = '103141'
SC_BERN = '/102126'
SCL_TIGERS = '/102127'
RAPPI_LAKERS = '/101060'
ZSC_LIONS = '/101139'
# positions
ALL_POSITIONS = '/all'
VERTEIDIGER = '/2'
STUERMER = '/3'
# licence 
ALL_LICENCES = '/all'
SWISS_LICENCE = '/1'
FOREIGN_LICENCE = '/2'
# more standard stuff
FILTERING_INFO = '&filterBy=Season,Phase,Team,Position,Licence'
STANDARD_ENDING = "&take=1000&callback=externalStatisticsCallback"
##########################################################################
# excel stuff
##########################################################################
CURRENT_DATE = datetime.datetime.now().strftime("%Y_%m_%d_%H_%M")
MY_STORAGEPATH = "C:/Users/Jessi/Documents/Hockey_stats/Stats_from_"
FILE_ENDING = ".xlsx"
MY_EXCELPATH = MY_STORAGEPATH + CURRENT_DATE + FILE_ENDING 
TITLE_HOME_FORMAT = {'bold': True
                     ,'align': 'center'
                     ,'font_color': 'White'
                     ,'font_size': 14
                     ,'bg_color': '#BE5014'}
TITLE_AWAY_FORMAT = {'bold': True
                     ,'align': 'center'
                     ,'font_color': 'White'
                     ,'font_size': 14
                     ,'bg_color': '#0B3040'}

##########################################################################
# tha functions
##########################################################################
#
# create the excel file
#
def create_excel_file():
    my_excelfile = xlsxwriter.Workbook(MY_EXCELPATH)
    standard_format = my_excelfile.add_format()
    standard_format.set_font_name('Aptos Narrow')
    standard_format.set_font_size(11)
    return my_excelfile
#
# send the request and return the data as dictionary (this is the heart of the whole exercise)
#
def send_request(request_url: str) -> dict: 
    # add some nice errorhandling (analog condition handling in deltamakro V2 where the except case is WHEN=(WRN|ERR,...))
    try:
     # make a GET request to an API endpoint
      response = requests.get(request_url)
    except: 
        raise Exception('Could not reach the API endpoint for the url: ' + request_url)
    try:
     # use regex to extract the JSON and remove the JS stuff
     # meaning we search for externalStatisticsCallback( some stuff ); 
      response_without_JS = re.search(r'externalStatisticsCallback\((.*)\);', response.text).group(1)
      # place the JSON into a dcitionary
      json_data = json.loads(response_without_JS)
      return json_data
    except:
        raise Exception('Could not parse the response into JSON for the url: ' + request_url)

#
# add a worhsheet to the excel file which gets created at the beginning (worksheet = tab in Excel)
#
def add_worksheet(my_excelfile: xlsxwriter.Workbook, sheet_name:str) -> xlsxwriter.worksheet:
    try: 
        worksheet = my_excelfile.add_worksheet(sheet_name)
        return worksheet
    except:
        raise Exception('Could not add the worksheet: ' + sheet_name)

#
# write the titles for the player stats to the desired worksheet 
#
def write_titels_to_worksheet_players(my_worksheet: xlsxwriter.worksheet, start_row: int, start_column: int) -> int:
    my_worksheet.write(start_row, start_column, 'Player')
    start_column +=1
    my_worksheet.write(start_row, start_column, 'GP')
    start_column +=1
    my_worksheet.write(start_row, start_column, 'Goals')
    start_column +=1
    my_worksheet.write(start_row, start_column, 'Assists')
    start_column +=1
    my_worksheet.write(start_row, start_column, 'Shots on Goal')
    start_column +=1
    my_worksheet.write(start_row, start_column, 'Schüsse aufs Tor aus Slot')
    start_column +=1
    my_worksheet.write(start_row, start_column, 'Schüsse nebes Tor')
    start_column +=1
    my_worksheet.write(start_row, start_column, 'Schüsse an die Torumrandung')
    start_column +=1
    my_worksheet.write(start_row, start_column, '+/-')
    start_column +=1
    my_worksheet.write(start_row, start_column, 'Bullys Total')
    start_column +=1
    my_worksheet.write(start_row, start_column, 'Gewonnene Bullys')
    start_column +=1
    my_worksheet.write(start_row, start_column, 'Verlorene Bullys')
    start_column +=1
    my_worksheet.write(start_row, start_column, 'Blocked Shots')
    start_column +=1
    my_worksheet.write(start_row, start_column, 'Shorthanders')
    start_column +=1
    my_worksheet.write(start_row, start_column, 'Shorthander Assists')
    start_column +=1
    my_worksheet.write(start_row, start_column, 'Powerplay Goals')
    start_column +=1
    my_worksheet.write(start_row, start_column, 'Powerplay Assists')
    start_column +=1
    my_worksheet.write(start_row, start_column, 'Penalty Minutes')
    return start_column

#
# create the request URLs and get the data 
#
def get_the_data():
  request_url = (BASE_URL + STAT_CACHE + SUMMARY_STAT + STANDARD_INFO + QUARY_BASIC 
               + CURRENT_SEASON + PHASE_REGULAR + ALL_TEAMS + ALL_POSITIONS 
               + ALL_LICENCES + FILTERING_INFO + STANDARD_SORTING + STANDARD_ENDING)
  general_data = send_request(request_url)
  request_url = (BASE_URL + STAT_CACHE + GOAL_STATS + STANDARD_INFO +QUARY_BASIC 
               + CURRENT_SEASON + PHASE_REGULAR + ALL_TEAMS 
               + ALL_POSITIONS + ALL_LICENCES + FILTERING_INFO + STANDARD_SORTING + STANDARD_ENDING)
  goal_data = send_request(request_url)
  request_url = (BASE_URL + STAT_CACHE + SHOT_STAT + STANDARD_INFO +QUARY_BASIC 
               + CURRENT_SEASON + PHASE_REGULAR + ALL_TEAMS + ALL_POSITIONS 
               + ALL_LICENCES + FILTERING_INFO + STANDARD_SORTING + STANDARD_ENDING)
  shot_data = send_request(request_url)
  request_url = (BASE_URL + STAT_CACHE + FACEOFF_STAT + STANDARD_INFO +QUARY_BASIC 
               + CURRENT_SEASON + PHASE_REGULAR + ALL_TEAMS + ALL_POSITIONS 
               + ALL_LICENCES + FILTERING_INFO + STANDARD_SORTING + STANDARD_ENDING)
  faceoff_data = send_request(request_url)
  request_url = (BASE_URL + STAT_CACHE + GOALIE_SUMMARY_STAT + STANDARD_INFO +QUARY_BASIC 
               + CURRENT_SEASON + PHASE_REGULAR + ALL_TEAMS + ALL_POSITIONS 
               + ALL_LICENCES + FILTERING_INFO + STANDARD_SORTING + STANDARD_ENDING)
  goalie_data = send_request(request_url)
  request_url = (BASE_URL + STAT_CACHE + GOALIE_SHOT_STAT + STANDARD_INFO +QUARY_BASIC + CURRENT_SEASON 
               + PHASE_REGULAR + ALL_TEAMS + ALL_POSITIONS + ALL_LICENCES + FILTERING_INFO 
               + STANDARD_SORTING + STANDARD_ENDING)
  goalie_shot_data = send_request(request_url)
  return general_data, goal_data, shot_data, faceoff_data, goalie_data, goalie_shot_data

#
# add the general data to a dict
#
def add_general_data(i: list, team_name: str) -> dict:
    # create a temporary dict to store the player data in a strructured way which is easy to extend
    # by additional fields later on as well retreval by key is possible (aka search) which makes the 
    # gathering and adding of data easier
    temp_dict = { 'Name': i[1]
                , 'Team': team_name
                , 'Position': i[3]
                , 'Games_played': i[4]
                , 'Goals': i[5]
                , 'Assists': i[6]
                , 'PIM': i[9]
                , '+/-': i[10]}
    return temp_dict

#
# add goal data
#
def add_goal_data(i: list) -> dict:
    temp_dict = { 'Assists': i[6]
                , 'PP_Goals': i[12]
                , 'PP_Assists': i[13]
                , 'Box_Play_Goals': i[14]
                , 'Box_Play_Assists': i[15]}
    return temp_dict

#
# add shot data
#
def add_shot_data(i: list) -> dict:
    temp_dict = { 'Shots_total': i[5]
                , 'Shots_from_slot': i[8]
                , 'Shots_missed': i[10]
                , 'Shots_on_border': i[12]
                , 'Blocked_shots': i[14]}
    return temp_dict

#
# add the goalie data to a dict
#
def add_goalie_data(i: list, j: list) -> dict:
    # create a temporary dict to store the player data in a strructured way which is easy to extend
    # by additional fields later on as well retreval by key is possible (aka search) which makes the 
    # gathering and adding of data easier
    temp_dict = { 'Name': i[1]
                , 'Team': i[2]['name']
                , 'Position': 'Torhüter'
                , 'Goals_agains': i[4]
                , 'Total_shots_agains': j[4]
                , 'Total_save%': i[9]
                , 'Shots_from_slot': j[6]
                , 'Slot_save%': j[12]
                , 'Missed_shots': j[8]
                , 'Shots_against_border': j[10]
                , 'PIM': i[11]}
    return temp_dict
#
# take the gathered data and sort them into dicts per team 
#
def create_team_data():
    # get the data from the API

    general_data, goal_data, shot_data, faceoff_data, goalie_data, goalie_shot_data = get_the_data()
    # create the dicts to tell python that they actually are dicts
    ehc_biel_data = {}
    ehc_kloten_data = {}
    ev_zug_data = {}
    fribourg_gotteron_data = {}
    genf_servette_data = {}
    hc_ajoie_data = {}
    hc_ambri_piotta_data = {}
    hc_davos_data = {}
    hc_lugano_data = {} 
    hc_lausanne_data = {}
    sc_bern_data = {}
    scl_tigers_data = {}
    rapperswil_jona_lakers_data = {}
    zsc_lions_data = {}

    # the gathered data is a huge JSON which contains several dicts and lists. We actually 
    # do not care about the headers and dropdown options and therefore we directly go to the key 'data' which then is 
    # a list represented by i. each of them contains the statistic data of one player and yet another dict for the team info. 
    # we extract this one to filter the players by team and add to the corresponding team dict where the player name is the key
    for i in general_data['data']:
      # the team info is a dictionary stored inthe third list item and we actually need it later to get some details 
      team_raw = i[2]
      # grab the n numeric ID of the team 
      team_id = team_raw['id']
      # match (aka evaluate or switch) for team ID and not teamname to avoid issues with spelling changes
      match team_id:
         # Biel
        case 102128:
            # player name is always the second item in the list and since we are here in Python and not in COBOL 
            # we start counting at 0.
            ehc_biel_data[i[1]] = add_general_data(i, team_raw['name'])
        # Kloten
        case 101149:
            ehc_kloten_data[i[1]] = add_general_data(i, team_raw['name'])
        # Zug
        case 101144:
            ev_zug_data[i[1]] = add_general_data(i, team_raw['name'])
        # Fribourg  
        case 103138:
            fribourg_gotteron_data[i[1]] = add_general_data(i, team_raw['name'])
        # Genf
        case 103140:
            genf_servette_data[i[1]] = add_general_data(i, team_raw['name'])
        # Ajoie
        case 103144:
            hc_ajoie_data[i[1]] = add_general_data(i, team_raw['name'])
        # Ambri
        case 101152:
            hc_ambri_piotta_data[i[1]] = add_general_data(i, team_raw['name'])
        # Davos
        case 101151:
            hc_davos_data[i[1]] = add_general_data(i, team_raw['name'])
        # Lugano
        case 101150:
            hc_lugano_data[i[1]] = add_general_data(i, team_raw['name'])
        # Lausanne
        case 103141:
            hc_lausanne_data[i[1]] = add_general_data(i, team_raw['name'])
        # Bern
        case 102126:
            sc_bern_data[i[1]] = add_general_data(i, team_raw['name'])
        # Tigers
        case 102127:
            scl_tigers_data[i[1]] = add_general_data(i, team_raw['name'])
        # Rappi
        case 101060:
            rapperswil_jona_lakers_data[i[1]] = add_general_data(i, team_raw['name'])
        # ZSC
        case 101139:
            zsc_lions_data[i[1]] = add_general_data(i, team_raw['name'])
    # the dicts are now created for each team and all players with the general data 
    for i in goal_data['data']:
        # since the players are already sorted into the team dicts, we once again need the team ID to then use 
        # a match case to update the correct dict with the gaol specific data. But this time we go directly for the ID
        # as we already have the team name in the player dict from the previous data aggregation
        team_id = i[2]['id']
        match team_id:
            # Biel
            case 102128:
                # I have to agree, this is kinda cool and would not work in COBOL
                # we take the team dict, access the player sub dict by the player name 
                # (which is the Key in the main dict) with [i[1]] (because we have to access the second list item for the name)
                # we then update the said sub dict with additional key value pairs which we receive as a dict 
                # by the function add_goal_data
                if i[1] in ehc_biel_data.keys():
                  ehc_biel_data[i[1]] |= add_goal_data(i)
            # Kloten
            case 101149:
                if i[1] in ehc_kloten_data.keys():
                  ehc_kloten_data[i[1]] |= add_goal_data(i)
            # Zug
            case 101144:
                if i[1] in ev_zug_data.keys():
                  ev_zug_data[i[1]] |= add_goal_data(i)
            # Fribourg
            case 103138:
                if i[1] in fribourg_gotteron_data.keys():
                  fribourg_gotteron_data[i[1]] |= add_goal_data(i)
            # Genf
            case 103140:
                if i[1] in genf_servette_data.keys(): 
                  genf_servette_data[i[1]] |= add_goal_data(i)
            # Ajoie
            case 103144:
                if i[1] in hc_ajoie_data.keys():
                  hc_ajoie_data[i[1]] |= add_goal_data(i)
            # Ambri
            case 101152:
                if i[1] in hc_ambri_piotta_data.keys():
                  hc_ambri_piotta_data[i[1]] |= add_goal_data(i)
            # Davos
            case 101151:
                if i[1] in hc_davos_data.keys():
                  hc_davos_data[i[1]] |= add_goal_data(i)
            # Lugano
            case 101150:
                if i[1] in hc_lugano_data.keys():
                  hc_lugano_data[i[1]] |= add_goal_data(i)
            # Lausanne
            case 103141:
                if i[1] in hc_lausanne_data.keys():
                  hc_lausanne_data[i[1]] |= add_goal_data(i)
            # Bern
            case 102126:
                if i[1] in sc_bern_data.keys():
                  sc_bern_data[i[1]] |= add_goal_data(i)
            # Tigers
            case 102127:
                if i[1] in scl_tigers_data.keys():
                  scl_tigers_data[i[1]] |= add_goal_data(i)
            # Rappi
            case 101060:
                if i[1] in rapperswil_jona_lakers_data.keys():
                  rapperswil_jona_lakers_data[i[1]] |= add_goal_data(i)
            # ZSC
            case 101139:
                if i[1] in zsc_lions_data.keys():
                  zsc_lions_data[i[1]] |= add_goal_data(i)
    # now the same for the shot data 
    for i in shot_data['data']:
        team_id = i[2]['id']
        match team_id:
            # Biel
            case 102128:
                if i[1] in ehc_biel_data.keys():
                  ehc_biel_data[i[1]] |= add_shot_data(i)
            # Kloten
            case 101149:
                if i[1] in ehc_kloten_data.keys():
                  ehc_kloten_data[i[1]] |= add_shot_data(i)
            # Zug
            case 101144:
                if i[1] in ev_zug_data.keys():
                  ev_zug_data[i[1]] |= add_shot_data(i)
            # Fribourg
            case 103138:
                if i[1] in fribourg_gotteron_data.keys():
                  fribourg_gotteron_data[i[1]] |= add_shot_data(i)
            # Genf
            case 103140:
                if i[1] in genf_servette_data.keys(): 
                  genf_servette_data[i[1]] |= add_shot_data(i)
            # Ajoie
            case 103144:
                if i[1] in hc_ajoie_data.keys():
                  hc_ajoie_data[i[1]] |= add_shot_data(i)
            # Ambri
            case 101152:
                if i[1] in hc_ambri_piotta_data.keys():
                  hc_ambri_piotta_data[i[1]] |= add_shot_data(i)
            # Davos
            case 101151:
                if i[1] in hc_davos_data.keys():
                  hc_davos_data[i[1]] |= add_shot_data(i)
            # Lugano
            case 101150:
                if i[1] in hc_lugano_data.keys():
                  hc_lugano_data[i[1]] |= add_shot_data(i)
            # Lausanne
            case 103141:
                if i[1] in hc_lausanne_data.keys():
                  hc_lausanne_data[i[1]] |= add_shot_data(i)
            # Bern
            case 102126:
                if i[1] in sc_bern_data.keys():
                  sc_bern_data[i[1]] |= add_shot_data(i)
            # Tigers
            case 102127:
                if i[1] in scl_tigers_data.keys():
                  scl_tigers_data[i[1]] |= add_shot_data(i)
            # Rappi
            case 101060:
                if i[1] in rapperswil_jona_lakers_data.keys():
                  rapperswil_jona_lakers_data[i[1]] |= add_shot_data(i)
            # ZSC
            case 101139:
                if i[1] in zsc_lions_data.keys():
                  zsc_lions_data[i[1]] |= add_shot_data(i)
    # now again the same for the faceoff data
    for i in faceoff_data['data']:
        team_id = i[2]['id']
        match team_id:
            # Biel
            case 102128:
                # we could actually use again the functions but since the added values are 
                # short enough to keep everything in one line, we stick to the update option 
                # to demonstrate that we not only can merge but also add key values
                if i[1] in ehc_biel_data.keys():
                  ehc_biel_data[i[1]].update({'FO_total': i[5], 'FO_won': i[8], 'FO_lost': i[9]})
            # Kloten 
            case 101149: 
                if i[1] in ehc_kloten_data.keys():
                  ehc_kloten_data[i[1]].update({'FO_total': i[5], 'FO_won': i[8], 'FO_lost': i[9]})
            # Zug
            case 101144:
                if i[1] in ev_zug_data.keys():
                  ev_zug_data[i[1]].update({'FO_total': i[5], 'FO_won': i[8], 'FO_lost': i[9]})
            # Fribourg 
            case 103138: 
                if i[1] in fribourg_gotteron_data.keys():
                  fribourg_gotteron_data[i[1]].update({'FO_total': i[5], 'FO_won': i[8], 'FO_lost': i[9]})
            # Genf
            case 103140: 
                if i[1] in genf_servette_data.keys(): 
                  genf_servette_data[i[1]].update({'FO_total': i[5], 'FO_won': i[8], 'FO_lost': i[9]})
            # Ajoie
            case 103144:
                if i[1] in hc_ajoie_data.keys():
                  hc_ajoie_data[i[1]].update({'FO_total': i[5], 'FO_won': i[8], 'FO_lost': i[9]})
            # Ambri 
            case 101152:
                if i[1] in hc_ambri_piotta_data.keys():
                  hc_ambri_piotta_data[i[1]].update({'FO_total': i[5], 'FO_won': i[8], 'FO_lost': i[9]})
            # Davos
            case 101151:
                if i[1] in hc_davos_data.keys():
                  hc_davos_data[i[1]].update({'FO_total': i[5], 'FO_won': i[8], 'FO_lost': i[9]})
            # Lugano
            case 101150:
                if i[1] in hc_lugano_data.keys():
                  hc_lugano_data[i[1]].update({'FO_total': i[5], 'FO_won': i[8], 'FO_lost': i[9]})
            # Lausanne
            case 103141: 
                if i[1] in hc_lausanne_data.keys():
                  hc_lausanne_data[i[1]].update({'FO_total': i[5], 'FO_won': i[8], 'FO_lost': i[9]})
            # Bern
            case 102126:
                if i[1] in sc_bern_data.keys():
                  sc_bern_data[i[1]].update({'FO_total': i[5], 'FO_won': i[8], 'FO_lost': i[9]})
            # Tigers
            case 102127:
                if i[1] in scl_tigers_data.keys():
                  scl_tigers_data[i[1]].update({'FO_total': i[5], 'FO_won': i[8], 'FO_lost': i[9]})
            # Rappi 
            case 101060:
                if i[1] in rapperswil_jona_lakers_data.keys():
                  rapperswil_jona_lakers_data[i[1]].update({'FO_total': i[5], 'FO_won': i[8], 'FO_lost': i[9]})
            # ZSC
            case 101139:
                if i[1] in zsc_lions_data.keys():
                  zsc_lions_data[i[1]].update({'FO_total': i[5], 'FO_won': i[8], 'FO_lost': i[9]})

    # and now the goalie data
    # different as above we go through both lists (summary and shots agains data) at the same time in one 
    # loop. To achieve this we cannot use a simple AND operation, but have to use the zip function, which
    # combines the two lists (it is an intelligent merge) but only works if both lists have the same sorting (or 
    # else we'll have garbage data). Apparently the output is not a list but a tuple and we access the data already
    # in the loop definition 
    # this method is chosen, as we have few goal keepers and some data of both lists we actually need (and also to 
    # show once more the different options we have)
    for i, j in zip(goalie_data['data'], goalie_shot_data['data']):
        team_id = i[2]['id']
        match team_id:
            # Biel
            case 102128:
                ehc_biel_data[i[1]] = add_goalie_data(i, j)
            # Kloten
            case 101149:
                ehc_kloten_data[i[1]] = add_goalie_data(i, j)
            # Zug
            case 101144:
                ev_zug_data[i[1]] = add_goalie_data(i, j)
            # Fribourg
            case 103138:
                fribourg_gotteron_data[i[1]] = add_goalie_data(i, j)
            # Genf
            case 103140:
                genf_servette_data[i[1]] = add_goalie_data(i, j)
            # Ajoie
            case 103144:
                hc_ajoie_data[i[1]] = add_goalie_data(i, j)
            # Ambri
            case 101152:
                hc_ambri_piotta_data[i[1]] = add_goalie_data(i, j)
            # Davos
            case 101151:
                hc_davos_data[i[1]] = add_goalie_data(i, j)
            # Lugano
            case 101150:
                hc_lugano_data[i[1]] = add_goalie_data(i, j)
            # Lausanne
            case 103141:
                hc_lausanne_data[i[1]] = add_goalie_data(i, j)
            # Bern
            case 102126:
                sc_bern_data[i[1]] = add_goalie_data(i, j)
            # Tigers
            case 102127:
                scl_tigers_data[i[1]] = add_goalie_data(i, j)
            # Rappi
            case 101060:
                rapperswil_jona_lakers_data[i[1]] = add_goalie_data(i, j)
            # ZSC
            case 101139:
                zsc_lions_data[i[1]] = add_goalie_data(i, j)

    # return the team dicts 
    return (ehc_biel_data, ehc_kloten_data, ev_zug_data, fribourg_gotteron_data, genf_servette_data,
            hc_ajoie_data, hc_ambri_piotta_data, hc_davos_data, hc_lugano_data, hc_lausanne_data,
            sc_bern_data, scl_tigers_data, rapperswil_jona_lakers_data, zsc_lions_data)
#
# get the upcoming games and all past games of the season
#
def get_game_info ():
    request_url = (BASE_URL + RESULT_CACHE + RESULT_PAGE + NEXT_GAME + GAME_QUERY
                   + GAME_FILTER + GAME_ORDER + STANDARD_ENDING + STANDARD_INFO)
    next_games = send_request(request_url)
    request_url = (BASE_URL + STAT_CACHE + RESULT_PAGE + ALL_GAMES_QUERY + FILTER_QUERY_PAST_GAMES 
                   + CURRENT_SEASON + ALL_TEAMS + SEASON_START_TIL_TODAY + FILTER_BY_PAST_GAMES
                   + ORDER_BY_PAST_GAMES + STANDARD_ENDING + STANDARD_INFO)
    past_games = send_request(request_url)
    return (next_games, past_games)
#
# create a dictionary for 1 of the past 5 games based on the 
# received game and the boolean which states if the team was playing
# at home or away. We also add this to the game information, as this 
# info will be passed to the output. You might wonder why a dictionary?
# Well i like dictionaries, as they provide a sense of order and control
# where as I can put anything into a list, in a dictionary I have some
# structure, which increases the readability for me at least.
#
def create_game(i: list, home: bool) -> dict:
    
    if home is True: 
        h = 'H'
     
        game = { 'Date': i[1]
               , 'Oponent': i[4]['name']
               , 'Home_away': h
               , 'Goald_scored': i[5]['homeTeam']
               , 'Goals_received': i[5]['awayTeam']}
    else:
        h = 'A'
        
        game = {  'Date': i[1]
                , 'Oponent': i[3]['name']
                , 'Home_away': h
                , 'Goald_scored': i[5]['awayTeam']
                , 'Goals_received': i[5]['homeTeam']}
    
    return game
#
#
#
def create_player_sheet(my_workbook: xlsxwriter.Workbook, my_worksheet: xlsxwriter.worksheet,home_team: team, away_team: team):
    start_row_title = 1
    start_column = 0
    my_column_pos = write_titels_to_worksheet_players(my_worksheet,start_row_title,start_column)

    home_format = my_workbook.add_format(TITLE_HOME_FORMAT)
    my_worksheet.merge_range(0,start_column,0,my_column_pos,'Stürmer',home_format)

    start_column = (my_column_pos + 2)
    my_column_pos = write_titels_to_worksheet_players(my_worksheet,start_row_title,start_column)

    home_format = my_workbook.add_format(TITLE_HOME_FORMAT)
    my_worksheet.merge_range(0,start_column,0,my_column_pos,'Verteidiger',home_format)
    

    count_forward = 1
    count_defense = 1
    for i in home_team.player_data: 

        match (home_team.player_data[i]['Position']):
            case 'Stürmer':
                start_write_col = 0
                count_forward += 1
                start_row = count_forward
            case 'Verteidiger':
                start_write_col = start_column
                count_defense += 1
                start_row = count_defense
            # the python equivalent to WHEN OTHER
            case _:
                # in contrast to the expectation the continue statement actually
                # ensures that loop skips the current iteration and goes to the next one
                continue
        try:
            my_worksheet.write(start_row,
                           start_write_col,
                           home_team.player_data[i]['Name'])
        except: 
           pass
        start_write_col += 1

        try:
            my_worksheet.write(start_row,
                           start_write_col,
                           home_team.player_data[i]['Games_played'])
        except: 
           pass
        start_write_col += 1

        try:
            my_worksheet.write(start_row,
                           start_write_col,
                           home_team.player_data[i]['Goals'])
        except: 
           pass
        start_write_col += 1

        try:
            my_worksheet.write(start_row,
                           start_write_col,
                           home_team.player_data[i]['Assists'])
        except: 
           pass
        start_write_col += 1

        try:
           my_worksheet.write(start_row,
                           start_write_col,
                           home_team.player_data[i]['Shots_total'])
        except: 
           pass
        start_write_col += 1

        try: 
            my_worksheet.write(start_row,
                           start_write_col,
                           home_team.player_data[i]['Shots_from_slot'])
        except:
           pass
        start_write_col += 1

        try: 
            my_worksheet.write(start_row,
                           start_write_col,
                           home_team.player_data[i]['Shots_missed'])
        except:
           pass
        start_write_col += 1
        try:
            my_worksheet.write(start_row,
                           start_write_col,
                           home_team.player_data[i]['Shots_on_border'])
        except:
            pass
        start_write_col += 1

        try:
            my_worksheet.write(start_row,
                           start_write_col,
                           home_team.player_data[i]['+/-'])
        except:
            pass
        start_write_col += 1

        try:
            my_worksheet.write(start_row,
                           start_write_col,
                           home_team.player_data[i]['FO_total'])
        except:
            pass
        start_write_col += 1

        try:
            my_worksheet.write(start_row,
                           start_write_col,
                           home_team.player_data[i]['FO_won'])
        except:
            pass
        start_write_col += 1

        try:
            my_worksheet.write(start_row,
                           start_write_col,
                           home_team.player_data[i]['FO_lost'])
        except:
           pass
        start_write_col += 1

        try:
            my_worksheet.write(start_row,
                           start_write_col,
                           home_team.player_data[i]['Blocked_shots'])
        except:
           pass
        start_write_col += 1

        try:
            my_worksheet.write(start_row,
                           start_write_col,
                           home_team.player_data[i]['Box_Play_Goals'])
        except:
            pass
        start_write_col += 1

        try:
            my_worksheet.write(start_row,
                           start_write_col,
                           home_team.player_data[i]['Box_Play_Assists'])
        except: 
           pass
        start_write_col += 1

        try:
            my_worksheet.write(start_row,
                           start_write_col,
                           home_team.player_data[i]['PP_Goals'])
        except: 
           pass
        start_write_col += 1

        try:
            my_worksheet.write(start_row,
                           start_write_col,
                           home_team.player_data[i]['PP_Assists'])
        except: 
           pass
        start_write_col += 1

        try:
            my_worksheet.write(start_row,
                           start_write_col,
                           home_team.player_data[i]['PIM'])
        except: 
           pass

    my_worksheet.add_table(1,0,count_forward,17,{'style': 'Table Style Light 17',
                                             'columns': [{'header': 'Player'},
                                                         {'header': 'GP'},
                                                         {'header': 'Goals'},
                                                         {'header': 'Assists'},
                                                         {'header': 'Shots on Goal'},
                                                         {'header': 'Schüsse aufs Tor aus Slot'},
                                                         {'header': 'Schüsse nebes Tor'},
                                                         {'header': 'Schüsse an die Torumrandung'},
                                                         {'header': '+/-'},
                                                         {'header': 'Bullys Total'},
                                                         {'header': 'Gewonnene Bullys'},
                                                         {'header': 'Verlorene Bullys'},
                                                         {'header': 'Blocked Shots'},
                                                         {'header': 'Shorthanders'},
                                                         {'header': 'Shorthander Assists'},
                                                         {'header': 'Powerplay Goals'},
                                                         {'header': 'Powerplay Assists'},
                                                         {'header': 'Penalty Minutes'},
                                                         ]})
    my_worksheet.add_table(1,start_column,count_defense,36,{'style': 'Table Style Light 17',
                                             'columns': [{'header': 'Player'},
                                                         {'header': 'GP'},
                                                         {'header': 'Goals'},
                                                         {'header': 'Assists'},
                                                         {'header': 'Shots on Goal'},
                                                         {'header': 'Schüsse aufs Tor aus Slot'},
                                                         {'header': 'Schüsse nebes Tor'},
                                                         {'header': 'Schüsse an die Torumrandung'},
                                                         {'header': '+/-'},
                                                         {'header': 'Bullys Total'},
                                                         {'header': 'Gewonnene Bullys'},
                                                         {'header': 'Verlorene Bullys'},
                                                         {'header': 'Blocked Shots'},
                                                         {'header': 'Shorthanders'},
                                                         {'header': 'Shorthander Assists'},
                                                         {'header': 'Powerplay Goals'},
                                                         {'header': 'Powerplay Assists'},
                                                         {'header': 'Penalty Minutes'},
                                                         ]})
    
    
    my_worksheet.autofit()
# 
# the actual main function
#
def main() -> None:
    my_excelfile = create_excel_file()
    (ehc_biel_data, ehc_kloten_data, ev_zug_data, fribourg_gotteron_data, genf_servette_data,
     hc_ajoie_data, hc_ambri_piotta_data, hc_davos_data, hc_lugano_data, hc_lausanne_data,
     sc_bern_data, scl_tigers_data, rapperswil_jona_lakers_data, zsc_lions_data) = create_team_data()
    # here we create the objects of the class team with handing over the Team ID (currently hard coded but we theoretically 
    # can also use the request response), the proper Team name and the player data
    # the team name is a bit silly on how we get that. we access the dictionary of the player data with translating
    # the dictionary to a list, which leaves us with a bunch of player names in a list (the sub dictionaries 
    # are not taken with in this transformation) and we take the very first entry to give the dictionary the key
    # or said differently - the dictionary tells itself what the key is
    # and yes we could allocate the objects dynamically but i am too lazy for this atm
    biel = team(102128,ehc_biel_data[list(ehc_biel_data)[0]]['Team'],ehc_biel_data)
    kloten = team(101149,ehc_kloten_data[list(ehc_kloten_data)[0]]['Team'],ehc_kloten_data)
    zug = team(101144,ev_zug_data[list(ev_zug_data)[0]]['Team'],ev_zug_data)
    fribourg = team(103138,fribourg_gotteron_data[list(fribourg_gotteron_data)[0]]['Team'],fribourg_gotteron_data)
    genf = team(103140,genf_servette_data[list(genf_servette_data)[0]]['Team'],genf_servette_data)
    ajoie = team(103144,hc_ajoie_data[list(hc_ajoie_data)[0]]['Team'],hc_ajoie_data)
    ambri = team(101152,hc_ambri_piotta_data[list(hc_ambri_piotta_data)[0]]['Team'],hc_ambri_piotta_data)
    davos = team(101151,hc_davos_data[list(hc_davos_data)[0]]['Team'],hc_davos_data)
    lugano = team(101150,hc_lugano_data[list(hc_lugano_data)[0]]['Team'],hc_lugano_data)
    lausanne = team(103141,hc_lausanne_data[list(hc_lausanne_data)[0]]['Team'],hc_lausanne_data)
    bern = team(102126,sc_bern_data[list(sc_bern_data)[0]]['Team'],sc_bern_data)
    langnau = team(102127,scl_tigers_data[list(scl_tigers_data)[0]]['Team'],scl_tigers_data)
    rapperswil = team(101060,rapperswil_jona_lakers_data[list(rapperswil_jona_lakers_data)[0]]['Team'],rapperswil_jona_lakers_data)
    zsc = team(101139,zsc_lions_data[list(zsc_lions_data)[0]]['Team'],zsc_lions_data)

    (next_games, past_games) = get_game_info()
    lists_to_fill = True
    # we create the dictionary with all teams where the id is the teamID 
    # in order to iterate over the teams / not to have to create the massive 
    # switches we have for the player data and address the objects directly without 
    # knowing directly from the code which object we are addressing at the moment 
    # we could have done that for the teams as well but then i would need to pass
    # the dictionary and to refactor everything for which i am too lazy atm
    team_dict = {biel.team_id: biel
                ,kloten.team_id: kloten
                ,zug.team_id: zug 
                ,fribourg.team_id: fribourg
                ,genf.team_id: genf
                ,ajoie.team_id: ajoie
                ,ambri.team_id: ambri
                ,davos.team_id: davos
                ,lugano.team_id: lugano
                ,lausanne.team_id: lausanne
                ,bern.team_id: bern
                ,langnau.team_id:langnau
                ,rapperswil.team_id: rapperswil
                ,zsc.team_id: zsc}
    
    # if all are full, no need to go further through the list of the past games
    while lists_to_fill:
        for i in (past_games['data']):
            # only if we still have space we actually add stuff to the list of past games
            if len(team_dict[i[3]['id']].past_5_games) < 5:
                # do the entry for the home team
                team_dict[i[3]['id']].past_5_games.append(create_game(i,True))
            if len(team_dict[i[4]['id']].past_5_games) < 5:
                # do the entry for the away team
                team_dict[i[4]['id']].past_5_games.append(create_game(i,False))
            # Geez this IF was killing me on figuring out how to do that
            # the all comand is basically a for iteration as we use it to iterate over the lsit items (for i in data['data'])
            # we then create for each entry in the dictionary teams the variable j of the type team for which we check the 
            # length of the list past_5_games where each entry in the list is actually a dictionary with the game data added above
            if all(len(j.past_5_games) >= 5 for j in team_dict.values()):
                # set the condition th break the while loop
                lists_to_fill = False
                # we need the break statement in order to escape the 'for loop' to come to the 'while loop'
                break
            
    for i in (next_games['data']):
        tab_name = ''
        tab_name = (team_dict[i[4]['id']].name) + "_" + (team_dict[i[5]['id']].name) + "_PD"
        my_worksheet = add_worksheet(my_excelfile, tab_name)
        create_player_sheet(my_excelfile,my_worksheet,team_dict[i[4]['id']],team_dict[i[5]['id']])
        break
    my_excelfile.close()
main()