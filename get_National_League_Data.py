##########################################################################
# the imports
##########################################################################
import requests
import json
import re 
import xlsxwriter
from xlsxwriter.utility import cell_autofit_width
import datetime 
import statistics
##########################################################################
# the objects
##########################################################################
class team:
    def __init__(self, team_id, name, player_data):
        self.team_id = team_id 
        self.name = name
        self.player_data = player_data
        self.past_5_games = []
        self.amount_def_min_5 = None 
        self.amount_forward_min_5 = None 
        self.amount_import_min_5 = None
        # data gets taken form seperate stat as it can be that players get
        # traded and the data situation is not the best in such cases
        # None is basically an Initialize of the container to tell that yes 
        # there will be a variable 
        self.amount_of_games = None 
        self.goals_made = None 
        self.goals_against = None 
        self.goals_made_forward = None 
        self.goals_made_defense = None 
        self.goals_made_imports = None 
        self.sog = None 
        self.sog_slot = None
        self.sob = None 
        self.shots_missed = None
        self.blocked_shots = None
        self.sog_received = None 
        self.sog_slot_received = None
        self.pp_op = None 
        self.time_in_pp = None
        self.pp_goals = None 
        self.pp_received_goals = None 
        self.pk_situations = None 
        self.pk_goals_received = None
        self.time_in_pk = None 
        self.shorthanders = None

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
TEAM_GOAL_OVERVIEW = 'alias=teamGoal'
TEAM_POSITION_GOALS = 'alias=teamGoalPosNation'
TEAM_SHOT_DETAILS = 'alias=teamShotDetail'
TEAM_PP_DATA = 'alias=teamPp'
TEAM_PK_DATA = 'alias=teamPk'
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
# sorting 
SORT_BY_TEAM = '&orderBy=team&orderByDescending=false'
# more standard stuff
FILTERING_INFO = '&filterBy=Season,Phase,Team,Position,Licence'
TEAM_STAT_FILTER_INFO = '&filterBy=Season,Phase'
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
TEAM_NAME_HOME_FORMAT = {'bold': True
                     ,'align': 'center'
                     ,'font_color': 'White'
                     ,'font_size': 17
                     ,'bg_color': "#943C01"}
TITLE_AWAY_FORMAT = {'bold': True
                     ,'align': 'center'
                     ,'font_color': 'White'
                     ,'font_size': 14
                     ,'bg_color': '#366092'}
TEAM_NAME_AWAY_FORMAT = {'bold': True
                        ,'align': 'center'
                        ,'font_color': 'White'
                        ,'font_size': 17
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
                , 'Games_played': i[3]
                , 'Position': 'Torhüter'
                , 'Goals_against': i[4]
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
    start_row_title = 2
    start_column = 0
    my_column_pos = write_titels_to_worksheet_players(my_worksheet,start_row_title,start_column)

    home_format = my_workbook.add_format(TITLE_HOME_FORMAT)
    my_worksheet.merge_range((start_row_title - 1),start_column,(start_row_title - 1),my_column_pos,'Stürmer',home_format)

    start_column = (my_column_pos + 2)
    my_column_pos = write_titels_to_worksheet_players(my_worksheet,start_row_title,start_column)

    my_worksheet.merge_range((start_row_title - 1),start_column,(start_row_title - 1),my_column_pos,'Verteidiger',home_format)
    
    goalie_start_col = (my_column_pos + 2)
    my_worksheet.merge_range((start_row_title - 1), goalie_start_col, (start_row_title - 1), (goalie_start_col + 11), 'Goalie', home_format)

    team_name_format = my_workbook.add_format(TEAM_NAME_HOME_FORMAT)
    my_worksheet.merge_range((start_row_title - 2),0,(start_row_title - 2),(goalie_start_col + 11),home_team.name,team_name_format)


    count_forward = start_row_title
    count_defense = start_row_title
    count_goalie = start_row_title
    # forward stat data definition
    average_gp_forward = 0
    median_gp_forward = []
    average_goals_forward = 0
    median_goals_forward = []
    average_assists_forward = 0
    median_assists_forward = []
    average_sog_forward = 0
    median_sog_forward = []
    average_slot_sog_forward = 0
    median_shot_missed_forward = []
    average_shot_missed_forward = 0
    median_sob_forward = []
    average_sob_forward = 0
    median_slot_sog_forward = []
    average_bilanz_forward = 0
    median_bilanz_forward = []
    average_bullies_forward = 0
    median_bullies_forward = []
    average_bullies_won_forward = 0
    median_bullies_won_forward = []
    average_bullies_lost_forward = 0
    median_bullies_lost_forward = []
    average_blocked_shots_forward = 0
    median_blocked_shots_forward = []
    average_bpa_forward = 0
    median_bpa_forward = []
    average_shs_forward = 0
    median_shs_forward = []
    average_ppg_forward = 0
    median_ppg_forward = []
    average_ppa_forward = 0
    median_ppa_forward = []
    average_pim_forward = 0
    median_pim_forward = []
    # defense stat data definition
    average_gp_defense = 0
    median_gp_defense = []
    average_goals_defense = 0
    median_goals_defense = []
    average_assists_defense = 0
    median_assists_defense = []
    average_sog_defense = 0
    median_sog_defense = []
    average_slot_sog_defense = 0
    median_shot_missed_defense = []
    average_shot_missed_defense = 0
    median_sob_defense = []
    average_sob_defense = 0
    median_slot_sog_defense = []
    average_bilanz_defense = 0
    median_bilanz_defense = []
    average_bullies_defense = 0
    median_bullies_defense = []
    average_bullies_won_defense = 0
    median_bullies_won_defense = []
    average_bullies_lost_defense = 0
    median_bullies_lost_defense = []
    average_blocked_shots_defense = 0
    median_blocked_shots_defense = []
    average_bpa_defense = 0
    median_bpa_defense = []
    average_shs_defense = 0
    median_shs_defense = []
    average_ppg_defense = 0
    median_ppg_defense = []
    average_ppa_defense = 0
    median_ppa_defense = []
    average_pim_defense = 0
    median_pim_defense = []
    # goalie stat data definition 
    total_gp = 0
    total_pim = 0
    total_received_goals = 0
    total_saves_overall = 0
    total_received_shots = 0
    total_save_rate_overall = 0
    total_shots_from_slot = 0
    total_saves_from_slot_overall = 0
    total_save_rate_slot = 0
    total_missed_shots = 0
    total_shots_on_frame = 0

    forward_player_average = 0
    defense_player_average = 0
    goalie_player_count = 0

    for i in home_team.player_data: 

        match (home_team.player_data[i]['Position']):
            case 'Stürmer':
                start_write_col = 0
                count_forward += 1
                start_row = count_forward
                try:
                   median_gp_forward.append(int(home_team.player_data[i]['Games_played']))
                except: 
                   median_gp_forward.append(int(0))
                try:
                   median_goals_forward.append(int(home_team.player_data[i]['Goals']))
                except: 
                   median_goals_forward.append(int(0))
                try:
                   median_assists_forward.append(int(home_team.player_data[i]['Assists']))
                except: 
                    median_assists_forward.append(int(0))      
                try:
                    median_sog_forward.append(int(home_team.player_data[i]['Shots_total']))
                except: 
                   median_sog_forward.append(int(0))
                try:
                    median_slot_sog_forward.append(int(home_team.player_data[i]['Shots_from_slot']))
                except:
                   median_slot_sog_forward.append(int(0))
                try:
                    median_shot_missed_forward.append(int(home_team.player_data[i]['Shots_missed']))
                except:
                   median_shot_missed_forward.append(int(0))
                try:
                    median_sob_forward.append(int(home_team.player_data[i]['Shots_on_border']))
                except:
                   median_sob_forward.append(int(0))
                try:
                    median_bilanz_forward.append(int(home_team.player_data[i]['+/-']))
                except:
                   median_bilanz_forward.append(int(0))
                try:
                    median_bullies_forward.append(int(home_team.player_data[i]['FO_total']))
                except:
                   median_bullies_forward.append(int(0))   
                try:
                    median_bullies_won_forward.append(int(home_team.player_data[i]['FO_won']))
                except:
                   median_bullies_won_forward.append(int(0))
                try:
                    median_bullies_lost_forward.append(int(home_team.player_data[i]['FO_lost']))
                except:
                   median_bullies_lost_forward.append(int(0))
                try:
                    median_blocked_shots_forward.append(int(home_team.player_data[i]['Blocked_shots']))
                except:
                   median_blocked_shots_forward.append(int(0))
                try:
                    median_shs_forward.append(int(home_team.player_data[i]['Box_Play_Goals']))
                except:
                   median_shs_forward.append(int(0))
                try:
                    median_bpa_forward.append(int(home_team.player_data[i]['Box_Play_Assists']))
                except:
                   median_bpa_forward.append(int(0))
                try:
                    median_ppg_forward.append(int(home_team.player_data[i]['PP_Goals']))
                except:
                   median_ppg_forward.append(int(0))
                try:
                    median_ppa_forward.append(int(home_team.player_data[i]['PP_Assists']))
                except:
                   median_ppa_forward.append(int(0))    
                try:
                    median_pim_forward.append(int(home_team.player_data[i]['PIM']))
                except:
                   median_pim_forward.append(int(0))
                if int(home_team.player_data[i]['Games_played']) >= 5:
                    forward_player_average += 1
                    try:
                       temp = float(home_team.player_data[i]['Games_played'])
                    except:
                       temp = 0
                    res = average_gp_forward + temp
                    average_gp_forward = res
                    try:
                       temp = float(home_team.player_data[i]['Goals'])
                    except:
                       temp = 0
                    res = average_goals_forward + temp
                    average_goals_forward = res
                    try:
                       temp = float(home_team.player_data[i]['Assists'])
                    except:
                       temp = 0
                    res = average_assists_forward + temp
                    average_assists_forward = res
                    try:
                       temp = float(home_team.player_data[i]['Shots_total'])
                    except:
                       temp = 0
                    res = temp + average_sog_forward
                    average_sog_forward = res
                    try:
                       temp = float(home_team.player_data[i]['Shots_from_slot'])
                    except:
                       temp = 0
                    res = temp + average_slot_sog_forward
                    average_slot_sog_forward = res
                    try:
                       temp = float(home_team.player_data[i]['Shots_missed'])
                    except:
                       temp = 0
                    res = temp + average_shot_missed_forward
                    average_shot_missed_forward = res
                    try:
                       temp = float(home_team.player_data[i]['Shots_on_border'])
                    except:
                       temp = 0
                    res = temp + average_sob_forward
                    average_sob_forward = res
                    try:
                       temp = float(home_team.player_data[i]['+/-'])
                    except:
                       temp = 0
                    res = temp + average_bilanz_forward
                    average_bilanz_forward = res
                    try:
                       temp = float(home_team.player_data[i]['FO_total'])
                    except:
                       temp = 0
                    res = temp + average_bullies_forward
                    average_bullies_forward = res
                    try:
                       temp = float(home_team.player_data[i]['FO_won'])
                    except:
                       temp = 0
                    res = temp + average_bullies_won_forward
                    average_bullies_won_forward = res
                    try:
                       temp = float(home_team.player_data[i]['FO_lost'])
                    except:
                       temp = 0
                    res = temp + average_bullies_lost_forward
                    average_bullies_lost_forward = res
                    try:
                       temp = float(home_team.player_data[i]['Blocked_shots'])
                    except:
                       temp = 0
                    res = temp + average_blocked_shots_forward
                    average_blocked_shots_forward = res
                    try:
                       temp = float(home_team.player_data[i]['Box_Play_Goals'])
                    except:
                       temp = 0
                    res = temp + average_shs_forward
                    average_shs_forward = res 
                    try:
                       temp = float(home_team.player_data[i]['Box_Play_Assists'])
                    except:
                       temp = 0
                    res = temp + average_bpa_forward
                    average_bpa_forward = res 
                    try:
                       temp = float(home_team.player_data[i]['PP_Goals'])
                    except:
                       temp = 0
                    res = temp + average_ppg_forward
                    average_ppg_forward = res 
                    try:
                       temp = float(home_team.player_data[i]['PP_Assists'])
                    except:
                       temp = 0
                    res = temp + average_ppa_forward
                    average_ppa_forward = res
                    try:
                       temp = float(home_team.player_data[i]['PIM'])
                    except:
                       temp = 0
                    res = temp + average_pim_forward
                    average_pim_forward = res 
            case 'Verteidiger':
                start_write_col = start_column
                count_defense += 1
                start_row = count_defense
                try:
                   median_gp_defense.append(int(home_team.player_data[i]['Games_played']))
                except: 
                   median_gp_defense.append(int(0))
                try:
                   median_goals_defense.append(int(home_team.player_data[i]['Goals']))
                except: 
                   median_goals_defense.append(int(0))
                try:
                   median_assists_defense.append(int(home_team.player_data[i]['Assists']))
                except: 
                    median_assists_defense.append(int(0))      
                try:
                    median_sog_defense.append(int(home_team.player_data[i]['Shots_total']))
                except: 
                   median_sog_defense.append(int(0))
                try:
                    median_slot_sog_defense.append(int(home_team.player_data[i]['Shots_from_slot']))
                except:
                   median_slot_sog_defense.append(int(0))
                try:
                    median_shot_missed_defense.append(int(home_team.player_data[i]['Shots_missed']))
                except:
                   median_shot_missed_defense.append(int(0))
                try:
                    median_sob_defense.append(int(home_team.player_data[i]['Shots_on_border']))
                except:
                   median_sob_defense.append(int(0))
                try:
                    median_bilanz_defense.append(int(home_team.player_data[i]['+/-']))
                except:
                   median_bilanz_defense.append(int(0))
                try:
                    median_bullies_defense.append(int(home_team.player_data[i]['FO_total']))
                except:
                   median_bullies_defense.append(int(0))   
                try:
                    median_bullies_won_defense.append(int(home_team.player_data[i]['FO_won']))
                except:
                   median_bullies_won_defense.append(int(0))
                try:
                    median_bullies_lost_defense.append(int(home_team.player_data[i]['FO_lost']))
                except:
                   median_bullies_lost_defense.append(int(0))
                try:
                    median_blocked_shots_defense.append(int(home_team.player_data[i]['Blocked_shots']))
                except:
                   median_blocked_shots_defense.append(int(0))
                try:
                    median_shs_defense.append(int(home_team.player_data[i]['Box_Play_Goals']))
                except:
                   median_shs_defense.append(int(0))
                try:
                    median_bpa_defense.append(int(home_team.player_data[i]['Box_Play_Assists']))
                except:
                   median_bpa_defense.append(int(0))
                try:
                    median_ppg_defense.append(int(home_team.player_data[i]['PP_Goals']))
                except:
                   median_ppg_defense.append(int(0))
                try:
                    median_ppa_defense.append(int(home_team.player_data[i]['PP_Assists']))
                except:
                   median_ppa_defense.append(int(0))    
                try:
                    median_pim_defense.append(int(home_team.player_data[i]['PIM']))
                except:
                   median_pim_defense.append(int(0))

                if int(home_team.player_data[i]['Games_played']) >= 5:
                    defense_player_average += 1
                    try:
                       temp = float(home_team.player_data[i]['Games_played'])
                    except:
                       temp = 0
                    res = average_gp_defense + temp
                    average_gp_defense = res
                    try:
                       temp = float(home_team.player_data[i]['Goals'])
                    except:
                       temp = 0
                    res = average_goals_defense + temp
                    average_goals_defense = res
                    try:
                       temp = float(home_team.player_data[i]['Assists'])
                    except:
                       temp = 0
                    res = average_assists_defense + temp
                    average_assists_defense = res
                    try:
                       temp = float(home_team.player_data[i]['Shots_total'])
                    except:
                       temp = 0
                    res = temp + average_sog_defense
                    average_sog_defense = res
                    try:
                       temp = float(home_team.player_data[i]['Shots_from_slot'])
                    except:
                       temp = 0
                    res = temp + average_slot_sog_defense
                    average_slot_sog_defense = res
                    try:
                       temp = float(home_team.player_data[i]['Shots_missed'])
                    except:
                       temp = 0
                    res = temp + average_shot_missed_defense
                    average_shot_missed_defense = res
                    try:
                       temp = float(home_team.player_data[i]['Shots_on_border'])
                    except:
                       temp = 0
                    res = temp + average_sob_defense
                    average_sob_defense = res
                    try:
                       temp = float(home_team.player_data[i]['+/-'])
                    except:
                       temp = 0
                    res = temp + average_bilanz_defense
                    average_bilanz_defense = res
                    try:
                       temp = float(home_team.player_data[i]['FO_total'])
                    except:
                       temp = 0
                    res = temp + average_bullies_defense
                    average_bullies_defense = res
                    try:
                       temp = float(home_team.player_data[i]['FO_won'])
                    except:
                       temp = 0
                    res = temp + average_bullies_won_defense
                    average_bullies_won_defense = res
                    try:
                       temp = float(home_team.player_data[i]['FO_lost'])
                    except:
                       temp = 0
                    res = temp + average_bullies_lost_defense
                    average_bullies_lost_defense = res
                    try:
                       temp = float(home_team.player_data[i]['Blocked_shots'])
                    except:
                       temp = 0
                    res = temp + average_blocked_shots_defense
                    average_blocked_shots_defense = res
                    try:
                       temp = float(home_team.player_data[i]['Box_Play_Goals'])
                    except:
                       temp = 0
                    res = temp + average_shs_defense
                    average_shs_defense = res 
                    try:
                       temp = float(home_team.player_data[i]['Box_Play_Assists'])
                    except:
                       temp = 0
                    res = temp + average_bpa_defense
                    average_bpa_defense = res 
                    try:
                       temp = float(home_team.player_data[i]['PP_Goals'])
                    except:
                       temp = 0
                    res = temp + average_ppg_defense
                    average_ppg_defense = res 
                    try:
                       temp = float(home_team.player_data[i]['PP_Assists'])
                    except:
                       temp = 0
                    res = temp + average_ppa_defense
                    average_ppa_defense = res
                    try:
                       temp = float(home_team.player_data[i]['PIM'])
                    except:
                       temp = 0
                    res = temp + average_pim_defense
                    average_pim_defense = res 
            case 'Torhüter':
              goalie_player_count += 1
              start_write_col = goalie_start_col
              count_goalie += 1
              start_row = count_goalie

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
                temp_gp = total_gp
                total_gp = temp_gp + int(home_team.player_data[i]['Games_played'])
              except: 
                pass
              start_write_col += 1
              try:
                my_worksheet.write(start_row,
                           start_write_col,
                           home_team.player_data[i]['Goals_against'])
                temp_goals = total_received_goals
                total_received_goals = temp_goals + int(home_team.player_data[i]['Goals_against'])
              except: 
                pass
              start_write_col += 1
              try:
                my_worksheet.write(start_row,
                           start_write_col,
                           home_team.player_data[i]['Total_shots_agains'])
                temp_shots = total_received_shots
                total_received_shots = temp_shots + int(home_team.player_data[i]['Total_shots_agains'])
              except: 
                pass
              start_write_col += 1
              total_shots = int(home_team.player_data[i]['Total_shots_agains'])
              total_save_rate = float(home_team.player_data[i]['Total_save%'])
              temp_rate = total_save_rate_overall
              total_save_rate_overall = temp_rate + total_save_rate
              total_saves = round((total_shots/100) * total_save_rate)
              temp_saves = total_saves_overall
              total_saves_overall = temp_saves + total_saves
              my_worksheet.write(start_row,
                           start_write_col,
                           total_saves)
              start_write_col += 1
              try:
                my_worksheet.write(start_row,
                           start_write_col,
                           home_team.player_data[i]['Total_save%'])
              except: 
                pass
              start_write_col += 1
              try:
                my_worksheet.write(start_row,
                           start_write_col,
                           home_team.player_data[i]['Shots_from_slot'])
                temp_shots = total_shots_from_slot
                total_shots_from_slot = temp_shots + int(home_team.player_data[i]['Shots_from_slot'])
              except: 
                pass
              start_write_col += 1
              shots_from_slot = int(home_team.player_data[i]['Shots_from_slot'])
              slot_save_rate = float(home_team.player_data[i]['Slot_save%'])
              temp_rate = total_save_rate_slot
              total_save_rate_slot = temp_rate + slot_save_rate

              slot_saves = round((shots_from_slot/100) * slot_save_rate)
              temp_saves = total_saves_from_slot_overall
              total_saves_from_slot_overall = temp_saves + slot_saves
              my_worksheet.write(start_row,
                           start_write_col,
                           slot_saves)
              start_write_col += 1
              try:
                my_worksheet.write(start_row,
                           start_write_col,
                           home_team.player_data[i]['Slot_save%'])
              except: 
                pass
              start_write_col += 1
              try:
                my_worksheet.write(start_row,
                           start_write_col,
                           home_team.player_data[i]['Missed_shots'])
                temp_shots = total_missed_shots
                total_missed_shots = temp_shots + int(home_team.player_data[i]['Missed_shots'])
              except: 
                pass
              start_write_col += 1
              try:
                my_worksheet.write(start_row,
                           start_write_col,
                           home_team.player_data[i]['Shots_against_border'])
                temp_shots = total_shots_on_frame
                total_shots_on_frame = temp_shots + int(home_team.player_data[i]['Shots_against_border'])
              except: 
                pass
              start_write_col += 1
              try:
                my_worksheet.write(start_row,
                           start_write_col,
                           home_team.player_data[i]['PIM'])
                temp_pim = total_pim
                total_pim = temp_pim + int(home_team.player_data[i]['PIM'])
              except: 
                pass
              
              continue
              
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
    # write the average line
    # first for the forward
    start_write_col = 0
    count_forward += 1
    start_row = count_forward
    my_worksheet.write(start_row,
                           start_write_col,
                           (home_team.name + " Stürmer Durchschnitt (min 5 Spiele)"))
    start_write_col += 1
    average_gp_forward = average_gp_forward / forward_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_gp_forward)
    start_write_col += 1
    average_goals_forward = average_goals_forward / forward_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_goals_forward)
    start_write_col += 1
    average_assists_forward = average_assists_forward / forward_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_assists_forward)
    start_write_col += 1
    average_sog_forward = average_sog_forward / forward_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_sog_forward)
    start_write_col += 1

    average_slot_sog_forward = average_slot_sog_forward / forward_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_slot_sog_forward)
    start_write_col += 1

    average_shot_missed_forward = average_shot_missed_forward / forward_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_shot_missed_forward)
    start_write_col += 1
    average_sob_forward = average_sob_forward / forward_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_sob_forward)
    start_write_col += 1
    average_bilanz_forward = average_bilanz_forward / forward_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_bilanz_forward)
    start_write_col += 1
    average_bullies_forward = average_bullies_forward / forward_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_bullies_forward)
    start_write_col += 1
    average_bullies_won_forward = average_bullies_won_forward / forward_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_bullies_won_forward)
    start_write_col += 1
    average_bullies_lost_forward = average_bullies_lost_forward / forward_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_bullies_lost_forward)
    start_write_col += 1
    average_blocked_shots_forward = average_blocked_shots_forward / forward_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_blocked_shots_forward)
    start_write_col += 1
    average_shs_forward = average_shs_forward / forward_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_shs_forward)
    start_write_col += 1
    average_bpa_forward = average_bpa_forward / forward_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_bpa_forward)
    start_write_col += 1
    average_ppg_forward = average_ppg_forward / forward_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_ppg_forward)
    start_write_col += 1
    average_ppa_forward = average_ppa_forward / forward_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_ppa_forward)
    start_write_col += 1
    average_pim_forward = average_pim_forward / forward_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_pim_forward)

    start_write_col = 0
    count_forward += 1
    start_row = count_forward
    my_worksheet.write(start_row,
                           start_write_col,
                           (home_team.name + " Stürmer Median"))
    start_write_col += 1

    temp = statistics.median(median_gp_forward)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_goals_forward)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_assists_forward)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    average_sog_forward = average_sog_forward / forward_player_average
    temp = statistics.median(median_sog_forward)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_slot_sog_forward)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_shot_missed_forward)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_sob_forward)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_bilanz_forward)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_bullies_forward)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_bullies_won_forward)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_bullies_lost_forward)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_blocked_shots_forward)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_shs_forward)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_bpa_forward)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_ppg_forward)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_ppg_forward)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_pim_forward)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    # and now the defense
    start_write_col = start_column
    count_defense += 1
    start_row = count_defense
    my_worksheet.write(start_row,
                           start_write_col,
                           (home_team.name + " Verteidiger Durchschnitt (min 5 Spiele)"))
    start_write_col += 1
    average_gp_defense = average_gp_defense / defense_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_gp_defense)
    start_write_col += 1
    average_goals_defense = average_goals_defense / defense_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_goals_defense)
    start_write_col += 1
    average_assists_defense = average_assists_defense / defense_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_assists_defense)
    start_write_col += 1
    average_sog_defense = average_sog_defense / defense_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_sog_defense)
    start_write_col += 1

    average_slot_sog_defense = average_slot_sog_defense / defense_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_slot_sog_defense)
    start_write_col += 1

    average_shot_missed_defense = average_shot_missed_defense / defense_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_shot_missed_defense)
    start_write_col += 1
    average_sob_defense = average_sob_defense / defense_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_sob_defense)
    start_write_col += 1
    average_bilanz_defense = average_bilanz_defense / defense_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_bilanz_defense)
    start_write_col += 1
    average_bullies_defense = average_bullies_defense / defense_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_bullies_defense)
    start_write_col += 1
    average_bullies_won_defense = average_bullies_won_defense / defense_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_bullies_won_defense)
    start_write_col += 1
    average_bullies_lost_defense = average_bullies_lost_defense / defense_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_bullies_lost_defense)
    start_write_col += 1
    average_blocked_shots_defense = average_blocked_shots_defense / defense_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_blocked_shots_defense)
    start_write_col += 1
    average_shs_defense = average_shs_defense / defense_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_shs_defense)
    start_write_col += 1
    average_bpa_defense = average_bpa_defense / defense_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_bpa_defense)
    start_write_col += 1
    average_ppg_defense = average_ppg_defense / defense_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_ppg_defense)
    start_write_col += 1
    average_ppa_defense = average_ppa_defense / defense_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_ppa_defense)
    start_write_col += 1
    average_pim_defense = average_pim_defense / defense_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_pim_defense)

    start_write_col = start_column
    count_defense += 1
    start_row = count_defense
    my_worksheet.write(start_row,
                           start_write_col,
                           (home_team.name + " Verteidiger Median"))
    start_write_col += 1

    temp = statistics.median(median_gp_defense)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_goals_defense)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_assists_defense)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    average_sog_defense = average_sog_defense / defense_player_average
    temp = statistics.median(median_sog_defense)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_slot_sog_defense)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_shot_missed_defense)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_sob_defense)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_bilanz_defense)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_bullies_defense)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_bullies_won_defense)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_bullies_lost_defense)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_blocked_shots_defense)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_shs_defense)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_bpa_defense)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_ppg_defense)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_ppg_defense)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_pim_defense)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    # and now the goalie total 
    start_write_col = goalie_start_col
    count_goalie += 1
    start_row = count_goalie

    my_worksheet.write(start_row,
                           start_write_col,
                           'Total')
    start_write_col += 1
    my_worksheet.write(start_row,
                           start_write_col,
                           total_gp)
    start_write_col += 1
    my_worksheet.write(start_row,
                           start_write_col,
                           total_received_goals)
    start_write_col += 1
    my_worksheet.write(start_row,
                           start_write_col,
                           total_received_shots)
    start_write_col += 1
    my_worksheet.write(start_row,
                           start_write_col,
                           total_saves_overall)
    
    start_write_col += 1
    total_save_rate_overall = total_save_rate_overall / goalie_player_count
    my_worksheet.write(start_row,
                           start_write_col,
                           total_save_rate_overall)
    start_write_col += 1
    my_worksheet.write(start_row,
                           start_write_col,
                           total_shots_from_slot)
    start_write_col += 1
    my_worksheet.write(start_row,
                           start_write_col,
                           total_saves_from_slot_overall)
    start_write_col += 1
    total_save_rate_slot = total_save_rate_slot / goalie_player_count
    my_worksheet.write(start_row,
                           start_write_col,
                           total_save_rate_slot)
    start_write_col += 1
    my_worksheet.write(start_row,
                           start_write_col,
                           total_missed_shots)
    start_write_col += 1
    my_worksheet.write(start_row,
                           start_write_col,
                           total_shots_on_frame)
    start_write_col += 1
    my_worksheet.write(start_row,
                           start_write_col,
                           total_pim)
    my_worksheet.add_table(start_row_title,0,count_forward,17,{'style': 'Table Style Light 21',
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
    my_worksheet.add_table(start_row_title,start_column,count_defense,my_column_pos,{'style': 'Table Style Light 21',
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
    my_worksheet.add_table(start_row_title,goalie_start_col,count_goalie,(goalie_start_col + 11),{'style': 'Table Style Light 21',
                                             'columns': [{'header': 'Player'},
                                                         {'header': 'GP'},
                                                         {'header': 'Erhaltene Tore'},
                                                         {'header': 'Schüsse aufs Tor'},
                                                         {'header': 'Saves'},
                                                         {'header': 'Saves %'},
                                                         {'header': 'Schüsse aus dem Slot'},
                                                         {'header': 'Slot Saves'},
                                                         {'header': 'Slot Save %'},
                                                         {'header': 'Schüsse neben das Tor'},
                                                         {'header': 'Schüsse an die Torumrandung'},
                                                         {'header': 'Penalty Minutes'}
                                                         ]})
    # now the same for the guest team
    start_column = 0
    start_row_title = (max(count_defense, count_forward)) + 4
    my_column_pos = write_titels_to_worksheet_players(my_worksheet,start_row_title, start_column)

    away_format = my_workbook.add_format(TITLE_AWAY_FORMAT)
    my_worksheet.merge_range((start_row_title - 1),start_column,(start_row_title - 1),my_column_pos,'Stürmer',away_format)

    start_column = (my_column_pos + 2)
    my_column_pos = write_titels_to_worksheet_players(my_worksheet,start_row_title, start_column)

    my_worksheet.merge_range((start_row_title - 1),start_column,(start_row_title - 1),my_column_pos,'Verteidiger',away_format)

    goalie_start_col = (my_column_pos + 2)
    my_worksheet.merge_range((start_row_title - 1), goalie_start_col, (start_row_title - 1), (goalie_start_col + 11), 'Goalie', away_format)

    team_name_format = my_workbook.add_format(TEAM_NAME_AWAY_FORMAT)
    my_worksheet.merge_range((start_row_title - 2),0,(start_row_title - 2),(goalie_start_col + 11),away_team.name,team_name_format)

    count_forward = start_row_title
    count_defense = start_row_title
    count_goalie = start_row_title
    # reset the stat data definitions
        # forward stat data definition
    average_gp_forward = 0
    median_gp_forward.clear()
    average_goals_forward = 0
    median_goals_forward.clear()
    average_assists_forward = 0
    median_assists_forward.clear()
    average_sog_forward = 0
    median_sog_forward.clear()
    average_slot_sog_forward = 0
    median_shot_missed_forward.clear()
    average_shot_missed_forward = 0
    median_sob_forward.clear()
    average_sob_forward = 0
    median_slot_sog_forward.clear()
    average_bilanz_forward = 0
    median_bilanz_forward.clear()
    average_bullies_forward = 0
    median_bullies_forward.clear()
    average_bullies_won_forward = 0
    median_bullies_won_forward.clear()
    average_bullies_lost_forward = 0
    median_bullies_lost_forward.clear()
    average_blocked_shots_forward = 0
    median_blocked_shots_forward.clear()
    average_bpa_forward = 0
    median_bpa_forward.clear()
    average_shs_forward = 0
    median_shs_forward.clear()
    average_ppg_forward = 0
    median_ppg_forward.clear()
    average_ppa_forward = 0
    median_ppa_forward.clear()
    average_pim_forward = 0
    median_pim_forward.clear()
    # defense stat data definition
    average_gp_defense = 0
    median_gp_defense.clear()
    average_goals_defense = 0
    median_goals_defense.clear()
    average_assists_defense = 0
    median_assists_defense.clear()
    average_sog_defense = 0
    median_sog_defense.clear()
    average_slot_sog_defense = 0
    median_shot_missed_defense.clear()
    average_shot_missed_defense = 0
    median_sob_defense.clear()
    average_sob_defense = 0
    median_slot_sog_defense.clear()
    average_bilanz_defense = 0
    median_bilanz_defense.clear()
    average_bullies_defense = 0
    median_bullies_defense.clear()
    average_bullies_won_defense = 0
    median_bullies_won_defense.clear()
    average_bullies_lost_defense = 0
    median_bullies_lost_defense.clear()
    average_blocked_shots_defense = 0
    median_blocked_shots_defense.clear()
    average_bpa_defense = 0
    median_bpa_defense.clear()
    average_shs_defense = 0
    median_shs_defense.clear()
    average_ppg_defense = 0
    median_ppg_defense.clear()
    average_ppa_defense = 0
    median_ppa_defense.clear()
    average_pim_defense = 0
    median_pim_defense.clear()
    # goalie stat data definition 
    total_gp = 0
    total_pim = 0
    total_received_goals = 0
    total_saves_overall = 0
    total_received_shots = 0
    total_save_rate_overall = 0
    total_shots_from_slot = 0
    total_saves_from_slot_overall = 0
    total_save_rate_slot = 0
    total_missed_shots = 0
    total_shots_on_frame = 0

    goalie_player_count = 0
    forward_player_average = 0
    defense_player_average = 0

    for i in away_team.player_data: 

        match (away_team.player_data[i]['Position']):
            case 'Stürmer':
                start_write_col = 0
                count_forward += 1
                start_row = count_forward
                
                try:
                   median_gp_forward.append(int(away_team.player_data[i]['Games_played']))
                except: 
                   median_gp_forward.append(int(0))
                try:
                   median_goals_forward.append(int(away_team.player_data[i]['Goals']))
                except: 
                   median_goals_forward.append(int(0))
                try:
                   median_assists_forward.append(int(away_team.player_data[i]['Assists']))
                except: 
                    median_assists_forward.append(int(0))      
                try:
                    median_sog_forward.append(int(away_team.player_data[i]['Shots_total']))
                except: 
                   median_sog_forward.append(int(0))
                try:
                    median_slot_sog_forward.append(int(away_team.player_data[i]['Shots_from_slot']))
                except:
                   median_slot_sog_forward.append(int(0))
                try:
                    median_shot_missed_forward.append(int(away_team.player_data[i]['Shots_missed']))
                except:
                   median_shot_missed_forward.append(int(0))
                try:
                    median_sob_forward.append(int(away_team.player_data[i]['Shots_on_border']))
                except:
                   median_sob_forward.append(int(0))
                try:
                    median_bilanz_forward.append(int(away_team.player_data[i]['+/-']))
                except:
                   median_bilanz_forward.append(int(0))
                try:
                    median_bullies_forward.append(int(away_team.player_data[i]['FO_total']))
                except:
                   median_bullies_forward.append(int(0))   
                try:
                    median_bullies_won_forward.append(int(away_team.player_data[i]['FO_won']))
                except:
                   median_bullies_won_forward.append(int(0))
                try:
                    median_bullies_lost_forward.append(int(away_team.player_data[i]['FO_lost']))
                except:
                   median_bullies_lost_forward.append(int(0))
                try:
                    median_blocked_shots_forward.append(int(away_team.player_data[i]['Blocked_shots']))
                except:
                   median_blocked_shots_forward.append(int(0))
                try:
                    median_shs_forward.append(int(away_team.player_data[i]['Box_Play_Goals']))
                except:
                   median_shs_forward.append(int(0))
                try:
                    median_bpa_forward.append(int(away_team.player_data[i]['Box_Play_Assists']))
                except:
                   median_bpa_forward.append(int(0))
                try:
                    median_ppg_forward.append(int(away_team.player_data[i]['PP_Goals']))
                except:
                   median_ppg_forward.append(int(0))
                try:
                    median_ppa_forward.append(int(away_team.player_data[i]['PP_Assists']))
                except:
                   median_ppa_forward.append(int(0))    
                try:
                    median_pim_forward.append(int(away_team.player_data[i]['PIM']))
                except:
                   median_pim_forward.append(int(0))

                if int(away_team.player_data[i]['Games_played']) >= 5:
                    forward_player_average += 1
                    temp = float(away_team.player_data[i]['Games_played'])
                    res = average_gp_forward + temp
                    average_gp_forward = res
                    temp = float(away_team.player_data[i]['Goals'])
                    res = average_goals_forward + temp
                    average_goals_forward = res
                    temp = float(away_team.player_data[i]['Assists'])
                    res = average_assists_forward + temp
                    average_assists_forward = res
                    try:
                        temp = float(away_team.player_data[i]['Shots_total'])
                    except:
                       temp = 0
                    try:
                        res = temp + average_sog_forward
                    except:
                       temp = 0
                    average_sog_forward = res
                    try:
                        temp = float(away_team.player_data[i]['Shots_from_slot'])
                    except:
                       temp = 0
                    res = temp + average_slot_sog_forward
                    average_slot_sog_forward = res
                    try:
                       temp = float(away_team.player_data[i]['Shots_missed'])
                    except:
                       temp = 0
                    res = temp + average_shot_missed_forward
                    average_shot_missed_forward = res
                    try:
                       temp = float(away_team.player_data[i]['Shots_on_border'])
                    except:
                       temp = 0
                    res = temp + average_sob_forward
                    average_sob_forward = res
                    try:
                       temp = float(away_team.player_data[i]['+/-'])
                    except:
                       temp = 0
                    res = temp + average_bilanz_forward
                    average_bilanz_forward = res
                    try:
                       temp = float(away_team.player_data[i]['FO_total'])
                    except:
                       temp = 0
                    res = temp + average_bullies_forward
                    average_bullies_forward = res
                    try:
                       temp = float(away_team.player_data[i]['FO_won'])
                    except:
                       temp = 0
                    res = temp + average_bullies_won_forward
                    average_bullies_won_forward = res
                    try:
                       temp = float(away_team.player_data[i]['FO_lost'])
                    except:
                       temp = 0
                    res = temp + average_bullies_lost_forward
                    average_bullies_lost_forward = res
                    try:
                       temp = float(away_team.player_data[i]['Blocked_shots'])
                    except:
                       temp = 0
                    res = temp + average_blocked_shots_forward
                    average_blocked_shots_forward = res
                    try:
                       temp = float(away_team.player_data[i]['Box_Play_Goals'])
                    except:
                       temp = 0
                    res = temp + average_shs_forward
                    average_shs_forward = res 
                    try:
                       temp = float(away_team.player_data[i]['Box_Play_Assists'])
                    except:
                       temp = 0
                    res = temp + average_bpa_forward
                    average_bpa_forward = res 
                    try:
                       temp = float(away_team.player_data[i]['PP_Goals'])
                    except:
                       temp = 0
                    res = temp + average_ppg_forward
                    average_ppg_forward = res 
                    try:
                       temp = float(away_team.player_data[i]['PP_Assists'])
                    except:
                       temp = 0
                    res = temp + average_ppa_forward
                    average_ppa_forward = res
                    try:
                       temp = float(away_team.player_data[i]['PIM'])
                    except:
                       temp = 0
                    res = temp + average_pim_forward
                    average_pim_forward = res 
            case 'Verteidiger':
                start_write_col = start_column
                count_defense += 1
                start_row = count_defense
                try:
                   median_gp_defense.append(int(away_team.player_data[i]['Games_played']))
                except: 
                   median_gp_defense.append(int(0))
                try:
                   median_goals_defense.append(int(away_team.player_data[i]['Goals']))
                except: 
                   median_goals_defense.append(int(0))
                try:
                   median_assists_defense.append(int(away_team.player_data[i]['Assists']))
                except: 
                    median_assists_defense.append(int(0))      
                try:
                    median_sog_defense.append(int(away_team.player_data[i]['Shots_total']))
                except: 
                   median_sog_defense.append(int(0))
                try:
                    median_slot_sog_defense.append(int(away_team.player_data[i]['Shots_from_slot']))
                except:
                   median_slot_sog_defense.append(int(0))
                try:
                    median_shot_missed_defense.append(int(away_team.player_data[i]['Shots_missed']))
                except:
                   median_shot_missed_defense.append(int(0))
                try:
                    median_sob_defense.append(int(away_team.player_data[i]['Shots_on_border']))
                except:
                   median_sob_defense.append(int(0))
                try:
                    median_bilanz_defense.append(int(away_team.player_data[i]['+/-']))
                except:
                   median_bilanz_defense.append(int(0))
                try:
                    median_bullies_defense.append(int(away_team.player_data[i]['FO_total']))
                except:
                   median_bullies_defense.append(int(0))   
                try:
                    median_bullies_won_defense.append(int(away_team.player_data[i]['FO_won']))
                except:
                   median_bullies_won_defense.append(int(0))
                try:
                    median_bullies_lost_defense.append(int(away_team.player_data[i]['FO_lost']))
                except:
                   median_bullies_lost_defense.append(int(0))
                try:
                    median_blocked_shots_defense.append(int(away_team.player_data[i]['Blocked_shots']))
                except:
                   median_blocked_shots_defense.append(int(0))
                try:
                    median_shs_defense.append(int(away_team.player_data[i]['Box_Play_Goals']))
                except:
                   median_shs_defense.append(int(0))
                try:
                    median_bpa_defense.append(int(away_team.player_data[i]['Box_Play_Assists']))
                except:
                   median_bpa_defense.append(int(0))
                try:
                    median_ppg_defense.append(int(away_team.player_data[i]['PP_Goals']))
                except:
                   median_ppg_defense.append(int(0))
                try:
                    median_ppa_defense.append(int(away_team.player_data[i]['PP_Assists']))
                except:
                   median_ppa_defense.append(int(0))    
                try:
                    median_pim_defense.append(int(away_team.player_data[i]['PIM']))
                except:
                   median_pim_defense.append(int(0))

                if int(away_team.player_data[i]['Games_played']) >= 5:
                    defense_player_average += 1
                    try:
                       temp = float(away_team.player_data[i]['Games_played'])
                    except:
                       temp = 0
                    res = average_gp_defense + temp
                    average_gp_defense = res
                    try:
                       temp = float(away_team.player_data[i]['Goals'])
                    except:
                       temp = 0
                    res = average_goals_defense + temp
                    average_goals_defense = res
                    try:
                       temp = float(away_team.player_data[i]['Assists'])
                    except:
                       temp = 0
                    res = average_assists_defense + temp
                    average_assists_defense = res
                    try:
                       temp = float(away_team.player_data[i]['Shots_total'])
                    except:
                       temp = 0
                    res = temp + average_sog_defense
                    average_sog_defense = res
                    try:
                       temp = float(away_team.player_data[i]['Shots_from_slot'])
                    except:
                       temp = 0
                    res = temp + average_slot_sog_defense
                    average_slot_sog_defense = res
                    try:
                       temp = float(away_team.player_data[i]['Shots_missed'])
                    except:
                       temp = 0
                    res = temp + average_shot_missed_defense
                    average_shot_missed_defense = res
                    try:
                       temp = float(away_team.player_data[i]['Shots_on_border'])
                    except:
                       temp = 0
                    res = temp + average_sob_defense
                    average_sob_defense = res
                    try:
                       temp = float(away_team.player_data[i]['+/-'])
                    except:
                       temp = 0
                    res = temp + average_bilanz_defense
                    average_bilanz_defense = res
                    try:
                       temp = float(away_team.player_data[i]['FO_total'])
                    except:
                       temp = 0
                    res = temp + average_bullies_defense
                    average_bullies_defense = res
                    try:
                       temp = float(away_team.player_data[i]['FO_won'])
                    except:
                       temp = 0
                    res = temp + average_bullies_won_defense
                    average_bullies_won_defense = res
                    try:
                       temp = float(away_team.player_data[i]['FO_lost'])
                    except:
                       temp = 0
                    res = temp + average_bullies_lost_defense
                    average_bullies_lost_defense = res
                    try:
                       temp = float(away_team.player_data[i]['Blocked_shots'])
                    except:
                       temp = 0
                    res = temp + average_blocked_shots_defense
                    average_blocked_shots_defense = res
                    try:
                       temp = float(away_team.player_data[i]['Box_Play_Goals'])
                    except:
                       temp = 0
                    res = temp + average_shs_defense
                    average_shs_defense = res 
                    try:
                       temp = float(away_team.player_data[i]['Box_Play_Assists'])
                    except:
                       temp = 0
                    res = temp + average_bpa_defense
                    average_bpa_defense = res 
                    try:
                       temp = float(away_team.player_data[i]['PP_Goals'])
                    except:
                       temp = 0
                    res = temp + average_ppg_defense
                    average_ppg_defense = res 
                    try:
                       temp = float(away_team.player_data[i]['PP_Assists'])
                    except:
                       temp = 0
                    res = temp + average_ppa_defense
                    average_ppa_defense = res
                    try:
                       temp = float(away_team.player_data[i]['PIM'])
                    except:
                       temp = 0
                    res = temp + average_pim_defense
                    average_pim_defense = res 
            case 'Torhüter':
              start_write_col = goalie_start_col
              count_goalie += 1
              start_row = count_goalie
              goalie_player_count += 1
              try:
                my_worksheet.write(start_row,
                           start_write_col,
                           away_team.player_data[i]['Name'])

              except: 
                pass
              start_write_col += 1
              try:
                my_worksheet.write(start_row,
                           start_write_col,
                           away_team.player_data[i]['Games_played'])
                temp_gp = total_gp
                total_gp = temp_gp + int(away_team.player_data[i]['Games_played'])
              except: 
                pass
              start_write_col += 1
              try:
                my_worksheet.write(start_row,
                           start_write_col,
                           away_team.player_data[i]['Goals_against'])
                temp_goals = total_received_goals
                total_received_goals = temp_goals + int(away_team.player_data[i]['Goals_against'])
              except: 
                pass
              start_write_col += 1
              try:
                my_worksheet.write(start_row,
                           start_write_col,
                           away_team.player_data[i]['Total_shots_agains'])
                temp_shots = total_received_shots
                total_received_shots = temp_shots + int(away_team.player_data[i]['Total_shots_agains'])
              except: 
                pass
              start_write_col += 1
              total_shots = int(away_team.player_data[i]['Total_shots_agains'])
              total_save_rate = float(away_team.player_data[i]['Total_save%'])
              temp_rate = total_save_rate_overall
              total_save_rate_overall = temp_rate + total_save_rate
              total_saves = round((total_shots/100) * total_save_rate)
              temp_saves = total_saves_overall
              total_saves_overall = temp_saves + total_saves
              my_worksheet.write(start_row,
                           start_write_col,
                           total_saves)
              start_write_col += 1
              try:
                my_worksheet.write(start_row,
                           start_write_col,
                           away_team.player_data[i]['Total_save%'])
              except: 
                pass
              start_write_col += 1
              try:
                my_worksheet.write(start_row,
                           start_write_col,
                           away_team.player_data[i]['Shots_from_slot'])
                temp_shots = total_shots_from_slot
                total_shots_from_slot = temp_shots + int(away_team.player_data[i]['Shots_from_slot'])
              except: 
                pass
              start_write_col += 1
              shots_from_slot = int(away_team.player_data[i]['Shots_from_slot'])
              slot_save_rate = float(away_team.player_data[i]['Slot_save%'])
              temp_rate = total_save_rate_slot
              total_save_rate_slot = temp_rate + slot_save_rate

              slot_saves = round((shots_from_slot/100) * slot_save_rate)
              temp_saves = total_saves_from_slot_overall
              total_saves_from_slot_overall = temp_saves + slot_saves
              my_worksheet.write(start_row,
                           start_write_col,
                           slot_saves)
              start_write_col += 1
              try:
                my_worksheet.write(start_row,
                           start_write_col,
                           away_team.player_data[i]['Slot_save%'])
              except: 
                pass
              start_write_col += 1
              try:
                my_worksheet.write(start_row,
                           start_write_col,
                           away_team.player_data[i]['Missed_shots'])
                temp_shots = total_missed_shots
                total_missed_shots = temp_shots + int(away_team.player_data[i]['Missed_shots'])
              except: 
                pass
              start_write_col += 1
              try:
                my_worksheet.write(start_row,
                           start_write_col,
                           away_team.player_data[i]['Shots_against_border'])
                temp_shots = total_shots_on_frame
                total_shots_on_frame = temp_shots + int(away_team.player_data[i]['Shots_against_border'])
              except: 
                pass
              start_write_col += 1
              try:
                my_worksheet.write(start_row,
                           start_write_col,
                           away_team.player_data[i]['PIM'])
                temp_pim = total_pim
                total_pim = temp_pim + int(away_team.player_data[i]['PIM'])
              except: 
                pass
              
              continue
            # the python equivalent to WHEN OTHER
            case _:
                # in contrast to the expectation the continue statement actually
                # ensures that loop skips the current iteration and goes to the next one
                continue
        try:
            my_worksheet.write(start_row,
                           start_write_col,
                           away_team.player_data[i]['Name'])
        except: 
           pass
        start_write_col += 1

        try:
            my_worksheet.write(start_row,
                           start_write_col,
                           away_team.player_data[i]['Games_played'])
        except: 
           pass
        start_write_col += 1

        try:
            my_worksheet.write(start_row,
                           start_write_col,
                           away_team.player_data[i]['Goals'])
        except: 
           pass
        start_write_col += 1

        try:
            my_worksheet.write(start_row,
                           start_write_col,
                           away_team.player_data[i]['Assists'])
        except: 
           pass
        start_write_col += 1

        try:
           my_worksheet.write(start_row,
                           start_write_col,
                           away_team.player_data[i]['Shots_total'])
        except: 
           pass
        start_write_col += 1

        try: 
            my_worksheet.write(start_row,
                           start_write_col,
                           away_team.player_data[i]['Shots_from_slot'])
        except:
           pass
        start_write_col += 1

        try: 
            my_worksheet.write(start_row,
                           start_write_col,
                           away_team.player_data[i]['Shots_missed'])
        except:
           pass
        start_write_col += 1
        try:
            my_worksheet.write(start_row,
                           start_write_col,
                           away_team.player_data[i]['Shots_on_border'])
        except:
            pass
        start_write_col += 1

        try:
            my_worksheet.write(start_row,
                           start_write_col,
                           away_team.player_data[i]['+/-'])
        except:
            pass
        start_write_col += 1

        try:
            my_worksheet.write(start_row,
                           start_write_col,
                           away_team.player_data[i]['FO_total'])
        except:
            pass
        start_write_col += 1

        try:
            my_worksheet.write(start_row,
                           start_write_col,
                           away_team.player_data[i]['FO_won'])
        except:
            pass
        start_write_col += 1

        try:
            my_worksheet.write(start_row,
                           start_write_col,
                           away_team.player_data[i]['FO_lost'])
        except:
           pass
        start_write_col += 1

        try:
            my_worksheet.write(start_row,
                           start_write_col,
                           away_team.player_data[i]['Blocked_shots'])
        except:
           pass
        start_write_col += 1

        try:
            my_worksheet.write(start_row,
                           start_write_col,
                           away_team.player_data[i]['Box_Play_Goals'])
        except:
            pass
        start_write_col += 1

        try:
            my_worksheet.write(start_row,
                           start_write_col,
                           away_team.player_data[i]['Box_Play_Assists'])
        except: 
           pass
        start_write_col += 1

        try:
            my_worksheet.write(start_row,
                           start_write_col,
                           away_team.player_data[i]['PP_Goals'])
        except: 
           pass
        start_write_col += 1

        try:
            my_worksheet.write(start_row,
                           start_write_col,
                           away_team.player_data[i]['PP_Assists'])
        except: 
           pass
        start_write_col += 1

        try:
            my_worksheet.write(start_row,
                           start_write_col,
                           away_team.player_data[i]['PIM'])
        except: 
           pass
    start_write_col = 0
    count_forward += 1
    start_row = count_forward
    my_worksheet.write(start_row,
                           start_write_col,
                           (away_team.name + " Stürmer Durchschnitt (min 5 Spiele)"))
    start_write_col += 1
    average_gp_forward = average_gp_forward / forward_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_gp_forward)
    start_write_col += 1
    average_goals_forward = average_goals_forward / forward_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_goals_forward)
    start_write_col += 1
    average_assists_forward = average_assists_forward / forward_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_assists_forward)
    start_write_col += 1
    average_sog_forward = average_sog_forward / forward_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_sog_forward)
    start_write_col += 1

    average_slot_sog_forward = average_slot_sog_forward / forward_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_slot_sog_forward)
    start_write_col += 1

    average_shot_missed_forward = average_shot_missed_forward / forward_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_shot_missed_forward)
    start_write_col += 1
    average_sob_forward = average_sob_forward / forward_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_sob_forward)
    start_write_col += 1
    average_bilanz_forward = average_bilanz_forward / forward_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_bilanz_forward)
    start_write_col += 1
    average_bullies_forward = average_bullies_forward / forward_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_bullies_forward)
    start_write_col += 1
    average_bullies_won_forward = average_bullies_won_forward / forward_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_bullies_won_forward)
    start_write_col += 1
    average_bullies_lost_forward = average_bullies_lost_forward / forward_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_bullies_lost_forward)
    start_write_col += 1
    average_blocked_shots_forward = average_blocked_shots_forward / forward_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_blocked_shots_forward)
    start_write_col += 1
    average_shs_forward = average_shs_forward / forward_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_shs_forward)
    start_write_col += 1
    average_bpa_forward = average_bpa_forward / forward_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_bpa_forward)
    start_write_col += 1
    average_ppg_forward = average_ppg_forward / forward_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_ppg_forward)
    start_write_col += 1
    average_ppa_forward = average_ppa_forward / forward_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_ppa_forward)
    start_write_col += 1
    average_pim_forward = average_pim_forward / forward_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_pim_forward)

    start_write_col = 0
    count_forward += 1
    start_row = count_forward
    my_worksheet.write(start_row,
                           start_write_col,
                           (away_team.name + " Stürmer Median"))
    start_write_col += 1

    temp = statistics.median(median_gp_forward)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_goals_forward)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_assists_forward)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    average_sog_forward = average_sog_forward / forward_player_average
    temp = statistics.median(median_sog_forward)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_slot_sog_forward)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_shot_missed_forward)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_sob_forward)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_bilanz_forward)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_bullies_forward)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_bullies_won_forward)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_bullies_lost_forward)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_blocked_shots_forward)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_shs_forward)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_bpa_forward)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_ppg_forward)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_ppg_forward)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_pim_forward)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    # and now the defense
    start_write_col = start_column
    count_defense += 1
    start_row = count_defense
    my_worksheet.write(start_row,
                           start_write_col,
                           (away_team.name + " Verteidiger Durchschnitt (min 5 Spiele)"))
    start_write_col += 1

    average_gp_defense = average_gp_defense / defense_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_gp_defense)
    start_write_col += 1
    average_goals_defense = average_goals_defense / defense_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_goals_defense)
    start_write_col += 1
    average_assists_defense = average_assists_defense / defense_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_assists_defense)
    start_write_col += 1
    average_sog_defense = average_sog_defense / defense_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_sog_defense)
    start_write_col += 1

    average_slot_sog_defense = average_slot_sog_defense / defense_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_slot_sog_defense)
    start_write_col += 1

    average_shot_missed_defense = average_shot_missed_defense / defense_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_shot_missed_defense)
    start_write_col += 1
    average_sob_defense = average_sob_defense / defense_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_sob_defense)
    start_write_col += 1
    average_bilanz_defense = average_bilanz_defense / defense_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_bilanz_defense)
    start_write_col += 1
    average_bullies_defense = average_bullies_defense / defense_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_bullies_defense)
    start_write_col += 1
    average_bullies_won_defense = average_bullies_won_defense / defense_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_bullies_won_defense)
    start_write_col += 1
    average_bullies_lost_defense = average_bullies_lost_defense / defense_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_bullies_lost_defense)
    start_write_col += 1
    average_blocked_shots_defense = average_blocked_shots_defense / defense_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_blocked_shots_defense)
    start_write_col += 1
    average_shs_defense = average_shs_defense / defense_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_shs_defense)
    start_write_col += 1
    average_bpa_defense = average_bpa_defense / defense_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_bpa_defense)
    start_write_col += 1
    average_ppg_defense = average_ppg_defense / defense_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_ppg_defense)
    start_write_col += 1
    average_ppa_defense = average_ppa_defense / defense_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_ppa_defense)
    start_write_col += 1
    average_pim_defense = average_pim_defense / defense_player_average
    my_worksheet.write(start_row,
                           start_write_col,
                           average_pim_defense)

    start_write_col = start_column
    count_defense += 1
    start_row = count_defense
    my_worksheet.write(start_row,
                           start_write_col,
                           (away_team.name + " Verteidiger Median"))
    start_write_col += 1

    temp = statistics.median(median_gp_defense)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_goals_defense)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_assists_defense)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    average_sog_defense = average_sog_defense / defense_player_average
    temp = statistics.median(median_sog_defense)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_slot_sog_defense)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_shot_missed_defense)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_sob_defense)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_bilanz_defense)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_bullies_defense)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_bullies_won_defense)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_bullies_lost_defense)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_blocked_shots_defense)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_shs_defense)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_bpa_defense)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_ppg_defense)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_ppg_defense)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    start_write_col += 1
    temp = statistics.median(median_pim_defense)
    my_worksheet.write(start_row,
                           start_write_col,
                           temp)
    # now the goalie total 
    start_write_col = goalie_start_col
    count_goalie += 1
    start_row = count_goalie

    my_worksheet.write(start_row,
                           start_write_col,
                           'Total')
    start_write_col += 1
    my_worksheet.write(start_row,
                           start_write_col,
                           total_gp)
    start_write_col += 1
    my_worksheet.write(start_row,
                           start_write_col,
                           total_received_goals)
    start_write_col += 1
    my_worksheet.write(start_row,
                           start_write_col,
                           total_received_shots)
    start_write_col += 1
    my_worksheet.write(start_row,
                           start_write_col,
                           total_saves_overall)
    
    start_write_col += 1
    total_save_rate_overall = total_save_rate_overall / goalie_player_count
    my_worksheet.write(start_row,
                           start_write_col,
                           total_save_rate_overall)
    start_write_col += 1
    my_worksheet.write(start_row,
                           start_write_col,
                           total_shots_from_slot)
    start_write_col += 1
    my_worksheet.write(start_row,
                           start_write_col,
                           total_saves_from_slot_overall)
    start_write_col += 1
    total_save_rate_slot = total_save_rate_slot / goalie_player_count
    my_worksheet.write(start_row,
                           start_write_col,
                           total_save_rate_slot)
    start_write_col += 1
    my_worksheet.write(start_row,
                           start_write_col,
                           total_missed_shots)
    start_write_col += 1
    my_worksheet.write(start_row,
                           start_write_col,
                           total_shots_on_frame)
    start_write_col += 1
    my_worksheet.write(start_row,
                           start_write_col,
                           total_pim)
    
    my_worksheet.add_table(start_row_title,0,count_forward,17,{'style': 'Table Style Light 16',
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
    my_worksheet.add_table(start_row_title,start_column,count_defense,my_column_pos,{'style': 'Table Style Light 16',
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
    
    my_worksheet.add_table(start_row_title,goalie_start_col,count_goalie,(goalie_start_col + 11),{'style': 'Table Style Light 16',
                                             'columns': [{'header': 'Player'},
                                                         {'header': 'GP'},
                                                         {'header': 'Erhaltene Tore'},
                                                         {'header': 'Schüsse aufs Tor'},
                                                         {'header': 'Saves'},
                                                         {'header': 'Saves %'},
                                                         {'header': 'Schüsse aus dem Slot'},
                                                         {'header': 'Slot Saves'},
                                                         {'header': 'Slot Save %'},
                                                         {'header': 'Schüsse neben das Tor'},
                                                         {'header': 'Schüsse an die Torumrandung'},
                                                         {'header': 'Penalty Minutes'}
                                                         ]})
    my_worksheet.autofit()
def get_team_stats():
     request_url = (BASE_URL + STAT_CACHE + TEAM_GOAL_OVERVIEW + QUARY_BASIC + CURRENT_SEASON + PHASE_REGULAR
                  + SORT_BY_TEAM + STANDARD_ENDING) 
     team_goal_data = send_request(request_url)

     request_url = (BASE_URL + STAT_CACHE + TEAM_POSITION_GOALS + QUARY_BASIC + CURRENT_SEASON + PHASE_REGULAR
                  + SORT_BY_TEAM + STANDARD_ENDING) 
     team_positoin_goal_data = send_request(request_url)
     request_url = (BASE_URL + STAT_CACHE + TEAM_SHOT_DETAILS + QUARY_BASIC + CURRENT_SEASON + PHASE_REGULAR
                  + SORT_BY_TEAM + STANDARD_ENDING) 
     team_shot_data = send_request(request_url)
     request_url = (BASE_URL + STAT_CACHE + TEAM_PP_DATA + QUARY_BASIC + CURRENT_SEASON + PHASE_REGULAR
                  + SORT_BY_TEAM + STANDARD_ENDING) 
     team_pp_data = send_request(request_url)
     request_url = (BASE_URL + STAT_CACHE + TEAM_PK_DATA + QUARY_BASIC + CURRENT_SEASON + PHASE_REGULAR
                  + SORT_BY_TEAM + STANDARD_ENDING) 
     team_pk_data = send_request(request_url)
     return (team_goal_data, team_positoin_goal_data, team_shot_data, team_pp_data, team_pk_data)
#
#
#
def fill_team_data(my_workbook: xlsxwriter.Workbook, my_worksheet: xlsxwriter.worksheet, team_dict: dict) -> None:
    # first we write the descriptions 
    start_row = 0
    start_column = 0
    header_format = my_workbook.add_format({'bg_color': '#BFBFBF', 
                                            'bold': True, 
                                            'valign': 'center',
                                            'align': 'center',
                                            'align': 'vcenter',
                                            'font_name': 'ADLaM Display',
                                            })
    
    goal__title_format = my_workbook.add_format({'bg_color': '#B5E6A2', 
                                                    'bold': True, 
                                                    'valign': 'center',
                                                    'align': 'center',
                                                    'align': 'vcenter',
                                                    'font_name': 'ADLaM Display',
                                                    })
    
    goal_format = my_workbook.add_format({'bg_color': '#DAF2D0', 
                                                    'bold': False, 
                                                    'valign': 'left',
                                                    'align': 'left',
                                                    'font_name': 'Aptos Narrow',
                                                    })
    
    shot_title_format = my_workbook.add_format({'bg_color': '#94DCF8', 
                                              'bold': True, 
                                              'valign': 'center',
                                              'align': 'center',
                                              'align': 'vcenter',
                                              'font_name': 'ADLaM Display',
                                              })
    
    shot_format = my_workbook.add_format({'bg_color': '#CAEDFB', 
                                        'bold': False, 
                                        'valign': 'left',
                                        'align': 'left',
                                        'font_name': 'Aptos Narrow',
                                        })
    
    pp_title_format = my_workbook.add_format({'bg_color': '#A6C9EC', 
                                              'bold': True, 
                                              'valign': 'center',
                                              'align': 'center',
                                              'align': 'vcenter',
                                              'font_name': 'ADLaM Display',
                                              })
    
    pp_format = my_workbook.add_format({'bg_color': '#DAE9F8', 
                                        'bold': False, 
                                        'valign': 'left',
                                        'align': 'left',
                                        'font_name': 'Aptos Narrow',
                                        })
    
    pk_title_format = my_workbook.add_format({'bg_color': '#A4ADEE', 
                                              'bold': True, 
                                              'valign': 'center',
                                              'align': 'center',
                                              'align': 'vcenter',
                                              'font_name': 'ADLaM Display',
                                              })
    
    pk_format = my_workbook.add_format({'bg_color': '#D5D9F7', 
                                        'bold': False, 
                                        'valign': 'left',
                                        'align': 'left',
                                        'font_name': 'Aptos Narrow',
                                        })
    # title row
    my_worksheet.write(start_row, start_column, 'Was', header_format)
    
    # title of goal group
    start_row +=1
    my_worksheet.write(start_row, start_column, 'Tore', goal__title_format)
    my_worksheet.set_row(start_row, 23.25)

    # goal group descriptions
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Gespielte Spiele', goal_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Total erhaltene Tore', goal_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Total erziehlte Tore', goal_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Tore erziehlt durch Verteidiger', goal_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Anteil an Toren Verteidiger %', goal_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Tore erziehlt durch Stürmer', goal_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Anteil an Toren Stürmer %', goal_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Tore erziehlt durch Imports', goal_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Anteil an Toren Imports %', goal_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Tore pro Verteidiger mit min 5 Spielen', goal_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Tore pro Stürmer mit min 5 Spielen', goal_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Tore pro Import mit min 5 Spielen', goal_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Tore pro Spiel', goal_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Erhaltene Tore pro Spiel', goal_format)

    # title of shot group
    start_row +=2
    my_worksheet.write(start_row, start_column, 'Schüsse', shot_title_format)
    my_worksheet.set_row(start_row, 23.25)
    # shot group descriptions
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Total Schüsse aufs Tor', shot_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Schüsse aufs Tor aus dem Slot', shot_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Schüsse an die Torumrandung', shot_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Schüsse neben das Tor', shot_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Geblockte Schüsse', shot_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Schüsse aufs Tor pro Spiel', shot_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Schüsse aufs Tor aus dem Slot pro Spiel', shot_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Schüsse and die Torumrandung pro Spiel', shot_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Schüsse neben das Tor pro Spiel', shot_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Schüsse Effizienz %', shot_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Geblockte Schüsse pro Spiel', shot_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Total erhaltene Schüsse aufs Tor', shot_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Erhaltene Schüsse aufs Tor aus dem Slot', shot_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Schusseffizienz der Gegner %', shot_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Erhaltene Schüsse aufs Tor pro Spiel', shot_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Erhaltene Schüsse aufs Tor aus dem Slot pro Spiel', shot_format)

    # title of power play group
    start_row +=2
    my_worksheet.write(start_row, start_column, 'Powerplay', pp_title_format)
    my_worksheet.set_row(start_row, 23.25)
    # powerplay group descriptions
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Powerplay Situationen', pp_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Minuten im Powerplay', pp_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Erziehlte Tore im Powerplay', pp_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Powerplay Situationen pro Spiel', pp_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Erziehlte Tore pro Powerplay', pp_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Durchnittliche Minuten im Powerplay bis zu einem Tor', pp_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Durchschnittliche Länge eines Powerplays', pp_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Powerplay Effizienz %', pp_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Erhaltene Tore im Powerplay', pp_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Erhaltene Tore pro Powerplay', pp_format)

    # title of boxplay group
    start_row +=2
    my_worksheet.write(start_row, start_column, 'Boxplay', pk_title_format)
    my_worksheet.set_row(start_row, 23.25)
    # boxplay group descriptions
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Boxplay Situationen', pk_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Minuten im Boxplay', pk_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Erhaltene Tore im Boxplay', pk_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Boxplay Situationen pro Spiel', pk_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Erhaltene Tore pro Boxplay', pk_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Minuten im Boxplay bis zu einem Tor', pk_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Durchnittliche Länge eines Boxplays', pk_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Boxplay Effizienz %', pk_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Erziehlte Tore im Boxplay', pk_format)
    start_row += 1
    my_worksheet.write(start_row, start_column, 'Erziehlte Tore pro Boxplay', pk_format)
    my_worksheet.autofit()
    # and now we have to fill the actuall data from the teams from the team dict
    for i in team_dict:
       start_column += 1
       start_row = 0
       # the title / team name
       my_worksheet.write(start_row, start_column, team_dict[i].name, header_format)
       # after some displays and some calculations i figured out that the cell_autofit_width
       # works with a dict / map where the width for each character is hard coded since 
       # true autofitting is acutally only possible in an excel runtime
       # for the font size and font type I use, the calculated width was too small 
       # and 1.34 will make the cells a bit too big (up to 2 pixels) but this is a compromise I
       # am willing to accept 
       max_width = round(cell_autofit_width(team_dict[i].name) * 1.34)
       my_worksheet.set_column_pixels(start_column, start_column,max_width)
       # the goal stats
       start_row += 2
       my_worksheet.write(start_row, start_column, team_dict[i].amount_of_games, goal_format)
       start_row += 1
       my_worksheet.write(start_row, start_column, team_dict[i].goals_against, goal_format)
       start_row += 1
       my_worksheet.write(start_row, start_column, team_dict[i].goals_made, goal_format)
       start_row += 1
       my_worksheet.write(start_row, start_column, team_dict[i].goals_made_defense, goal_format)
       start_row += 1
       temp = round(100/team_dict[i].goals_made * team_dict[i].goals_made_defense, 2)
       my_worksheet.write(start_row, start_column, temp, goal_format)
       start_row += 1
       my_worksheet.write(start_row, start_column, team_dict[i].goals_made_forward, goal_format)
       start_row += 1
       temp = round(100/team_dict[i].goals_made * team_dict[i].goals_made_forward, 2)
       my_worksheet.write(start_row, start_column, temp, goal_format)
       start_row += 1
       my_worksheet.write(start_row, start_column, team_dict[i].goals_made_imports, goal_format)
       start_row += 1
       temp = round(100/team_dict[i].goals_made * team_dict[i].goals_made_imports, 2)
       my_worksheet.write(start_row, start_column, temp, goal_format)
       start_row += 1
       my_worksheet.write(start_row, start_column, 'Will come with the next release', goal_format)
       start_row += 1
       my_worksheet.write(start_row, start_column, 'Will come with the next release', goal_format)
       start_row += 1
       my_worksheet.write(start_row, start_column, 'Will come with the next release', goal_format)
       start_row += 1
       temp = round(team_dict[i].goals_made / team_dict[i].amount_of_games, 2)
       my_worksheet.write(start_row, start_column, temp, goal_format)
       start_row += 1
       temp = round(team_dict[i].goals_against / team_dict[i].amount_of_games, 2)
       my_worksheet.write(start_row, start_column, temp, goal_format)
       # shot data
       start_row += 3
       my_worksheet.write(start_row, start_column, team_dict[i].sog, shot_format)
       start_row += 1
       my_worksheet.write(start_row, start_column, team_dict[i].sog_slot, shot_format)
       start_row += 1
       my_worksheet.write(start_row, start_column, team_dict[i].sob, shot_format)
       start_row += 1
       my_worksheet.write(start_row, start_column, team_dict[i].shots_missed, shot_format)
       start_row += 1
       my_worksheet.write(start_row, start_column, team_dict[i].blocked_shots, shot_format)
       start_row += 1
       temp = round(team_dict[i].sog / team_dict[i].amount_of_games, 2)
       my_worksheet.write(start_row, start_column, temp, shot_format)
       start_row += 1
       temp = round(team_dict[i].sog_slot / team_dict[i].amount_of_games, 2)
       my_worksheet.write(start_row, start_column, temp, shot_format)
       start_row += 1
       temp = round(team_dict[i].sob / team_dict[i].amount_of_games, 2)
       my_worksheet.write(start_row, start_column, temp, shot_format)
       start_row += 1
       temp = round(team_dict[i].shots_missed / team_dict[i].amount_of_games, 2)
       my_worksheet.write(start_row, start_column, temp, shot_format)
       start_row += 1
       temp = round(100/team_dict[i].sog * team_dict[i].goals_made, 2)
       my_worksheet.write(start_row, start_column, temp, shot_format)
       start_row += 1
       temp = round(team_dict[i].blocked_shots / team_dict[i].amount_of_games, 2)
       my_worksheet.write(start_row, start_column, temp, shot_format)
       start_row += 1
       my_worksheet.write(start_row, start_column, team_dict[i].sog_received, shot_format)
       start_row += 1
       my_worksheet.write(start_row, start_column, team_dict[i].sog_slot_received, shot_format)
       start_row += 1
       temp = round(100/team_dict[i].sog_received * team_dict[i].goals_against, 2)
       my_worksheet.write(start_row, start_column, temp, shot_format)
       start_row += 1
       temp = round(team_dict[i].sog_received / team_dict[i].amount_of_games, 2)
       my_worksheet.write(start_row, start_column, temp, shot_format)
       start_row += 1
       temp = round(team_dict[i].sog_slot_received / team_dict[i].amount_of_games, 2)
       my_worksheet.write(start_row, start_column, temp, shot_format)
       # powerplay stats
       start_row += 3
       my_worksheet.write(start_row, start_column, team_dict[i].pp_op, pp_format)
       start_row += 1
       my_worksheet.write(start_row, start_column, team_dict[i].time_in_pp, pp_format)
       start_row += 1
       my_worksheet.write(start_row, start_column, team_dict[i].pp_goals, pp_format)
       start_row += 1
       temp = round(team_dict[i].pp_op / team_dict[i].amount_of_games, 2)
       my_worksheet.write(start_row, start_column, temp, pp_format)
       start_row += 1
       temp = round(team_dict[i].pp_goals / team_dict[i].pp_op, 2)
       my_worksheet.write(start_row, start_column, temp, pp_format)
       # apparently i can use split() to create a list form my string and I can even provide a Delimitter!
       # so instead of INSPECT Var FOR CHARACTERS BEFORE INITIAL : TALLYING Counter
       # I receive a List with 1 - n entries excluding the Delimitter 
       # Example: 
       # 1234:5678 -> ['1234', '5678']
       temp_sec = 0
       temp_time = team_dict[i].time_in_pp.split(':')
      # convert minutes to seconds
       temp_sec = int(temp_time[0]) * 60
       # convert the seconds seperatly to an int to avoid the variable type of temp_sec
       # to turn into a tuple 
       temp_sec_bc_python_stopid = int(temp_time[1])
       temp_sec += temp_sec_bc_python_stopid
       temp_sec = temp_sec / team_dict[i].pp_goals
       # we could also do that with the divmod() function, since // gives you the quotient
       # and % the remainder
       temp_len_min = temp_sec // 60
       temp_len_sec = temp_sec % 60
       # add the decimal positions of the minutes to the calculated seconds
       temp_len_sec += temp_len_min % 1
       # turn the results into integers to cut off anything which is beyond the decimal point
       temp_len_min_rounded = int(temp_len_min)
       temp_len_sec_roundded = int(temp_len_sec)
       # create a string because i tried to do it first with join and there it needs a list as 
       # input and xlsx is not capable to receive too much as the input 
       temp_string_bc_xls_stopid = f"{temp_len_min_rounded:02}" + ':' + f"{temp_len_sec_roundded:02}"
       start_row += 1
       my_worksheet.write(start_row, start_column, temp_string_bc_xls_stopid, pp_format)

       # convert minutes to seconds
       temp_sec = int(temp_time[0]) * 60
       # convert the seconds seperatly to an int to avoid the variable type of temp_sec
       # to turn into a tuple 
       temp_sec_bc_python_stopid = int(temp_time[1])
       temp_sec += temp_sec_bc_python_stopid
       temp_sec = temp_sec / team_dict[i].pp_op
       # we could also do that with the divmod() function, since // gives you the quotient
       # and % the remainder
       temp_len_min = temp_sec // 60
       temp_len_sec = temp_sec % 60
       # add the decimal positions of the minutes to the calculated seconds
       temp_len_sec += temp_len_min % 1
       # turn the results into integers to cut off anything which is beyond the decimal point
       temp_len_min_rounded = int(temp_len_min)
       temp_len_sec_roundded = int(temp_len_sec)
       # create a string because i tried to do it first with join and there it needs a list as 
       # input and xlsx is not capable to receive too much as the input 
       # with the f"{x:02} we force x to turn into a string with a certain format
       temp_string_bc_xls_stopid = f"{temp_len_min_rounded:02}" + ':' + f"{temp_len_sec_roundded:02}"
       start_row += 1
       my_worksheet.write(start_row, start_column, temp_string_bc_xls_stopid, pp_format)
       start_row += 1
       temp = round(100 / team_dict[i].pp_op * team_dict[i].pp_goals, 2)
       my_worksheet.write(start_row, start_column, temp, pp_format)
       start_row += 1
       my_worksheet.write(start_row, start_column, team_dict[i].pp_received_goals, pp_format)
       start_row += 1
       temp = round(team_dict[i].pp_received_goals / team_dict[i].pp_op, 2)
       my_worksheet.write(start_row, start_column, temp, pp_format)
       # and now the boxplay stats
       start_row += 3
       my_worksheet.write(start_row, start_column, team_dict[i].pk_situations, pk_format)
       start_row += 1
       my_worksheet.write(start_row, start_column, team_dict[i].time_in_pk, pk_format)
       start_row += 1
       my_worksheet.write(start_row, start_column, team_dict[i].pk_goals_received, pk_format)
       start_row += 1
       temp = round(team_dict[i].pk_situations / team_dict[i].amount_of_games, 2)
       my_worksheet.write(start_row, start_column, temp, pk_format)
       start_row += 1
       temp = round(team_dict[i].pk_goals_received / team_dict[i].pk_situations, 2)
       my_worksheet.write(start_row, start_column, temp, pk_format)
       temp_sec = 0
       temp_time = team_dict[i].time_in_pk.split(':')
       temp_sec = int(temp_time[0]) * 60
       temp_sec_bc_python_stopid = int(temp_time[1])
       temp_sec += temp_sec_bc_python_stopid
       temp_sec = temp_sec / team_dict[i].pk_goals_received
       temp_len_min = temp_sec // 60
       temp_len_sec = temp_sec % 60
       temp_len_sec += temp_len_min % 1
       temp_len_min_rounded = int(temp_len_min)
       temp_len_sec_roundded = int(temp_len_sec)
       temp_string_bc_xls_stopid = f"{temp_len_min_rounded:02}" + ':' + f"{temp_len_sec_roundded:02}"
       start_row += 1
       my_worksheet.write(start_row, start_column, temp_string_bc_xls_stopid, pk_format)
       temp_sec = int(temp_time[0]) * 60
       temp_sec_bc_python_stopid = int(temp_time[1])
       temp_sec += temp_sec_bc_python_stopid
       temp_sec = temp_sec / team_dict[i].pk_situations
       temp_len_min = temp_sec // 60
       temp_len_sec = temp_sec % 60
       temp_len_sec += temp_len_min % 1
       temp_len_min_rounded = int(temp_len_min)
       temp_len_sec_roundded = int(temp_len_sec)
       temp_string_bc_xls_stopid = f"{temp_len_min_rounded:02}" + ':' + f"{temp_len_sec_roundded:02}"
       start_row += 1
       my_worksheet.write(start_row, start_column, temp_string_bc_xls_stopid, pk_format)
       start_row += 1
       temp = round(100 / team_dict[i].pk_situations * team_dict[i].pk_goals_received, 2)
       my_worksheet.write(start_row, start_column, temp, pk_format)
       start_row += 1
       my_worksheet.write(start_row, start_column, team_dict[i].shorthanders, pk_format)
       start_row += 1
       temp = round(team_dict[i].shorthanders / team_dict[i].pk_situations, 2)
       my_worksheet.write(start_row, start_column, temp, pk_format)

    # we still keep the autofit in order to get a nice fit for the first column 
    # and the other ones are not getting overwritten 
    my_worksheet.autofit()
    # documentation states that it starts the index at 0 but for some reason 
    # 1, 0 is required in order to fix the first row to the sheet
    # maybe it is to ensure that no column gets fixed together 
    # however, since it grew a bit I decided to fix column and row
    my_worksheet.freeze_panes(1,1)
    my_worksheet.set_row(0, 23.25)
    
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

    # iterate over the upcoming games, create a worksheet (also refered as tab) for each game in the excel
    # and fill the player data
    for i in (next_games['data']):
        tab_name = ''
        tab_name = (team_dict[i[4]['id']].name) + "_" + (team_dict[i[5]['id']].name) + "_PD"
        my_worksheet = add_worksheet(my_excelfile, tab_name)
        # the function does not return anything but we need to hand over the home as well as the away team 
        # and since we have the ID as the key in the dictionary, it is quite easy to pass it over
        create_player_sheet(my_excelfile,my_worksheet,team_dict[i[4]['id']],team_dict[i[5]['id']])
      
    # create now the team stats
    (team_goal_data, team_positoin_goal_data, team_shot_data
   , team_pp_data, team_pk_data) = get_team_stats()

    for i in (team_goal_data['data']):
      if i[1]['id'] == 100000:
         continue 
      team_dict[i[1]['id']].amount_of_games = int(i[2])
      team_dict[i[1]['id']].goals_made = int(i[3])
      team_dict[i[1]['id']].goals_against = int(i[5])
      team_dict[i[1]['id']].pp_goals = int(i[11])
      team_dict[i[1]['id']].pp_received_goals = int(i[13])
      team_dict[i[1]['id']].shorthanders = int(i[15])
      team_dict[i[1]['id']].pk_goals_received = int(i[17])

    for i in (team_positoin_goal_data['data']):
      if i[1]['id'] == 100000:
         continue
      team_dict[i[1]['id']].goals_made_forward = int(i[4])
      team_dict[i[1]['id']].goals_made_defense = int(i[6])
      team_dict[i[1]['id']].goals_made_imports = int(i[8])
    
    for i in (team_shot_data['data']):
      if i[1]['id'] == 100000:
         continue 
      team_dict[i[1]['id']].sog = int(i[3])
      team_dict[i[1]['id']].sog_slot = int(i[6])
      team_dict[i[1]['id']].shots_missed = int(i[8])
      team_dict[i[1]['id']].sob = int(i[10])
      team_dict[i[1]['id']].blocked_shots = int(i[12])
      team_dict[i[1]['id']].sog_received = int(i[14])
      team_dict[i[1]['id']].sog_slot_received = int(i[16])

    for i in (team_pp_data['data']):
      if i[1]['id'] == 100000:
         continue 
      team_dict[i[1]['id']].pp_op = int(i[3])
      team_dict[i[1]['id']].time_in_pp = i[8]

    for i in (team_pk_data['data']):
      if i[1]['id'] == 100000:
         continue 
      team_dict[i[1]['id']].pk_situations = int(i[3])
      team_dict[i[1]['id']].time_in_pk = i[8]

    tab_name = 'Team Stats'  
    team_tab = add_worksheet(my_excelfile, tab_name)
    fill_team_data(my_excelfile, team_tab, team_dict) 

    my_excelfile.close()
main()