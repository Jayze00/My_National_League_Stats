##########################################################################
# the imports
##########################################################################
import requests
import json
import re 
import xlsxwriter
import datetime 

##########################################################################
# the variables 
##########################################################################
base_url = 'https://data.sihf.ch/Statistic/api/cms/cache300?'
#kind of stat 
summary_stat = 'alias=player'
goal_stats = 'alias=playerGoalAssist'
shot_stat = 'alias=playerShotDetail'
faceoff_stat = 'alias=playerFaceoff'
goalie_summary_stat = 'alias=goalie'
goalie_shot_stat = 'alias=goalieShotDetail'
# strandard stuff for language, sorting and query definition 
standard_sorting = '&orderBy=player&orderByDescending=false'
standard_info = '&skip=-1&language=de'
quary_basic = '&searchQuery=1//1&filterQuery=' 
##########################################################################
# FILTERS                                                                #
##########################################################################
# season 
current_season = '2026'
# phase
phase_regular = '/4940'
# Teams 
all_teams = '/all'
biel = '/102128'
ehc_kloten = '/101149'
ev_zug = '/101144'
fribourg_gotteron = '103138'
genf_servette = '/103140'
ajoie = '/103144'
ambri_piotta = '/101152'
davos = '/101151'
lugano = '/101150'
lausanne = '103141'
sc_bern = '/102126'
scl_tigers = '/102127'
rappi_lakers = '/101060'
zsc_lions = '/101139'
# positions
all_positions = '/all'
verteidiger = '/2'
stuermer = '/3'
# licence 
all_licences = '/all'
swiss_licence = '/1'
foreign_licence = '/2'
# more standard stuff
filtering_info = '&filterBy=Season,Phase,Team,Position,Licence'
standard_ending = "&take=1000&callback=externalStatisticsCallback"
##########################################################################
# excel stuff
##########################################################################
current_date = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
my_storagepath = "C:/Users/Jessi/Documents/Hockey_stats/Stats_from_T"
file_ending = ".xlsx"
my_excelpath = my_storagepath + current_date + file_ending 
# create the excel file
my_excelfile = xlsxwriter.Workbook(my_excelpath)
##########################################################################
# tha functions
##########################################################################
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
def write_titels_to_worksheet_players(my_worksheet: xlsxwriter.worksheet, start_row: int) -> None:
    my_worksheet.write(start_row, 0, 'Player')
    my_worksheet.write(start_row, 1, 'GP')
    my_worksheet.write(start_row, 2, 'Goals')
    my_worksheet.write(start_row, 3, 'Assists')
    my_worksheet.write(start_row, 4, 'Shots on Goal')
    my_worksheet.write(start_row, 5, 'Schüsse aufs Tor aus Slot')
    my_worksheet.write(start_row, 6, '+/-')
    my_worksheet.write(start_row, 7, 'Bullys Total')
    my_worksheet.write(start_row, 8, 'Gewonnene Bullys')
    my_worksheet.write(start_row, 9, 'Blocked Shots')
    my_worksheet.write(start_row, 10, 'Shorthanders')
    my_worksheet.write(start_row, 11, 'Powerplay Goals')
    my_worksheet.write(start_row, 12, 'Penalty Minutes')

#
# create the request URLs and get the data 
#
def get_the_data():
  request_url = (base_url + summary_stat + standard_info + quary_basic 
               + current_season + phase_regular + all_teams + all_positions 
               + all_licences + filtering_info + standard_sorting + standard_ending)
  general_data = send_request(request_url)
  request_url = (base_url + goal_stats + standard_info +quary_basic 
               + current_season + phase_regular + all_teams 
               + all_positions + all_licences + filtering_info + standard_sorting + standard_ending)
  goal_data = send_request(request_url)
  request_url = (base_url + shot_stat + standard_info +quary_basic 
               + current_season + phase_regular + all_teams + all_positions 
               + all_licences + filtering_info + standard_sorting + standard_ending)
  shot_data = send_request(request_url)
  request_url = (base_url + faceoff_stat + standard_info +quary_basic 
               + current_season + phase_regular + all_teams + all_positions 
               + all_licences + filtering_info + standard_sorting + standard_ending)
  faceoff_data = send_request(request_url)
  request_url = (base_url + goalie_summary_stat + standard_info +quary_basic 
               + current_season + phase_regular + all_teams + all_positions 
               + all_licences + filtering_info + standard_sorting + standard_ending)
  #goalie_data = send_request(request_url)
  request_url = (base_url + goalie_shot_stat + standard_info +quary_basic + current_season 
               + phase_regular + all_teams + all_positions + all_licences + filtering_info 
               + standard_sorting + standard_ending)
  #goalie_shot_data = send_request(request_url)
  return general_data, goal_data, shot_data, faceoff_data
#
# take the gathered data and sort them into dicts per team 
#
def create_team_data():
    # get the data from the API
    general_data, goal_data, shot_data, faceoff_data = get_the_data()
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
            temp_key = i[1]
            # create a temporary dict to store the player data in a strructured way which is easy to extend
            # by additional fields later on as well retreval by key is possible (aka search) which makes the 
            # gathering and adding of data easier
            temp_dict = {'Name': i[1]
                              # access the team name from the dictionary which we extracted before
                            , 'Team': team_raw['name']
                            , 'Position': i[3]
                            , 'Games Played': i[4]
                            , 'Goals': i[5]
                            , 'Assists': i[6]
                            , 'PIM': i[9]
                            , '+/-': i[10]}
            ehc_biel_data[temp_key] = temp_dict
        # Kloten
        case 101149:
            temp_key = i[1]
            temp_dict ={'Name': i[1]
                            , 'Team': team_raw['name'] 
                            , 'Position': i[3]
                            , 'Games Played': i[4]
                            , 'Goals': i[5]
                            , 'Assists': i[6]
                            , 'PIM': i[9]
                            , '+/-': i[10]}
            ehc_kloten_data[temp_key] = temp_dict
        # Zug
        case 101144:
            temp_key = i[1]
            temp_dict = {'Name': i[1]
                            , 'Team': team_raw['name'] 
                            , 'Position': i[3]
                            , 'Games Played': i[4]
                            , 'Goals': i[5]
                            , 'Assists': i[6]
                            , 'PIM': i[9]
                            , '+/-': i[10]}
            ev_zug_data[temp_key] = temp_dict
        # Fribourg  
        case 103138:
            temp_key = i[1]
            temp_dict = {'Name': i[1]  
                            , 'Team': team_raw['name']
                            , 'Position': i[3]
                            , 'Games Played': i[4]
                            , 'Goals': i[5]
                            , 'Assists': i[6]
                            , 'PIM': i[9]
                            , '+/-': i[10]}
            fribourg_gotteron_data[temp_key] = temp_dict
        # Genf
        case 103140:
            temp_key = i[1]
            temp_dict = {'Name': i[1]
                            , 'Team': team_raw['name']
                            , 'Position': i[3]
                            , 'Games Played': i[4]
                            , 'Goals': i[5]
                            , 'Assists': i[6]
                            , 'PIM': i[9]
                            , '+/-': i[10]}
            genf_servette_data[temp_key] = temp_dict
        # Ajoie
        case 103144:
            temp_key = i[1]
            temp_dict = {'Name': i[1]
                            , 'Team': team_raw['name']
                            , 'Position': i[3]
                            , 'Games Played': i[4]
                            , 'Goals': i[5]
                            , 'Assists': i[6]
                            , 'PIM': i[9]
                            , '+/-': i[10]}
            hc_ajoie_data[temp_key] = temp_dict
        # Ambri
        case 101152:
            temp_key = i[1]
            temp_dict = {'Name': i[1]
                            , 'Team': team_raw['name']
                            , 'Position': i[3]
                            , 'Games Played': i[4]
                            , 'Goals': i[5]
                            , 'Assists': i[6]
                            , 'PIM': i[9]
                            , '+/-': i[10]}
            hc_ambri_piotta_data[temp_key] = temp_dict
        # Davos
        case 101151:
            temp_key = i[1]
            temp_dict = {'Name': i[1]
                            , 'Team': team_raw['name']
                            , 'Position': i[3]
                            , 'Games Played': i[4]
                            , 'Goals': i[5]
                            , 'Assists': i[6]
                            , 'PIM': i[9]
                            , '+/-': i[10]}
            hc_davos_data[temp_key] = temp_dict
        # Lugano
        case 101150:
            temp_key = i[1]
            temp_dict = {'Name': i[1]
                            , 'Team': team_raw['name']
                            , 'Position': i[3]
                            , 'Games Played': i[4]
                            , 'Goals': i[5]
                            , 'Assists': i[6]
                            , 'PIM': i[9]
                            , '+/-': i[10]}
            hc_lugano_data[temp_key] = temp_dict
        # Lausanne
        case 103141:
            temp_key = i[1]
            temp_dict ={'Name': i[1]
                            , 'Team': team_raw['name']
                            , 'Position': i[3]  
                            , 'Games Played': i[4]
                            , 'Goals': i[5]
                            , 'Assists': i[6]
                            , 'PIM': i[9]
                            , '+/-': i[10]}
            hc_lausanne_data[temp_key] = temp_dict
        # Bern
        case 102126:
            temp_key = i[1]
            temp_dict = {'Name': i[1]
                            , 'Team': team_raw['name']
                            , 'Position': i[3]  
                            , 'Games Played': i[4]
                            , 'Goals': i[5]
                            , 'Assists': i[6]
                            , 'PIM': i[9]
                            , '+/-': i[10]}
            sc_bern_data[temp_key] = temp_dict
        # Tigers
        case 102127:
            temp_key = i[1]
            temp_dict = {'Name': i[1]
                            , 'Team': team_raw['name']  
                            , 'Position': i[3]
                            , 'Games Played': i[4]
                            , 'Goals': i[5]
                            , 'Assists': i[6]
                            , 'PIM': i[9]
                            , '+/-': i[10]}
            scl_tigers_data[temp_key] = temp_dict
        # Rappi
        case 101060:
            temp_key = i[1]
            temp_dict ={'Name': i[1]
                            , 'Team': team_raw['name']
                            , 'Position': i[3]
                            , 'Games Played': i[4]
                            , 'Goals': i[5]
                            , 'Assists': i[6]
                            , 'PIM': i[9]   
                            , '+/-': i[10]}
            rapperswil_jona_lakers_data[temp_key] = temp_dict
        # ZSC
        case 101139:
            temp_key = i[1]
            temp_dict = { 'Name': i[1]
                            , 'Team': team_raw['name']
                            , 'Position': i[3]  
                            , 'Games Played': i[4]
                            , 'Goals': i[5]
                            , 'Assists': i[6]
                            , 'PIM': i[9]
                            , '+/-': i[10]}
            zsc_lions_data[temp_key] = temp_dict
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
                # we then update the said sub dict with additional key value pairs 
                ehc_biel_data[i[1]].update({'Assists': i[6], 'PP_Goals': i[12], 'PP_Assists': i[13]
                                            , 'Box_Play_Goals': i[14], 'Box_Play_Assists': i[15]})
            # Kloten
            case 101149:
                ehc_kloten_data[i[1]].update({'Assists': i[6], 'PP_Goals': i[12], 'PP_Assists': i[13]
                                            , 'Box_Play_Goals': i[14], 'Box_Play_Assists': i[15]})
            # Zug
            case 101144:
                ev_zug_data[i[1]].update({'Assists': i[6], 'PP_Goals': i[12], 'PP_Assists': i[13]
                                            , 'Box_Play_Goals': i[14], 'Box_Play_Assists': i[15]})
            # Fribourg
            case 103138:
                fribourg_gotteron_data[i[1]].update({'Assists': i[6], 'PP_Goals': i[12], 'PP_Assists': i[13]
                                            , 'Box_Play_Goals': i[14], 'Box_Play_Assists': i[15]})
            # Genf
            case 103140:
                genf_servette_data[i[1]].update({'Assists': i[6], 'PP_Goals': i[12], 'PP_Assists': i[13]
                                            , 'Box_Play_Goals': i[14], 'Box_Play_Assists': i[15]})
            # Ajoie
            case 103144:
                hc_ajoie_data[i[1]].update({'Assists': i[6], 'PP_Goals': i[12], 'PP_Assists': i[13]
                                            , 'Box_Play_Goals': i[14], 'Box_Play_Assists': i[15]})
            # Ambri
            case 101152:
                hc_ambri_piotta_data[i[1]].update({'Assists': i[6], 'PP_Goals': i[12], 'PP_Assists': i[13]
                                            , 'Box_Play_Goals': i[14], 'Box_Play_Assists': i[15]})
            # Davos
            case 101151:
                hc_davos_data[i[1]].update({'Assists': i[6], 'PP_Goals': i[12], 'PP_Assists': i[13]
                                            , 'Box_Play_Goals': i[14], 'Box_Play_Assists': i[15]})
            # Lugano
            case 101150:
                hc_lugano_data[i[1]].update({'Assists': i[6], 'PP_Goals': i[12], 'PP_Assists': i[13]
                                            , 'Box_Play_Goals': i[14], 'Box_Play_Assists': i[15]})
            # Lausanne
            case 103141:
                hc_lausanne_data[i[1]].update({'Assists': i[6], 'PP_Goals': i[12], 'PP_Assists': i[13]
                                            , 'Box_Play_Goals': i[14], 'Box_Play_Assists': i[15]})
            # Bern
            case 102126:
                sc_bern_data[i[1]].update({'Assists': i[6], 'PP_Goals': i[12], 'PP_Assists': i[13]
                                            , 'Box_Play_Goals': i[14], 'Box_Play_Assists': i[15]})
            # Tigers
            case 102127:
                scl_tigers_data[i[1]].update({'Assists': i[6], 'PP_Goals': i[12], 'PP_Assists': i[13]
                                            , 'Box_Play_Goals': i[14], 'Box_Play_Assists': i[15]})
            # Rappi
            case 101060:
                rapperswil_jona_lakers_data[i[1]].update({'Assists': i[6], 'PP_Goals': i[12], 'PP_Assists': i[13]
                                            , 'Box_Play_Goals': i[14], 'Box_Play_Assists': i[15]})
            # ZSC
            case 101139:
                zsc_lions_data[i[1]].update({'Assists': i[6], 'PP_Goals': i[12], 'PP_Assists': i[13]
                                            , 'Box_Play_Goals': i[14], 'Box_Play_Assists': i[15]})
    # now the same for the shot data 
    for i in shot_data['data']:
        team_id = i[2]['id']
        match team_id:
            # Biel
            case 102128:
                ehc_biel_data[i[1]].update({'Shots_total': i[5], 'Shots_from_slot': i[8], 'Shots_missed': i[10]
                                            , 'Shots_on_border': i[12], 'Blocked_shots': i[14]})
            # Kloten
            case 101149:
                ehc_kloten_data[i[1]].update({'Shots_total': i[5], 'Shots_from_slot': i[8], 'Shots_missed': i[10]
                                            , 'Shots_on_border': i[12], 'Blocked_shots': i[14]})
            # Zug
            case 101144:
                ev_zug_data[i[1]].update({'Shots_total': i[5], 'Shots_from_slot': i[8], 'Shots_missed': i[10]
                                            , 'Shots_on_border': i[12], 'Blocked_shots': i[14]})
            # Fribourg
            case 103138:
                fribourg_gotteron_data[i[1]].update({'Shots_total': i[5], 'Shots_from_slot': i[8], 'Shots_missed': i[10]
                                            , 'Shots_on_border': i[12], 'Blocked_shots': i[14]})
            # Genf
            case 103140:
                genf_servette_data[i[1]].update({'Shots_total': i[5], 'Shots_from_slot': i[8], 'Shots_missed': i[10]
                                            , 'Shots_on_border': i[12], 'Blocked_shots': i[14]})
            # Ajoie
            case 103144:
                hc_ajoie_data[i[1]].update({'Shots_total': i[5], 'Shots_from_slot': i[8], 'Shots_missed': i[10]
                                            , 'Shots_on_border': i[12], 'Blocked_shots': i[14]})
            # Ambri
            case 101152:
                hc_ambri_piotta_data[i[1]].update({'Shots_total': i[5], 'Shots_from_slot': i[8], 'Shots_missed': i[10]
                                            , 'Shots_on_border': i[12], 'Blocked_shots': i[14]})
            # Davos
            case 101151:
                hc_davos_data[i[1]].update({'Shots_total': i[5], 'Shots_from_slot': i[8], 'Shots_missed': i[10]
                                            , 'Shots_on_border': i[12], 'Blocked_shots': i[14]})
            # Lugano
            case 101150:
                hc_lugano_data[i[1]].update({'Shots_total': i[5], 'Shots_from_slot': i[8], 'Shots_missed': i[10]
                                            , 'Shots_on_border': i[12], 'Blocked_shots': i[14]})
            # Lausanne
            case 103141:
                hc_lausanne_data[i[1]].update({'Shots_total': i[5], 'Shots_from_slot': i[8], 'Shots_missed': i[10]
                                            , 'Shots_on_border': i[12], 'Blocked_shots': i[14]})
            # Bern
            case 102126:
                sc_bern_data[i[1]].update({'Shots_total': i[5], 'Shots_from_slot': i[8], 'Shots_missed': i[10]
                                            , 'Shots_on_border': i[12], 'Blocked_shots': i[14]})
            # Tigers
            case 102127:
                scl_tigers_data[i[1]].update({'Shots_total': i[5], 'Shots_from_slot': i[8], 'Shots_missed': i[10]
                                            , 'Shots_on_border': i[12], 'Blocked_shots': i[14]})
            # Rappi
            case 101060:
                rapperswil_jona_lakers_data[i[1]].update({'Shots_total': i[5], 'Shots_from_slot': i[8], 'Shots_missed': i[10]
                                            , 'Shots_on_border': i[12], 'Blocked_shots': i[14]})
            # ZSC
            case 101139:
                zsc_lions_data[i[1]].update({'Shots_total': i[5], 'Shots_from_slot': i[8], 'Shots_missed': i[10]
                                            , 'Shots_on_border': i[12], 'Blocked_shots': i[14]})
    # now again the same for the faceoff data
    print(faceoff_data)

    # return the team dicts 
    return (ehc_biel_data, ehc_kloten_data, ev_zug_data, fribourg_gotteron_data, genf_servette_data,
            hc_ajoie_data, hc_ambri_piotta_data, hc_davos_data, hc_lugano_data, hc_lausanne_data,
            sc_bern_data, scl_tigers_data, rapperswil_jona_lakers_data, zsc_lions_data)
#
#
#
def main():
    (ehc_biel_data, ehc_kloten_data, ev_zug_data, fribourg_gotteron_data, genf_servette_data,
     hc_ajoie_data, hc_ambri_piotta_data, hc_davos_data, hc_lugano_data, hc_lausanne_data,
     sc_bern_data, scl_tigers_data, rapperswil_jona_lakers_data, zsc_lions_data) = create_team_data()
    # Further processing and writing to Excel can be done here

    my_excelfile.close()
#
# ZSC Lions Stürmer
#
main()
zsc_stuermer = add_worksheet(my_excelfile, 'ZSC Stürmer')
start_row = 0
write_titels_to_worksheet_players(zsc_stuermer, start_row)

        #zsc_stuermer.write(start_row, 0, i['playerName'])
        #zsc_stuermer.write(start_row, 1, i['gamesPlayed'])
# zsc_stuermer_data = {general_data['4940'] | 
#                      goal_data['data']['playerGoalAssistStats']['4940']['101139']['3'] | 
#                      shot_data['data']['playerShotDetailStats']['4940']['101139']['3'] | 
#                      faceoff_data['data']['playerFaceoffStats']['4940']['101139']['3']}
# for player_id, player_stats in zsc_stuermer_data.items():
#     start_row += 1
#     zsc_stuermer.write(start_row, 0, player_stats['playerName'])
#     zsc_stuermer.write(start_row, 1, player_stats['gamesPlayed'])
#     zsc_stuermer.write(start_row, 2, player_stats['goals'])
#     zsc_stuermer.write(start_row, 3, player_stats['assists'])
#     zsc_stuermer.write(start_row, 4, player_stats['shotsOnGoal'])
#     zsc_stuermer.write(start_row, 5, player_stats['shotsOnGoalFromSlot'])
#     zsc_stuermer.write(start_row, 6, player_stats['plusMinus'])
#     zsc_stuermer.write(start_row, 7, player_stats['faceoffsTotal'])
#     zsc_stuermer.write(start_row, 8, player_stats['faceoffsWon'])
#     zsc_stuermer.write(start_row, 9, player_stats['blockedShots'])
#     zsc_stuermer.write(start_row, 10, player_stats['shorthandedGoals'])
#     zsc_stuermer.write(start_row, 11, player_stats['powerplayGoals'])
#     zsc_stuermer.write(start_row, 12, player_stats['penaltyMinutes'])


def main():
    (ehc_biel_data, ehc_kloten_data, ev_zug_data, fribourg_gotteron_data, genf_servette_data,
     hc_ajoie_data, hc_ambri_piotta_data, hc_davos_data, hc_lugano_data, hc_lausanne_data,
     sc_bern_data, scl_tigers_data, rapperswil_jona_lakers_data, zsc_lions_data) = create_team_data()
    print(zsc_lions_data)
    # Further processing and writing to Excel can be done here

    my_excelfile.close()