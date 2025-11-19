###############################################################################
# the imports
###############################################################################
import datetime
import requests
import json
import re
import warnings
###############################################################################
# the objects 
###############################################################################
# define first the constant dictionary for league translation 
# placed for now and will come later
LEAGUES = {
   'national_league'  :   1,
   'swiss_league'     :   2, 
   'my_hockey_league' : 100,
   'pf_women_league'  :  42,
   'swhl_b'           :  43,
   'swhl_c'           : 101,
   'swhl_d'           : 104,
   'erst_liga'        :   3,
   'zweit_liga'       :  10,
   'dritt_liga'       :  18,
   'viert_liga'       :  19
}
class player_filters: 
   def __init__(self, season, phase, team="all", licence="all", position="all"):
      self.season   = season
      self.phase    = phase 
      self.team     = team
      self.licence  = licence 
      self.position = position
###############################################################################
# the function used to identify what the current season is
###############################################################################
def get_current_season() -> str:
    # format of datetime.now() is YYYY-MM-DD HH:MM:SS.hs
    # since we are anyway only interested in the year and month
    # we get the current date, transform it the way we get it into a string 
    # and use the .split() method to concatinate it given the delimitter -
    # which leaves us with a list where entry 0 contains the year
    # entry 1 contains the month and entry 2 day and time 
    current_date = str(datetime.datetime.now()).split('-')
    # take the two attributes we are interessted in and transform them into 
    # integers, since we are going to perform some maths with them 
    season_year = int(current_date[0])
    current_month = int(current_date[1])
    if current_month >= 9:
        season_year += 1
    # we are kinda lazy and do not even allocate a seperate variable for the
    # return statement since this will be the same outcome but with more MIPs
    # and memory used #COBOLThings
    return str(season_year)
###############################################################################
# the variables
###############################################################################
# contsants 
# swiss icehockey statistic and data base URL
BASE_URL = 'https://data.sihf.ch/Statistic/api/cms/cache'
# stats
STAT_CACHE = '300?'
# kind of stat
# player and team stats
# field player stats
SUMMARY_STAT                    = 'alias=player'
GOAL_STATS                      = 'alias=playerGoalAssist'
SHOT_STAT                       = 'alias=playerShotDetail'
PENALTY_STAT                    = 'alias=playerFoul'
SHOOTOUT_STAT                   = 'alias=playerShootout'
FACEOFF_SUMMARY_STAT            = 'alias=playerFaceoff'
FACEOFF_ZONE_STAT               = 'playerFaceoffZone'
FACEOFF_ZONE_PER_GAME_STAT      = 'alias=playerFaceoffZoneGame'
TIME_ON_ICE_STAT                = 'alias=playerTimeOnIce'
# goalie stats 
GOALIE_SUMMARY_STAT             = 'alias=goalkeeper'
GOALIE_SHOT_STAT                = 'alias=goalkeeperShotDetail'
GOALIE_WIN_STAT                 = 'alias=goalkeeperWin'
GOALIE_SHOOTOUT_STAT            = 'alias=goalkeeperShootout'
# team stats 
TEAM_GOAL_SUMMARY_STAT          = 'alias=teamGoal'
TEAM_POSITION_GOAL_STAT         = 'alias=teamGoalPosNation'
TEAM_SHOT_STAT                  = 'alias=teamShotDetail'
TEAM_PP_STAT                    = 'alias=teamPp'
TEAM_PK_STAT                    = 'alias=teamPk'
TEAM_PENALTY_STAT               = 'alias=teamFoul'
TEAM_SHOOTOUT_STAT              = 'alias=teamShootout'
TEAM_FACEOFF_SUMMARY_STAT       = 'alias=teamFaceoff'
TEAM_FACEOFF_ZONE_STAT          = 'alias=teamFaceoffZone'
TEAM_FACEOFF_ZONE_PER_GAME_STAT = 'alias=teamFaceoffZoneGame'
TEAM_ATTENDANCE_STAT            = 'alias=teamSpectator'
# data
DATA_CACHE = '600?'
# postfinance Top Scorer
TOPSCORER_DATA = 'alias=pftopscorer'
# results
RESULTS = 'alias=results'
# filters for the stats
# season (only stat filter we will use, as there will be an initial call 
# to get the possible filter options) 
CURRENT_SEASON   = get_current_season()
# search defnition for stats
SEARCHQUERY = '&searchQuery='
FILTERQUERY = '&filterQuery='
FILTERBY = '&filterBy='
#
FILTER_NAMES = 'Season,Phase,Team,Position,Licence'
# query ending
RECORDS_TO_BE_RETURNED = '&take='
STANDARD_ENDING = '&callback=externalStatisticsCallback&skip=-1&language=de'

def send_request(request_url: str) -> dict: 
    try:
     # make a GET request to an API endpoint
      response = requests.get(request_url)
    except: 
        raise Exception('Could not reach the API endpoint for the url: ' + request_url)
    try:
     # use regex to extract the JSON and remove the JS stuff
     # meaning we search for externalStatisticsCallback( some stuff ); and we take what is inbetween
      response_without_JS = re.search(r'externalStatisticsCallback\((.*)\);', response.text).group(1)
      # place the JSON into a dcitionary
      json_data = json.loads(response_without_JS)
      return json_data
    except:
        raise Exception('Could not parse the response into JSON for the url: ' + request_url)
    
def get_filter_options (season: str):
    """
    If no season or a string with more than 4 characters gets provided it will
    be replaced with the current season. 
    A request to the swiss icehockey API gets sent and all possible filters
    (eg. season, teams, phases, positions, etc.) will be added as dictionary
    within a dictionary. The 'alias' values are the required ones for further
    processing. 

    Args: 
        season (str): A 4 digit year in string format 

    Return: 
        filter_options: A dictionary with n entries where the values are again a 
        dictionary
    """
    # at first we are going to build a request URL which we then will send.
    # since we are only interessted in the filter options and not in the actuall 
    # data we will keep the query short and ensure as little as possible will be 
    # returned
    if len(season) == 0 or len(season) > 4: 
        warnings.warn('Changed provided season to current season')
        season = CURRENT_SEASON
    # since we do not want to do a lot of input validation we just try to send 
    # the request with the input data, and it is not possible, we use the current
    # season 
    try:
       request_url = (BASE_URL + STAT_CACHE + SUMMARY_STAT + SEARCHQUERY + '1//1'
                    + FILTERQUERY + season + FILTERBY + FILTER_NAMES + RECORDS_TO_BE_RETURNED 
                    + '1' + STANDARD_ENDING)

       raw_data = send_request(request_url)
    except: 
       warnings.warn('Something went wrong / No data was present the provided season. ' \
                     'received data will be based of the current season')
       request_url = (BASE_URL + STAT_CACHE + SUMMARY_STAT + SEARCHQUERY + '1//1'
                    + FILTERQUERY + CURRENT_SEASON + FILTERBY + FILTER_NAMES 
                    + RECORDS_TO_BE_RETURNED + '1' + STANDARD_ENDING)

       raw_data = send_request(request_url) 
    # predefine an empty dict (yet another COBOL thing)
    filter_options = {}
    if (raw_data['filters'][0]['selected'] != season):
        warnings.warn('Something went wrong / No data was present the provided season. Received data will be based of the current season')
    try: 
        raw_data['filters'][2]['title'] != 'Team'
    except:
        # no teams yet present so we need to send another request
        # but we should have received some filters, if not we crash
        temp_phase = ''
        for i in (raw_data['filters'][1]['entries']):
          if i['name'] == 'Regular Season':
             temp_phase = i['alias']
             break
        if len(temp_phase) == 0: 
           raise Exception('No full filter list could be loaded for your selected season: ' + season)
        request_url =(BASE_URL + STAT_CACHE + SUMMARY_STAT + SEARCHQUERY + '1//1'
                    + FILTERQUERY + season + '/' + temp_phase + FILTERBY + FILTER_NAMES 
                    + RECORDS_TO_BE_RETURNED + '1' + STANDARD_ENDING)
        raw_data.clear
        raw_data = send_request(request_url)
    # add all options for filtering to the dictionary
    for i in (raw_data['filters']):
        filter_options[i['title']] = i['entries']


    return filter_options

###############################################################################
# field player stats 
###############################################################################

def get_summary_player_stats (filters: player_filters) -> list:
   # Input validation for the fields where no standard value is present
   if len(filters.phase) == 0: 
      raise Exception('Phase is mandatory for the summary stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the summary stat')
   
   # creation of the request URL with the help of the constants and the 
   # filter class where defaults are set (and where this is not possible
   # the input was already checked)
   request_url =(BASE_URL + STAT_CACHE + SUMMARY_STAT + SEARCHQUERY + '1//1'
               + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
               + filters.team + '/' + filters.position + '/' + filters.licence
               + FILTERBY + FILTER_NAMES + RECORDS_TO_BE_RETURNED + '1000' 
               + STANDARD_ENDING)
   # get the whole data (including page headers of the official website, 
   # all possible filters and the actual data)
   raw_data = send_request(request_url)

   # only return the actual data since the rest is obsolet for us 
   summary_data  = raw_data['data']
   summary_data.append(raw_data['header'])
   # at this state summary_data is a list with 0_n dictionaries as elements 
   return summary_data

def get_player_goal_stats (filters: player_filters) -> list: 
   # Input validation for the fields where no standard value is present
   if len(filters.phase) == 0: 
      raise Exception('Phase is mandatory for the goal/assist stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the goal/assist stat')
   
   # creation of the request URL with the help of the constants and the 
   # filter class where defaults are set (and where this is not possible
   # the input was already checked)
   request_url =(BASE_URL + STAT_CACHE + GOAL_STATS + SEARCHQUERY + '1//1'
               + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
               + filters.team + '/' + filters.position + '/' + filters.licence
               + FILTERBY + FILTER_NAMES + RECORDS_TO_BE_RETURNED + '1000' 
               + STANDARD_ENDING)
   # get the whole data (including page headers of the official website, 
   # all possible filters and the actual data)
   raw_data = send_request(request_url)

   # only return the actual data since the rest is obsolet for us 
   goal_data  = raw_data['data']
   goal_data.append(raw_data['header'])
   # at this state goal_data is a list with 0_n dictionaries as elements 
   return goal_data

def get_player_shot_stats (filters: player_filters) -> list: 
   # Input validation for the fields where no standard value is present
   if len(filters.phase) == 0: 
      raise Exception('Phase is mandatory for the shot on goal stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the shot on goal stat')
   
   # creation of the request URL with the help of the constants and the 
   # filter class where defaults are set (and where this is not possible
   # the input was already checked)
   request_url =(BASE_URL + STAT_CACHE + SHOT_STAT + SEARCHQUERY + '1//1'
               + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
               + filters.team + '/' + filters.position + '/' + filters.licence
               + FILTERBY + FILTER_NAMES + RECORDS_TO_BE_RETURNED + '1000' 
               + STANDARD_ENDING)
   # get the whole data (including page headers of the official website, 
   # all possible filters and the actual data)
   raw_data = send_request(request_url)

   # only return the actual data since the rest is obsolet for us 
   shot_data  = raw_data['data']
   shot_data.append(raw_data['header'])
   # at this state shot_data is a list with 0_n dictionaries as elements 
   return shot_data

def get_player_shootout_stats (filters: player_filters) -> list: 
   # Input validation for the fields where no standard value is present
   if len(filters.phase) == 0: 
      raise Exception('Phase is mandatory for the shootout stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the shootout stat')
   
   # creation of the request URL with the help of the constants and the 
   # filter class where defaults are set (and where this is not possible
   # the input was already checked)
   request_url =(BASE_URL + STAT_CACHE + SHOOTOUT_STAT + SEARCHQUERY + '1//1'
               + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
               + filters.team + '/' + filters.position + '/' + filters.licence
               + FILTERBY + FILTER_NAMES + RECORDS_TO_BE_RETURNED + '1000' 
               + STANDARD_ENDING)
   # get the whole data (including page headers of the official website, 
   # all possible filters and the actual data)
   raw_data = send_request(request_url)

   # only return the actual data since the rest is obsolet for us 
   shootout_data  = raw_data['data']
   shootout_data.append(raw_data['header'])
   # at this state shootout_data is a list with 0_n dictionaries as elements 
   return shootout_data

def get_player_faceoff_summary_stats (filters: player_filters) -> list: 
   # Input validation for the fields where no standard value is present
   if len(filters.phase) == 0: 
      raise Exception('Phase is mandatory for the faceoff stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the faceoff stat')
   
   # creation of the request URL with the help of the constants and the 
   # filter class where defaults are set (and where this is not possible
   # the input was already checked)
   request_url =(BASE_URL + STAT_CACHE + FACEOFF_SUMMARY_STAT + SEARCHQUERY + '1//1'
               + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
               + filters.team + '/' + filters.position + '/' + filters.licence
               + FILTERBY + FILTER_NAMES + RECORDS_TO_BE_RETURNED + '1000' 
               + STANDARD_ENDING)
   # get the whole data (including page headers of the official website, 
   # all possible filters and the actual data)
   raw_data = send_request(request_url)

   # only return the actual data since the rest is obsolet for us 
   faceoff_data  = raw_data['data']
   faceoff_data.append(raw_data['header'])
   # at this state faceoff_data is a list with 0_n dictionaries as elements 
   return faceoff_data

def get_player_faceoff_zone_stats (filters: player_filters) -> list: 
   # Input validation for the fields where no standard value is present
   if len(filters.phase) == 0: 
      raise Exception('Phase is mandatory for the faceoff zone stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the faceoff zone stat')
   
   # creation of the request URL with the help of the constants and the 
   # filter class where defaults are set (and where this is not possible
   # the input was already checked)
   request_url =(BASE_URL + STAT_CACHE + FACEOFF_ZONE_STAT + SEARCHQUERY + '1//1'
               + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
               + filters.team + '/' + filters.position + '/' + filters.licence
               + FILTERBY + FILTER_NAMES + RECORDS_TO_BE_RETURNED + '1000' 
               + STANDARD_ENDING)
   # get the whole data (including page headers of the official website, 
   # all possible filters and the actual data)
   raw_data = send_request(request_url)

   # only return the actual data since the rest is obsolet for us 
   faceoff_data  = raw_data['data']
   faceoff_data.append(raw_data['header'])
   # at this state faceoff_data is a list with 0_n dictionaries as elements 
   return faceoff_data

def get_player_faceoff_zone_pg_stats (filters: player_filters) -> list: 
   # Input validation for the fields where no standard value is present
   if len(filters.phase) == 0: 
      raise Exception('Phase is mandatory for the faceoff zone per game stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the faceoff zone per game stat')
   
   # creation of the request URL with the help of the constants and the 
   # filter class where defaults are set (and where this is not possible
   # the input was already checked)
   request_url =(BASE_URL + STAT_CACHE + FACEOFF_ZONE_PER_GAME_STAT + SEARCHQUERY + '1//1'
               + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
               + filters.team + '/' + filters.position + '/' + filters.licence
               + FILTERBY + FILTER_NAMES + RECORDS_TO_BE_RETURNED + '1000' 
               + STANDARD_ENDING)
   # get the whole data (including page headers of the official website, 
   # all possible filters and the actual data)
   raw_data = send_request(request_url)

   # only return the actual data since the rest is obsolet for us 
   faceoff_data  = raw_data['data']
   faceoff_data.append(raw_data['header'])
   # at this state faceoff_data is a list with 0_n dictionaries as elements 
   return faceoff_data

def get_player_time_on_ice_stats (filters: player_filters) -> list: 
   # Input validation for the fields where no standard value is present
   if len(filters.phase) == 0: 
      raise Exception('Phase is mandatory for the summary stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the summary stat')
   
   # creation of the request URL with the help of the constants and the 
   # filter class where defaults are set (and where this is not possible
   # the input was already checked)
   request_url =(BASE_URL + STAT_CACHE + TIME_ON_ICE_STAT + SEARCHQUERY + '1//1'
               + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
               + filters.team + '/' + filters.position + '/' + filters.licence
               + FILTERBY + FILTER_NAMES + RECORDS_TO_BE_RETURNED + '1000' 
               + STANDARD_ENDING)
   # get the whole data (including page headers of the official website, 
   # all possible filters and the actual data)
   raw_data = send_request(request_url)

   # only return the actual data since the rest is obsolet for us 
   toi_data  = raw_data['data']
   toi_data.append(raw_data['header'])
   # at this state toi_data is a list with 0_n dictionaries as elements 
   return toi_data

###############################################################################
# Goalkeeper stats 
###############################################################################

def get_goalie_summary_stats (filters: player_filters) -> list: 
   # Input validation for the fields where no standard value is present
   if len(filters.phase) == 0: 
      raise Exception('Phase is mandatory for the goalie summary stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the goalie summary stat')
   
   # creation of the request URL with the help of the constants and the 
   # filter class where defaults are set (and where this is not possible
   # the input was already checked)
   request_url =(BASE_URL + STAT_CACHE + GOALIE_SUMMARY_STAT + SEARCHQUERY + '1//1'
               + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
               + filters.team + '/' + filters.licence + FILTERBY + FILTER_NAMES 
               + RECORDS_TO_BE_RETURNED + '1000' + STANDARD_ENDING)
   # get the whole data (including page headers of the official website, 
   # all possible filters and the actual data)
   raw_data = send_request(request_url)

   # only return the actual data since the rest is obsolet for us 
   goalie_data  = raw_data['data']
   goalie_data.append(raw_data['header'])
   # at this state goalie_data is a list with 0_n dictionaries as elements 
   return goalie_data

def get_goalie_shots_against_stats (filters: player_filters) -> list: 
   # Input validation for the fields where no standard value is present
   if len(filters.phase) == 0: 
      raise Exception('Phase is mandatory for the goalie shots against stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the goalie shots against stat')
   
   # creation of the request URL with the help of the constants and the 
   # filter class where defaults are set (and where this is not possible
   # the input was already checked)
   request_url =(BASE_URL + STAT_CACHE + GOALIE_SHOT_STAT + SEARCHQUERY + '1//1'
               + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
               + filters.team + '/' + filters.licence + FILTERBY + FILTER_NAMES 
               + RECORDS_TO_BE_RETURNED + '1000' + STANDARD_ENDING)
   # get the whole data (including page headers of the official website, 
   # all possible filters and the actual data)
   raw_data = send_request(request_url)

   # only return the actual data since the rest is obsolet for us 
   goalie_shot_data  = raw_data['data']
   goalie_shot_data.append(raw_data['header'])
   # at this state goalie_shot_data is a list with 0_n dictionaries as elements 
   return goalie_shot_data

def get_goalie_win_stats (filters: player_filters) -> list: 
   # Input validation for the fields where no standard value is present
   if len(filters.phase) == 0: 
      raise Exception('Phase is mandatory for the goalie win stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the goalie win stat')
   
   # creation of the request URL with the help of the constants and the 
   # filter class where defaults are set (and where this is not possible
   # the input was already checked)
   request_url =(BASE_URL + STAT_CACHE + GOALIE_WIN_STAT + SEARCHQUERY + '1//1'
               + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
               + filters.team + '/' + filters.licence + FILTERBY + FILTER_NAMES 
               + RECORDS_TO_BE_RETURNED + '1000' + STANDARD_ENDING)
   # get the whole data (including page headers of the official website, 
   # all possible filters and the actual data)
   raw_data = send_request(request_url)

   # only return the actual data since the rest is obsolet for us 
   goalie_win_data  = raw_data['data']
   goalie_win_data.append(raw_data['header'])
   # at this state goalie_win_data is a list with 0_n dictionaries as elements 
   return goalie_win_data

def get_goalie_shootout_stats (filters: player_filters) -> list: 
   # Input validation for the fields where no standard value is present
   if len(filters.phase) == 0: 
      raise Exception('Phase is mandatory for the goalie shootout stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the goalie shootout stat')
   
   # creation of the request URL with the help of the constants and the 
   # filter class where defaults are set (and where this is not possible
   # the input was already checked)
   request_url =(BASE_URL + STAT_CACHE + GOALIE_SHOOTOUT_STAT + SEARCHQUERY + '1//1'
               + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
               + filters.team + '/' + filters.licence + FILTERBY + FILTER_NAMES 
               + RECORDS_TO_BE_RETURNED + '1000' + STANDARD_ENDING)
   # get the whole data (including page headers of the official website, 
   # all possible filters and the actual data)
   raw_data = send_request(request_url)

   # only return the actual data since the rest is obsolet for us 
   goalie_shootout_data  = raw_data['data']
   # add the headers as last entry to the list (so they can be poped before a loop)
   goalie_shootout_data.append(raw_data['header'])
   # at this state goalie_win_data is a list with 0_n dictionaries as elements 
   return goalie_shootout_data

###############################################################################
# Team stats 
###############################################################################
def get_team_goal_summary_stats(filters: player_filters) -> list:
   if len(filters.phase) == 0:
      raise Exception('Phase is mandatory for the team goal summary stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the team goal summary stat')
   
   request_url =(BASE_URL + STAT_CACHE + TEAM_GOAL_SUMMARY_STAT + SEARCHQUERY + '1//1'
               + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
               + filters.team + '/' + filters.licence + FILTERBY + FILTER_NAMES 
               + RECORDS_TO_BE_RETURNED + '1000' + STANDARD_ENDING)
   raw_data = send_request(request_url)

   team_goal_summary_data = raw_data['data']
   team_goal_summary_data.append(raw_data['header'])
   return team_goal_summary_data

def get_team_goal_pos_stats(filters: player_filters) -> list:
   if len(filters.phase) == 0:
      raise Exception('Phase is mandatory for the team goal position and licence stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the team goal position and licence stat')
   
   request_url =(BASE_URL + STAT_CACHE + TEAM_POSITION_GOAL_STAT + SEARCHQUERY + '1//1'
               + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
               + filters.team + '/' + filters.licence + FILTERBY + FILTER_NAMES 
               + RECORDS_TO_BE_RETURNED + '1000' + STANDARD_ENDING)
   raw_data = send_request(request_url)

   team_goal_pos_data = raw_data['data']
   team_goal_pos_data.append(raw_data['header'])

   return team_goal_pos_data

def get_team_shot_stats(filters: player_filters) -> list:
   if len(filters.phase) == 0:
      raise Exception('Phase is mandatory for the team shot stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the team shot stat')
   
   request_url =(BASE_URL + STAT_CACHE + TEAM_SHOT_STAT + SEARCHQUERY + '1//1'
               + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
               + filters.team + '/' + filters.licence + FILTERBY + FILTER_NAMES 
               + RECORDS_TO_BE_RETURNED + '1000' + STANDARD_ENDING)
   raw_data = send_request(request_url)

   team_shot_data = raw_data['data']
   team_shot_data.append(raw_data['header'])
   return team_shot_data

def get_team_pp_stats(filters: player_filters) -> list:
   if len(filters.phase) == 0:
      raise Exception('Phase is mandatory for the team powerplay stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the team powerplay stat')
   
   request_url =(BASE_URL + STAT_CACHE + TEAM_PP_STAT + SEARCHQUERY + '1//1'
               + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
               + filters.team + '/' + filters.licence + FILTERBY + FILTER_NAMES 
               + RECORDS_TO_BE_RETURNED + '1000' + STANDARD_ENDING)
   raw_data = send_request(request_url)

   team_pp_data = raw_data['data']
   team_pp_data.append(raw_data['header'])
   return team_pp_data

def get_team_pk_stats(filters: player_filters) -> list:
   if len(filters.phase) == 0:
      raise Exception('Phase is mandatory for the team boxplay stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the team boxplay stat')
   
   request_url =(BASE_URL + STAT_CACHE + TEAM_PK_STAT + SEARCHQUERY + '1//1'
               + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
               + filters.team + '/' + filters.licence + FILTERBY + FILTER_NAMES 
               + RECORDS_TO_BE_RETURNED + '1000' + STANDARD_ENDING)
   raw_data = send_request(request_url)

   team_pk_data = raw_data['data']
   team_pk_data.append(raw_data['header'])
   return team_pk_data

def get_team_foul_stats(filters: player_filters) -> list:
   if len(filters.phase) == 0:
      raise Exception('Phase is mandatory for the team foul stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the team foul stat')
   
   request_url =(BASE_URL + STAT_CACHE + TEAM_PENALTY_STAT + SEARCHQUERY + '1//1'
               + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
               + filters.team + '/' + filters.licence + FILTERBY + FILTER_NAMES 
               + RECORDS_TO_BE_RETURNED + '1000' + STANDARD_ENDING)
   raw_data = send_request(request_url)

   team_foul_data = raw_data['data']
   team_foul_data.append(raw_data['header'])
   return team_foul_data

def get_team_shootout_stats(filters: player_filters) -> list:
   if len(filters.phase) == 0:
      raise Exception('Phase is mandatory for the team shootout stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the team shootout stat')
   
   request_url =(BASE_URL + STAT_CACHE + TEAM_SHOOTOUT_STAT + SEARCHQUERY + '1//1'
               + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
               + filters.team + '/' + filters.licence + FILTERBY + FILTER_NAMES 
               + RECORDS_TO_BE_RETURNED + '1000' + STANDARD_ENDING)
   raw_data = send_request(request_url)

   team_shootout_data = raw_data['data']
   team_shootout_data.append(raw_data['header'])
   return team_shootout_data

def get_team_faceoff_summary_stats(filters: player_filters) -> list:
   if len(filters.phase) == 0:
      raise Exception('Phase is mandatory for the team faceoff summary stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the team faceoff summary stat')
   
   request_url =(BASE_URL + STAT_CACHE + TEAM_FACEOFF_SUMMARY_STAT + SEARCHQUERY + '1//1'
               + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
               + filters.team + '/' + filters.licence + FILTERBY + FILTER_NAMES 
               + RECORDS_TO_BE_RETURNED + '1000' + STANDARD_ENDING)
   raw_data = send_request(request_url)

   team_faceoff_summary_data = raw_data['data']
   team_faceoff_summary_data.append(raw_data['header'])
   return team_faceoff_summary_data

def get_team_faceoff_zone_per_game_stats(filters: player_filters) -> list:
   if len(filters.phase) == 0:
      raise Exception('Phase is mandatory for the team faceoff zone stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the team faceoff zone stat')
   
   request_url =(BASE_URL + STAT_CACHE + TEAM_FACEOFF_ZONE_STAT + SEARCHQUERY + '1//1'
               + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
               + filters.team + '/' + filters.licence + FILTERBY + FILTER_NAMES 
               + RECORDS_TO_BE_RETURNED + '1000' + STANDARD_ENDING)
   raw_data = send_request(request_url)

   team_faceoff_zone_data = raw_data['data']
   team_faceoff_zone_data.append(raw_data['header'])
   return team_faceoff_zone_data

def get_team_faceoff_zone_per_game_stats(filters: player_filters) -> list:
   if len(filters.phase) == 0:
      raise Exception('Phase is mandatory for the team faceoff zone per game stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the team faceoff zone per game stat')
   
   request_url =(BASE_URL + STAT_CACHE + TEAM_FACEOFF_ZONE_PER_GAME_STAT + SEARCHQUERY + '1//1'
               + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
               + filters.team + '/' + filters.licence + FILTERBY + FILTER_NAMES 
               + RECORDS_TO_BE_RETURNED + '1000' + STANDARD_ENDING)
   raw_data = send_request(request_url)

   team_faceoff_zpg_data = raw_data['data']
   team_faceoff_zpg_data.append(raw_data['header'])
   return team_faceoff_zpg_data

def get_team_spectators_stats(filters: player_filters) -> list:
   if len(filters.phase) == 0:
      raise Exception('Phase is mandatory for the team attendance stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the team attendance stat')
   
   request_url =(BASE_URL + STAT_CACHE + TEAM_ATTENDANCE_STAT + SEARCHQUERY + '1//1'
               + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
               + filters.team + '/' + filters.licence + FILTERBY + FILTER_NAMES 
               + RECORDS_TO_BE_RETURNED + '1000' + STANDARD_ENDING)
   raw_data = send_request(request_url)

   team_attendance_data = raw_data['data']
   team_attendance_data.append(raw_data['header'])
   return team_attendance_data

# test data to see if the new fancy stuff works 
def test_stats_current_season () -> None: 
    test_filters = player_filters('2026', '4940', '101139')
    print(get_goalie_shootout_stats(test_filters))

def test_get_filters() -> None:
   print(get_filter_options('2025'))

def test_goalie_shootout() -> None:
   test_filters = player_filters('2025', '4595', '101139')
   print(get_goalie_shootout_stats(test_filters))

def test_team_stat() -> None:
   test_filters = player_filters('2025', '4595','', '')
   print(get_team_goal_pos_stats(test_filters))

test_stats_current_season()