###############################################################################
# the imports
###############################################################################
import datetime
import requests
import json
import re
import warnings
import numpy as np
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
class Planned_game: 
   """
   This object contains the relevant information of a planned game of the SIHF

   Attributes
   ----------
   league : str
      A string which is part of the constant LEAGUES which then gets 
      translated into the numeric representation used by SIHF
      only used as a help in case of troubles, should not have any logic attached to it.
   home_team : str
      A block of numbers representing the home team.
   away_team : str
      A block of numbers representing the home team.
   start_date : str
      A string representing the date of when the game takes place in the format DD.MM.YYYY.
   start_time : str
      A string representing the time at which the game starts in the format HH:MM (24h system used = no AM/PM).

   """

   def __init__(self, league, home_team, away_team, start_date, start_time):
      self.league     = league
      self.home_team  = home_team
      self.away_team  = away_team 
      self.start_date = start_date
      self.start_time = start_time

class Game_detail:
   """
   This object contains the relevant information of a completed / finished game of the SIHF

   Attributes
   ----------
   game_id: str
      A string representing the official SIHF GameID, typically containing the season. This 
      attribute is used for the communication with the API.
   league : str
      A string which is part of the constant LEAGUES which then gets 
      translated into the numeric representation used by SIHF
      only used as a help in case of troubles, should not have any logic attached to it.
   home_team_id : str
      A block of numbers representing the home team. Used to create the attribute home_team.
   away_team_id : str
      A block of numbers representing the away team. Used to create the attribute away_team.
   date : str
      A string representing the date of when the game took place in the format DD.MM.YYYY
   home_team : Team_game_detail()
      A variable which contains all stats of the game of the home team
   away_team : Team_game_detail()
      A variable which contains all stats of the game of the away team

   """

   def __init__(self, game_id, league, home_team_id, away_team_id, date):
      """
      Initializes a player_filter object.
      
      Parameters
      -----------
         game_id: str
                  A string representing the official SIHF GameID, typically containing 
                  the season. This attribute is used for the communication with the API.
         league: str 
                 A string, which contains a league name in lower case and underscore
                 instead of space format. The value then gets used to check it in the 
                 dictionary LEAGUES and retreive the numeric representation used by SIHF
         home_team_id: str
                       A block of numbers representing the home team of a game. 
                       Used to create the Team_game_detail object for the home team.
         away_team_id: str
                       A block of numbers representing the away team of a game. 
                       Used to create the Team_game_detail object for the away team.
         date: str 
               A string representing the date of when the game took place 
               in the format DD.MM.YYYY
      Returns
      -------
         Nothing

      See also
      --------
         swissicehockey_stats LEAGUES
         swissicehockey_stats.Game_details.Team_game_detail

      """
      self.game_id    = game_id
      self.league     = league
      self.date       = date
      self.home_team  = self.Team_game_detail(home_team_id, True)
      self.away_team  = self.Team_game_detail(away_team_id, False)

   class Team_game_detail:
      """
      This object contains the stats of one team for one finished / completed game 

      Attributes
      ----------
      team_id: str
         A block of numbers representing the home team.
      hometeam : bool
         A boolean which tells if the team represented by this object was the home team on the
         game (=true) or not (=false)
      goals : str
         Amount of goals scored by the team stored as a string.
      goals_received : str
         Amount of goals against the team stored as a string.
      sog : str
         Amount of shot on goal performed by the team stored as a string
      shots_missed : str
         Amount of shots by the team, which missed the goal, stored as a string
      sob : str
         Amount of shots against the goal border by a team stored as a string
      blocked_shots: str
         Amount of shots blocked by a team stored as a string
      fo_won: str
         Amount of face offs won by a team stored as a string
      fo_lost: str
         Amount of face offs lost by a team stored as a string
      fo_tot: str
         Total amount of performed face offs by a team stored as a string
      fo_oz_won_rate: str
         Percentage rate rounded to two decimal positions of face offs won by a 
         team in the offensive zone, stored as a string
      fo_nz_won_rate: str
         Percentage rate rounded to two decimal positions of face offs won by a 
         team in the neutral zone, stored as a string
      fo_dz_won_rate: str
         Percentage rate rounded to two decimal positions of face offs won by a 
         team in the defensive zone, stored as a string
      pp_time: str
         Time spent in power play in the format MM:SS stored as a string
      pk_time: str
         Time spent in box play (aka penalty killing) in the format MM:SS stored as a string
      pp_ops: str
         Amount of power play opportunities for a team stored as a string
      pk_sit: str
         Amount of box play situations for a team stored as a string
      pp_goals: str
         Amount of scored goals by a team in power play stored as a string
      pk_goals: str
         Amount of scored goals by a team in box play (aka short handers) stored as a string
      pp_goals_received: str
         Amount of goals against a team in power play stored as a string
      pk_goals_received: str
         Amount of goals against a team in box play stored as a string
      pim: str
         Amount of penalty minutes received by a team stored as a string

      """
      def __init__(self, team_id: str, hometeam: bool):
         """
         Initializes a player_filter object.
      
         Parameters
         -----------
            team_id: str
                     A block of numbers representing the home team.
            hometeam: bool 
                      A boolean which tells if the team represented by this object was the home team on the
                      game (=true) or not (=false)
         Return
         -------
            Nothing

         """
         self.team_id           = team_id
         self.hometeam          = hometeam
         self.goals             = None  
         self.goals_received    = None  
         self.sog               = None 
         self.shots_missed      = None 
         self.sob               = None
         self.blocked_shots     = None
         self.fo_won            = None
         self.fo_lost           = None
         self.fo_tot            = None
         self.fo_oz_won_rate    = None
         self.fo_nz_won_rate    = None
         self.fo_dz_won_rate    = None
         self.pp_time           = None
         self.pk_time           = None
         self.pp_ops            = None
         self.pk_sit            = None
         self.pp_goals          = None 
         self.pk_goals          = None 
         self.pp_goals_received = None
         self.pk_goals_received = None
         self.pim               = None

class Player_filters: 
   """
   A set of of possible filters for the SIHF player and team statistics

   Attributes
   ----------
   season : str
      A 4 character represenation of the season (in SIHF format) which is always the 
      year, when the play offs take place (e.g season 2025 took place from 17.09.2024 - 24.04.2025)
   league : str 
      A string which is part of the constant LEAGUES which then gets 
      translated into the numeric representation used by SIHF
   phase : str
      A block of numbers representing the phase of a season. 
      Phases can be regular season, play-in, play-off, play-out, etc.
   team : str
      A numeric block representing the team you would like to get the data
      from (default all)
   licence : str
      A string representing a flag if you want to filter for Swiss or abroad 
      licences (only applies for player filters) (default all)
   position: str
      A string representing a flag which position you want to filter
      (only applies for player filters) (default all)

   """

   # the ="all" indicates that if not provided, this is the standard value of the
   # given variable 
   def __init__(self, season, league, phase, team="all", licence="all", position="all"):
      """
      Initializes a Player_filter object.
      
      Parameters
      ----------
         season: str
                 A 4 character represenation of the season (in SIHF format) 
                 which is always the year, when the play offs take place 
                 (e.g season 2025 took place from 17.09.2024 - 24.04.2025)
         league: str
                 A string, which contains a league name in lower case and underscore
                 instead of space format. The value then gets used to check it in the 
                 dictionary LEAGUES and retreive the numeric representation used by SIHF
         phase: str
                A block of numbers representing the phase of a season. 
                Phases can be regular season, play-in, play-off, play-out, etc.
         team: str, default=all 
               A numeric block representing the team you would like to get the data
               from (default all)
         licence: str, default=all
                  A string representing a flag if you want to filter for Swiss or abroad 
                  licences (only applies for player filters) (default all)
         position: str, default=all
                   A string representing a flag which position you want to filter
                   (only applies for player filters) (default all)
      Returns
      -------
         Nothing
      
      See also
      --------
         swissicehockey_stats LEAGUES

      """
      self.season   = season
      # translate the provided string into the number needed by swiss ice hockey
      # using the predefined dictionary - we directly transform it into a string 
      # as we only use the league parameter for the URL creation 
      self.league   = str(LEAGUES[league])
      self.phase    = phase 
      self.team     = team
      self.licence  = licence 
      self.position = position
###############################################################################
# the function used to identify what the current season is
###############################################################################
def get_current_season() -> str:
    """
    Gets the current season in SIHF format based on the current date and the current month.
    
    :return: A 4 character represenation of the season (in SIHF format) which is always the 
             year, when the play offs take place (e.g season 2025 took place from 
             17.09.2024 - 24.04.2025).
    :rtype: str

    """
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
# game detail base URL
GAME_URL = 'https://data.sihf.ch/statistic/api/cms/gameoverview?'
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
FACEOFF_ZONE_STAT               = 'alias=playerFaceoffZone'
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
# today cache
TODAY_CACHE = '30?'
# postfinance Top Scorer
TOPSCORER_DATA = 'alias=pftopscorer'
# results
RESULTS = 'alias=results'
# todays games
GAMES_TODAY = 'alias=today'
# game detail
GAME_DETAIL = 'alias=gameDetail'
# precice which games
TODAY_AND_DELAYED_GAMES = 'todayDelayed'
NEXT_AND_DELAYED_GAMES = 'nextDayDelayed'
CURRENT_SEASON_AND_DELAYED_GAMES = 'currentSeasonDelayed'
# size of games to be returned?
TODAY_SIZE = '&size=today'
NEXT_DAY_SIZE = '&size=nextday'
# filters for the stats
# season (only stat filter we will use, as there will be an initial call 
# to get the possible filter options) 
CURRENT_SEASON   = get_current_season()
# search definition 
SEARCHQUERY = '&searchQuery='
FILTERQUERY = '&filterQuery='
FILTERBY = '&filterBy='
# order by
ORDERBY = '&orderBy='
# filter names
FILTER_NAMES = 'Season,Phase,Team,Position,Licence'
GAME_FILER_NAMES = 'League,Team,deferredState'
PAST_GAMES_FILTER_NAMES = 'season,phase,date,deferredState,team1,team2'
# order names 
ORDER_NAMES = 'League'
ORDER_DATE = 'date'
ORDER_ASC = '&orderByDescending=false'
ORDER_DESC = '&orderByDescending=true'
# query ending
RECORDS_TO_BE_RETURNED = '&take='
STANDARD_ENDING = '&callback=externalStatisticsCallback&skip=-1&language=de'
################################################################################
# the functions 
################################################################################
################################################################################
# base functions
################################################################################
def send_request(request_url: str) -> dict:
    """
    Performs a HTTP/HTTPS GET request towards the SIHF data API with the provided 
    URL and extracts the JavaScript container arround it with the help of regex

    Parameters
    ----------
        request_url: str
                     A HTTP or HTTPS URL with the endpoint at an API 
                     to perform a GET Request

    Returns
    ------ 
        json_data: dict
                   A dictionary with n entries which contain the result without 
                   JavaScript in JSON format (that's why it is in a dictionary),
                   where you have a filters, data and header in case you have a 
                   result
    Raises
    ------
    Exception
      If the API endpoint could not be reached (URL is provided in the exception)
    Exception
      If the regex extraction or the json load of the response caused issues
      (URL is provided in the exception)

    """
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
    
def get_filter_options (league: str, season=' ') -> dict:
    """
    Get the filter options for given league and season and return them in a dict

    If no season or a string with more than 4 characters gets provided it will
    be replaced with the current season. 
    A request to the swiss icehockey API gets sent and all possible filters
    (eg. season, teams, phases, positions, etc.) will be added as dictionary
    within a dictionary. The 'alias' values are the required ones for further
    processing. 

    Parameters
    ---------- 
        season: str, optional
                A 4 digit year in string format 
        league: str 
                A string parameter which is part of the constant LEAGUES 
                and gets translated into the numeric representation used 
                by the siwss ice hockey fundation 

    Returns
    ------- 
        filter_options: dict
                        A dictionary with n entries where the values 
                        are again a dictionary
    Raises
    ------
    Exception 
        If no filters could be retreived for the given season

    Warnings
    --------
        If the provided season got changed to the current season

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
    league_num = str(LEAGUES[league])
    try:
       request_url = (BASE_URL + STAT_CACHE + SUMMARY_STAT + SEARCHQUERY + '1//' + league_num
                    + FILTERQUERY + season + FILTERBY + FILTER_NAMES + RECORDS_TO_BE_RETURNED 
                    + '1' + STANDARD_ENDING)

       raw_data = send_request(request_url)
    except: 
       warnings.warn('Something went wrong / No data was present the provided season. ' \
                     'received data will be based of the current season')
       request_url = (BASE_URL + STAT_CACHE + SUMMARY_STAT + SEARCHQUERY + '1//' + league_num
                    + FILTERQUERY + CURRENT_SEASON + FILTERBY + FILTER_NAMES 
                    + RECORDS_TO_BE_RETURNED + '1' + STANDARD_ENDING)

       raw_data = send_request(request_url) 
    # predefine an empty dict (yet another COBOL thing)
    filter_options = {}
    if (raw_data['filters'][0]['selected'] != season):
        warnings.warn('Something went wrong / No data was present the provided season. ' \
                      'Received data will be based of the current season')
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
        request_url =(BASE_URL + STAT_CACHE + SUMMARY_STAT + SEARCHQUERY + '1//' + league_num
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

def get_summary_player_stats (filters: Player_filters) -> list:
   """
   Get the summary stat of field players using the provided filters.

   The summary stat is the exact same as if you click on player stats ->
   Summary on the SIHF web page.

   Parameters
   ----------
      filters: Player_filters
               A set of filters which can be passed / used to get the 
               desired data.
   Returns
   -------
      summary_data: list of list of str or dict
                    A list containing the statistic data where one entry 
                    represents one player as a list which contains the actual 
                    data or a dictionary in case of the team.
                    The very last entry of the list currently still contains
                    the headers for an easier implementation.
   Raises
   ------
   Exception
             If phase is not provided.
   Exception
             If season is not provided.
   
   See also
   --------
   swissicehockey_stats.Player_filters
   https://m.sihf.ch/de/game-center/national-league/#/players/points/desc/page/0/

   """
   # Input validation for the fields where no standard value is present
   if len(filters.phase) == 0: 
      raise Exception('Phase is mandatory for the summary stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the summary stat')
   
   # creation of the request URL with the help of the constants and the 
   # filter class where defaults are set (and where this is not possible
   # the input was already checked)
   request_url =(BASE_URL + STAT_CACHE + SUMMARY_STAT + SEARCHQUERY + '1//'
               + filters.league + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
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

def get_player_goal_stats (filters: Player_filters) -> list: 
   """
   Get the goal stat of field players using the provided filters.

   The goal stat is the exact same as if you click on player stats ->
   Goals/Assists on the SIHF web page.

   Parameters
   ----------
      filters: Player_filters
               A set of filters which can be passed / used to get the 
               desired data.
   Returns
   -------
      goal_data: list of list of str or dict
                 A list containing the statistic data where one entry 
                 represents one player as a list which contains the actual 
                 data or a dictionary in case of the team.
                 The very last entry of the list currently still contains
                 the headers for an easier implementation.
   Raises
   ------
   Exception
             If phase is not provided.
   Exception
             If season is not provided.
   
   See also
   --------
   swissicehockey_stats.Player_filters
   https://m.sihf.ch/de/game-center/national-league/#/mashup/players/playerGoalAssist/goals/desc/page/1/

   """
   # Input validation for the fields where no standard value is present
   if len(filters.phase) == 0: 
      raise Exception('Phase is mandatory for the goal/assist stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the goal/assist stat')
   
   # creation of the request URL with the help of the constants and the 
   # filter class where defaults are set (and where this is not possible
   # the input was already checked)
   request_url =(BASE_URL + STAT_CACHE + GOAL_STATS + SEARCHQUERY + '1//'
               + filters.league + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
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

def get_player_shot_stats (filters: Player_filters) -> list: 
   """
   Get the shot stat of field players using the provided filters.

   The shot stat is the exact same as if you click on player stats ->
   Goals/Assists on the SIHF web page.

   Parameters
   ----------
      filters: Player_filters
               A set of filters which can be passed / used to get the 
               desired data.
   Returns
   -------
      shot_data: list of list of str or dict
                 A list containing the statistic data where one entry 
                 represents one player as a list which contains the actual 
                 data or a dictionary in case of the team.
                 The very last entry of the list currently still contains
                 the headers for an easier implementation.
   Raises
   ------
   Exception
             If phase is not provided.
   Exception
             If season is not provided.
   
   See also
   --------
   swissicehockey_stats.Player_filters
   https://m.sihf.ch/de/game-center/national-league/#/mashup/players/playerShotDetail/sogSog/desc/page/1/

   """
   # Input validation for the fields where no standard value is present
   if len(filters.phase) == 0: 
      raise Exception('Phase is mandatory for the shot on goal stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the shot on goal stat')
   
   # creation of the request URL with the help of the constants and the 
   # filter class where defaults are set (and where this is not possible
   # the input was already checked)
   request_url =(BASE_URL + STAT_CACHE + SHOT_STAT + SEARCHQUERY + '1//'
               + filters.league + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
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

def get_player_penlaties_stats (filters: Player_filters) -> list: 
   """
   Get the penalties stat of field players using the provided filters.

   The shot stat is the exact same as if you click on player stats ->
   Penalties on the SIHF web page.

   Parameters
   ----------
      filters: Player_filters
               A set of filters which can be passed / used to get the 
               desired data.
   Returns
   -------
      pim_data: list of list of str or dict
                A list containing the statistic data where one entry 
                represents one player as a list which contains the actual 
                data or a dictionary in case of the team.
                The very last entry of the list currently still contains
                the headers for an easier implementation.
   Raises
   ------
   Exception
             If phase is not provided.
   Exception
             If season is not provided.
   
   See also
   --------
   pim = penalty minutes
   swissicehockey_stats.Player_filters
   https://m.sihf.ch/de/game-center/national-league/#/mashup/players/playerFoul/pimTotal/desc/page/1/

   """
   # Input validation for the fields where no standard value is present
   if len(filters.phase) == 0: 
      raise Exception('Phase is mandatory for the shot on goal stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the shot on goal stat')
   
   # creation of the request URL with the help of the constants and the 
   # filter class where defaults are set (and where this is not possible
   # the input was already checked)
   request_url =(BASE_URL + STAT_CACHE + PENALTY_STAT + SEARCHQUERY + '1//'
               + filters.league + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
               + filters.team + '/' + filters.position + '/' + filters.licence
               + FILTERBY + FILTER_NAMES + RECORDS_TO_BE_RETURNED + '1000' 
               + STANDARD_ENDING)
   # get the whole data (including page headers of the official website, 
   # all possible filters and the actual data)
   raw_data = send_request(request_url)

   # only return the actual data since the rest is obsolet for us 
   pim_data  = raw_data['data']
   pim_data.append(raw_data['header'])
   # at this state pim_data is a list with 0_n dictionaries as elements 
   return pim_data

def get_player_shootout_stats (filters: Player_filters) -> list: 
   """
   Get the shootout stat of field players using the provided filters.

   The shootout stat is the exact same as if you click on player stats ->
   Shootouts on the SIHF web page.

   Parameters
   ----------
      filters: Player_filters
               A set of filters which can be passed / used to get the 
               desired data.
   Returns
   -------
      shootout_data: list of list of str or dict
                     A list containing the statistic data where one entry 
                     represents one player as a list which contains the actual 
                     data or a dictionary in case of the team.
                     The very last entry of the list currently still contains
                     the headers for an easier implementation.
   Raises
   ------
   Exception
             If phase is not provided.
   Exception
             If season is not provided.
   
   See also
   --------
   swissicehockey_stats.Player_filters
   https://m.sihf.ch/de/game-center/national-league/#/mashup/players/playerShootout/penShots/desc/page/1/

   """
   # Input validation for the fields where no standard value is present
   if len(filters.phase) == 0: 
      raise Exception('Phase is mandatory for the shootout stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the shootout stat')
   
   # creation of the request URL with the help of the constants and the 
   # filter class where defaults are set (and where this is not possible
   # the input was already checked)
   request_url =(BASE_URL + STAT_CACHE + SHOOTOUT_STAT + SEARCHQUERY + '1//'
               + filters.league + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
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

def get_player_faceoff_summary_stats (filters: Player_filters) -> list: 
   """
   Get the faceoffs summary stat of field players using the provided filters.

   The faceoffs summary stat is the exact same as if you click on player stats ->
   Faceoffs Summary on the SIHF web page.

   Parameters
   ----------
      filters: Player_filters
               A set of filters which can be passed / used to get the 
               desired data.
   Returns
   -------
      faceoff_data: list of list of str or dict
                    A list containing the statistic data where one entry 
                    represents one player as a list which contains the actual 
                    data or a dictionary in case of the team.
                    The very last entry of the list currently still contains
                    the headers for an easier implementation.
   Raises
   ------
   Exception
             If phase is not provided.
   Exception
             If season is not provided.
   
   See also
   --------
   swissicehockey_stats.Player_filters
   https://m.sihf.ch/de/game-center/national-league/#/mashup/players/playerFaceoff/faceoffs/desc/page/1/

   """
   # Input validation for the fields where no standard value is present
   if len(filters.phase) == 0: 
      raise Exception('Phase is mandatory for the faceoff stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the faceoff stat')
   
   # creation of the request URL with the help of the constants and the 
   # filter class where defaults are set (and where this is not possible
   # the input was already checked)
   request_url =(BASE_URL + STAT_CACHE + FACEOFF_SUMMARY_STAT + SEARCHQUERY + '1//'
               + filters.league + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
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

def get_player_faceoff_zone_stats (filters: Player_filters) -> list: 
   """
   Get the faceoffs per zone stat of field players using the provided filters.

   The faceoffs per zone stat is the exact same as if you click on player stats ->
   Faceoffs/Zone on the SIHF web page.

   Parameters
   ----------
      filters: Player_filters
               A set of filters which can be passed / used to get the 
               desired data.
   Returns
   -------
      faceoff_data: list of list of str or dict
                    A list containing the statistic data where one entry 
                    represents one player as a list which contains the actual 
                    data or a dictionary in case of the team.
                    The very last entry of the list currently still contains
                    the headers for an easier implementation.
   Raises
   ------
   Exception
             If phase is not provided.
   Exception
             If season is not provided.
   
   See also
   --------
   swissicehockey_stats.Player_filters
   https://m.sihf.ch/de/game-center/national-league/#/mashup/players/playerFaceoffZone/faceoffsOffensive/desc/page/1/

   """
   # Input validation for the fields where no standard value is present
   if len(filters.phase) == 0: 
      raise Exception('Phase is mandatory for the faceoff zone stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the faceoff zone stat')
   
   # creation of the request URL with the help of the constants and the 
   # filter class where defaults are set (and where this is not possible
   # the input was already checked)
   request_url =(BASE_URL + STAT_CACHE + FACEOFF_ZONE_STAT + SEARCHQUERY + '1//'
               + filters.league + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
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

def get_player_faceoff_zone_pg_stats (filters: Player_filters) -> list: 
   """
   Get the faceoffs per game stat of field players using the provided filters.

   The faceoffs per game stat is the exact same as if you click on player stats ->
   Faceoffs/Spiel on the SIHF web page.

   Parameters
   ----------
      filters: Player_filters
               A set of filters which can be passed / used to get the 
               desired data.
   Returns
   -------
      faceoff_data: list of list of str or dict
                    A list containing the statistic data where one entry 
                    represents one player as a list which contains the actual 
                    data or a dictionary in case of the team.
                    The very last entry of the list currently still contains
                    the headers for an easier implementation.
   Raises
   ------
   Exception
             If phase is not provided.
   Exception
             If season is not provided.
   
   See also
   --------
   swissicehockey_stats.Player_filters
   https://m.sihf.ch/de/game-center/national-league/#/mashup/players/playerFaceoffZoneGame/faceoffsOffensivePerGame/desc/page/1/

   """
   # Input validation for the fields where no standard value is present
   if len(filters.phase) == 0: 
      raise Exception('Phase is mandatory for the faceoff zone per game stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the faceoff zone per game stat')
   
   # creation of the request URL with the help of the constants and the 
   # filter class where defaults are set (and where this is not possible
   # the input was already checked)
   request_url =(BASE_URL + STAT_CACHE + FACEOFF_ZONE_PER_GAME_STAT + SEARCHQUERY + '1//'
               + filters.league + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
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

def get_player_time_on_ice_stats (filters: Player_filters) -> list: 
   """
   Get the time on ice stat of field players using the provided filters.

   The time on ice stat is the exact same as if you click on player stats ->
   Time on ice on the SIHF web page.

   Parameters
   ----------
      filters: Player_filters
               A set of filters which can be passed / used to get the 
               desired data.
   Returns
   -------
      toi_data: list of list of str or dict
                A list containing the statistic data where one entry 
                represents one player as a list which contains the actual 
                data or a dictionary in case of the team.
                The very last entry of the list currently still contains
                the headers for an easier implementation.
   Raises
   ------
   Exception
             If phase is not provided.
   Exception
             If season is not provided.
   
   See also
   --------
   toi = time on ice
   swissicehockey_stats.Player_filters
   https://m.sihf.ch/de/game-center/national-league/#/mashup/players/playerTimeOnIce/timeOnIcePerGame/desc/page/1/

   """
   # Input validation for the fields where no standard value is present
   if len(filters.phase) == 0: 
      raise Exception('Phase is mandatory for the summary stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the summary stat')
   
   # creation of the request URL with the help of the constants and the 
   # filter class where defaults are set (and where this is not possible
   # the input was already checked)
   request_url =(BASE_URL + STAT_CACHE + TIME_ON_ICE_STAT + SEARCHQUERY + '1//'
               + filters.league + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
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

def get_goalie_summary_stats (filters: Player_filters) -> list: 
   # Input validation for the fields where no standard value is present
   if len(filters.phase) == 0: 
      raise Exception('Phase is mandatory for the goalie summary stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the goalie summary stat')
   
   # creation of the request URL with the help of the constants and the 
   # filter class where defaults are set (and where this is not possible
   # the input was already checked)
   request_url =(BASE_URL + STAT_CACHE + GOALIE_SUMMARY_STAT + SEARCHQUERY + '1//'
               + filters.league + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
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

def get_goalie_shots_against_stats (filters: Player_filters) -> list: 
   # Input validation for the fields where no standard value is present
   if len(filters.phase) == 0: 
      raise Exception('Phase is mandatory for the goalie shots against stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the goalie shots against stat')
   
   # creation of the request URL with the help of the constants and the 
   # filter class where defaults are set (and where this is not possible
   # the input was already checked)
   request_url =(BASE_URL + STAT_CACHE + GOALIE_SHOT_STAT + SEARCHQUERY + '1//'
               + filters.league + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
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

def get_goalie_win_stats (filters: Player_filters) -> list: 
   # Input validation for the fields where no standard value is present
   if len(filters.phase) == 0: 
      raise Exception('Phase is mandatory for the goalie win stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the goalie win stat')
   
   # creation of the request URL with the help of the constants and the 
   # filter class where defaults are set (and where this is not possible
   # the input was already checked)
   request_url =(BASE_URL + STAT_CACHE + GOALIE_WIN_STAT + SEARCHQUERY + '1//'
               + filters.league + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
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

def get_goalie_shootout_stats (filters: Player_filters) -> list: 
   # Input validation for the fields where no standard value is present
   if len(filters.phase) == 0: 
      raise Exception('Phase is mandatory for the goalie shootout stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the goalie shootout stat')
   
   # creation of the request URL with the help of the constants and the 
   # filter class where defaults are set (and where this is not possible
   # the input was already checked)
   request_url =(BASE_URL + STAT_CACHE + GOALIE_SHOOTOUT_STAT + SEARCHQUERY + '1//'
               + filters.league + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
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
def get_team_goal_summary_stats(filters: Player_filters) -> list:
   if len(filters.phase) == 0:
      raise Exception('Phase is mandatory for the team goal summary stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the team goal summary stat')
   
   request_url =(BASE_URL + STAT_CACHE + TEAM_GOAL_SUMMARY_STAT + SEARCHQUERY + '1//'
               + filters.league + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
               + filters.team + '/' + filters.licence + FILTERBY + FILTER_NAMES 
               + RECORDS_TO_BE_RETURNED + '1000' + STANDARD_ENDING)
   raw_data = send_request(request_url)

   team_goal_summary_data = raw_data['data']
   team_goal_summary_data.append(raw_data['header'])
   return team_goal_summary_data

def get_team_goal_pos_stats(filters: Player_filters) -> list:
   if len(filters.phase) == 0:
      raise Exception('Phase is mandatory for the team goal position and licence stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the team goal position and licence stat')
   
   request_url =(BASE_URL + STAT_CACHE + TEAM_POSITION_GOAL_STAT + SEARCHQUERY + '1//'
               + filters.league + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
               + filters.team + '/' + filters.licence + FILTERBY + FILTER_NAMES 
               + RECORDS_TO_BE_RETURNED + '1000' + STANDARD_ENDING)
   raw_data = send_request(request_url)

   team_goal_pos_data = raw_data['data']
   team_goal_pos_data.append(raw_data['header'])

   return team_goal_pos_data

def get_team_shot_stats(filters: Player_filters) -> list:
   if len(filters.phase) == 0:
      raise Exception('Phase is mandatory for the team shot stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the team shot stat')
   
   request_url =(BASE_URL + STAT_CACHE + TEAM_SHOT_STAT + SEARCHQUERY + '1//'
               + filters.league + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
               + filters.team + '/' + filters.licence + FILTERBY + FILTER_NAMES 
               + RECORDS_TO_BE_RETURNED + '1000' + STANDARD_ENDING)
   raw_data = send_request(request_url)

   team_shot_data = raw_data['data']
   team_shot_data.append(raw_data['header'])
   return team_shot_data

def get_team_pp_stats(filters: Player_filters) -> list:
   if len(filters.phase) == 0:
      raise Exception('Phase is mandatory for the team powerplay stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the team powerplay stat')
   
   request_url =(BASE_URL + STAT_CACHE + TEAM_PP_STAT + SEARCHQUERY + '1//'
               + filters.league + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
               + filters.team + '/' + filters.licence + FILTERBY + FILTER_NAMES 
               + RECORDS_TO_BE_RETURNED + '1000' + STANDARD_ENDING)
   raw_data = send_request(request_url)

   team_pp_data = raw_data['data']
   team_pp_data.append(raw_data['header'])
   return team_pp_data

def get_team_pk_stats(filters: Player_filters) -> list:
   if len(filters.phase) == 0:
      raise Exception('Phase is mandatory for the team boxplay stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the team boxplay stat')
   
   request_url =(BASE_URL + STAT_CACHE + TEAM_PK_STAT + SEARCHQUERY + '1//'
               + filters.league + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
               + filters.team + '/' + filters.licence + FILTERBY + FILTER_NAMES 
               + RECORDS_TO_BE_RETURNED + '1000' + STANDARD_ENDING)
   raw_data = send_request(request_url)

   team_pk_data = raw_data['data']
   team_pk_data.append(raw_data['header'])
   return team_pk_data

def get_team_foul_stats(filters: Player_filters) -> list:
   if len(filters.phase) == 0:
      raise Exception('Phase is mandatory for the team foul stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the team foul stat')
   
   request_url =(BASE_URL + STAT_CACHE + TEAM_PENALTY_STAT + SEARCHQUERY + '1//'
               + filters.league + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
               + filters.team + '/' + filters.licence + FILTERBY + FILTER_NAMES 
               + RECORDS_TO_BE_RETURNED + '1000' + STANDARD_ENDING)
   raw_data = send_request(request_url)

   team_foul_data = raw_data['data']
   team_foul_data.append(raw_data['header'])
   return team_foul_data

def get_team_shootout_stats(filters: Player_filters) -> list:
   if len(filters.phase) == 0:
      raise Exception('Phase is mandatory for the team shootout stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the team shootout stat')
   
   request_url =(BASE_URL + STAT_CACHE + TEAM_SHOOTOUT_STAT + SEARCHQUERY + '1//'
               + filters.league + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
               + filters.team + '/' + filters.licence + FILTERBY + FILTER_NAMES 
               + RECORDS_TO_BE_RETURNED + '1000' + STANDARD_ENDING)
   raw_data = send_request(request_url)

   team_shootout_data = raw_data['data']
   team_shootout_data.append(raw_data['header'])
   return team_shootout_data

def get_team_faceoff_summary_stats(filters: Player_filters) -> list:
   if len(filters.phase) == 0:
      raise Exception('Phase is mandatory for the team faceoff summary stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the team faceoff summary stat')
   
   request_url =(BASE_URL + STAT_CACHE + TEAM_FACEOFF_SUMMARY_STAT + SEARCHQUERY + '1//'
               + filters.league + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
               + filters.team + '/' + filters.licence + FILTERBY + FILTER_NAMES 
               + RECORDS_TO_BE_RETURNED + '1000' + STANDARD_ENDING)
   raw_data = send_request(request_url)

   team_faceoff_summary_data = raw_data['data']
   team_faceoff_summary_data.append(raw_data['header'])
   return team_faceoff_summary_data

def get_team_faceoff_zone_per_game_stats(filters: Player_filters) -> list:
   if len(filters.phase) == 0:
      raise Exception('Phase is mandatory for the team faceoff zone stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the team faceoff zone stat')
   
   request_url =(BASE_URL + STAT_CACHE + TEAM_FACEOFF_ZONE_STAT + SEARCHQUERY + '1//'
               + filters.league + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
               + filters.team + '/' + filters.licence + FILTERBY + FILTER_NAMES 
               + RECORDS_TO_BE_RETURNED + '1000' + STANDARD_ENDING)
   raw_data = send_request(request_url)

   team_faceoff_zone_data = raw_data['data']
   team_faceoff_zone_data.append(raw_data['header'])
   return team_faceoff_zone_data

def get_team_faceoff_zone_per_game_stats(filters: Player_filters) -> list:
   if len(filters.phase) == 0:
      raise Exception('Phase is mandatory for the team faceoff zone per game stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the team faceoff zone per game stat')
   
   request_url =(BASE_URL + STAT_CACHE + TEAM_FACEOFF_ZONE_PER_GAME_STAT + SEARCHQUERY + '1//'
               + filters.league + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
               + filters.team + '/' + filters.licence + FILTERBY + FILTER_NAMES 
               + RECORDS_TO_BE_RETURNED + '1000' + STANDARD_ENDING)
   raw_data = send_request(request_url)

   team_faceoff_zpg_data = raw_data['data']
   team_faceoff_zpg_data.append(raw_data['header'])
   return team_faceoff_zpg_data

def get_team_spectators_stats(filters: Player_filters) -> list:
   if len(filters.phase) == 0:
      raise Exception('Phase is mandatory for the team attendance stat')
   
   if len(filters.season) == 0:
      raise Exception('Season is mandatory for the team attendance stat')
   
   request_url =(BASE_URL + STAT_CACHE + TEAM_ATTENDANCE_STAT + SEARCHQUERY + '1//'
               + filters.league + FILTERQUERY + filters.season + '/' + filters.phase + '/' 
               + filters.team + '/' + filters.licence + FILTERBY + FILTER_NAMES 
               + RECORDS_TO_BE_RETURNED + '1000' + STANDARD_ENDING)
   raw_data = send_request(request_url)

   team_attendance_data = raw_data['data']
   team_attendance_data.append(raw_data['header'])
   return team_attendance_data
###############################################################################
# Top Scorer Data
###############################################################################
def get_pf_top_scorer_data(filters: Player_filters) -> list:
   league_filter = '1//'

   if len(filters.league) == 0: 
      warnings.warn('data of the NL will be returned, since no league was provided')
      league_filter = league_filter + '1'
   else:
      league_filter = league_filter + filters.league
   
   filter_values = ''

   if len(filters.season) == 0:
      warnings.warn('data of the current season will be returned, since no season was provided')
   else:
      filter_values = filters.season
   
   if len(filters.phase) > 0: 
      filter_values = filter_values + filters.phase
   
   
   request_url =(BASE_URL + DATA_CACHE + TOPSCORER_DATA + SEARCHQUERY + league_filter
               + FILTERQUERY + filter_values + RECORDS_TO_BE_RETURNED + '99' + FILTERBY
               + FILTER_NAMES + STANDARD_ENDING)
   raw_data = send_request(request_url)

   top_scorer_data = raw_data['data']
   top_scorer_data.append(raw_data['header'])
   return top_scorer_data

###############################################################################
# Game Data
###############################################################################
# get todays games
def get_todays_games (filters: Player_filters) -> list:
   league_filter = '1,2,4,10//'
 
   if len(filters.league) == 0: 
      warnings.warn('data of the NL will be returned, since no league was provided')
      league_filter = league_filter + '1'
      used_league = '1'
   else:
      league_filter = league_filter + filters.league
      used_league = filters.league
   league_filter = league_filter + '/////'
   

   request_url = (BASE_URL + TODAY_CACHE + GAMES_TODAY + SEARCHQUERY + league_filter
                + TODAY_AND_DELAYED_GAMES + FILTERQUERY + ORDERBY + ORDER_NAMES
                + RECORDS_TO_BE_RETURNED + '99' + FILTERBY + GAME_FILER_NAMES  + STANDARD_ENDING)
   raw_data = send_request(request_url)
   # tell python that we want a list
   todays_games = []
   # for every element in data we create a object of the type planned game
   # and add it to the created list
   for i in (raw_data['data']):
      # if we do not have any games, do not create the objects
      if i[0] == 'Heute sind keine Spiele!':
         break
      todays_games.append(Planned_game(used_league, i[3]['id'], i[4]['id'],i[1], i[2]))

   return todays_games

# get next round games 
def get_next_games (filters: Player_filters) -> list:
   league_filter = '1,2,4,10//'

   if len(filters.league) == 0: 
      warnings.warn('data of the NL will be returned, since no league was provided')
      league_filter = league_filter + '1'
      used_league = '1'
   else:
      league_filter = league_filter + filters.league
      used_league = filters.league

   league_filter = league_filter + '/////'
   request_url = (BASE_URL + DATA_CACHE + RESULTS + NEXT_DAY_SIZE + SEARCHQUERY + league_filter
                + NEXT_AND_DELAYED_GAMES + FILTERQUERY + RECORDS_TO_BE_RETURNED + '99' 
                + FILTERBY + GAME_FILER_NAMES + STANDARD_ENDING)
   raw_data = send_request(request_url)
   next_games = []
   for i in (raw_data['data']): 
      next_games.append(Planned_game(used_league,i[4]['id'], i[5]['id'], i[2], i[3]))

   return next_games

def get_past_n_games_of_team(filters: Player_filters, amount_of_games: int) -> list:
   if len(filters.team) == 0: 
     raise Exception('No team was provided to search for!')
   if amount_of_games == 0:
     warnings.warn('No amount of games to retrive was provied, assume default value of 5')
     amount_of_games = 5
   league_filter = '1,10//'
   if len(filters.season) == 0:
      warnings.warn('No season provided, data returned will be of current season')
      used_season = get_current_season()
   else: 
      used_season = filters.season
   if len(filters.league) == 0: 
      warnings.warn('data of the NL will be returned, since no league was provided')
      league_filter = league_filter + '1'
      used_league = '1'
   else:
      league_filter = league_filter + filters.league
      used_league = filters.league
   if len(filters.phase) == 0: 
      used_phase = 'all'
   else:
      used_phase = filters.phase
   daterange= ''
   dates = get_game_plan_of_season(used_season, used_league)
   if int(used_season) >  int((str(datetime.datetime.now()).split('-'))[0]):
      daterange = str(dates[0]) + '-' + datetime.datetime.now().strftime("%d.%m.%Y")
   else:
      daterange = str(dates[0]) + '-' + str(dates[1])
   request_url = (BASE_URL + STAT_CACHE + RESULTS + SEARCHQUERY + league_filter 
                + FILTERQUERY + used_season + '/' + used_phase + '/' + daterange 
                + '/all/' + filters.team + '/all' + FILTERBY + PAST_GAMES_FILTER_NAMES 
                + ORDERBY + ORDER_DATE + ORDER_DESC +  RECORDS_TO_BE_RETURNED 
                + str(amount_of_games) + STANDARD_ENDING)
   raw_data = send_request(request_url)
   past_games_data = raw_data['data']
   past_games = []
   for i in past_games_data:
      game = get_game_detail(i[9]['gameId'])
      past_games.append(game)
   return past_games

def get_game_plan_of_season(season: str, league: str) -> list:
   """
   Helper mehtod which gets the whole game plan of a season of a league, takes the first and the last
   games of the list, gets their dates in DD.MM.YYYY format and adds them into a list

   Parameters
   ----------
   season (str): A 4 character represenation of the season (in SIHF format) which is always the 
                 year, when the play offs take place (e.g season 2025 took place from 
                 17.09.2024 - 24.04.2025)
   league (str): A one - three character string of numbers representing the league, as defined
                 in LEAGUES constant

   Return
   ------ 
   dates_of_season (list): A list, where the first entry represents the start of the season
                           provided and the second (and last) entry represents the last game 
                           of the season. The list element is a string containing a date in the
                           format DD.MM.YYYY

   """
   # Input validation
   # season is mandatory (only because it is an internal method)
   if len(season) == 0: 
     raise Exception('No Season was provided!')
   league_filter = '1,10//'
   if len(league) == 0: 
      warnings.warn('data of the NL will be returned, since no league was provided')
      league_filter = league_filter + '1'
   else:
      league_filter = league_filter + league
   request_url = (BASE_URL + STAT_CACHE + RESULTS + SEARCHQUERY + league_filter 
                 + FILTERQUERY + season + '/all/all/all/all' + FILTERBY + PAST_GAMES_FILTER_NAMES
                 + ORDERBY + ORDER_DATE + ORDER_ASC +  RECORDS_TO_BE_RETURNED + '999' + STANDARD_ENDING)
   raw_data = send_request(request_url)
   games_played = raw_data['data']
   dates_of_season = []
   dates_of_season.append(games_played[0][1])
   # dear COBOL 85, you could have done it like that (yes I agree that going backwards
   # in the memory where all of a sudden you ended in the AGRs of the calling program was not the 
   # best way to handle it)
   dates_of_season.append(games_played[-1][1])
   return dates_of_season

def get_game_detail(game_id: str) -> Game_detail:
   if len(game_id) == 0:
      raise Exception('No game ID was provided!')
   request_url = (GAME_URL + GAME_DETAIL + SEARCHQUERY + game_id + STANDARD_ENDING)
   raw_data = send_request(request_url)
   # do not ask why this works, but here we go, get the team stat detail of a game without header
   # a list with a lot of litst as list objects where each sub list has exactly 3 entries: 
   # 0 = the title
   # 1 = Home team Value
   # 2 = away team value
   game_data_array = np.array(raw_data['stats'][4]['data'])
   game_data = raw_data['stats'][4]['data']
   # create the base object which is to be returned, to ensure that the team classes get
   # allocated, in order to be able to feed the data
   game_detail = Game_detail(game_id, raw_data['league']['id']
                           , raw_data['details']['homeTeam']['id']
                           , raw_data['details']['awayTeam']['id']
                           , (raw_data['startDateTime'].split('T'))[0])
   # feed the teams 
   game_detail.home_team.goals             = raw_data['result']['homeTeam']
   game_detail.away_team.goals             = raw_data['result']['awayTeam']
   game_detail.home_team.goals_received    = raw_data['result']['awayTeam']
   game_detail.away_team.goals_received    = raw_data['result']['homeTeam']
   # now this is super duper very cool; 
   # we take game data and added is as an array for the numpy module
   # due to that we can now use numpy with our created array and search 
   # for values (np.where) and this returns us two arrays where the first
   # array contains the position in the outer list (e.g. BkS is on index 15
   # for game 20261105000203) and the second array contains the position in the
   # inner list (which is always 0). This had to be built because in case
   # of overtimes, dynamically more entries are created for every stat.
   # I am quite sure that this could be done in a smoother way but this is
   # already kinda cool
   list_pos_dyn = 0
   list_pos_dyn = int(((np.where(game_data_array == 'SOG Total'))[0])[0])
   game_detail.home_team.sog               = game_data[list_pos_dyn][1]
   game_detail.away_team.sog               = game_data[list_pos_dyn][2]
   list_pos_dyn = 0
   list_pos_dyn = int(((np.where(game_data_array == 'SHM Total'))[0])[0])
   game_detail.home_team.shots_missed      = game_data[list_pos_dyn][1]
   game_detail.away_team.shots_missed      = game_data[list_pos_dyn][2]
   list_pos_dyn = 0
   list_pos_dyn = int(((np.where(game_data_array == 'SHP Total'))[0])[0])
   game_detail.home_team.sob               = game_data[list_pos_dyn][1]
   game_detail.away_team.sob               = game_data[list_pos_dyn][2]
   list_pos_dyn = 0
   list_pos_dyn = int(((np.where(game_data_array == 'BkS'))[0])[0])
   game_detail.home_team.blocked_shots     = game_data[list_pos_dyn][1]
   game_detail.away_team.blocked_shots     = game_data[list_pos_dyn][2]
   list_pos_dyn = 0
   list_pos_dyn = int(((np.where(game_data_array == 'FOW Total'))[0])[0])
   game_detail.home_team.fo_won            = game_data[list_pos_dyn][1]
   game_detail.away_team.fo_won            = game_data[list_pos_dyn][2]
   list_pos_dyn = 0
   list_pos_dyn = int(((np.where(game_data_array == 'FOL Total'))[0])[0])
   game_detail.home_team.fo_lost           = game_data[list_pos_dyn][1]
   game_detail.away_team.fo_lost           = game_data[list_pos_dyn][2]
   list_pos_dyn = 0
   list_pos_dyn = int(((np.where(game_data_array == 'FO Total'))[0])[0])
   game_detail.home_team.fo_tot            = game_data[list_pos_dyn][1]
   game_detail.away_team.fo_tot            = game_data[list_pos_dyn][2]
   list_pos_dyn = 0
   list_pos_dyn = int(((np.where(game_data_array == 'FO% oz Total'))[0])[0])
   game_detail.home_team.fo_oz_won_rate    = game_data[list_pos_dyn][1]
   game_detail.away_team.fo_oz_won_rate    = game_data[list_pos_dyn][2]
   list_pos_dyn = 0
   list_pos_dyn = int(((np.where(game_data_array == 'FO% nz Total'))[0])[0])
   game_detail.home_team.fo_nz_won_rate    = game_data[list_pos_dyn][1]
   game_detail.away_team.fo_nz_won_rate    = game_data[list_pos_dyn][2]
   list_pos_dyn = 0
   list_pos_dyn = int(((np.where(game_data_array == 'FO% dz Total'))[0])[0])
   game_detail.home_team.fo_dz_won_rate    = game_data[list_pos_dyn][1]
   game_detail.away_team.fo_dz_won_rate    = game_data[list_pos_dyn][2]
   list_pos_dyn = 0
   list_pos_dyn = int(((np.where(game_data_array == 'PPT'))[0])[0])
   game_detail.home_team.pp_time           = game_data[list_pos_dyn][1]
   game_detail.away_team.pp_time           = game_data[list_pos_dyn][2]
   list_pos_dyn = 0
   list_pos_dyn = int(((np.where(game_data_array == 'PKT'))[0])[0])
   game_detail.home_team.pk_time           = game_data[list_pos_dyn][1]
   game_detail.away_team.pk_time           = game_data[list_pos_dyn][2]
   list_pos_dyn = 0
   list_pos_dyn = int(((np.where(game_data_array == 'PP OP'))[0])[0])
   game_detail.home_team.pp_ops            = game_data[list_pos_dyn][1]
   game_detail.away_team.pp_ops            = game_data[list_pos_dyn][2]
   list_pos_dyn = 0
   list_pos_dyn = int(((np.where(game_data_array == 'PK SI'))[0])[0])
   game_detail.home_team.pk_sit            = game_data[list_pos_dyn][1]
   game_detail.away_team.pk_sit            = game_data[list_pos_dyn][2]
   list_pos_dyn = 0
   list_pos_dyn = int(((np.where(game_data_array == 'PPG'))[0])[0])
   game_detail.home_team.pp_goals          = game_data[list_pos_dyn][1]
   game_detail.away_team.pp_goals          = game_data[list_pos_dyn][2]
   list_pos_dyn = 0
   list_pos_dyn = int(((np.where(game_data_array == 'SHG'))[0])[0])
   game_detail.home_team.pk_goals          = game_data[list_pos_dyn][1]
   game_detail.away_team.pk_goals          = game_data[list_pos_dyn][2]
   list_pos_dyn = 0
   list_pos_dyn = int(((np.where(game_data_array == 'PP GA'))[0])[0])
   game_detail.home_team.pp_goals_received = game_data[list_pos_dyn][1]
   game_detail.away_team.pp_goals_received = game_data[list_pos_dyn][2]
   list_pos_dyn = 0
   list_pos_dyn = int(((np.where(game_data_array == 'PK GA'))[0])[0])
   game_detail.home_team.pk_goals_received = game_data[list_pos_dyn][1]
   game_detail.away_team.pk_goals_received = game_data[list_pos_dyn][2]
   list_pos_dyn = int(((np.where(game_data_array == 'PIM Total'))[0])[0])
   game_detail.home_team.pim               = game_data[list_pos_dyn][1]
   game_detail.away_team.pim               = game_data[list_pos_dyn][2]

   return game_detail
#      
# test data to see if the new fancy stuff works 
#
def test_stats_current_season () -> None: 
    test_filters = Player_filters('2026', 'national_league', '4940', '101139')
    print(get_goalie_shootout_stats(test_filters))

def test_get_filters() -> None:
   print(get_filter_options('swiss_league', '2025'))

def test_goalie_shootout() -> None:
   test_filters = Player_filters('2025','national_league', '4595', '101139')
   print(get_goalie_shootout_stats(test_filters))

def test_team_stat() -> None:
   test_filters = Player_filters('2025', 'national_league', '4595')
   print(get_team_goal_pos_stats(test_filters))

def test_ts_data() -> None:
   test_filters = Player_filters('', 'swiss_league', '', '101139')
   print(get_pf_top_scorer_data(test_filters))
   test_filters2 = Player_filters('', 'national_league', '', '101139')
   print(get_pf_top_scorer_data(test_filters2))

def test_today_game_data() -> None:
   test_filters = Player_filters('', 'swiss_league', '', '101139')
   print(get_todays_games(test_filters))
   test_filters2 = Player_filters('', 'national_league', '', '101139')
   print(get_todays_games(test_filters2))

def test_next_game_data() -> None:
   test_filters = Player_filters('', 'swiss_league', '', '101139')
   my_list = get_next_games(test_filters)
   for i in my_list:
      print(i.start_date)
      print(i.home_team)
      print(i.start_time)
   test_filters2 = Player_filters('', 'national_league', '', '101139')
   my_list.clear
   my_list = get_next_games(test_filters2)
   for i in my_list:
      print(i.start_date)
      print(i.home_team)
      print(i.start_time)

def test_past_x_game_data() -> None:
   test_filters = Player_filters('', 'national_league', '', '101139')
   amt = 10
   print((get_past_n_games_of_team(test_filters, amt))[9].away_team.sog)
   test_filters2 = Player_filters('', 'swiss_league', '', '102129')
   amt = 5
   print((get_past_n_games_of_team(test_filters2, amt))[3].home_team.sog)

def test_get_game_det() -> None:
   game = get_game_detail('20261105000203')
   print(game.home_team.sob)
   print(game.away_team.shots_missed)
   print(game.away_team.pp_goals)
test_past_x_game_data()