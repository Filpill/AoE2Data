from pyparsing import match_only_at_col
import requests
import sys
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
from matplotlib.offsetbox import AnnotationBbox, OffsetImage
from datetime import date
from datetime import datetime
from json import loads
from PIL import Image

def api2df(requestAddr):

    #Making the API Request
    print(f'Making API Call to: {requestAddr}')
    response = requests.get(requestAddr)
    if response.status_code == 200:
        print('API Request Successful')
    elif response.status_code == 404:
        print('API Request Failed')
        KeyboardInterrupt

    #Converting JSON to Dict and Parsing API Data into a Dataframe
    dict = loads(response.text)

    return dict


#Function to Detect Operating System and Adjust Pathing to Respective Filesystem
def pathing(filename):

    #Windows Operating System
    if 'win' in sys.platform:
        filepath = f'{sys.path[0]}\\{filename}'
    #Linux/Mac Operating Sytem
    else:
        filepath = f'{sys.path[0]}/{filename}'

    return filepath

def main():

    #-------------------------------API Request and Sorting Data-------------------------------------

    #---------------------------------------------
    #----------------Date Variables---------------
    #---------------------------------------------
    date_today = date.today()
    date_today_unix = int(datetime.now().timestamp())
    date_today_str = date_today.strftime("%d-%B-%Y")

    #--------------------------------------
    #--------------Game Data---------------
    #--------------------------------------

    #Address for Requesting Age of Empires 2 Leaderboard Data
    data='strings'     #Return Keys For the Game String Data
    requestAddr = f'https://aoe2.net/api/{data}?game=aoe2de'
    dict = api2df(requestAddr)
    df = pd.DataFrame.from_dict(dict,orient='index')

    #Parsing the civilisation data into a dataframe
    dict = df.loc['civ',0]
    df_civ = pd.DataFrame.from_dict(dict)

    #--------------------------------------
    #--------------Leaderboard-------------
    #--------------------------------------

    #Address for Requesting Age of Empires 2 Leaderboard Data
    leaderboard_id = 3 #3 = Random Map Leaderboard
    count = 25         #Extracting this many entries from table
    requestAddr = f'https://aoe2.net/api/leaderboard?game=aoe2de&leaderboard_id={leaderboard_id}&start=1&count={count}'
    dict = api2df(requestAddr)
    df = pd.DataFrame.from_dict(dict,orient='index')

    #Parsing the leaderboard data into a dataframe
    dict = df.loc['leaderboard',0]
    leaderboard = pd.DataFrame.from_dict(dict)

    #Taking the Intersection of Column Headers and grabbing top 10 players
    rating_top = leaderboard[leaderboard.columns.intersection(['steam_id','name','rating'])]
    rating_top = rating_top.head(15)
    rating_top = rating_top.iloc[::-1] #Flipping table upside down

    #------------------------------------------------#

    #-------------------------------------------------------------------------------------------------------------------------------------

    #--------------Player Rating History-------------#
    data = 'ratinghistory'  #Returning Rating Data
    leaderboard_id = 3      #3 = Random Map Leaderboard
    count = 500             #Extracting this many entries from table
    top_player_id = rating_top.loc[0,'steam_id'] #Nesting the top player id from previous API call into the following request
    top_player_name = rating_top.loc[0,'name']   #Grabbing name to place into chart title
    requestAddr = f'https://aoe2.net/api/player/{data}?game=aoe2de&leaderboard_id={leaderboard_id}&steam_id={top_player_id}&count={count}'
    dict = api2df(requestAddr)

    #Parsing Player Rating History into Table
    df_player_hist = pd.DataFrame.from_dict(dict,orient='columns')
    df_player_hist['match_date'] = ''
    df_player_hist['win_loss'] = ''

    #Looping through DataFrame to prepare more data for analysis
    for i in range(df_player_hist.shape[0]):

        #Since the timestamps are UNIX based, we need to convert them into DateTime Format so they are readable
        ts = df_player_hist.iat[i,df_player_hist.columns.get_loc('timestamp')]
        df_player_hist.iat[i,df_player_hist.columns.get_loc('match_date')] = datetime.fromtimestamp(ts)

        #Determine if match was a win or a loss (represented by an W or L string)
        i_b = i+1 #Create index value for match b to grab previous match
        if i_b != df_player_hist.shape[0]:

            match_a = df_player_hist.iat[i,df_player_hist.columns.get_loc('num_wins')]
            match_b = df_player_hist.iat[i_b,df_player_hist.columns.get_loc('num_wins')]

            if match_a > match_b:
                df_player_hist.iat[i,df_player_hist.columns.get_loc('win_loss')] = 'W'
            else:
                df_player_hist.iat[i,df_player_hist.columns.get_loc('win_loss')] = 'L'

        #Can't determine a win/loss for the final iteration so we have to drop the last row
        else:
            df_player_hist = df_player_hist[:-1]
        
    #Performing Groupby Operation To Determine Monthly Statistics
    df_player_hist = df_player_hist.groupby(['match_date','win_loss'])['win_loss'].count()
    df_player_hist = df_player_hist.groupby([pd.Grouper(level='match_date',freq='1M'),'win_loss']).sum().unstack('win_loss')
    df_player_hist.index = pd.to_datetime(df_player_hist.index).strftime('%b-%y')

    #------------------------------------------------#

    #-------------------------------------------------------------------------------------------------------------------------------------

    #--------------Player Match History-------------#
    data = 'matches'        #Returning Match Data
    leaderboard_id = 3      #3 = Random Map Leaderboard
    count = 10             #Extracting this many entries from table
    since = date_today_unix #Extracting Match History From This Date
    top_player_id = rating_top.loc[0,'steam_id'] #Nesting the top player id from previous API call into the following request
    top_player_name = rating_top.loc[0,'name']   #Grabbing name to place into chart title
    requestAddr = f'https://aoe2.net/api/player/{data}?game=aoe2de&leaderboard_id={leaderboard_id}&steam_id={top_player_id}&count={count}&since={since}'
    dict = api2df(requestAddr)

    #Parsing Player Match History into Table
    df_player_matches = pd.DataFrame.from_dict(dict,orient='columns')
    filename='matches.csv'
    filepath=pathing(filename)
    df_player_matches.to_csv(filepath)
    print(df_player_matches)

    #Parsing the civilisation data into a dataframe
    # dict = df.loc[:,'players']
    # df_player_matches = pd.DataFrame.from_dict(dict)
    # print(df_player_matches)

    #-----------------------------------------------#

    #-------------------------------------------------------------------------------------------------------------------------------------

    #-------------------------------Plotting and Configuring Chart-------------------------------------
    #Making a colormap to map onto the bar chart
    clist = [(0, "green"), (0.5, "orange"), (1, "red")]
    rvb = mcolors.LinearSegmentedColormap.from_list('',clist)
    df_arange = np.arange(rating_top.shape[0]).astype(float)
    
    #Setting Axis Limits
    lhs_axis_lim = 2450
    rhs_axis_lim = 2575

    #Plotting a Leaderboard/Ratings Chart
    fig, ax = plt.subplots(figsize=(12,7))
    ax = rating_top.plot('name','rating',kind='barh',ax=ax,width=0.8,color=rvb(df_arange/rating_top.shape[0]),edgecolor='darkslategrey')
    ax.set_title(f'AoE2DE Ratings - Random Map - Top 15 Players',loc='left',fontsize=16)
    plt.text(rhs_axis_lim,15,f'Ratings Extracted on {date_today_str}',fontsize =9,horizontalalignment='right')

    #Inserting AoE2 Logo on Existing Chart
    filename = 'AoE2Logo.png'
    filepath = pathing(filename)
    image = Image.open(filepath) #Reading Image Data
    [width,height] = image.size
    im_zoom = 0.15
    imagebox = OffsetImage(image,zoom=im_zoom)
    aoe2Logo = AnnotationBbox(imagebox 
                            ,[rhs_axis_lim,0],xybox=(-width*im_zoom,height*im_zoom/1.225)
                            , boxcoords='offset points', box_alignment=(0,1)
                            ,pad=0
                            )
    ax.add_artist(aoe2Logo)

    #Adjusting Plot Parameters
    ax.get_legend().remove()
    ax.set_ylabel('')
    plt.xlim(lhs_axis_lim,rhs_axis_lim)
    plt.xticks([])
    plt.subplots_adjust(left=0.25) #adjusting plot to prevent y-tick cutoff

    #----Data Labels-----
    text_offset = 0.175 #Text Offset Distance
    for i in range(rating_top.shape[0]):
        rating_value = rating_top.iat[i,rating_top.columns.get_loc('rating')]
        rating_pos = rating_value - 6.1
        if rating_value != 0:
            rating_value = int(rating_value)
            rating_pos = int(rating_pos)
            ax.text(rating_pos,i-text_offset,rating_value,fontsize=10,color='white',weight='bold')
    #----Data Labels-----

    #Saving graph
    filename = 'AoE2_Ratings.png'
    filepath = pathing(filename)
    plt.savefig(filepath,dpi=100)
    # plt.show()

    #-------------------------------------------------------------------------------------------------------------------------------------

    #Plotting a Player History Chart
    upper_ylim = 60 #Setting upper bound of y-axis

    fig, ax = plt.subplots(figsize=(12,5))
    ax = df_player_hist.plot(kind='bar',ax=ax,width=0.8,edgecolor='darkslategrey')
    ax.set_title(f'AoE2DE Match History - Random Map - {top_player_name}',loc='left',fontsize=16)
    plt.text(df_player_hist.shape[0]-0.35,upper_ylim+1,f'Match History Extracted on {date_today_str}',fontsize =9,horizontalalignment='right')
    ax.set_ylabel('Number of Matches', fontsize=10)
    ax.set_xlabel('')
    ax.tick_params(axis='x',labelrotation= 0,labelsize=10)
    plt.ylim(0,upper_ylim) #Setting y-lim
    plt.show()


if __name__ == "__main__":
    main()