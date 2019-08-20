
# coding: utf-8

# In[1]:


# YT
from apiclient.discovery import build #pip install google-api-python-client
from apiclient.errors import HttpError #pip install google-api-python-client
from oauth2client.tools import argparser #pip install oauth2client
from datetime import datetime
import pandas as pd #pip install pandas
import numpy as np
import win32com.client
import requests
import json
# + TW
import time
import sys
import tweepy # pip install tweepy
import urllib.request
import openpyxl
import xlsxwriter
# + YT comment_analytics
from collections import Counter
import operator
#import nltk
#from nltk.corpus import stopwords # pip install nltk
import re


# In[2]:


# API를 사용하기 위한 key들
if np.random.randint(1, 3, size=1)[0] == 1:
    yt_api_key = "AIzaSyDFx9O3qfqtT-j-VR9vqd9mSXOFGAB7Ick" 
else:
    yt_api_key = "AIzaSyA6tLM4V5Aj90YXfCiQMnIv9dGq0kQMOIk"

tw_consumer_key="125Z3IxmGV1ZObDeIzDVvLyI0"
tw_consumer_secret="BxZCgYCnzE9OHrRH7NVUJUCob65p1RqypjHrpk5yseZCfuiC8A"
tw_access_token="249621931-dXFJt8Y5ZsGHAQwI8ygFsAHJHeh7vyDA93vKlGZZ"
tw_access_token_secret="thHCJMGNJV10iB83YDTbv2CG15pgTpAunGDslgyTBx9Ko"
tw_auth = tweepy.OAuthHandler(tw_consumer_key, tw_consumer_secret)
tw_auth.set_access_token(tw_access_token, tw_access_token_secret)
tw_api = tweepy.API(tw_auth)


# In[3]:


# TWITTER

def post_stat(url, user, timeline, num, row):
       
    favorite_count = []
    retweet_count = []
    reply_count = []
    date = []
    address = []
    sa_count = []
    sub_count = []
    
    for k in range(0,num):
        try:
            status = timeline[k]
        except:
            continue
        numofpost = k + 1
        
        j_results = json.loads(json.dumps(status._json))   
        sub_count.append(j_results['user']['followers_count'])
        favorite_count.append(int(j_results['favorite_count']))
        retweet_count.append(int(j_results['retweet_count']))
        date.append(time.strftime('%Y-%m-%d %H:%M:%S', time.strptime(j_results['created_at'],'%a %b %d %H:%M:%S +0000 %Y')))
        if num == 1:
            date = str(date)[2:12]
            address.append(url.split('https://')[1])
        else:
            address.append('https://twitter.com/' + user.screen_name +'/status/' + j_results['id_str'])
            #print('https://twitter.com/' + user.screen_name +'/status/' + j_results['id_str'])
        
        request = urllib.request.Request('https://twitter.com/' + str(j_results['user']['id']) + '/status/' + str(j_results['id']))
        response = urllib.request.urlopen(request)
        temp = response.read().decode('utf-8')
        
        try:
            reply_count.append(int((temp.split('reply-count-aria-' + str(j_results['id']) + '" data-aria-label-part>답글 ')[1].split('개')[0]).replace(',','')))
        except:
            reply_count.append(int(temp.split('reply-count-aria-' + str(j_results['id']) + '" >답글 ')[1].split('개')[0]))
            
        sa_count.append(int(j_results['favorite_count']) + int(j_results['retweet_count']) + int(reply_count[k]))
             
    try:    
        name = [user.screen_name] * numofpost
    except:
        name = [user]

    dump = [0] * numofpost
    Row = [row] * numofpost
    
    if num == 1:
        final = pd.DataFrame(np.column_stack([name, date[:10], '', 'N/A', favorite_count, reply_count, retweet_count, address]), 
                         columns = ['name','post_date','UpdateDate','view_count','like_count','comment_count','share_count', 'url'])
    else:
        final = pd.DataFrame(np.column_stack([Row, dump, dump, sub_count, dump, dump, dump, dump, dump, dump, dump, sa_count]), 
                         columns = ['number','YT', 'IG', 'TW', 'FB', 'UpdateDate', 'View (YT)','View (IG)','YT SA','IG SA(Image)','IG SA(Video)','TW SA'])
        final = final.groupby('number').mean()
        
    return final

def comment_tw(url):
    
    try:
        if url.rsplit('?',1)[1] != None:
            url = url.rsplit('?',1)[0]
    except:
        pass
    
    text = []
    name_list = []
    
    name = url.split('https://twitter.com/')[1].split('/')[0]
    post_id = url.rsplit('/',1)[1]
    
    request = urllib.request.Request('https://twitter.com/' + name + '/status/' + post_id)
    response = urllib.request.urlopen(request)
    #response.body.scrollTop = response.body.scrollHeight;
    #print(response)
    temp = response.read().decode('utf-8')
    
    
    try:
        temp = temp.split('data-aria-label-part="0">')
        length = len(temp)
        if length > 2:
            for i in range(1, length):      
                text.append(temp[i].split('</p>',1)[0])
                name_list.append(name)

            final = pd.DataFrame(np.column_stack([name_list, text])
                                 , columns = ['name', 'comment'])
        else:
            final = pd.DataFrame(np.column_stack([name, 'N/A'])
                                 , columns = ['name', 'comment'])
    except:
        final = pd.DataFrame(np.column_stack([name, 'N/A'])
                             , columns = ['name', 'comment'])
        
    return final


def tw(url, num, row):
    
    try:
        if url.rsplit('?',1)[1] != None:
            url = url.rsplit('?',1)[0]
    except:
        pass
    #try:
    if url.rsplit('/',2)[1] == 'status':
        try:
            data = tw_api.get_status(id = url.rsplit('/',1)[1])
            timeline = []
            timeline.append(data)
            result = post_stat(url, url.rsplit('/',3)[1], timeline, 1, row)
        except:
            result = pd.DataFrame(np.column_stack(['error', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', url]), 
                         columns = ['name','post_date','UpdateDate','view_count','like_count','comment_count','share_count', 'url']) 
            print('error in ' + url)
            
    else:
        try:
            name = url.split('https://twitter.com/')[1].split('/')[0]
            user = tw_api.get_user(name)
            timeline = tw_api.user_timeline(screen_name = name, count = num*3, include_rts = False, include_entities = True)
            result = post_stat(url, user, timeline, num, row)
        except:
            result = pd.DataFrame(np.column_stack([row, 0, 0, -3, 0, 0, 0, 0, 0, 0, 0, 0]), 
                         columns = ['number','YT', 'IG', 'TW', 'FB', 'UpdateDate', 'View (YT)','View (IG)','YT SA','IG SA(Image)','IG SA(Video)','TW SA'])
            result = result.groupby('number').mean()
            print('error in ' + url)
    
    return result


# In[4]:


# INSTAGRAM

def to_json(url):
    data = urllib.request.urlopen(url).read().decode('utf-8')
    return json.loads(data)

def insta_channel_stat(url, num, row):
    view_count = []
    like_count = []
    comment_count = []
    sub_count = []
    date = []
    address = []
    content_type = []
    sa_i = []
    sa_v = []
    image_count = 0
    video_count = 0
    url_temp = url
        
    count = 0
    url = url + "?__a=1" # ?__a=1 는 JSON으로 변환시켜줌
     
    try:
        while count <5 and (image_count < num or video_count < num): # count는 페이지를 로드하는 숫자.
            # the count clicks "Load more page" 
            if count  == 0: # first iteration

                data = to_json(url)
                sub = int(data['user']['followed_by']['count'])
                
                cursor = data['user']['media']['page_info']['end_cursor'] # 'Load more page' cursor. this allows to load more pages

                for i in range(0,12): # The first page shows 12 pictures
                    
                    df = data['user']['media']['nodes'][i] # for each post
                    
                    if df['is_video'] == False and image_count < num:     
                        like_count.append(int(df['likes']['count']))
                        comment_count.append(int(df['comments']['count']))
                        date.append(pd.to_datetime(df['date'], unit = 's'))
                        address.append('https://www.instagram.com/p/' + df['code'])
                        sub_count.append(sub)
                        view_count.append(0)                       
                        content_type.append(0)
                        sa_i.append(int(df['likes']['count']) + int(df['comments']['count']))
                        sa_v.append(0)
                        image_count += 1

                    elif df['is_video'] == True and video_count < num:
                        like_count.append(int(df['likes']['count']))
                        comment_count.append(int(df['comments']['count']))
                        date.append(pd.to_datetime(df['date'], unit = 's'))
                        address.append('https://www.instagram.com/p/' + df['code'])
                        sub_count.append(sub)
                        view_count.append(int(df['video_views']))
                        content_type.append(1)
                        sa_i.append(0)
                        sa_v.append(int(df['likes']['count']) + int(df['comments']['count']))
                        video_count += 1

                    else:
                        continue
                        

            else:   
                url = url + "&max_id=" + cursor # 'Load More pagge' 를 눌러서 load more posts 
                        # &max_id와 cursor를 사용해서 페이지를 더 당겨옴

                data = to_json(url)

                cursor = data['user']['media']['page_info']['end_cursor'] # iter through end cursor 

                for i in range(0,11):

                    df = data['user']['media']['nodes'][i] # for each post
                    
                    if df['is_video'] == False and image_count < num:                      
                        like_count.append(int(df['likes']['count']))
                        comment_count.append(int(df['comments']['count']))
                        date.append(pd.to_datetime(df['date'], unit = 's'))
                        address.append('https://www.instagram.com/p/' + df['code'])
                        sub_count.append(sub)
                        view_count.append(0)                       
                        content_type.append(0)
                        sa_i.append(int(df['likes']['count']) + int(df['comments']['count']))
                        sa_v.append(0)
                        image_count += 1

                    elif df['is_video'] == True and video_count < num:                    
                        like_count.append(int(df['likes']['count']))
                        comment_count.append(int(df['comments']['count']))
                        date.append(pd.to_datetime(df['date'], unit = 's'))
                        address.append('https://www.instagram.com/p/' + df['code'])
                        sub_count.append(sub)
                        view_count.append(int(df['video_views']))
                        content_type.append(1)
                        sa_i.append(0)
                        sa_v.append(int(df['likes']['count']) + int(df['comments']['count']))
                        video_count += 1

                    else:
                        continue

            count += 1
    except:
        pass
    
    dump = [0] * (image_count + video_count)
    Row = [row] * (image_count + video_count)
    
    
    final = pd.DataFrame(np.column_stack([Row, dump, sub_count, dump, dump, dump, dump, view_count, dump, sa_i, sa_v, dump, content_type]), 
                         columns = ['number','YT', 'IG', 'TW', 'FB', 'UpdateDate', 'View (YT)','View (IG)','YT SA','IG SA(Image)','IG SA(Video)','TW SA', 'type'])
    
    # 이미지와 비디오를 각각 나누어 평균을 냄
    final = final.groupby('type').mean()
    # 나눈 후에 두 이미지와 비디오 그룹을 합침
    
    if len(final) == 2:
        final['IG'].loc[0] /= 2
        final['IG'].loc[1] /= 2
    
    final = final.groupby('number').sum()
    
    
    return final

def insta_single_stat(url):
    url = url.split("?")[0]
    address = url

    url = url + "?__a=1" # JSON으로 변환
        #cursor = data['graphql']['shortcode_media']['edge_media_to_comment']['page_info']['end_cursor'] #cursor
        
    data = to_json(url)['graphql']['shortcode_media']
    comment_count = int(data['edge_media_to_comment']['count'])        
    like_count = int(data['edge_media_preview_like']['count'])
    name = data['owner']['username']
    try:
        date = data['edge_media_to_comment']['edges'][0]['node']['created_at']
        #print('created')
    except:
        date = data['taken_at_timestamp']
        #print('timestamp')
        
    #print(type(date))
    #print(date)
    
    date = str(pd.to_datetime(date, unit = 's'))[:10]
    
    #df['just_date'] = df['dates'].dt.date
    
    if data['is_video'] == True:
        content_type = 'video'
        view_count = int(data['video_view_count'])
    else:
        content_type = 'image'
        view_count = 'N/A'     
    
    #print(type(date))
    
    final = pd.DataFrame(np.column_stack([name, date, '', view_count, like_count, comment_count, 'N/A', address]), 
                         columns = ['name','post_date','UpdateDate','view_count','like_count','comment_count','share_count', 'url']) 
   
    return final

def comment_ig(url):
    
    channel_url = url
    url = url.split("?")[0]

    url = url + "?__a=1" # JSON으로 변환

    data = to_json(url)
    comment_text = data['graphql']['shortcode_media']['edge_media_to_comment']['edges']
    name = data['graphql']['shortcode_media']['owner']['username']
    
    id_name = []
    url_list = []
    comment_date = []
    author = []
    text = []


    for i in range(0,1):
    # 댓글 저장
        if i != 0:
            cursor = data['graphql']['shortcode_media']['edge_media_to_comment']['page_info']['end_cursor']
            temp_url = url + "&max_id=" + cursor
            data = to_json(temp_url)
            comment_text = data['graphql']['shortcode_media']['edge_media_to_comment']['edges']

            for i in range(0, len(comment_text)):
                text.append(comment_text[i]['node']['text']) # Append each comment
                id_name.append(name)
                url_list.append(channel_url)
                comment_date.append(str(pd.to_datetime(comment_text[i]['node']['created_at'], unit = 's'))[:10])
                author.append(comment_text[i]['node']['owner']['username'])

        else:
            data = to_json(url)
            comment_text = data['graphql']['shortcode_media']['edge_media_to_comment']['edges']
            for i in range(0, len(comment_text)):
                text.append(comment_text[i]['node']['text']) # Append each comment       
                id_name.append(name)
                url_list.append(channel_url)
                comment_date.append(str(pd.to_datetime(comment_text[i]['node']['created_at'], unit = 's'))[:10])
                author.append(comment_text[i]['node']['owner']['username'])
                

    final = pd.DataFrame(np.column_stack([id_name, url_list, comment_date, author,text])
                         , columns = ['id', 'url', 'date', 'author_name','comment'])
        
    return final

def ig(url, num, row):
    
    try:
        if url.rsplit('?',1)[1] != None:
            url = url.rsplit('?',1)[0]            
    except:
        pass
    
    
    if url.split('https://www.instagram.com/')[1].split('/')[0] == 'p':
        try:
            result = insta_single_stat(url)
        except:
            result = pd.DataFrame(np.column_stack(['error', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', url]), 
                         columns = ['name','post_date','UpdateDate','view_count','like_count','comment_count','share_count', 'url']) 
            print('error in ' + url)
    else:
        try:
            result = insta_channel_stat(url, num, row)
        except:
            result = pd.DataFrame(np.column_stack([row, 0, -3, 0, 0, 0, 0, 0, 0, 0, 0, 0]), 
                         columns = ['number','YT', 'IG', 'TW', 'FB', 'UpdateDate', 'View (YT)','View (IG)','YT SA','IG SA(Image)','IG SA(Video)','TW SA'])
            result = result.groupby('number').mean()
            print('error in ' + url)
  
    return result
        


# In[5]:


# YOUTUBE

def channel(user_name, api_key, num):
    # Get unique id for each user.
    
    parameters = {
                      'forUsername': user_name,
                      'part': 'contentDetails',
                      "maxResults" : num,
                      "key": api_key

                      }
    url = "https://www.googleapis.com/youtube/v3/channels"  # This is the url to extract the data from YouTUbe

    try:
        
        page = requests.request(method="get", url=url, params=parameters) # Get JSON data
        j_results = json.loads(page.text) # Convert JSON to dictionary
        return j_results['items'].pop()['contentDetails']['relatedPlaylists']['uploads'] # This returns the unique ID of the user

    except:
        parameters = {
          'id': user_name, # id is required instead of forUsername
          'part': 'contentDetails',
          "maxResults" : num,
          "key": api_key
          }
        
        page = requests.request(method="get", url=url, params=parameters)
        j_results = json.loads(page.text)
        return j_results['items'].pop()['contentDetails']['relatedPlaylists']['uploads']
    
def playlist(user_id, api_key, num):
    
    # Get playlists 
    
    parameters = {"part": "id",
              "part": "snippet",
              'part': 'contentDetails',
              "playlistId" : user_id,
              "maxResults" : num,
              "key": api_key
              }

    url = "https://www.googleapis.com/youtube/v3/playlistItems" # URL to get playlists
    page = requests.request(method="get", url=url, params=parameters) # Get JSON data
    j_results = json.loads(page.text) # Convert JSON to dictionary

    video_id = [] # Video id
    video_date = [] # video published date
    
    for i in range(0,num):
        
        video_id.append(j_results['items'][i]['contentDetails'].get("videoId")) # Get video id 
        video_date.append(j_results['items'][i]['contentDetails'].get("videoPublishedAt")[:-14]) # Get video id 

    return video_id, video_date

def video_stat(video_id, video_date, user_name, api_key, num, row):
    
    # id 가져옴
    parameters2 = {
                      'forUsername': user_name,
                      'part': 'contentDetails',
                      "maxResults" : num,
                      "key": api_key

                      }
    url2 = "https://www.googleapis.com/youtube/v3/channels"  # This is the url to extract the data from YouTUbe

    try:
        
        page2 = requests.request(method="get", url=url2, params=parameters2) # Get JSON data
        j_results2 = json.loads(page2.text) # Convert JSON to dictionary
        user_id = j_results2['items'].pop()['id'] # This returns the unique ID of the user

    except:
        parameters2 = {
          'id': user_name, # id is required instead of forUsername
          'part': 'contentDetails',
          "maxResults" : num,
          "key": api_key
          }
        
        page2 = requests.request(method="get", url=url2, params=parameters2)
        j_results2 = json.loads(page2.text)
        user_id = j_results2['items'].pop()['id']
    
    # subscriber count 가져옴
    request = urllib.request.Request("https://www.googleapis.com/youtube/v3/channels?part=statistics&id=" + user_id + "&key=" + api_key)
    response = urllib.request.urlopen(request)
    j = json.loads(response.read().decode('utf-8'))
    sub = j['items'].pop()['statistics']['subscriberCount']
    
    # Get each video data
    
    parameters = {"part": "id",
              'part': 'contentDetails',
              'part': 'statistics',
              "id" : "",
              "maxResults" : num,
              "key": api_key
              }

    url = "https://www.googleapis.com/youtube/v3/videos" # URL to get video detail
#    page = requests.request(method="get", url=url, params=parameters) # Get JSON 
#    j_results = json.loads(page.text) # Convert JSON to dictionary

    view_count = [] # View Count list
    like_count = [] # Like Count LIst
    dislike_count = [] #Dislike Count List
    favorite_count = [] # Favorite count list
    comment_count = [] # comment count list
    sa_count = []
    address = []
    sub_count = []
    name = [user_name] * len(video_id) # Generate a list of repeating names of the user
    dump = [0] * num
    Row = [row] * num

    for vi_id in video_id: # Loop through each video ID
        parameters['id'] = vi_id # 
        page = requests.request(method="get", url=url, params=parameters) # Get JSON
        j_results = json.loads(page.text) # Convert JSON to dictionary
        
        j_results = j_results['items'].pop() # pop out the list
        
        # Append the each item to the corresponding lists       
        view_count.append(int(j_results['statistics']['viewCount'])) 
        like_count.append(int(j_results['statistics']['likeCount']))
        dislike_count.append(int(j_results['statistics']['dislikeCount']))
        favorite_count.append(int(j_results['statistics']['favoriteCount']))
        comment_count.append(int(j_results['statistics']['commentCount']))
        address.append('https://www.youtube.com/watch?v=' + j_results['id'])
        sa_count.append(int(j_results['statistics']['likeCount']) + int(j_results['statistics']['commentCount']))
        sub_count.append(int(sub))
        
    # Create dataframe using the lists
    final = pd.DataFrame(np.column_stack([Row, sub_count, dump, dump, dump, dump, view_count, dump, sa_count, dump, dump, dump]), 
                         columns = ['number','YT', 'IG', 'TW', 'FB', 'UpdateDate', 'View (YT)','View (IG)','YT SA','IG SA(Image)','IG SA(Video)','TW SA'])
    
    final = final.groupby('number').mean()
    return final #return the dataframe

def single_stat(url, api_key):
    parameters = {"part": "id",
              'part': 'contentDetails',
              'part': 'statistics',
              "id" : "",
              "maxResults" : num,
              "key": api_key
              }
        
    parameters['id'] = url.split('=')[1]
    url = "https://www.googleapis.com/youtube/v3/videos?id="+ url.rsplit('=',1)[1] +"&key=" + api_key + "&fields=items(id,snippet(channelTitle,publishedAt,channelId,title,categoryId),statistics)&part=snippet,statistics"
    page = requests.request(method="get", url=url, params=parameters) # Get JSON
    j_results = json.loads(page.text) # Convert JSON to dictionary
    j_results = j_results['items'].pop() # pop out the list
    
    name = j_results['snippet']['channelTitle']
    video_date = j_results['snippet']['publishedAt'][:-14]
    view_count = j_results['statistics']['viewCount']
    like_count = j_results['statistics']['likeCount']
    dislike_count = j_results['statistics']['dislikeCount']
    favorite_count = j_results['statistics']['favoriteCount']
    comment_count = j_results['statistics']['commentCount']
    address = 'https://www.youtube.com/watch?v=' + j_results['id']

    final = pd.DataFrame(np.column_stack([name, video_date, '', view_count, like_count, comment_count, 'N/A', address]), 
                         columns = ['name','post_date','UpdateDate','view_count','like_count','comment_count','share_count', 'url']) 
   
    return final #return the dataframe


# In[6]:


def comment_yt(video_url, api_key): # Video ID 와 api_key를 받음
    url = "https://www.googleapis.com/youtube/v3/commentThreads" # youtube comment api
    
    try:
        if video_url.rsplit('&t=', 1)[1] != None:
                video_url = video_url.rsplit('&t=',1)[0]
    except:
        pass
    
    try:
        if url.rsplit('&',1)[1][:7] == 'feature':
            url = url.rsplit('&',1)[0]
    except:
        pass
    
    if url.rsplit('/',1)[1] == 'featured' or url.rsplit('/',1)[1][:6] == 'videos' or url[-1] == '/':
        url = url.rsplit('/',1)[0]
    
    video_id = video_url.split("watch?v=")[1]
    #print(video_id)
    count = 0
    comment_list = []
    comment_date = []
    video_id_name = []
    author = []
    url_list = []
    while True:
        parameters = {
                     'part': 'snippet',
                      "videoId" : video_id,
                      "key": api_key,
                        "pageToken":""
        }

        # 첫번째 페이지는 로드가 필요 없어서 "" 로 유지
        if count== 0:
            parameters['pageToken'] = ""
        else:
            # After the first loop, grab cursor to scroll down more for more results
            parameters['pageToken'] = cursor


        count = 1 # Count는 첫번째 페이지인가 아닌가를 판단
                # Count가 1이 되어서 첫번째 패이지가 아닌걸로 판단해서 Cursor를 붙힘

        try:
            # Load Page
            page = requests.request(method="get", url=url, params=parameters)
            j_results = json.loads(page.text)  
            # Loop through items for each comment
            for i in range(0,len(j_results['items'])):
           
                # 작석자 ID
                author.append(j_results['items'][i]['snippet']['topLevelComment']['snippet']['authorDisplayName'])
                #댓글
                comment_list.append(j_results['items'][i]['snippet']['topLevelComment']['snippet']['textDisplay'])
                #댓글 단 날짜
                comment_date.append(j_results['items'][i]['snippet']['topLevelComment']['snippet']['updatedAt'])
                video_id_name.append(video_id) # 비디오 ID 
                url_list.append(video_url)
                
            cursor = j_results['nextPageToken'] # Create Cursor

        except:
            break

    final = pd.DataFrame(np.column_stack([video_id_name, url_list, comment_date, author,comment_list])
                         , columns = ['video_id', 'url', 'date', 'author_name','comment'])
    
    return final

def comment_to_word(data): # def YouTube_comment의 return value를 사용
    # Special Character모두 지우기
    data['comment_special'] = data['comment'].map(lambda x: re.sub('[^a-zA-Z0-9 ]', '', x))
    
    # 문장을 한개의 list로 합침
    text = " ".join(data['comment_special'].str.lower()).split(" ")
    text_count = Counter(text) # 단어 숫자를 dictionary로 변환
    #stop = set(stopwords.words('english')) # stopword 
    stop = {'a', 'about', 'above', 'after', 'again', 'against', 'ain', 'all', 'am', 'an', 'and', 'any', 'are', 'aren', 'as', 'at',
            'be', 'because', 'been', 'before', 'being', 'below', 'between', 'both', 'but', 'by', 'can', 'couldn', 'd', 'did', 'didn',
            'do', 'does', 'doesn', 'doing', 'don', 'down', 'during', 'each', 'few', 'for', 'from', 'further', 'had', 'hadn', 'has',
            'hasn', 'have', 'haven', 'having', 'he', 'her', 'here', 'hers', 'herself', 'him', 'himself', 'his', 'how', 'i', 'if',
            'in', 'into', 'is', 'isn', 'it', 'its', 'itself', 'just', 'll', 'm', 'ma', 'me', 'mightn', 'more', 'most', 'mustn', 'my',
            'myself', 'needn', 'no', 'nor', 'not', 'now', 'o', 'of', 'off', 'on', 'once', 'only', 'or', 'other', 'our', 'ours', 'ourselves',
            'out', 'over', 'own', 're', 's', 'same', 'shan', 'she', 'should', 'shouldn', 'so', 'some', 'such', 't', 'than', 'that', 'the',
            'their', 'theirs', 'them', 'themselves', 'then', 'there', 'these', 'they', 'this', 'those', 'through', 'to', 'too', 'under',
            'until', 'up', 've', 'very', 'was', 'wasn', 'we', 'were', 'weren', 'what', 'when', 'where', 'which', 'while', 'who', 'whom',
            'why', 'will', 'with', 'won', 'wouldn', 'y', 'you', 'your', 'yours', 'yourself', 'yourselves'}
    
    del data['comment_special']
    
    # Stopword인지 아닌지 구분해서 stopword가 아니면 dictionary 만듬
    text_final = {}
    for key, value in text_count.items():
        if key not in stop:
            text_final[key] = value
    # sort
    text_final_2 = sorted(text_final.items(), reverse=True, key = operator.itemgetter(1))
    # Create dataframe
    final =pd.DataFrame.from_records(text_final_2,columns = ['word','number'])
    final['number'] = pd.to_numeric(final['number'])
    
    name_list = []
    for k in range(0,len(final)):
        name_list.append((data['video_id'][0]))
        
    final['name'] = pd.DataFrame(np.column_stack([name_list]), 
                         columns = ['name'])
        
    final = final[['name', 'word', 'number']]
    
    return final

def yt(url, api_key, num, row):

    # https://www.youtube.com/user/DJCHETASOFFICIAL/adfads 같은 것들을 https://www.youtube.com/user/DJCHETASOFFICIAL로 바꿈
    try:
        if url.rsplit('&t=', 1)[1] != None:
                url = url.rsplit('&t=',1)[0]
    except:
        pass
    
    try:
        if url.rsplit('&',1)[1][:7] == 'feature':
            url = url.rsplit('&',1)[0]
    except:
        pass
    
    if url.rsplit('/',1)[1] == 'featured' or url.rsplit('/',1)[1][:6] == 'videos' or url[-1] == '/':
        url = url.rsplit('/',1)[0]

        
#    user_url = [] # URL
#    user_name = [] # User name ex) DJCHETASOFFICIAL, ZEDDVEVO ...
#    user_type = [] # User type ex) user, channel, watch
        
    # /를 기준으로 2번 나눠줌 ex) https://www.youtube.com/user/DJCHETASOFFICIAL->'https://www.youtube.com', 'user', 'DJCHETASOFFICIAL'
    user_url = url.rsplit('/',2) 
             
    # Type이 user 나 channel일 경우 user_name과 user_type에 알맞은 값을 넣어준다
    if user_url[1] == 'user' or user_url[1] == 'channel':
        user_name = user_url[2]
        user_type = user_url[1]        

    # Type이 watch인 경우 name에는  - type에는 watch를 넣어준다 
    elif user_url[2].split('?')[0] == 'watch':
        user_name = ('-')
        user_type = ('watch')       
        
    # user, channel, watch 다 아닐 경우 name과 type 모두에 undefine을 넣어준다
    else:
        user_name = ('undefine')
        user_type = ('undefine')
    
    result = 0 #init
    
    if(user_type == 'user' or user_type == 'channel'):
        try:
            user_ids = channel(user_name, api_key, num)
            video_id, video_date = playlist(user_ids, api_key, num)
            result = video_stat(video_id,video_date, user_name, api_key, num, row)
            return result
        except:
            result = pd.DataFrame(np.column_stack([row,-3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]), 
                         columns = ['number','YT', 'IG', 'TW', 'FB', 'UpdateDate', 'View (YT)','View (IG)','YT SA','IG SA(Image)','IG SA(Video)','TW SA'])
            result = result.groupby('number').mean()
            print('error in ' + url)
        
    elif(user_type == 'watch'):
        try:
            result = single_stat(url, api_key)
        except:
            result = pd.DataFrame(np.column_stack(['error', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', url]), 
                         columns = ['name','post_date','UpdateDate','view_count','like_count','comment_count','share_count', 'url'])
            print('error in ' + url)
        
    return result


# In[12]:


# 'C:\\py_do_youtube\\input.xlsx'을 열어 활용할 수 있도록 세팅
excel = win32com.client.Dispatch("Excel.Application")
#excel.Visible = False
wb = excel.Workbooks.Open('C:\\SocialMediaData\\input.xlsx')
ws = wb.ActiveSheet


# 빈 셀이 나올 때 까지
url = []
url_num = 0
row_offset = 1 # 첫 번째 url의 행 번호
col_offset = 1 # 첫 번째 url의 열 번호
# excel(row_offset,col_offset)부터 url 시작 url에 모든 excel의 모든 url을 저장해둠
url.append(ws.Cells(url_num+row_offset,col_offset).Value)
url.append(ws.Cells(url_num+row_offset,col_offset+1).Value)
url.append(ws.Cells(url_num+row_offset,col_offset+2).Value)
loop = True
is_post = 0
is_comment = 0

while loop == True:
     # 다음 셀을 확인하기 위해 num++를 해줌
    url_num += 1
    url.append(ws.Cells(url_num+row_offset,col_offset).Value)
    url.append(ws.Cells(url_num+row_offset,col_offset+1).Value)
    url.append(ws.Cells(url_num+row_offset,col_offset+2).Value)

    if url[url_num * 3 + 1] == None and url[url_num * 3 + 2] == None:
        is_post += 1
    if url[url_num * 3] == None and url[url_num * 3 + 1] == None and url[url_num * 3 + 2] == None:
        loop = False

# 모든 url이 첫 번째 col에만 있다면 개별 포스팅 url이고 아니라면 채널 url이라고 판단
if is_post != url_num:
    num = int(input("Enter a number: "))
    is_comment = 'n'
    
else:
    while is_comment != 'y' and is_comment != 'n':
        is_comment = input("Do you want to extract COMMENT data? (y/n)")
    num = 5 # 아무 값이나 넣음
    
    
    
init = False
############################################
comment_youtube = []
comment_analytics_youtube = []
comment_instagram = []
comment_twitter = []
############################################
for count in range(0,url_num*3):
    # row는 몇번째 row의 url값인지 저장해두는 변수
    row = int(count / 3) + 1

    try:
        if url[count].find('https://www.youtube.com') != -1:
            platform = 'youtube'
        elif url[count].find('https://twitter.com') != -1:
            platform = 'twitter'
        elif url[count].find('https://www.instagram.com') != -1:
            platform = 'instagram'
        else:
            print('error: ' + url[count])
            if is_post == url_num:
                final = final.append(pd.DataFrame(np.column_stack(['error', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', url[count]]), 
                         columns = ['name','post_date','UpdateDate','view_count','like_count','comment_count','share_count', 'url']))
            else:
                final = final.append(pd.DataFrame(np.column_stack([row, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]), 
                         columns = ['number','YT', 'IG', 'TW', 'FB', 'UpdateDate', 'View (YT)','View (IG)','YT SA','IG SA(Image)','IG SA(Video)','TW SA']).groupby('number').mean())
            continue
    except:
        continue
    
    if is_post != url_num:
        print(url[count] + '\t' + str(count + 1) + '/' + str(url_num * 3))
    
    else:
        print(url[count] + '\t' + str(int(count/3) + 1) + '/' + str(url_num))
    
    
    
    # 주소의 문제가 아닌 이유 모를 오류가 날 때가 있으니 에러의 경우 같은 작업을 한번 더 해봄
    try:
        if count == 0 or init == False:
            init  = True        
            if platform == 'youtube':
                final = yt(url[count], yt_api_key, num, row)                   
            elif platform == 'twitter':
                final = tw(url[count], num, row)            
            elif platform == 'instagram':
                final = ig(url[count], num, row)                 
            else:
                init = False
                continue
        else:
            if platform == 'youtube':           
                final = final.append(yt(url[count], yt_api_key, num, row))                      
            elif platform == 'twitter':
                final = final.append(tw(url[count], num, row))                
            elif platform == 'instagram':
                final = final.append(ig(url[count], num, row))            
            else:
                continue
    except:
        if count == 0 or init == False:
            init  = True        
            if platform == 'youtube':
                final = yt(url[count], yt_api_key, num, row)                   
            elif platform == 'twitter':
                final = tw(url[count], num, row)            
            elif platform == 'instagram':
                final = ig(url[count], num, row)                 
            else:
                init = False
                continue
        else:
            if platform == 'youtube':           
                final = final.append(yt(url[count], yt_api_key, num, row))                      
            elif platform == 'twitter':
                final = final.append(tw(url[count], num, row))                
            elif platform == 'instagram':
                final = final.append(ig(url[count], num, row))            
            else:
                continue
            
    if is_comment == 'y':
        try:
            if platform == 'youtube':
                if len(comment_youtube) == 0:
                    save = comment_yt(url[count], yt_api_key)
                    comment_youtube = save
                    comment_analytics_youtube = comment_to_word(save)
                else:
                    save = comment_yt(url[count], yt_api_key)
                    comment_youtube = comment_youtube.append(save)   
                    comment_analytics_youtube = comment_analytics_youtube.append(comment_to_word(save))
            elif platform == 'twitter':
                if len(comment_twitter) == 0:
                    comment_twitter = comment_tw(url[count])
                else:
                    comment_twitter = comment_twitter.append(comment_tw(url[count]))         
            elif platform == 'instagram':      
                if len(comment_instagram) == 0:
                    comment_instagram = comment_ig(url[count])           
                else:
                    comment_instagram = comment_instagram.append(comment_ig(url[count]))       
            else:
                continue
        except:
            pass
 #################################################################### """      
        
        
for col in final.columns:
    final[col] = pd.to_numeric(final[col], errors = "ignore")

pd.options.display.float_format = '{:,.0f}'.format
pd.set_option('max_rows', 1000)
            
try:           
    group = final.groupby('number').sum()
    group['UpdateDate'] = datetime.today().strftime("%Y-%m-%d")
    group['View (YT)'].astype(int)
    group['View (IG)'].astype(int)
    group['YT SA'].astype(int)
    group['IG SA(Image)'].astype(int)
    group['IG SA(Video)'].astype(int)
    group['TW SA'].astype(int)
    writer = pd.ExcelWriter('C:\\SocialMediaData\\output.xlsx', engine='xlsxwriter')
    group.to_excel(writer, sheet_name='Sheet1')
    print('done')
except:
    final['UpdateDate'] = datetime.today().strftime("%Y-%m-%d")
    writer = pd.ExcelWriter('C:\\SocialMediaData\\output.xlsx', engine='xlsxwriter')
    final.to_excel(writer, sheet_name='Sheet1')
    print('done')


try:
    comment_youtube.to_excel(writer, sheet_name='YT')
except:
    pass

try:
    comment_instagram.to_excel(writer, sheet_name='IG')
except:
    pass

try:
    comment_twitter.to_excel(writer, sheet_name='TW')
except:
    pass

try:
    comment_analytics_youtube.to_excel(writer, sheet_name='YT_WORD')
except:
    pass


writer.save()

k=input("press close to exit") 


# In[ ]:





# In[ ]:




