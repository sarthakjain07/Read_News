from win32com.client import Dispatch # used to make pronounce function using sapi voice
import requests # used to request data from WEB
import json # used to convert strings in python datatypes
import datetime
def pronounce(str):
    '''This function is used to make pronounce any string given to it'''
    speak = Dispatch("SAPI.SpVoice")
    speak.Speak(str)

def wishMe(name):
    '''This function is used to make the program wishing u according to time'''
    current_hour=int(datetime.datetime.now().hour)
    if current_hour>=0 and current_hour<=12:
        pronounce(f"Good morning {name}")
    elif current_hour>=12 and current_hour<=18:
        pronounce(f"Good afternoon {name}")
    else:
        pronounce(f"Good evening {name}")    

if __name__ == '__main__':
    wishMe("Sarthak")
    print("News headlines for today...Please listen carefully:")
    pronounce("News headlines for today...Please listen carefully")
    url="http://newsapi.org/v2/top-headlines?country=in&apiKey=d2d7006237a34aa3984f3f925648c4a5" # This is the link of News API
    news=requests.get(url).text # converting the request made in the form of text
    news_dict=json.loads(news) # converting news into python object
    arts=news_dict['articles'] # accessing articles...
    for article in arts:
        print(article['title']) # getting title from articles
        pronounce(article['title'])
        pronounce("Moving on to next headline....")
    pronounce("Thanks for listening")