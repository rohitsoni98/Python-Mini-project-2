import requests
import json


def s(str):
    from win32com.client import Dispatch

    s = Dispatch("SAPI.SpVoice")

    s.Speak(str)

if __name__ == '__main__':
    s("news for today.. Lets begin")
    url = "https://newsapi.org/v2/top-headlines?country=in&apiKey=37e19dbce4cc4aaaa11e6d27d954cbdd"
    news = requests.get(url).text
    news_dict = json.loads(news)
    print(news_dict["articles"])
    arts = news_dict['articles']
    for article in arts:
        s(article['title'])
        s('moving on to the next news')
    s('Thanks for listening')
