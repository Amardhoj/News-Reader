import requests
import json

def speak(str):
    from win32com.client import Dispatch
    speak = Dispatch("SAPI.SpVoice")
    speak.Speak(str)

if __name__ == '__main__':
    speak("News for today")
    url = "https://newsapi.org/v2/top-headlines?country=us&apiKey=################################"
    news = requests.get(url).text
    news_dict = json.loads(news)
    data = news_dict["articles"]

    for index, article in enumerate(data):
        print(article["title"])
        speak(article["title"])
        if index != (len(data)-1):
            speak("Moving on to the next news...")

    speak("That's it for the day. Thanks for listening")