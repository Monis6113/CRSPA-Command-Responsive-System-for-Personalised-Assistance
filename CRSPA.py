import ctypes# lock or shut down pc
import json
import os
import random #
import re #to find integer
import subprocess
import webbrowser
from datetime import date as D, datetime as DT
from functools import lru_cache, reduce #to subtract
from time import time
import time #for sleep function
import keyboard #to press space so that it can play song
import pygame
import requests #to get news and articles
import speech_recognition as sr
import wikipedia
import winshell #to empty recycle bin and lock / shut down pc
from ecapture import ecapture as ec #to record audio and video and take photo
from requests.exceptions import ConnectionError, Timeout
from voice import SpeakText #def function to speak my custom made
import wolframalpha

# Initialize the recognizer
r = sr.Recognizer()

# Function to set an alarm
def setalarm(alarmtime):
    while True:
        now = DT.now()
        if now.strftime("%H:%M") == alarmtime:
            SpeakText("Time's up!")
            break
        time.sleep(1)

#api for google news
api_key = '415b5d7afbb84a9f8dc2038d2121bfea'

def get_news(category=None):
    if category:
        url = f'https://newsapi.org/v2/top-headlines?country=in&category={category}&apiKey={api_key}'
    else:
        url = f'https://newsapi.org/v2/top-headlines?country=in&apiKey={api_key}'
    try:
        response = requests.get(url)
        data = response.json()
        for i in range(5):
            SpeakText(data['articles'][i]['title'])
    except requests.exceptions.RequestException as e:
        SpeakText("Error: Could not connect to the server")
    except ValueError as e:
        SpeakText("Error: Invalid response received from the server")

# Get current date and time
today = D.today()
now = DT.now()
current_time = now.strftime("%H:%M:%S")

SpeakText("Hello sir! Welcome, Today is {} and the time is {}".format(today, current_time))

# Loop infinitely for user to speak
while(1):
    try:
        # use the microphone as source for input.
        with sr.Microphone() as source:
            r.adjust_for_ambient_noise(source, duration=0.2)
            audio = r.listen(source , timeout = 3)
            M = r.recognize_google(audio).lower()
            SpeakText(M)
            print("Processing task " + M)

            #data file handeling
            file = open("data.txt", "a")
            file.write("\n")
            file.write(M)
            file.close()

            # Websites
            if "open" in M:
                if "youtube" in M:
                    SpeakText("opening youtube")
                    webbrowser.open("www.youtube.com")
                elif "instagram" in M:
                    SpeakText("opening instagram")
                    webbrowser.open("https://www.instagram.com",)
                elif "facebook" in M:
                    SpeakText("opening facebook")
                    webbrowser.open("https://www.facebook.com",)
                elif "whatsapp" in M:
                    SpeakText("opening whatsapp")
                    webbrowser.open("https://web.whatsapp.com/")
                elif "twitter" in M:
                    SpeakText("opening twitter")
                    webbrowser.open("https://www.twitter.com/")
                elif "netflix" in M:
                    SpeakText("opening netflix")
                    webbrowser.open("https://www.netflix.com/in/")
                elif "tik tok" in M:
                    SpeakText("opeaning tik tok")
                    webbrowser.open("https://www.tiktok.com/en/")
                
                #apps
                elif "excel" in M:
                    SpeakText("opening excel")
                    os.startfile("C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE")
                elif "ms word" in M:
                    SpeakText("opening word")
                    os.startfile("C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE")
                elif "powerpoint" in M:
                    SpeakText("opening powerpoint")
                    os.startfile("C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE")
                elif "chrome" in M:
                    SpeakText("opening chrome")
                    os.startfile("C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe")
                elif "edge" in M:
                    SpeakText("opening edge")
                    os.startfile("C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe")
                elif "OBS" in M:
                    SpeakText("opeaning OBS") 
                    os.startfile("C:\\Program Files\\obs-studio\\bin\\64bit\\obs64.exe")
                elif "task manager" in M:
                    SpeakText("opening task manager")
                    os.startfile("C:\\WINDOWS\\system32\\Taskmgr.exe")
                elif "my sql" in M:
                    SpeakText("opening my SQL")
                    os.startfile("C:\\Program Files\\MySQL\\MySQL Server 8.0\\bin\\mysql.exe")
            # Time
            elif "time" in M:
                SpeakText(current_time)

            # Date
            elif "date" in M:
                SpeakText(today)

            #calculator
            number = re.findall(r'-?\d+',M)
            number = [int(x) for x in number]
            if "sum" in M or "+" in M:
                d = float(sum(number))
                c = "the sum is" or "your result is"
                SpeakText(d)
                print(f"{c} {d}")
            
            elif "-" in M :
                d = float(reduce(lambda x, y: x - y, number))
                c = "the substraction is" or "your result is"
                SpeakText(d)
                print(f"{c} {d}")
            
            elif  "subtract" in M:
                d = float(reduce(lambda x, y: y - x, number))
                c = "the substraction is" or "your result is"
                SpeakText(d)
                print(f"{c} {d}")
            
            elif "multiply" in M or "multiplication" in M:
                d = float(reduce(lambda x, y: x * y, number))
                c = "the multiplication is" or "your result is"
                SpeakText(d)
                print(f"{c} {d}")
                
            elif "divide" in M or "/" in M:
                d = float(reduce(lambda x, y: x // y, number))
                c = "the division is" or "your result is"
                SpeakText(d)
                print(f"{c} {d}")
                
            #music
            elif "spotify" in M:
                M = "https://open.spotify.com/collection/tracks"
                webbrowser.open(M)
                try :
                    time.sleep(2)
                    keyboard.press_and_release('space')
                except :
                    time.sleep(4)
                    keyboard.press_and_release('space')
                finally :
                    time.sleep(6)
                    keyboard.press_and_release('space')
            
            elif "music" in M:
                path = "C:\\Users\\RAJEEV\\Music\\Playlists"
                files = [f for f in os.listdir(path) if f.endswith(".mp3")]
                pygame.init()
                pygame.mixer.init()

                while True:
                    music = random.choice(files)
                    pygame.mixer.music.load(os.path.join(path, music))
                    pygame.mixer.music.play()

                    while pygame.mixer.music.get_busy():
                        if keyboard.is_pressed("p"):
                            pygame.mixer.music.pause()
                        elif keyboard.is_pressed("r"):
                            pygame.mixer.music.unpause()
                        elif keyboard.is_pressed("n"):
                            pygame.mixer.music.fadeout(2000)
                        elif keyboard.is_pressed("s"):
                            pygame.quit()        
                            break            
            #search on web
            elif "search" in M:
                if "youtube" in M:
                    words = ["youtube", "search", "on"]

                    for word in words:
                        M = M.replace(word, "")
                    print(M)
                    SpeakText("searching on youtube")
                    M = ("https://www.youtube.com/results?search_query="+M)
                    webbrowser.open(M)
                
                elif "wikipedia" in M:
                    try:
                        words = ["wikipedia", "search", "on","about"]

                        for word in words:
                            M = M.replace(word, "")
                        print(M)
                        # Search for an article
                        result = wikipedia.search(M)


                        # Extract the summary of the first article
                        summary = wikipedia.summary(result[0])
                        print(summary)
                        SpeakText(summary)
                    except sr.exceptions.PageError:
                        SpeakText("sorry I cannot find that")
                        

                #google search
                elif "google" in M:
                    words = ["google", "search", "on"]
                    for word in words:
                        M = M.replace(word, "")
                    print(M)
                    SpeakText("searching on google")
                    M = ("https://www.google.com/search?q="+M)
                    webbrowser.open(M)
                elif "bing" in M:
                    words = ["bing", "search", "on"]
                    for word in words:
                        M = M.replace(word, "")
                    print(M)
                    SpeakText("searching on bing")
                    M = ("https://www.bing.com/search?q="+M)
                    webbrowser.open(M)
                elif "article" in M:
                    words = ["article", "search", "about"]
                    for word in words:
                        M = M.replace(word, "")
                    print(M)
                    SpeakText("searching article on the web")
                    # Make the API call
                    url = f'https://newsapi.org/v2/everything?q={M}&apiKey={api_key}'
                    response = requests.get(url) 
                    S = response.json()
                    # Parse the JSON response
                    data = json.loads(response.text)

                    # Extract the articles from the response
                    articles = data['articles']

                    # Iterate through the articles
                    for article in articles:
                        # Extract the content of each article
                        content = article['content']
                    print (content)
                    SpeakText(content) 
            
            elif "what is" in M or "who is" in M:
                client = wolframalpha.Client("93HG7A-5L72G7J6VY")
                res = client.query(M)
                
                try:
                    print (next(res.results).text)
                    SpeakText (next(res.results).text)
                except StopIteration:
                    print ("No results")              
           
            elif "news" in M:
                if "business" in M:
                    get_news("business")
                elif "entertainment" in M:
                    get_news("entertainment")
                elif "general" in M:
                    get_news("general")
                elif "health" in M:
                    get_news("health")
                elif "science" in M:
                    get_news("science")
                elif "sports" in M:
                    get_news("sports")
                elif "technology" in M:
                    get_news("technology")
                else:
                    get_news()


            #game not complete
            elif "bored" in M and "i" in M:
                SpeakText("here are some games for you that you might like") 
                webbrowser.open("https://www.crazygames.com/")

            # Alarm
            elif "set alarm" in M:
                SpeakText("What time do you want to set the alarm for?")
                alarmtime = input("What time do you want to set the alarm for?")
                setalarm(alarmtime)  
            
            #extras
            elif 'empty recycle bin' in M:
                winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
                SpeakText("Recycle Bin Recycled")
            elif 'lock window' in M:
                SpeakText("locking the device")
                ctypes.windll.user32.LockWorkStation()
            elif 'shutdown' in M:
                SpeakText("Hold On a Sec ! Your system is on its way to shut down")
                subprocess.call('shutdown / p /f')
            #record
            elif "take" in M:
                if "photo" in M or "picture" in M:
                    ec.capture(0, "Camera ", "img.jpg") 
                elif "video" in M: 
                    ec.record("video.mp4")
                    time.sleep(5)
                    ec.stop()        
                elif "audio" in M:
                    ec.record_audio("audio.mp3")
                    time.sleep(5)
                    ec.stop()    
            elif "record" in M: 
                if "video" in M: 
                    ec.record("video.mp4")
                    time.sleep(5)
                    ec.stop()        
                elif "audio" in M:
                    ec.record_audio("audio.mp3")
                    time.sleep(5)
                    ec.stop()
                elif "photo" in M or "picture" in M:
                    ec.capture(0, "Jarvis Camera ", "img.jpg")    
            #stop the program
            elif "stop" in M or "sleep" in M or "shut down" in M:
                SpeakText("bye sir")
                os.quit() 
               
            SpeakText("complete")
    except sr.RequestError as e:
        print("Could not request results; {0}".format(e))
        
    except sr.UnknownValueError:
        print("unknown error occured") 
    except sr.WaitTimeoutError :
        print("time out")
