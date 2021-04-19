'''
----------------------------------------------------------------------------------VICTOR PROJECT-----------------------------------------------------------------------










VICTOR is a virtual assistant, also called AI assistant or digital assistant, is an application program that understands natural language voice commands and completes tasks for the user.









'''


import pyttsx3
import datetime as dt
import wikipedia
import webbrowser
import os
import random
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import sys
import smtplib
import subprocess 
import wolframalpha 
import pyjokes 
import bs4
from bs4 import BeautifulSoup as soup
import ctypes 
import json
import shutil 
import win32com.client as wincl
from urllib.request import urlopen 
import feedparser
import requests 
import scipy
import pywhatkit
import speech_recognition as sr

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)
#print(voices[1].id)




def speak(audio):
    '''
    Speaks the text passed as argument
    '''
    engine.say(audio)
    engine.runAndWait()
    pass

def GreetMe():
    '''
    Greets user
    '''
    hour = dt.datetime.now().hour
    if hour>=0 and hour<12:
        speak("Good Morning Sir")
    elif hour>=12 and hour<18:
        speak('Good Afternoon Sir')
    else:
        speak("Good Evening Sir")   


    speak("I am Rick, a virtual assistant! Please tell me how may I help you")
    print('I am Rick, a virtual assistant! Please tell me how may I help you')

def takeCommand():
    #It takes microphone input from the user and returns string output

    r = sr.Recognizer()
    
    with sr.Microphone() as source:
        r.adjust_for_ambient_noise(source,duration=1)
        print("Listening...")
        r.pause_threshold = 1
        r.energy_threshold = 500
       
        audio = r.listen(source)

    try:
        print("Recognizing...")    
        query = r.recognize_google(audio, language='en-in')

        print(f"User said: {query}\n")

    except Exception as e:
        print(e)    
        speak("Say that again please...")  
        print("Say that again please...")
        return "None"
    return query

def sendEmail(to,content):
    server = smtplib.SMTP('smtp.gmail.com',587)
    server.ehlo()
    server.starttls()
    server.login('mdshah930@gmail.com','your_password_here')
    server.sendmail('mdshah930@gmail.com',to,content)
    server.close()


def getNews():
    j=0
    news_url="https://news.google.com/news/rss"
    #news_url = "https://news.google.com/topstories?hl=en-IN&gl=IN&ceid=IN:en"
    Client=urlopen(news_url)
    xml_page=Client.read()
    Client.close()
    news_data=[]
    soup_page=soup(xml_page,"lxml") #soup_page=soup(xml_page,"xml")
    news_list=soup_page.findAll("item")
    # Print news title, url and publish date
    for news in news_list:
        news_data.append(news.title.text)
    for i in news_data:
        j+=1
        if(j>3):
            break
        print(i)
        speak(i)
        

def getWeather(city_name):
    city_name=str(city_name)
    api_key = "ac9f74ab44a45da1f300afc1bed68d03"
    base_url = "http://api.openweathermap.org/data/2.5/weather?"

    complete_url = base_url + "appid=" + api_key + "&q=" + city_name 
    response = requests.get(complete_url) 
    
    x = response.json() 
    
    if x["cod"] != "404": 
    
         
        y = x["main"] 
    
        
        current_temperature = y["temp"] 
    
        current_pressure = y["pressure"] 
    
        current_humidity = y["humidity"] 
    
        z = x["weather"] 
    
        weather_description = z[0]["description"] 
    
        c_temp = current_temperature - 273.5
        c_temp = round(c_temp,3)
        speak(f"Temperature is {c_temp} degree celsius")
        speak(f"Current Pressure in hPa {current_pressure}")
        speak(f"Current Humidity in percentage {current_humidity}")
        speak(f"Overall the weather is {weather_description}")

        if ('rain' in weather_description):
            speak("Sir you might consider taking an umbrella if moving out")
        
        print(" Temperature in celsius  = " +
                        str(c_temp) + 
            "\n atmospheric pressure (in hPa unit) = " +
                        str(current_pressure) +
            "\n humidity (in percentage) = " +
                        str(current_humidity) +
            "\n description = " +
                        str(weather_description)) 
    
    else: 
        print(" City Not Found ") 

        
def whatsapp_msg(to,content):
    to = to.lower()
    a=  dt.datetime.now().hour
    b = dt.datetime.now().minute + 2
    if to == 'person1':
        person = '+91number1'
    elif to == 'person2':
        person = '+91number2'
    elif to == 'person3':
        person = '+91number3'
    # hour = int(datetime.now().hour)
    # minute = int(datetime.now().minute) + 2

    pywhatkit.sendwhatmsg(str(person),content,a,b)




if __name__ == '__main__':
    #a={'Noun': ["the outermost region of the sun's atmosphere; visible as a white halo during a solar eclipse", '(botany', 'an electrical discharge accompanied by ionization of surrounding atmosphere', 'one or more circles of light seen around a luminous object', '(anatomy', 'a long cigar with blunt ends']}
    
    GreetMe()
    while True:
        query = takeCommand().lower()

        #Logic for executing tasks based on queries
        if 'tell me something about' in query or 'define' in query :
            speak("Searching Wikipedia...")
            if 'tell me something about' in query:
                query = query.replace('tell me something about',' ')
            elif 'define' in query :
                 query = query.replace('define',' ')
            result = wikipedia.summary(query,sentences=2)
            speak("Here's what I found")
            print("Here's what I found ...\n")
            print(result)
            speak(result)
        elif 'open youtube' in query :
            speak("Opening Sir")
            webbrowser.open("https://www.youtube.com/")
        elif 'open google' in query:
            speak("Opening Sir")
            webbrowser.open("https://www.google.com/")
        elif 'open gmail' in query:
            speak("Opening Sir")
            webbrowser.open("https://mail.google.com/mail/u/0/#inbox")

        elif 'open netflix' in query:
            speak("Opening Sir")
            webbrowser.open("https://www.netflix.com/browse")
      
        elif 'music' in query:
            speak("Shuffling through your directory")
            music_dir = 'D:\\songs\\english'
            songs = os.listdir(music_dir)
            i = random.randint(0,len(songs))
            os.startfile(os.path.join(music_dir,songs[i]))

        elif 'the time' in query:
            strTime = dt.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir the time is {strTime}")

        elif 'open whatsapp' in query:
            speak("Opening Sir")
            whatsapp_path = "C:\\Users\\Mihir\\AppData\\Local\\WhatsApp\\WhatsApp.exe"
            os.startfile(whatsapp_path)

        elif 'open zoom' in query:
            speak("Opening Sir")
            zoom_path = "C:\\Users\\Mihir\\AppData\\Roaming\\Zoom\\bin\\Zoom.exe"
            os.startfile(zoom_path)

        elif 'open vscode' in query:
            speak("Opening Sir")
            code_path = "C:\\Users\\Mihir\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"

        elif 'send email' in query:
            try:
                speak("To whom sir?")
                target = takeCommand()
                speak("What should I say")
                con = takeCommand()
                to = "mdshah930@gmail.com"
                sendEmail(to,con)
                speak("Email has been sent successfully Sir")
            except Exception as e:
                print(e)
                speak("Due to some technical difficulties I could not send this email at the moment")
        
        elif 'weather today' in query:
            getWeather('Mumbai')


        elif 'weather condition' in query or 'weather in other city' in query:
            speak("Yes Sir!, Which city's weather forecast shall I read for you?")
            cityy = takeCommand()
            getWeather(cityy)

        elif 'whatsapp message' in query:
            speak("Yes Sir! Who is the recipient?")
            rec = takeCommand()
            speak("Sir what is the message :")
            msg = takeCommand()
            whatsapp_msg(rec,msg)
            speak("Message has been sent successfully!")

        elif 'on youtube' in query:
            query = query.replace('play','')
            print(query)
            pywhatkit.playonyt(query)

        elif 'who is your dad' in query:
            speak("I have you, thats enough for me")
            print('I have you, thats enough for me')
        
        elif 'what are you doing' in query:
            speak("What am I doing? Well I am talking with you Sir")
            print("What am I doing? Well I am talking with you Sir")
        
        elif 'what can you do' in query:
            print("Sir I can do variety of things like sending emails, open websites, search through wikipedia, play some music for you, get weather conditions of any city and many other tasks. I was made in 2020, I consider myself pretty smart at this age")
            speak("Sir I can do variety of things like sending emails, open websites, search through wikipedia, play some music for you, get weather conditions of any city and many other tasks. I was made in 2020, I consider myself pretty smart at this age")

        elif 'bye' in query:
            print("It was my pleasure serving you Sir")
            speak("It was my pleasure serving you Sir")
            break

        elif 'who are you' in query:
            print("Sir, I am a virtual assistant created by Mr.Mihir Shah")
            speak("Sir, I am a virtual assistant created by Mr.Mihir Shah")

        elif 'lock' in query: 
            speak("locking the device") 
            ctypes.windll.user32.LockWorkStation()

        elif 'how are you' in query: 
            print("I am fine, Thank you") 
            print("How are you, Sir")
            speak("I am fine, Thank you") 
            speak("How are you, Sir") 

        elif 'joke' in query:
            print(a)
            a=pyjokes.get_joke(language = 'en', category = 'all')
            speak(a) 
            
        
        elif 'stressed' in query or 'bad mood' in query:
            speak('Sir, would you like to hear a joke or maybe hear some music ?')
            print('Sir, would you like to hear a joke or maybe hear some music ?')
        
        elif 'news' in query:
            getNews()

        #Just for fun 
        elif "i love you" in query: 
            speak("It's hard to understand")
            print("It's hard to understand")
        
        elif 'great' in query:
            speak("Thank you Sir!")

        elif 'hello' in query:
            speak("Hello Sir, you seem to be in a good mood today")

        elif 'fine' in query:
            speak("Great! What can I do for you Sir?")
        
''' OUTPUT 
This is just an example of output of few tasks of victor - 
You can find a full demonstration of all tasks at the link given below
https://youtu.be/OOlHs4TtGd4

Listening...
Recognizing...
User said: who are you

Listening...
Recognizing...
User said: what can you do

Listening...
Recognizing...
User said: tell me the latest news

"Party Ticket Being Given To Rapist": Congress Worker Thrashed In UP - NDTV
CBI registers FIR against accused in Hathras case - Times of India
‘Pathway to self-reliant rural India’: PM Modi’s top quotes from SVAMITVA event - Hindustan Times
Listening...
Recognizing...
User said: define artificial intelligence

Artificial intelligence (AI), sometimes called machine intelligence, is intelligence demonstrated by machines, unlike the natural intelligence displayed by humans and animals. Leading AI textbooks define the field as the study of "intelligent agents": any device that perceives its environment and takes actions that maximize its chance of successfully achieving its goals.
Listening...
Recognizing...
User said: how is the weather today

 Temperature in celsius  = 29.65
 atmospheric pressure (in hPa unit) = 1003
 humidity (in percentage) = 84
 description = haze
Listening...
Recognizing...
User said: what are you doing

Listening...
Recognizing...
User said: bye
'''