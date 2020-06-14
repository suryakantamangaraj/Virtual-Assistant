#import pyttsx3 #microsoft text to speech
import speech_recognition as sr # importing speech recognition package from google api
import datetime #importing date and time 
import wikipedia #importing wikipedia
import webbrowser #importing webbrowser
import os #importing os to save/open files
import smtplib #access google mail through smtp
# from pygame import mixer
import playsound #to play saved mp3 file
from gtts import gTTS #google text to speech
import wolframalpha #to calculate strings into formula, its a website which provides api, 100 times per day
import ssl #to acess ssl
from selenium import webdriver  #to control browser operations
from selenium.webdriver.common.keys import Keys
from io import BytesIO
from io import StringIO
import warnings
import calendar
import random
import sys
import re
import requests
import subprocess
from pyowm import OWM
import youtube_dl
#import vlc
import urllib
#import urllib2
import json
from bs4 import BeautifulSoup as soup
from urllib.request import urlopen
#from urllib2 import urlopen
import wikipedia
from time import strftime

# Ignore any warning messages
warnings.filterwarnings('ignore')

ssl._create_default_https_context = ssl._create_unverified_context
num = 1

# Function to get the virtual assistant response
def assistant_speaks(output):
    global num
    num +=1
    print("Lima : ", output)

    # Convert the text to speech
    toSpeak = gTTS(text=output, lang='en-US', slow=False)

    # Save the converted audio to a file
    file = str(num)+".mp3"
    toSpeak.save(file)

    # mp3_fp = BytesIO()
    # toSpeak = gTTS(output, 'en', slow=False)
    # toSpeak.write_to_fp(mp3_fp)
    # os.system("D:\PeRSon\\audio\spoken.mp3")
    '''mixer.init()
    mixer.music.load('D:\PeRSon\\audio\spoken.mp3')
    mixer.music.play()
    time.sleep(5)
    mixer.music.stop()'''
    # song = AudioSegment.from_file(mp3_fp, format="mp3")
    # playsound.playsound(mp3_fp)

    # Play the converted file
    playsound.playsound(file, True)

    #Remove that temp file
    os.remove(file)

# Record audio and return it as a string
def get_audio():
    # Create a recognizer object named r
    r = sr.Recognizer()
    # Open the microphone and start recording 
    #NOTE: # The with statement itself ensures proper acquisition and release of resources
    audio = '' # Speech recognition using Google's Speech Recognition
    with sr.Microphone() as source: # Creates a new Microphone instance, which represents a physical microphone on the computer
        print("Listening...")
        audio = r.listen(source, phrase_time_limit=5) # Records a single phrase 
    print("Recognising...")
    try: #Try to get google to recognize the audio NOTE: The try block lets you test a block of code for errors
        text = r.recognize_google(audio,language='en-US')
        print("You : ", text)
        return text
    except:
        assistant_speaks("Could not understand your audio, PLease try again!")
        return 0

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        assistant_speaks("Good Morning!")

    elif hour>=12 and hour<18:
       assistant_speaks("Good Afternoon!")   

    else:
       assistant_speaks("Good Evening!")  
    
def search_web(input):
    driver = webdriver.Firefox()
    driver.implicitly_wait(1)
    driver.maximize_window()
    if 'youtube' in input.lower():
        assistant_speaks("Opening in youtube")
        indx = input.lower().split().index('youtube')
        query = input.split()[indx+1:]
        driver.get("http://www.youtube.com/results?search_query=" + '+'.join(query))
        return

    else:
        assistant_speaks("Searching in google")
        indx = input.lower().split().index('google')
        query = input.split()[indx + 2:]
        driver.get("https://www.google.com/search?q=" + '+'.join(query))
        return


def open_application(input):
    if "visual studio" in input:
        assistant_speaks("Opening Microsoft Visual Code")
        os.startfile('C:\\Program Files (x86)\\Microsoft Visual Studio\\2019\\Community\\Common7\\IDE\\devenv.exe')
        return
    elif "firefox" in input or "mozilla" in input:
        assistant_speaks("Opening Mozilla Firefox")
        os.startfile('C:\\Program Files\\Mozilla Firefox\\firefox.exe')
        return
    elif "word" in input:
        assistant_speaks("Opening Microsoft Word")
        os.startfile('C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE')
        return
    elif "excel" in input:
        assistant_speaks("Opening Microsoft Excel")
        os.startfile('C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE')
        return
    elif "ppt" in input or "powerpoint" in input:
        assistant_speaks("Opening Microsoft Powerpoint")
        os.startfile('C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE')
        return
    else:
        assistant_speaks("Application is not available or, Unable to access system files, it may make your system vulnerable")
        return


def process_text(input):
    try:
        if "who are you" in input or "define yourself" in input or "who is Lima" in input:
            speak = '''Hello, I am Lima. Your personal Assistant.
            I am here to make your life easier. 
            You can command me to perform various tasks'''
            assistant_speaks(speak)
            return
        elif "who made you" in input or "created you" in input:
            speak = "I have been created by Surya Raj."
            assistant_speaks(speak)
            return
        elif "who is Surya Raj" in input:
            speak = "Thanks for asking about Surya. He is the creator of me."
            assistant_speaks(speak)
            return
        elif "what can you do" in input:
            speak = '''Something you can ask me like: search web, play youtube, open apps, calculations
           and Q&A. Also, you can add new functionalities as a developer.'''
            assistant_speaks(speak)
            return
        elif "crazy" in input:
            speak = """Well, there are 2 mental asylums in India."""
            assistant_speaks(speak)
            return
        elif "calculate" in input.lower():
            app_id= "UP67VA-28RPELRA52"
            client = wolframalpha.Client(app_id)

            indx = input.lower().split().index('calculate')
            query = input.split()[indx + 1:]
            res = client.query(' '.join(query))
            answer = next(res.results).text
            assistant_speaks("The answer is " + answer)
            return
        elif 'the time' in input.lower():
            strTime = datetime.datetime.now().strftime("%H:%M:%S")    
            assistant_speaks(f"Sir, the time is {strTime}")
        elif 'open' in input:
            open_application(input.lower())
            sys.exit()
        elif 'search' in input or 'play' in input:
            search_web(input.lower())
            sys.exit()
            
        elif 'wikipedia' in input.lower():
            assistant_speaks('Searching Wikipedia...')
            query = input.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=5)
            assistant_speaks("According to Wikipedia")
            assistant_speaks(results)
            return
        else:
            assistant_speaks("I can search the web for you, Do you want to continue?")
            ans = get_audio()
            if 'yes' in str(ans) or 'yeah' in str(ans):
                search_web(input)
            else:
                return
    except Exception as e:
        print(e)
        assistant_speaks("I don't understand, I can search the web for you, Do you want to continue?")
        ans = get_audio()
        if 'yes' in str(ans) or 'yeah' in str(ans):
            search_web(input)


if __name__ == "__main__":
    #assistant_speaks("What's your name, Human?")
    name ='Sir'
    #name = get_audio()
    assistant_speaks("Hello, " + name + '.')
    wishMe()
    #assistant_speaks("How can I help you?")
    while(1):
        #assistant_speaks("How can I help you?")
        text = get_audio().lower()
        if text == 0:
            continue
        #assistant_speaks("What else can I do for you?")
        #assistant_speaks(text)
        if "exit" in str(text) or "bye" in str(text) or "go " in str(text) or "sleep" in str(text):
            assistant_speaks("Ok bye, "+ name+'.')
            sys.exit()
        process_text(text)
