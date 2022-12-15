import pyttsx3 
import speech_recognition as sr 
import webbrowser   
import datetime   
import wikipedia    
import random
import os
import json
#import winshell
import pyjokes
import smtplib
import time
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
import pywhatkit as pwt
import wolframalpha as wf

def takeCommand():  
    r = sr.Recognizer() 
    with sr.Microphone() as source: 
        print('Listening') 
        r.pause_threshold = 0.5
        audio = r.listen(source)           
        try: 
            print("Recognizing") 
            Query = r.recognize_google(audio, language='en-in') 
            print("You said ", Query) 
        except Exception as e: 
            print(e) 
            print("Say that again sir") 
            return "None"
        return Query   

def Hello(): 
    speak("hello, I am your desktop assistant. Tell me how may I help you!")

def speak(audio):      
    engine = pyttsx3.init() 
    voices = engine.getProperty('voices')       
    engine.setProperty('voice', voices[1].id) # Method for the speaking of the the assistant 
    engine.say(audio)    # Blocks while processing all the currently; queued commands 
    engine.runAndWait() 

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>= 0 and hour<12:
        speak("Good Morning Sir !")
  
    elif hour>= 12 and hour<18:
        speak("Good Afternoon Sir !")  
  
    else:
        speak("Good Evening Sir !") 

def tellJoke():
    jok_dict = {1: 'Hear about the new restaurant called Karma?There’s no menu: You get what you deserve.haha',
                2: 'Did you hear about the claustrophobic astronaut?He just needed a little space.haha',
                3: 'A man tells his doctor, “Doc, help me. I’m addicted to Twitter!”The doctor replies, “Sorry, I don’t follow you …”haha',
                4: 'Why are football stadiums so cool? Because every seat has a fan in it!haha',
                5: 'Where do frogs keep their money?The river bank!haha',
                6: 'Did you hear the guy who invented the knock knock joke?He won the "no-bell" prize. haha',
                7: 'Why can\'nt eggs keep a secret?They tend to crack under pressure. haha',
                8: 'Why don’t Calculus majors throw house parties?Because you should never drink and derive.haha',
                9: 'What happened when the geese fell down the stairs?They all got goose bumps.haha',
                10: 'How do you make a squid laugh?With ten-tickles. haha'}
    p = random.randint(1,10)
    joke = jok_dict[p]
    print("here's the joke::\n",joke)
    speak(joke)
  
def tellDay(): 
    day = datetime.datetime.today().weekday() + 1
    Day_dict = {1: 'Monday', 2: 'Tuesday',  
                3: 'Wednesday', 4: 'Thursday',  
                5: 'Friday', 6: 'Saturday', 
                7: 'Sunday'} 
    if day in Day_dict.keys(): 
        day_of_the_week = Day_dict[day] 
        print(day_of_the_week) 
        speak("The day is " + day_of_the_week) 
  
def tellTime(): 
    time = str(datetime.datetime.now())
    print(time) 
    hour = time[11:13] 
    min = time[14:16] 
    speak("The time is " + hour + " hours and " + min + "Minutes")     

def usrname():
    speak("What should i call you sir")
    uname = takeCommand()
    speak("Welcome "+uname+"How can i Help you")

def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls() # these 2 lines are used to check the connection with smtp server of gmail
    server.login('harshithamuvvala2@gmail.com', 'akhilesh123') # Enable low security in gmail
    server.sendmail('harshithamuvvala2@gmail.com', to, content)
    server.close()
 
def Take_query(): 
    Hello() 
    while(True): 
        query = takeCommand().lower() 
        if "open geeksforgeeks" in query: 
            speak("Opening GeeksforGeeks ") 
            webbrowser.open("www.geeksforgeeks.com") 
            continue
        
        elif "wish me" in query: 
            wishMe() 
            continue
        
        elif "open google" in query: 
            speak("Opening Google ") 
            webbrowser.open("www.google.com") 
            continue
        
        elif "open youtube" in query: 
            speak("Opening Youtube ") 
            webbrowser.open("www.youtube.com") 
            continue
        
        elif "which day it is" in query: 
            tellDay() 
            continue
        
        elif "time" in query: 
            tellTime() 
            continue
        
        elif "tell me a joke" in query:
            tellJoke()
            continue
        
        elif "one more" in query:
            tellJoke()
            continue
        
        elif "bye" in query: 
            speak("Bye. Checking out") 
            exit() 
        
        elif "from wikipedia" in query: 
            speak("Checking the wikipedia ") 
            query = query.replace("wikipedia", "") # "" contains our search word
            result = wikipedia.summary(query, sentences=3) 
            speak("According to wikipedia") 
            speak(result) 
        
        elif "tell me your name" in query: 
            speak("I am Adam. Your deskstop Assistant") 
        
        elif "take notes" in query:
            speak("Yes I can take notes. Tell me what to note down")
            data=takeCommand()
            notes=open("data.txt","w")
            notes.write(data)
            notes.close()
            speak("notes taken")
        elif "repeat notes" in query:
            notes=open("data.txt","r")
            speak("Your notes are "+notes.read())

        elif 'send email' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                #speak("whom should i send")
                to = "harshithamuvvala1@gmail.com"   
                sendEmail(to, content)
                speak("Email has been sent !")
            except Exception as e:
                print(e)
                speak("I am not able to send this email")
        
        elif "whatsapp" in query:
            pwt.sendwhatmsg("+919492534480","welcome",20,41)
        
        elif 'search' in query or 'play' in query:
            query = query.replace("search", "")
            query = query.replace("play", "")         
            webbrowser.open("www.google.com/search?q="+query)

        elif 'what is' in query or 'who is' in query:
            client=wf.Client('YPAGX9-7PJPTPV4TK')
            while True:
                query=str(input('query'))
                res=client.query(query)
                output=next(res.results).text
                print(output)
                speak(output)

        
        elif "how are you" in query:
            speak("I'm fine, glad you asked")
 
if __name__ == '__main__':
    #usrname() 
    #print(sr.Microphone.list_microphone_names())
    Take_query()
