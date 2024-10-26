import pyttsx3
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import smtplib
from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtCore import QTimer, QTime, QDate, Qt
from PyQt5.QtGui import QMovie
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.uic import loadUiType
from buddyf import Ui_missminute
import sys
import random
import xlwt
from xlwt import Workbook
from bs4 import BeautifulSoup
import requests
import time
import keyboard
from openpyxl import Workbook, load_workbook
from autocorrect import Speller
import nltk
from textblob import TextBlob
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
rate = engine.getProperty('rate')
engine.setProperty('rate', 200)
engine.setProperty('voices', voices[1].id)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def get_weather(city_name):
    base_url = "https://www.google.com/search?q=" + city_name + "+weather"
    headers = {'User-Agent': 'Mozilla/5.0'}
    response = requests.get(base_url, headers=headers)
    soup = BeautifulSoup(response.text, 'html.parser')
    temperature = soup.find('div', attrs={'class': 'BNeawe iBp4i AP7Wnd'}).text
    weather = soup.find('div', attrs={'class': 'BNeawe tAd8D AP7Wnd'}).text
    return f"Temperature : {temperature}\nWeather : {weather}"

Alphanum = {
    'one': '1',
    'two': '2',
    'three': '3',
    'four': '4',
    'five': '5',
    'six': '6',
    'seven': '7',
    'eight': '8',
    'nine': '9',
    'zero': '0'
    }
rower = ['one', 'two', 'three', 'four', 'five' , 'six', 'seven' , 'eight', 'nine' , 'ten' , 'eleven', 'twelve','thirteen']

def wish():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good Morning,")

    elif hour >= 12 and hour < 18:
        speak("Good Afternoon,")

    else:
        speak("Good Evening!")

    speak("sir")
class MainThread(QThread):
    def __init__(self):
        super(MainThread,self).__init__()
        
    def run(self):
        self.TaskExecution()

    def takeCommand(self):
        r = sr.Recognizer()
        with sr.Microphone() as source:
            speak("Listening....")
            r.pause_threshold = 1
            r.adjust_for_ambient_noise(source)
            audio = r.listen(source)
        try:
            list11 = ['m','l', 'k','n','o','p']
            am=random.choice(list11)
            if am=='l':
                speak("i am on it, sir")
            elif am=='m':
                speak('okay sir, i am working on it')
            elif am=='k':
                speak('just a second sir')
            elif am=='o':
                speak('hang on sir')
            elif am=='p':
                speak('wait sir')
            elephant = r.recognize_google(audio, language='en-in')
        except Exception as e:
            speak("pardon please...")
            return "None"
        spell = Speller(lang='en')
        if 'notpad' in elephant or 'Notepad' in elephant:
            query=elephant
        elif 'XL' in elephant:
            query=elephant.replace('XL','excel')
        else:
            query=spell(elephant)
        print(query)
        return query
    
    def startCommand(self):
        r = sr.Recognizer()
        with sr.Microphone() as source:
            r.pause_threshold = 1
            r.adjust_for_ambient_noise(source)
            audio = r.listen(source)
        try:
            quer = r.recognize_google(audio, language='en-in')
        except Exception as e:
            return "None"
        print(quer)
        return quer

    def TaskExecution(self):
        file= 'xwlt example.xls'
        while True :
            self.quer = self.startCommand()
            print('done')
            if 'hey' in self.quer.lower():
                print('done')
                wish()
                my_file1 = open('onetimecommand.txt','w+')
                mma = self.takeCommand()
                my_file1.write(mma)
                my_file1.close()
                chop=(' and ',' then ',' or ',' than ')
                myfile2 = open('onetimecommand.txt','r+')
                jjk= myfile2.read()
                jjk=jjk.lower()
                for i in chop:
                    if i in jjk:
                        jjk = jjk.replace(i,' . ')
                if ' you ' in jjk:
                    jjk = jjk.replace(' you ',' . you ')
                if ' i ' in jjk:
                    jjk = jjk.replace(' i ',' . i ')
                if ' we ' in jjk:
                    jjk = jjk.replace(' we ',' . we ')
                if ' . . ' in jjk:
                    jjk = jjk.replace(' . . ',' . ')
                print(jjk)

                krk=TextBlob(jjk)
                katy=krk.sentences
                length=len(katy)
                i=0
                j=length-1
                for i in range (length):
                    command = katy[i]
                    self.query = command
                    
                    if 'open youtube' in self.query.lower():
                        speak('openning youtube')
                        webbrowser.open("youtube.com")
                        
                    elif "close notepad" in self.query.lower()or "quit notepad" in self.query.lower()or "exit notepad" in self.query.lower() :
                        speak('closing Notepad file')
                        try :
                            os.system('TASKKILL /F /IM notepad.exe')
                        except Exception as e:
                            return
                        
                    elif"quit excel" in self.query.lower() or "exit excel" in self.query.lower() or "close excel" in self.query.lower():
                        speak('closing Excel file')
                        try :
                            os.system('TASKKILL /F /IM EXCEL.EXE')
                        except Exception as e:
                            return
                        
                    elif 'explain' in self.query.lower():
                        speak('i was built as a friend for human which is totally based on pre feeded command, to hence human productivity and efficiency')

                    elif 'open google' in self.query.lower():
                        speak('openning google')
                        webbrowser.open("google.com")
                        
                    elif 'hello' in self.query:
                        speak('hey there')
                        
                    elif"open excel" in self.query.lower():
                        speak('opening excel file')
                        keyboard.press_and_release('win + e')
                        time.sleep(1)
                        keyboard.press_and_release('ctrl + f')
                        time.sleep(1)
                        keyboard.write(file)
                        time.sleep(1)
                        keyboard.press_and_release('enter')
                        time.sleep(1)
                        keyboard.press_and_release('tab')
                        time.sleep(1)
                        keyboard.press_and_release('tab')
                        time.sleep(1)
                        keyboard.press_and_release('tab')
                        time.sleep(1)
                        keyboard.press_and_release('tab')
                        time.sleep(1)
                        keyboard.press_and_release('down')
                        time.sleep(1)
                        keyboard.press_and_release('up')
                        time.sleep(1)
                        keyboard.press_and_release('enter')
                        
                    elif"open notepad" in self.query.lower():
                        speak('opening notepad file')
                        excel = "C:/Users/Dell/Videos/upgrade/datamanager.txt"
                        os.startfile(excel)
            
                    elif 'stone paper scissor' in self.query.lower():
                        speak('lets play it')
                        for i in range (3):
                            list1 = [0, 1, 2]
                        a=random.choice(list1)
                        if a==0 :
                            speak('stone')
                            speak ('what did u get sir')
                            self.query = self.takeCommand()
                            if 'paper' in self.query:
                                speak('u ,wwoonn ,sir')
                            elif 'scissor' in self.query:
                                speak ('i ,wwoonn ')
                            else:
                                speak('ask me again sir')
                                break
                        elif a==1:
                            speak('paper')
                            speak ('what did u get sir')
                            self.query = self.takeCommand()
                            if 'scissor' in self.query:
                                speak('congrats sir u won ,sir')
                            elif 'stone' in self.query:
                                speak ('i, wwoonn, sir')
                            else:
                                speak('ask me again sir')
                                break
                        elif a==2 :
                            speak('scissor')
                            speak ('what did u get sir')
                            self.query = self.takeCommand()
                            if 'stone' in self.query:
                                speak('u wwoonn ,sir')
                            elif 'paper' in self.query:
                                speak ('i wwoonn ')
                            else:
                                speak('ask me again sir')
                                break
                            
                    elif 'search' in self.query.lower():
                        speak('Searching Wikipedia...')
                        query = self.query.replace("wikipedia", "")
                        results = wikipedia.summary(query, sentences=2)
                        speak("According to Wikipedia")
                        speak(results)
                        
                    elif 'introduce' in self.query.lower():
                        speak('Hello guys, my name is buddy, i am your personal virtual assistant , which play games, i am version 3 point O')

                    elif 'play music' in self.query.lower():
                        webbrowser.open('https://www.youtube.com/watch?v=jADTdg-o8i0')

                    elif 'mind' in self.query.lower():
                        webbrowser.open('https://www.youtube.com/watch?v=SDXHcR5AL6E')
                        speak ('sir, sit down, close your eyes,and meditate')

                    elif 'more about' in self.query.lower():
                        speak('ITs the last 2 years (during covid ), all of us faced many challenges. then online class, work from home stuff like those came into play which affects our working of mind, to resolve this. i was built by student of class 12 Jatin Garg, i can do lots of work by your voice. even i can send emails which reduces your screen time. currently i can play stone paper sccisor , and search anything on wikipedia but soon i will be able to do whatever stuff u want ')

                    elif "joke" in self.query.lower():
                        speak('what do u call a train carrying bubblegum?.................. a chew chew train')

                    elif "namaste" in self.query.lower():
                        speak('nameste ji, ma apki kya sahaeta kar sakti hu')

                    elif "raise ho" in self.query.lower():
                        speak('mm theek hu, aap btao')

                    elif "kar sakte" in self.query.lower():
                        speak('m bohot kuch kr skti hu jaise ganne bjana aur apki problem solve krna')

                    elif "notepad" in self.query.lower():
                        speak('wait...., notepad mode is loading')
                        self.query = self.takeCommand()
                        my_file = open('datamanager.txt','w+')
                        my_file.write(self.query)
                        my_file.close()
                        
                    elif "weather" in self.query.lower():
                        speak('ok sir, sir, please tell me your city name')
                        self.query = self.takeCommand()
                        try:
                            city_name = self.query
                            speak(get_weather(city_name))
                        except Exception as e:
                            speak('sorry sir, I can\'t find the city')
                            return
                    
                    elif "timer" in self.query.lower():
                        speak('what shall i remind you about')
                        self.query = self.takeCommand()
                        text = str(self.query)
                        speak ('sir, after how many minute should i remind you?')
                        self.query = self.takeCommand()
                        try:
                            m=int(self.query)
                        except e as exception:
                            speak('sorry sir, can\'t proceed, try again later')
                            break
                        speak ('sir , i will tell u in ')
                        speak(m,'min')
                        local_time = m
                        local_time = local_time*60
                        time.sleep(local_time)
                        speak('time is over sir')
                        speak(text)
                        
                    elif "excel" in self.query.lower():
                        speak('excel mode is loading')
                        try :
                            os.system('TASKKILL /F /IM EXCEL.EXE')
                        except Exception as e:
                            return
                        speak('Do you wnna make new file or edit')
                        ex=self.takeCommand()
                        if 'new' in ex.lower():
                            wb = xlwt.Workbook()
                            sheet1 = wb.add_sheet('Sheet 1')
                            speak('how many columns you want sir')
                            self.query = self.takeCommand()
                            print(type(self.query))
                            if  "1" in self.query:
                                column=1
                            elif  "2" in self.query:
                                column=2
                            elif "3" in self.query:
                                column=3
                            elif  "4" in self.query:
                                column=4
                            elif "5" in self.query:
                                column=5
                            elif "6" in self.query:
                                column=6
                            elif "7" in self.query:
                                column=7
                            elif  "8" in self.query:
                                column=8
                            elif  "9" in self.query:
                                column=9
                            else:
                                speak('do it from start')
                                break
                            speak('how many rows you want sir')
                            self.query = self.takeCommand()
                            if '1' in self.query:
                                row=1
                            elif "2"in self.query:
                                row=2
                            elif "3"in self.query:
                                row=3
                            elif "4"in self.query:
                                row=4
                            elif  "5"in self.query:
                                row=5
                            elif "6"in self.query:
                                row=6
                            elif "7"in self.query:
                                row=7
                            elif "8" in self.query:
                                row=8
                            elif "9" in self.query:
                                row=9
                            else:
                                speak('do it from start')
                                break
                            speak('sir now tell me your entries column wise')
                            t=0
                            print(column,row)
                            for t in range (1,column+1):
                                for j in range (1,row+1):
                                    print('done')
                                    self.query = self.takeCommand()
                                    print('done')
                                    print(type(self.query))
                                    print('done')
                                    entry=str(self.query)
                                    print('done')
                                    sheet1.write(j,t,entry)
                                    print('done')
                            speak('what should be the file name')
                            nam=self.takeCommand()
                            file= nam + '.xls'
                            wb.save(file)
                            speak('completed')
                        elif 'edit' in ex.lower():
                            speak('ok sir')
                            speak('opening excel file')
                            keyboard.press_and_release('win + e')
                            time.sleep(1)
                            keyboard.press_and_release('ctrl + f')
                            time.sleep(1)
                            keyboard.write(file)
                            time.sleep(1)
                            keyboard.press_and_release('enter')
                            time.sleep(1)
                            keyboard.press_and_release('tab')
                            time.sleep(1)
                            keyboard.press_and_release('down')
                            time.sleep(1)
                            keyboard.press_and_release('up')
                            time.sleep(1)
                            keyboard.press_and_release('enter')
                            time.sleep(20)
                            speak('do you wnna make heading bold')
                            b=self.takeCommand()
                            if 'yes' in b:
                                keyboard.press_and_release('down')
                                time.sleep(0.2)
                                keyboard.press_and_release('right')
                                time.sleep(0.2)
                                keyboard.press_and_release('ctrl+b')
                                time.sleep(1)
                                keyboard.press_and_release('right')
                                time.sleep(0.2)
                                keyboard.press_and_release('ctrl+b')
                                time.sleep(1)
                                keyboard.press_and_release('right')
                                time.sleep(0.2)
                                keyboard.press_and_release('ctrl+b')
                                time.sleep(1)
                                keyboard.press_and_release('right')
                                time.sleep(0.2)
                                keyboard.press_and_release('ctrl+b')
                                time.sleep(1)
                                keyboard.press_and_release('right')
                                time.sleep(0.2)
                                keyboard.press_and_release('ctrl+b')
                                time.sleep(1)
                                keyboard.press_and_release('right')
                                time.sleep(0.2)
                                keyboard.press_and_release('ctrl+b')
                                time.sleep(1)
                                keyboard.press_and_release('right')
                                time.sleep(0.2)
                                keyboard.press_and_release('ctrl+b')
                                time.sleep(1)
                                keyboard.press_and_release('right')
                                time.sleep(0.2)
                                keyboard.press_and_release('ctrl+b')
                                time.sleep(1)
                                keyboard.press_and_release('right')
                                time.sleep(0.2)
                                keyboard.press_and_release('ctrl+b')
                                time.sleep(1)
                                keyboard.press_and_release('left')
                                time.sleep(0.2)
                                keyboard.press_and_release('left')
                                time.sleep(0.2)
                                keyboard.press_and_release('left')
                                time.sleep(0.2)
                                keyboard.press_and_release('left')
                                time.sleep(0.2)
                                keyboard.press_and_release('left')
                                time.sleep(0.2)
                                keyboard.press_and_release('left')
                                time.sleep(0.2)
                                keyboard.press_and_release('left')
                                time.sleep(0.2)
                                keyboard.press_and_release('left')
                                time.sleep(0.2)
                                keyboard.press_and_release('left')
                                time.sleep(0.2)
                                keyboard.press_and_release('up')
                                time.sleep(0.2)
                            elif 'no' in b:
                                speak('')
                    elif"stop speaking" in self.query.lower() or 'quite' in self.query.lower():
                        speak('ok sir')
                        break
                            
                    elif"stop" in self.query.lower() or "pause" in self.query.lower():
                        speak('stoping video')
                        keyboard.press_and_release('k')

                    elif "play" in self.query.lower():
                        speak('playing video')
                        keyboard.press_and_release('space bar')
                        
                    elif"main screen" in self.query.lower():
                        speak('getting back to main screen')
                        keyboard.press('alt+f4')
                        keyboard.release('alt+f4')
                        time.sleep(1)
                        keyboard.press('alt+f4')
                        time.sleep(1)
                        keyboard.release('alt+f4')
                        time.sleep(1)
                        keyboard.press('alt+f4')
                        time.sleep(1)
                        keyboard.release('alt+f4')
                        time.sleep(1)
                        keyboard.press_and_release('enter')
                        time.sleep(1)
                        keyboard.press('alt+f4')
                        time.sleep(1)
                        keyboard.release('alt+f4')
                        time.sleep(1)
                        keyboard.press('alt+f4')
                        time.sleep(1)
                        keyboard.release('alt+f4')
                        time.sleep(1)
                        keyboard.press_and_release('enter')
                        time.sleep(1)
                        keyboard.press('alt+f4')
                        time.sleep(1)
                        keyboard.release('alt+f4')
                        time.sleep(1)
                        keyboard.press_and_release('enter')
                        time.sleep(1)
                        keyboard.press('alt+f4')
                        time.sleep(1)
                        keyboard.release('alt+f4')
                        time.sleep(1)
                        keyboard.press_and_release('enter')
                        time.sleep(1)
                        keyboard.press('alt+f4')
                        time.sleep(1)
                        keyboard.release('alt+f4')
                        time.sleep(1)
                        keyboard.press_and_release('enter')
                        time.sleep(1)
                        keyboard.press('alt+f4')
                        time.sleep(1)
                        keyboard.release('alt+f4')
                        time.sleep(1)
                        keyboard.press_and_release('enter')
                        time.sleep(1)
                        keyboard.press('alt+f4')
                        time.sleep(1)
                        keyboard.release('alt+f4')
                        time.sleep(1)
                        keyboard.press_and_release('enter')
                        time.sleep(1)
                        keyboard.release('alt+f4')
                        time.sleep(1)
                        keyboard.press_and_release('enter')
                        time.sleep(1)
                        keyboard.release('alt+f4')
                        time.sleep(1)
                        keyboard.press_and_release('enter')
                        time.sleep(1)
                        keyboard.release('alt+f4')
                        time.sleep(1)
                        keyboard.press_and_release('enter')
                        time.sleep(1)
                        keyboard.release('alt+f4')
                        time.sleep(1)
                        keyboard.press_and_release('enter')
                        time.sleep(1)
                        keyboard.release('alt+f4')
                        time.sleep(1)
                        keyboard.press_and_release('enter')
                        time.sleep(1)
                        keyboard.release('alt+f4')
                        time.sleep(1)
                        keyboard.press_and_release('enter')
                        time.sleep(1)
                        keyboard.release('alt+f4')
                        time.sleep(1)
                        keyboard.press_and_release('enter')
                        time.sleep(1)
                        keyboard.release('alt+f4')
                        time.sleep(1)
                        keyboard.press_and_release('enter')
                        time.sleep(1)
                
                    
                    elif"articles" in self.query.lower():
                        speak('sir, opening all articles pdf')
                        webbrowser.open('https://legislative.gov.in/sites/default/files/COI.pdf')
                        
                    elif "scroll down" in self.query.lower():
                        keyboard.press_and_release('pagedown')
                        speak('sir that much')
                        
                    elif "scroll up" in self.query.lower():
                        keyboard.press_and_release('pageup')
                        speak('sir that much')
                        
                    elif"run" in self.query.lower():
                        speak('Searching  in run...')
                        speak('what do you want to search in run')
                        query2 = self.takeCommand()
                        keyboard.press_and_release('win+r')
                        time.sleep(1)
                        keyboard.write(query2)
                        time.sleep(1)
                        keyboard.press_and_release('enter')
                        
                    elif"find" in self.query.lower():
                        speak('Searching  in files...')
                        speak('what do you want to search in file explorer')
                        query2 = self.takeCommand()
                        keyboard.press_and_release('win+e')
                        time.sleep(3)
                        keyboard.press_and_release('ctrl+f')
                        time.sleep(2)
                        keyboard.write(query2)
                        time.sleep(2)
                        keyboard.press_and_release('enter')
                        time.sleep(4)
                        speak('sir which number of file')
                        self.query= self.takeCommand()
                        keyboard.press_and_release('tab')
                        if  "1" in self.query:
                            keyboard.press_and_release('down')
                            keyboard.press_and_release('up')
                            keyboard.press_and_release('enter')
                        elif  "2" in self.query:
                            keyboard.press_and_release('down')
                            keyboard.press_and_release('enter')
                        elif "3" in self.query:
                            keyboard.press_and_release('down')
                            keyboard.press_and_release('down')
                            keyboard.press_and_release('enter')
                        elif  "4" in self.query:
                            keyboard.press_and_release('down')
                            keyboard.press_and_release('down')
                            keyboard.press_and_release('down')
                            keyboard.press_and_release('enter')
                        elif "5" in self.query:
                            keyboard.press_and_release('down')
                            keyboard.press_and_release('down')
                            keyboard.press_and_release('down')
                            keyboard.press_and_release('down')
                            keyboard.press_and_release('enter')
                        elif "6" in self.query:
                            keyboard.press_and_release('down')
                            keyboard.press_and_release('down')
                            keyboard.press_and_release('down')
                            keyboard.press_and_release('down')
                            keyboard.press_and_release('down')
                            keyboard.press_and_release('enter')
                        elif "7" in self.query:
                            keyboard.press_and_release('down')
                            keyboard.press_and_release('down')
                            keyboard.press_and_release('down')
                            keyboard.press_and_release('down')
                            keyboard.press_and_release('down')
                            keyboard.press_and_release('down')
                            keyboard.press_and_release('enter')
                        elif  "8" in self.query:
                            keyboard.press_and_release('down')
                            keyboard.press_and_release('down')
                            keyboard.press_and_release('down')
                            keyboard.press_and_release('down')
                            keyboard.press_and_release('down')
                            keyboard.press_and_release('down')
                            keyboard.press_and_release('down')
                            keyboard.press_and_release('enter')
                        elif  "9" in self.query:
                            keyboard.press_and_release('down')
                            keyboard.press_and_release('down')
                            keyboard.press_and_release('down')
                            keyboard.press_and_release('down')
                            keyboard.press_and_release('down')
                            keyboard.press_and_release('down')
                            keyboard.press_and_release('down')
                            keyboard.press_and_release('down')
                            keyboard.press_and_release('enter')
                        else:
                            keyboard.press_and_release('down')
                            keyboard.press_and_release('up')
                            keyboard.press_and_release('enter')
                            
                    elif"invitation" in self.query.lower():
                        speak('wait sir lemme open outlook')
                        keyboard.press_and_release('win+r')
                        time.sleep(1)
                        keyboard.write('outlook')
                        time.sleep(1)
                        keyboard.press_and_release('enter')
                        time.sleep(8)
                        keyboard.press_and_release('ctrl+n')
                        time.sleep(1)
                        speak('sir tell me the email address')
                        query2 = self.takeCommand()
                        if 'Hotmail' in query2 :
                            w= '@hotmail.com'
                        elif 'Rediff' in query2:
                            w= '@rediff.com'
                        elif 'Outlook' in query2:
                            w = '@outlook.com'
                        elif 'Yahoo' in query2:
                            w= '@yahoo.com'
                        else:
                            w='@gmail.com'
                        speak('sir now tell me the email id')
                        query1 = self.takeCommand()
                        x=query1.lower()
                        z= x+w
                        z=z.replace(' ','')
                        keyboard.write(z)
                        time.sleep(2)
                        keyboard.press_and_release('enter')
                        time.sleep(2)
                        keyboard.press_and_release('tab')
                        print('done')
                        time.sleep(1)
                        speak('is there any caption,, yes or no' )
                        cc= self.takeCommand()
                        if 'yes'in cc:
                            speak('sir tell me your caption')
                            k=self.takeCommand()
                            keyboard.write(k)
                            time.sleep(1)
                        else:
                            keyboard.write('')
                        time.sleep(1)
                        keyboard.press_and_release('tab')
                        time.sleep(1)
                        speak('is there any subject,, yes or no' )
                        sub= self.takeCommand()
                        if 'yes'in sub:
                            speak('sir tell me your subject')
                            k=self.takeCommand()
                            keyboard.write(k)
                            time.sleep(1)
                        else:
                            keyboard.write('')
                        keyboard.press_and_release('tab')
                        speak('what is your message')
                        mess= self.takeCommand()
                        keyboard.write(mess)
                        while 1==1:
                            speak('sir is that enough')
                            append=self.takeCommand()
                            if 'no' in append:
                                mess=self.takeCommand()
                                keyboard.write(mess)
                            else:
                                break
                        speak('sending this email')
                        keyboard.press_and_release('shift+tab')
                        time.sleep(0.2)
                        keyboard.press_and_release('shift+tab')
                        time.sleep(0.2)
                        keyboard.press_and_release('shift+tab')
                        time.sleep(0.2)
                        keyboard.press_and_release('shift+tab')
                        time.sleep(0.2)
                        keyboard.press_and_release('shift+tab')
                        time.sleep(0.2)
                        keyboard.press_and_release('shift+tab')
                        time.sleep(0.2)
                        keyboard.press_and_release('enter')
                        time.sleep(0.2)
                        keyboard.press_and_release('right')
                        time.sleep(0.2)
                        keyboard.press_and_release('enter')
                        time.sleep(0.2)
                        try :
                            os.system('TASKKILL /F /IM OUTLOOK.EXE')
                        except Exception as e:
                            return
                        
                    elif"task manager" in self.query.lower():
                        speak('sir you want to make new task or ,,old ,,or percentage')
                        query2=self.takeCommand()
                        if 'new' in query2:
                            speak('sir making new task')
                            speak('sir, tell me the task')
                            mainhead=self.takeCommand()
                            speak('ok sir')
                            speak('how many sub task are there')
                            self.query=self.takeCommand()
                            if '1' in self.query:
                                no=1
                                speak('sir tell me your sub task orderwise')
                                speak('sir tell me your first task')
                                a=self.takeCommand()
                                m=[a]
                            elif "2"in self.query:
                                no=2
                                speak('sir tell me your sub task orderwise')
                                speak('sir tell me your first task')
                                a=self.takeCommand()
                                speak('sir tell me your second task ')
                                b=self.takeCommand()
                                m=[a,b]
                            elif "3"in self.query:
                                no=3
                                speak('sir tell me your sub task orderwise')
                                speak('sir tell me your first task')
                                a=self.takeCommand()
                                speak('sir tell me your second task ')
                                b=self.takeCommand()
                                speak('sir tell me your third task ')
                                c=self.takeCommand()
                                m=[a,b,c]
                            elif "4"in self.query:
                                no=4
                                speak('sir tell me your sub task orderwise')
                                speak('sir tell me your first task')
                                a=self.takeCommand()
                                speak('sir tell me your second task ')
                                b=self.takeCommand()
                                speak('sir tell me your third task ')
                                c=self.takeCommand()
                                speak('sir tell me your fourth task ')
                                d=self.takeCommand()
                                m=[a,b,c,d]
                            elif  "5"in self.query:
                                no=5
                                speak('sir tell me your sub task orderwise')
                                speak('sir tell me your first task')
                                a=self.takeCommand()
                                speak('sir tell me your second task ')
                                b=self.takeCommand()
                                speak('sir tell me your third task ')
                                c=self.takeCommand()
                                speak('sir tell me your fourth task ')
                                d=self.takeCommand()
                                speak('sir tell me your fifth task ')
                                e=self.takeCommand()
                                m=[a,b,c,d,e]
                            elif "6"in self.query:
                                no=6
                                speak('sir tell me your sub task orderwise')
                                speak('sir tell me your first task')
                                a=self.takeCommand()
                                speak('sir tell me your second task ')
                                b=self.takeCommand()
                                speak('sir tell me your third task ')
                                c=self.takeCommand()
                                speak('sir tell me your fourth task ')
                                d=self.takeCommand()
                                speak('sir tell me your fifth task ')
                                e=self.takeCommand()
                                speak('sir tell me your sixth task ')
                                f=self.takeCommand()
                                m=[a,b,c,d,e,f]
                            elif "7"in self.query:
                                no=7
                                speak('sir tell me your sub task orderwise')
                                speak('sir tell me your first task')
                                a=self.takeCommand()
                                speak('sir tell me your second task ')
                                b=self.takeCommand()
                                speak('sir tell me your third task ')
                                c=self.takeCommand()
                                speak('sir tell me your fourth task ')
                                d=self.takeCommand()
                                speak('sir tell me your fifth task ')
                                e=self.takeCommand()
                                speak('sir tell me your sixth task ')
                                f=self.takeCommand()
                                speak('sir tell me your seventh task ')
                                g=self.takeCommand()
                                m=[a,b,c,d,e,f,g]
                            elif "8" in self.query:
                                no=8
                                speak('sir tell me your sub task orderwise')
                                speak('sir tell me your first task')
                                a=self.takeCommand()
                                speak('sir tell me your second task ')
                                b=self.takeCommand()
                                speak('sir tell me your third task ')
                                c=self.takeCommand()
                                speak('sir tell me your fourth task ')
                                d=self.takeCommand()
                                speak('sir tell me your fifth task ')
                                e=self.takeCommand()
                                speak('sir tell me your sixth task ')
                                f=self.takeCommand()
                                speak('sir tell me your seventh task ')
                                g=self.takeCommand()
                                speak('sir tell me your eighth task ')
                                h=self.takeCommand()
                                m=[a,b,c,d,e,f,g,h]
                            elif "9" in self.query:
                                no=9
                                speak('sir tell me your sub task orderwise')
                                speak('sir tell me your first task')
                                a=self.takeCommand()
                                speak('sir tell me your second task ')
                                b=self.takeCommand()
                                speak('sir tell me your third task ')
                                c=self.takeCommand()
                                speak('sir tell me your fourth task ')
                                d=self.takeCommand()
                                speak('sir tell me your fifth task ')
                                e=self.takeCommand()
                                speak('sir tell me your sixth task ')
                                f=self.takeCommand()
                                speak('sir tell me your seventh task ')
                                g=self.takeCommand()
                                speak('sir tell me your eighth task ')
                                h=self.takeCommand()
                                speak('sir tell me your ninth task ')
                                i=self.takeCommand()
                                m=[a,b,c,d,e,f,g,h,i]
                        elif 'old' in query2:
                            li=m
                            st=[]
                            for i in range(0,len(m)):
                                speak(m[i])
                                
                                ja=m[i]
                                speak('is it completed?')
                                yes=self.takeCommand()
                                if 'yes' in yes:
                                    speak('it is removed')
                                else:
                                    speak('ok sir')
                                    st.append(m[i])
                            print(st)
                            m=st
                            print(m)
                        elif 'show' in query2:
                            per=len(m)
                            k=per/no
                            real=int(k*100)
                            speak('you work is')
                            speak(real)
                            speak('percent incompleted')       
                        else:
                            speak('do it from start')
                            break
                        
                    elif"current" in self.query.lower():
                        speak('opening best site for current affairs')
                        webbrowser.open('https://www.gktoday.in/current-affairs/category/sports-current-affairs/')

                    elif "teach" in self.query.lower():
                        speak('sir currently i can teach you coding')
                        speak('is that ok sir')
                        teac=self.takeCommand()
                        if 'yes' in teac:
                            speak('sir lets start it with very basic')
                            speak('but sir have you installed python in your computer')
                            teach=self.takeCommand()
                            if 'no' in teach:
                                speak('sir, download it from here')
                                webbrowser.open('https://www.python.org/downloads/')
                                time.sleep(3)
                                speak('sir, click on download python which is in yellow border')
                                speak('sir run it and come again in teach mode')
                                break
                            else:
                                speak('sir you must be a keen learner')
                            speak('sir, python has two types of modes,, one is interactive and the other one is scriptive')
                            speak('lets look at interactive mode')
                            keyboard.press_and_release('windows')
                            time.sleep(0.2)
                            keyboard.write('python')
                            time.sleep(0.2)
                            keyboard.press_and_release('enter')
                            time.sleep(2)
                            speak('sir, it is an interactive mode where you can get result after every line')
                            speak('lets try it with some basic addition and subtraction')
                            speak('sir, you can use it like this')
                            speak('like we have to add 2 numbers')
                            speak('suppose we will add 3 and 4 with it')
                            time.sleep(1)
                            speak('we will write like this')
                            keyboard.write('3')
                            time.sleep(1)
                            keyboard.write('+')
                            time.sleep(1)
                            keyboard.write('4')
                            time.sleep(1)
                            speak('and press enter like this')
                            keyboard.press_and_release('enter')
                            speak('boom sir now you know how to add on python')
                            speak('now you have to try it yourself, multiply 5 by 4')
                            speak('your lecture 1 is completed')
                            
                    elif"meeting" in self.query.lower():
                        speak('sir, tell me the topic of meeting')
                        header= self.takeCommand()
                        speak('lets write it in front of you')
                        keyboard.press_and_release('win')
                        time.sleep(0.2)
                        keyboard.write('wordpad')
                        time.sleep(0.2)
                        keyboard.press_and_release('enter')
                        time.sleep(2)
                        keyboard.press_and_release('ctrl+e')
                        time.sleep(0.2)
                        keyboard.press_and_release('ctrl+b')
                        time.sleep(0.2)
                        keyboard.write(header)
                        keyboard.press_and_release('enter')
                        time.sleep(0.2)
                        keyboard.press_and_release('ctrl+l')
                        time.sleep(0.2)
                        keyboard.press_and_release('ctrl+b')
                        time.sleep(0.2)
                        speak('your solution')
                        body=self.takeCommand()
                        keyboard.write(body)
                        while 1==1:
                            speak('sir, is that enough')
                            yes=self.takeCommand()
                            if 'yes' in yes:
                                keyboard.press_and_release('alt+f4')
                                time.sleep(1)
                                keyboard.press_and_release('enter')
                                time.sleep(1)
                                speak('file name')
                                namee=self.takeCommand()
                                keyboard.write(namee)
                                time.sleep(1)
                                keyboard.press_and_release('enter')
                                time.sleep(1)
                                keyboard.press_and_release('enter')
                                break
                            else:
                                speak('writing')
                                body=self.takeCommand()
                                keyboard.write(body)
                        
                    elif"bye" in self.query.lower():
                        speak('bye,, sir')
                        
                    elif"good boy" in self.query.lower():
                        speak('thankyou sir')
                        
                    elif"ms word" in self.query.lower():
                        speak('sir i am pro helper at ms word')
                        speak('sir opening up M S word')
                        keyboard.press_and_release('win')
                        time.sleep(1)
                        keyboard.write('word')
                        time.sleep(1)
                        keyboard.press_and_release('enter')
                        time.sleep(10)
                        speak ('sir i am selecting blank document')
                        time.sleep(1)
                        keyboard.press_and_release('enter')
                        speak('now tell me what to do')
                        while 1==1:
                            book=self.takeCommand()
                            if 'bold' in book:
                                speak('sir bold is now clicked')
                                keyboard.press_and_release('ctrl + b')
                            elif 'note' in book:
                                while 1==1:
                                    speak('tell me sir what to write')
                                    content=self.takeCommand()
                                    time.sleep(2)
                                    keyboard.write(content)
                                    time.sleep(1)
                                    speak('sir is that enough')
                                    yes=self.takeCommand()
                                    if 'yes' in yes:
                                        speak('ok sir')
                                        break
                                    else:
                                        speak('')
                            elif 'italic' in book:
                                speak('sir italic is now clicked')
                                keyboard.press_and_release('ctrl + i')
                            elif 'underline' in book :
                                speak('sir underline is clicked')
                                keyboard.press_and_release('ctrl+u')
                            elif 'exit' in book:
                                speak('sir exiting')
                                break
                            elif 'increase font' in book:
                                speak('increasing font size with one')
                                keyboard.press_and_release('ctrl+]')
                            elif 'decrease font' in book:
                                speak('decreasing font size with one')
                                keyboard.press_and_release('ctrl+[')
                            elif 'centre' in book:
                                speak('center align is clicked')
                                keyboard.press_and_release('ctrl+e')
                            elif 'left' in book:
                                speak('left align is clicked')
                                keyboard.press_and_release('ctrl+l')
                            elif 'right' in book:
                                speak('right align is clicked')
                                keyboard.press_and_release('ctrl+r')
                            elif 'undo' in book:
                                speak('ok sir,undo')
                                keyboard.press_and_release('ctrl+z')
                            elif 'redo' in book:
                                speak('ok sir re doing')
                                keyboard.press_and_release('ctrl+y')
                            elif 'zoom' in book:
                                speak('ok sir zooming')
                                keyboard.press_and_release('alt+w')
                                time.sleep(1)
                                keyboard.press_and_release('alt+q')
                                time.sleep(1)
                                speak('sir tell me wht percentage you want 100 ,, 75,, 200')
                                perce=self.takeCommand()
                                if '1' in perce:
                                    speak('ok sir making it 100 percent')
                                    keyboard.press_and_release('1')
                                    time.sleep(1)
                                    keyboard.press_and_release('enter')
                                elif '2' in perce:
                                    speak('ok sir making it 200 percent')
                                    keyboard.press_and_release('2')
                                    time.sleep(1)
                                    keyboard.press_and_release('enter')
                                elif '7' in perce:
                                    speak('ok sir making it 75 percent')
                                    keyboard.press_and_release('7')
                                    time.sleep(1)
                                    keyboard.press_and_release('enter')
                                else:
                                    speak('this is not in option')
                            elif 'insert pic' in book:
                                speak('ok sir')
                                keyboard.press_and_release('alt+n')
                                time.sleep(1)
                                keyboard.press_and_release('alt+p')
                                time.sleep(1)
                                speak('you can choose sir')  
                            elif 'insert' in book:
                                speak('ok sir going to insert')
                                keyboard.press_and_release('alt+n')
                            
                            elif 'next line' in book:
                                speak('ok sir')
                                keyboard.press_and_release('enter')
                            elif 'save' in book:
                                speak('ok sir')
                                keyboard.press_and_release('alt+f4')
                                keyboard.press_and_release('enter')
                                keyboard.press_and_release('enter')
                                keyboard.press_and_release('enter')
                            elif '' in book:
                                speak('')
                        
                    elif"window" in self.query.lower():
                        speak('opening window')
                        keyboard.press_and_release('windows')
                        speak('what do you wnna search, sir')
                        m=self.takeCommand()
                        keyboard.write(m)
                        time.sleep(1)
                        keyboard.press_and_release('enter')
                        
                    elif"whatsapp" in self.query.lower():
                        speak('ok sir opening whatsapp')
                        keyboard.press_and_release('windows')
                        time.sleep(1)
                        keyboard.write('whatsapp')
                        time.sleep(1)
                        keyboard.press_and_release('enter')
                        time.sleep(6)
                        speak('whom do you wnna message')
                        k=self.takeCommand()
                        keyboard.write(k)
                        time.sleep(1)
                        keyboard.press_and_release('down')
                        time.sleep(1)
                        keyboard.press_and_release('enter')
                        time.sleep(1)
                        keyboard.press_and_release('tab')
                        time.sleep(0.2)
                        keyboard.press_and_release('tab')
                        time.sleep(0.2)
                        keyboard.press_and_release('tab')
                        time.sleep(0.2)
                        keyboard.press_and_release('tab')
                        time.sleep(0.2)
                        keyboard.press_and_release('tab')
                        time.sleep(0.2)
                        keyboard.press_and_release('tab')
                        time.sleep(0.2)
                        keyboard.press_and_release('tab')
                        time.sleep(0.2)
                        keyboard.press_and_release('tab')
                        time.sleep(0.2)
                        keyboard.press_and_release('tab')
                        time.sleep(0.2)
                        keyboard.press_and_release('tab')
                        time.sleep(0.2)
                        keyboard.press_and_release('tab')
                        time.sleep(0.2)
                        keyboard.press_and_release('tab')
                        time.sleep(0.2)
                        keyboard.press_and_release('tab')
                        time.sleep(0.2)
                        speak('tell me sir what to send')
                        mess=self.takeCommand()
                        keyboard.write(mess)
                        time.sleep(1)
                        keyboard.press_and_release('enter')
                        speak('sir closing whatsapp')
                        keyboard.press_and_release('alt+f4')
                        
                    elif"hack" in self.query.lower():
                        speak('sir thats illegal to hack anything so, sir be a nice guy')
                        
                    elif"emergency" in self.query.lower():
                        keyboard.press_and_release('windows')
                        time.sleep(1)
                        keyboard.write('whatsapp')
                        time.sleep(1)
                        keyboard.press_and_release('enter')
                        time.sleep(5)
                        keyboard.write('mummy')
                        time.sleep(1)
                        keyboard.press_and_release('down')
                        time.sleep(1)
                        keyboard.press_and_release('enter')
                        time.sleep(1)
                        keyboard.press_and_release('tab')
                        time.sleep(0.2)
                        keyboard.press_and_release('tab')
                        time.sleep(0.2)
                        keyboard.press_and_release('tab')
                        time.sleep(0.2)
                        keyboard.press_and_release('tab')
                        time.sleep(0.2)
                        keyboard.press_and_release('tab')
                        time.sleep(0.2)
                        keyboard.press_and_release('tab')
                        time.sleep(0.2)
                        keyboard.press_and_release('tab')
                        time.sleep(0.2)
                        keyboard.press_and_release('tab')
                        time.sleep(0.2)
                        keyboard.press_and_release('tab')
                        time.sleep(0.2)
                        keyboard.press_and_release('tab')
                        time.sleep(0.2)
                        keyboard.press_and_release('tab')
                        time.sleep(0.2)
                        keyboard.press_and_release('tab')
                        time.sleep(0.2)
                        keyboard.press_and_release('tab')
                        time.sleep(0.2)
                        keyboard.write('help me')
                        time.sleep(4)
                        keyboard.press_and_release('enter')
                        time.sleep(1)
                        keyboard.press_and_release('alt+f4')
                        
                    elif"help me" in self.query.lower():
                        speak('activating')
                        os.startfile('C:/Users/Dell/Pictures/presentation.mp3')
                        time.sleep(3)
                        keyboard.press_and_release('ctrl+t')
                        time.sleep(1)
                        keyboard.press_and_release('alt+tab')
                        
                    elif"homework" in self.query.lower():
                        speak('multiply 4 into 5')
                        events = []
                        keyboard.hook(events.append)
                        keyboard.wait("enter")
                        keyboard.unhook(events.append)
                        string = ""
                        for i in events:
                            a = str(i)
                            if 'down' in a:
                                temp = a[a.index("(")+1: a.index(" down)")]
                                if temp in "asdfghjklqwertyuiopzxcvbnm1234567890!@#$%^&*()+=-_":
                                    string += temp
                                elif temp == "backspace":
                                    string = string.removesuffix(string[-1])
                        string=str(string)
                        if string == '4*5':
                            speak('sir you are absolutly correct')
                        else:
                            speak('you are wrong sir, learn from lecture 1 sir')
                            
                    elif"exit window" in self.query.lower():
                        keyboard.press_and_release('alt+f4')
                        
                    elif"switch tab" in self.query.lower() or 'change tab' in self.query.lower():
                        keyboard.press_and_release('alt+tab')
                        
                    elif "close" in self.query.lower() or'exit' in self.query.lower()or 'quit' in self.query.lower():
                        speak('closing')
                        self.close
                        
                    elif"" in self.query:
                        speak('')
      
startExecution = MainThread()

class Main(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_missminute()
        self.ui.setupUi(self)
        self.ui.pushButton.clicked.connect(self.startTask)
        self.ui.pushButton_2.clicked.connect(self.close)

    def startTask(self):
        startExecution.start()
        self.ui.movie1 = QtGui.QMovie("C://bodymin.gif")
        self.ui.label.setMovie(self.ui.movie1)
        self.ui.movie1.start()
        self.ui.movie2 = QtGui.QMovie("C://c302f8106533937.5f91cf38ecb08.gif")
        self.ui.label_3.setMovie(self.ui.movie2)
        self.ui.movie2.start()
        self.ui.movie3 = QtGui.QMovie("C://mouthmin.gif")
        self.ui.label_2.setMovie(self.ui.movie3)
        self.ui.movie3.start()
        timer= QTimer(self)
        timer.timeout.connect(self.showTime)
        timer.start(1000)

    def showTime(self):
        current_time = QTime.currentTime()
        label_time = current_time.toString('hh:mm:ss')
        self.ui.textBrowser.setText('I AM SUPERHERO......')

app = QApplication(sys.argv)
main = Main()
main.show()
sys.exit(app.exec_())
