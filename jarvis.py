import pyttsx3

import speech_recognition as sr 
import datetime
import wikipedia  
import webbrowser
import random
import sys
import time 
import os
import os.path
import requests
import cv2  
from requests import get  
# import pywhatkit as kit     
import smtplib  
import pyjokes  
import pyautogui  
import PyPDF2
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import instaloader  
import operator  
from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtCore import QTimer, QTime, QDate, Qt
from PyQt5.QtGui import QMovie
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.uic import loadUiType
from JarvisUi import Ui_JarvisUi
from bs4 import BeautifulSoup

"""
IN PLACEOF PYTTSX3 WE CAN ALSO USE WIN32COM.CLIENT

# Python program to convert 
# text to speech 

# import the required module from text to speech conversion 
import win32com.client 

# Calling the Disptach method of the module which  
# interact with Microsoft Speech SDK to speak 
# the given input from the keyboard 

speaker = win32com.client.Dispatch("SAPI.SpVoice") 

while 1: 
    print("Enter the word you want to speak it out by computer") 
    s = input() 
    speaker.Speak(s) 

# To stop the program press 
# CTRL + Z 
"""

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices');
# print(voices[0].id)
engine.setProperty('voices', voices[1].id)


# text to speech
def speak(audio):
    engine.say(audio)
    print(audio)
    engine.runAndWait()


# def speak(audio):
#     speaker = Dispatch("SAPI.SpVoice")
#     print(audio)
#     speaker.Speak(audio)


# to wish
def wish():
    hour = int(datetime.datetime.now().hour)
    tt = time.strftime("%I:%M %p")

    if hour >= 0 and hour <= 12:
        speak(f"good morning Mam, its {tt}")
    elif hour >= 12 and hour <= 18:
        speak(f"good afternoon Mam, its {tt}")
    else:
        speak(f"good evening Mam, its {tt}")
    speak("I am jarvis, please tell me how can i help you")


# to send email
def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('YOUR EMAIL ADDRESS', 'YOUR PASSWORD')
    server.sendmail('YOUR EMAIL ADDRESS', to, content)
    server.close()


# for news updates
def news():
    main_url = 'http://newsapi.org/v2/top-headlines?sources=techcrunch&apiKey=83263a48521a48a797182dbc3926e513'

    main_page = requests.get(main_url).json()
    print(main_page)
    articles = main_page["articles"]
    print(articles)
    head = []
    day = ["first", "second", "third", "fourth", "fifth", "sixth", "seventh", "eighth", "ninth", "tenth"]
    for ar in articles:
        head.append(ar["title"])
    for i in range(len(day)):
        speak(f"today's {day[i]} news is: {head[i]}")
    
class MainThread(QThread):
    def __init__(self):
        super(MainThread, self).__init__()

    def takecommand(self):
        try:
            r = sr.Recognizer()
            with sr.Microphone() as source:
                print("listening...")
                r.pause_threshold = 0.5
                audio = r.listen(source, timeout=10, phrase_time_limit=12)
        except AttributeError as e:
            engine.say("Time out thank you for using Jarvis",e)
        except sr.WaitTimeoutError as e:
            engine.say("something went wrong try again",e)

            # r.pause_threshold = 1
            # r.adjust_for_ambient_noise(source)
            # audio = r.listen(source)
            # audio = r.listen(source,timeout=4,phrase_time_limit=7)

        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language='en-in')
            print(f"user said: {query}")

        except Exception as e:
            # speak("Say that again please...")
            return "none"
        query = query.lower()
        return query

    def run(self):
        self.TaskExecution()
        speak("please say wakeup to continue")
        while True:
             self.query = self.takecommand()
             if "wake up" in self.query or "are you there" in self.query or "hello" in self.query:
                 self.TaskExecution()

    def TaskExecution(self):
        wish()
        while True:
            self.query = self.takecommand()

            # logic building for tasks

            if "open notepad" in self.query:
                npath = "C:\\Windows\\system32\\notepad.exe"
                os.startfile(npath)

            # to close notepad application
            elif "close notepad" in self.query:
                speak("okay mam, closing notepad")
                os.system("taskkill /f /im notepad.exe")

            elif "open adobe reader" in self.query or "pdf reader" in self.query:
                apath = "C:\\Program Files (x86)\\Adobe\\Reader 11.0\\Reader\\AcroRd32.exe"
                os.startfile(apath)

            # to close adobe reader application
            elif "close adobe reader" in self.query or "pdf reader" in self.query:
                speak("okay mam, closing adobe")
                os.system("taskkill /f /im AcroRd32.exe")
                continue

            elif "open command prompt" in self.query:
                os.system("start cmd")
                continue

            elif "open camera" in self.query or "webcam" in self.query:
                cap = cv2.VideoCapture(0)
        
                while True:
                    ret, img = cap.read()
                    cv2.imshow('webcam', img)
                    k = cv2.waitKey(50)
                    if k == 27:
                        break;
                cap.release()
                cv2.destroyAllWindows()
                continue    # use esc button to close camera

            elif "ip address" in self.query:
                ip = get('https://api.ipify.org').text
                speak(f"your IP address is {ip}")
                continue

            elif "wikipedia" in self.query or "search on wikipedia" in self.query:
                speak("searching wikipedia....")
                speak("What topic would you like to search on Wikipedia?")
                query = self.takecommand()
                webbrowser.open(f"https://www.wikipedia.org//{query}")
                results = wikipedia.search(query)
                speak("according to wikipedia")
                speak(results)
                continue
                # print(results)

            elif "open youtube" in self.query or "play music on youtube" in self.query or "play music" in self.query:
                #webbrowser.open("www.youtube.com")
                speak("Sure,which video you wish to play")
                query=self.takecommand()
                #video = query.split('')[1]
                try:
                    speak(f"Okay mam, playing {query} on youtube")
                    url = f"https://www.youtube.com/results?search_query={query}"
                    webbrowser.open(url)
                except Exception as e:
                    print(e)
                    speak("sorry Mam, i am not able to find this video on youtube")
                continue
            
            # to close youtube application
            elif "close youtube" in self.query:
                speak("okay mam, closing youtube")
                os.system("taskkill /f /im msedge.exe")
                continue

            elif "open facebook" in self.query:
                webbrowser.open("www.facebook.com")

            elif "open stackoverflow" in self.query:
                webbrowser.open("www.stackoverflow.com")

            elif "open google" in self.query or "search on google" in self.query:
                speak("mam, what should i search on google")
                cm = self.takecommand()
                webbrowser.open(f"{cm}")
                continue

            # elif "send whatsapp message" in self.query:
            #     kit.sendwhatmsg("+91 user_number", "your_message",4,13)
            #     time.sleep(120)
            #     speak("message has been sent")


            # elif "email to savi" in self.query:
            #     try:
            #         speak("what should i say?")
            #         content = takecommand()
            #         to = "EMAIL OF THE OTHER PERSON"
            #         sendEmail(to,content)
            #         speak("Email has been sent to savi")

            #     except Exception as e:
            #         print(e)
            #         speak("sorry Mam, i am not able to sent this mail to avi")

            elif "you can sleep" in self.query or "sleep now" in self.query:
                speak("okay mam, i am going to sleep you can call me anytime.")
                # sys.exit()
                # gifThread.exit()
                break



            # ----to check power bckup percentage-------
            elif "how much power left" in self.query or "how much power we have" in self.query or "battery" in self.query:
                import psutil
                battery=psutil.sensors_battery()
                percentage=battery.percent
                speak(f"Mam your system {percentage} percent battery")
                if percentage>=75:
                    speak("we have enough power to continue our work")
                elif percentage>=40 and percentage<=75:
                    speak("we should connect our system to charging point to charge our battery")
                elif percentage<=15 and percentage<=30:
                    speak("we dont have enough power to work, please connect to charging")
                elif percentage<=15:
                    speak("we have very low power, please connect to charge the system will shutdown very soon")

               # ------- to check internet speed-------
                # elif "internet speed" in self.query:
                #     import speedtest
                #     st=speedtest.Speedtest()
                #     d1=st.download()
                #     up=st.upload()
                #     speak(f"sir we have {d1} bit per second downloading speed andd {up} bit per second uploading speed")
             
   

            # to find a joke
            elif "tell me some joke" in self.query:
                joke = pyjokes.get_joke()
                speak(joke)
                

            # elif "shut down the system" in self.query:
            #     os.system("shutdown /s /t 5")

            # elif "restart the system" in self.query:
            #     os.system("shutdown /r /t 5")

            # elif "sleep the system" in self.query:
            #     os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")

            elif "hello" in self.query or "hey" in self.query:
                speak("hello mam, may i help you with something.")

            elif "how are you" in self.query or "How you doing jarvis" in self.query:
                speak("i am fine mam, what about you.")
                
                if self.takecommand():
                    speak("Nice to hear that")
                else:
                    speak("what can I do for you mam?")
                continue

            elif "thank you" in self.query or "thanks" in self.query:
                speak("it's my pleasure mam.")

            elif "goodbye jarvis" in self.query or "offline" in self.query or "bye jarvis" in self.query:
                speak("Alright bye Mam, going offline. It was nice talking with you, Have a great day")
                sys.exit()
            ###################################################################################################################################
            ###########################################################################################################################################

            elif 'switch the window' in self.query or "chnage the window" in self.query:
                pyautogui.keyDown("alt")
                pyautogui.press("tab")
                time.sleep(1)
                pyautogui.keyUp("alt")


            elif "tell me news" in self.query or "tell me todays news" in self.query:
                speak("please wait mam, feteching the latest news")
                # if self.takecommand() == "stop":
                #     audio = r.listen(source)
                #     sys.exit
                # else:
                news()


            elif "email to savi" in self.query:

                speak("mam what should i say")
                self.query = self.takecommand()
                if "send a file" in self.query:
                    email = 'vsavi026@gmail.com'  # Your email
                    password = 'yourpass@123'  # Your email account password
                    send_to_email = 'vsavi026@gmail.com'  # Whom you are sending the message to
                    speak("okay mam, what is the subject for this email")
                    self.query = self.takecommand()
                    subject = self.query  # The Subject in the email
                    speak("and mam, what is the message for this email")
                    self.query2 = self.takecommand()
                    message = self.query2  # The message in the email
                    speak("mam please enter the correct path of the file into the shell")
                    file_location = input("please enter the path here")  # The File attachment in the email

                    speak("please wait,i am sending email now")

                    msg = MIMEMultipart()
                    msg['From'] = email
                    msg['To'] = send_to_email
                    msg['Subject'] = subject

                    msg.attach(MIMEText(message, 'plain'))

                    # Setup the attachment
                    filename = os.path.basename(file_location)
                    attachment = open(file_location, "rb")
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)

                    # Attach the attachment to the MIMEMultipart object
                    msg.attach(part)

                    server = smtplib.SMTP('smtp.gmail.com', 587)
                    server.starttls()
                    server.login(email, password)
                    text = msg.as_string()
                    server.sendmail(email, send_to_email, text)
                    server.quit()
                    speak("email has been sent to savi")

                else:
                    email = 'your@gmail.com'  # Your email
                    password = 'your_pass@123'  # Your email account password
                    send_to_email = 'to_person@gmail.com'  # Whom you are sending the message to
                    message = query  # The message in the email

                    server = smtplib.SMTP('smtp.gmail.com', 587)  # Connect to the server
                    server.starttls()  # Use TLS
                    server.login(email, password)  # Login to the email server
                    server.sendmail(email, send_to_email, message)  # Send the email
                    server.quit()  # Logout of the email server
                    speak("email has been sent to savi")


            ##########################################################################################################################################
            ###########################################################################################################################################

            # ----------calculation---------------
          
            elif "open calculator" in self.query or "can you calculate" in self.query:
   
                # r = sr.Recognizer()
                # with sr.Microphone() as source:
                #     speak("what you want to calculate")
                #     print("Listening...")
                #     audio = r.listen(source)
                # def get_operator_fn(op):
                #     return {
                #         '+': operator.add,
                #         '-': operator.sub,
                #         'x' or '*': operator.mul,
                #         'divided': operator.__truediv__,
                # }[op]
                # try:
                #     command = r.recognize_google(audio)
                #     #print("You said: " + command)
                
                #     result = eval(command)
                #     engine.say("The answer is " + str(result))
                #     print(str(result))
                #     engine.runAndWait()
                # except sr.UnknownValueError:
                #     print("Sorry, I could not understand your command.")
                #     engine.say("Sorry, I could not understand your command.")
                #     engine.runAndWait()
                r = sr.Recognizer()
                with sr.Microphone() as source:
                    speak("Say what you want to calculate, example: 3 plus 3")
                    print("listening.....")
                    r.adjust_for_ambient_noise(source)
                    audio = r.listen(source)
                try:
                    my_string = r.recognize_google(audio)
                    speak(my_string)

                    def get_operator_fn(op):
                        return {
                            '+': operator.add,
                            '-': operator.sub,
                            'x' or '*': operator.mul,
                            '/': operator.__truediv__,
                        }[op]
        
                    def eval_binary_expr(op1, oper, op2):
                        op1, op2 = int(op1), int(op2)
                        return get_operator_fn(oper)(op1, op2)

                    print(eval_binary_expr(*(my_string.split())))
                    cal=eval_binary_expr(*(my_string.split()))
                    speak(f"calculated answer is {cal}")
                except TypeError:
                    engine.say("Sorry, Invalid type error")
                except SyntaxError:
                    engine.say("Sorry, invalid syntax error.")
                except sr.UnknownValueError as e:
                    print(e)
                    engine.say("Sorry, I could not understand your command.")
                    engine.runAndWait()
                finally:
                    engine.say("Thank you for using calculator")
                continue

            # -----------------To find my location using IP Address

            elif "where i am" in self.query or "where we are" in self.query:
                speak("wait mam, let me check")
                try:
                    ipAdd = requests.get('https://api.ipify.org').text
                    print(ipAdd)
                    url = 'https://get.geojs.io/v1/ip/geo/' + ipAdd + '.json'
                    geo_requests = requests.get(url)
                    geo_data = geo_requests.json()
                    # print(geo_data)
                    city = geo_data['city']
                    # state = geo_data['state']
                    country = geo_data['country']
                    speak(f"mam i am not sure, but i think we are in {city} city of {country} country")
                except Exception as e:
                    speak("sorry mam, Due to network issue i am not able to find where we are.")
                    pass




            # -------------------To check a instagram profile----
            elif "instagram profile" in self.query or "search instagram profile" in self.query:
                speak("mam please enter the user name correctly.")
                name = input("Enter username here:")
                webbrowser.open(f"www.instagram.com/{name}")
                speak(f"Mam here is the profile of the user {name}")
                time.sleep(5)
                speak("Mam would you like to download profile picture of this account.")
                condition = self.takecommand()
                if "yes" in condition:
                    mod = instaloader.Instaloader()  # pip install instadownloader
                    mod.download_profile(name, profile_pic_only=True)
                    speak("i am done mam, profile picture is saved in our main folder. now i am ready for next command")
                else:
                    pass

            # -------------------  To take screenshot -------------
            elif "take screenshot" in self.query or "take a screenshot" in self.query:
                speak("mam, please tell me the name for this screenshot file")
                name = self.takecommand()
                speak("please mam hold the screen for few seconds, i am taking sreenshot")
                time.sleep(3)
                img = pyautogui.screenshot()
                img.save(f"{name}.png")
                speak("i am done mam, the screenshot is saved in our main folder. now i am ready for next command")
                continue
            elif "show me the screenshot" in self.query:
                try:
                    img = img.open('C://Users//Acer//Downloads//Jarvis Project' + name)
                    img.show(img)
                    speak("Here it is sir")
                    time.sleep(2)

                except IOError:
                    speak("Sorry sir, I am unable to display the screenshot")

            # speak("mam, do you have any other work")

            # -------------------  To Read PDF file -------------
            elif "read pdf" in self.query or "read pdf book" in self.query or "read pdf file" in self.query:
                try:
                    with open('The Book of Life_ A Novel.pdf', 'rb') as book:
                        pdf_reader = PyPDF2.PdfReader(book)
                        num_pages = len(pdf_reader.pages)
                        speak(f"Total numbers of pages in this book {num_pages} ")
                        speak("Mam please enter the page number i have to read")
                        start_page = int(input("Enter the starting page number: "))
                        for page_num in range(start_page, num_pages):
                            page = pdf_reader.pages[start_page]
                            text = page.extract_text()
                            speak(text)            
                            speak("Do you want to continue reading? ")
                            condition = self.takecommand()
                            speak(condition)
                            if 'yes' in condition:
                                start_page=start_page + 1                                
                            elif 'No' in condition:
                                speak("Thank you for reading book")
                                break
                except FileNotFoundError as e:
                        speak("Sorry, file not found")

            # --------------------- To Hide files and folder ---------------
            elif "hide all files" in self.query or "hide this folder" in self.query or "visible for everyone" in self.query:
                speak("Mam please tell me you want to hide this folder or make it visible for everyone")
                condition = self.takecommand()
                if "hide" in condition:
                    os.system("attrib +h /s /d")  # os module
                    speak("Mam, all the files in this folder are now hidden.")

                elif "visible" in condition:
                    os.system("attrib -h /s /d")
                    speak(
                        "Mam, all the files in this folder are now visible to everyone. i wish you are taking this decision in your own peace.")

                elif "leave it" in condition or "leave for now" in condition:
                    speak("Ok Mam")

            elif "temperature" in self.query or "what is the temprature" in self.query:
                api_key = "9c34a686e228221cab8cae501acf373d"
                base_url = "http://api.openweathermap.org/data/2.5/weather?"
                speak("Sure, please tell the city you want to know the temperature")
                city_name=self.takecommand()
                complete_url = base_url + "appid=" + api_key + "&q=" + city_name
                response = requests.get(complete_url)
                data = response.json()
                temperature = data['main']['temp']
                temperature = temperature - 273.15
                print("Temperature of", city_name, "is", round(temperature, 2), "Celsius.")
                text="Temperature of", city_name, "is", round(temperature, 2), "Celsius."
                speak(text)


startExecution = MainThread()


class Main(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_JarvisUi()
        self.ui.setupUi(self)
        self.ui.pushButton.clicked.connect(self.startTask)
        self.ui.pushButton_2.clicked.connect(self.close)

    def startTask(self):
        self.ui.movie = QtGui.QMovie("iron.gif.gif")
        self.ui.label.setMovie(self.ui.movie)
        self.ui.movie.start()
        self.ui.movie = QtGui.QMovie("T8bahf.gif")
        self.ui.label_2.setMovie(self.ui.movie)
        self.ui.movie.start()
        timer = QTimer(self)
        timer.timeout.connect(self.showTime)
        timer.start(1000)
        startExecution.start()

    def showTime(self):
        current_time = QTime.currentTime()
        current_date = QDate.currentDate()
        label_time = current_time.toString('hh:mm:ss')
        label_date = current_date.toString(Qt.ISODate)
        self.ui.textBrowser.setText(label_date)
        self.ui.textBrowser_2.setText(label_time)


# self.textBrowser.setText("Hello world")
#       self.textBrowser.setAlignment(QtCore.Qt.AlignCenter)

app = QApplication(sys.argv)
Jarvis = Main()
Jarvis.show()
sys.exit(app.exec_())
