from PyQt5.QtCore import QDate, QTime, QDateTime, Qt
import pyautogui
import subprocess,shutil
from playsound import playsound
import pyttsx3
import time,os
import selenium
from selenium import webdriver
from plyer import notification
import datetime
import requests
from win32com.client import Dispatch
import urllib3
from urllib3 import PoolManager
import bs4
import speech_recognition as sr    
import webbrowser,wikipedia    
import winsound

      
def listen():
  try:
    count =0
    microphone=sr.Recognizer()
    with sr.Microphone() as source:
        print("Vivek ,Please repeate the action you want to perform")
        audio=microphone.listen(source)
        print("User Said" +microphone.recognize_google(audio))
    
    userSaid=microphone.recognize_google(audio)
    return userSaid
   
  except sr.UnknownValueError:
    count=count+1  
    print("count ki value itni hai",count)
    alternative_speak("please please could you repeate again vivek")
    
    
    if count==3:
     alternative_speak("I think you are busy Let me Help you sir")   
     chooseDefaultAction()
    
  except sr.RequestError:
    alternative_speak("Sorry Internet connection is not presend..Switching to manual Mode")
    checkNetConnectionAndThenPlaySong()    
    
#   except sr.requestError:      
#     alternative_speak("I think NET Connection is not proper")  

    


def alternative_speak(str):
    engine=pyttsx3.init()
    voices=engine.getProperty('voices')
    engine.setProperty('voice',voices[1].id)
    engine.setProperty('rate',150)
    engine.say(str)
    engine.runAndWait()




def chooseDefaultAction():
    alternative_speak("No Voice Response Received")
    alternative_speak("Press 1 for play songs and 2 for play news")
    choice = input(int("Enter 1.play songs and 2 for play news"))
    
    if choice==1:
        checkNetConnectionAndThenPlaySong()
    
    elif choice==2:
        playNews()     
    
    


#Testring google Search using Voice 

def googleSerchUsingVoice():
    alternative_speak("Vivek Could you Please Specify what do you want from google")
    
    audio=listen()
    
    convertedToAudio=str(audio)
    
    query=f"search{audio}"
    
    print(query)
    
    
    alternative_speak("Here's what i found on Google Sir")  
    webbrowser.open_new_tab(f"https://www.google.com/search?q={convertedToAudio}")
    
    
    

#Play songs using voice 
def searchSongandPlaySongUsingVoice():
   
        driver=webdriver.Chrome(r"C:\Windows\System32\chromedriver_win32\chromedriver.exe")
        
        alternative_speak("Which Song May Please You Sir?")
        #songToPlay=input("Enter Which Song to Please You Sir").strip()
    
        
        try: 
         audio=listen()
         query=str(audio) 
         
        except sr.exceptions:
         checkNetConnectionAndThenPlaySong()   
                
        print(query)
    
    
        if(query):
        
         alternative_speak(f"{query} is coming up next for you...Enjoy sir")    
        
         driver.get(f"https://www.youtube.com/results?search_query={query}")

    
    
            # driver.implicitly_wait(5)
            
            # searchButton =driver.find_element_by_class_name('gNO89b')
            # searchButton.click()
         driver.implicitly_wait(5)
         videoToPlay=driver.find_element_by_xpath('//*[@id="video-title"]/yt-formatted-string')
            
         videoToPlay.click()
            
        else:
            alternative_speak("Sorry.Unable to hear what you say vivek sir")
            
        
        input()
        

#Search On Wikipedia
def serachonWikipedUsingVoice():
    alternative_speak("Vivek Could you Please Specify what to search on Wikipedia")
    
    audio=listen()
    
    convertedToAudio=str(audio)
    
    query=f"search on wikipedia {audio}"
    
    
    
    information=wikipedia.summary(title=f"{convertedToAudio}",sentences=10) 
    alternative_speak(information) 
    
#   alternative_speak("sorry.Can't reach to wikipedia at this moment vivek") 



#Starting Point
def GreetVivekSir():
    alternative_speak("Hello vivek sir")
    
    currentTime = datetime.datetime.now()
    Current_hour=currentTime.hour
    
    # date=datetime.datetime.today().strftime("%x")
    # month=datetime.datetime.today().strftime("%B")
    # dayOfWeek=datetime.datetime.today().strftime("%A")
    
    
    now=QDate.currentDate()
    
    
    
    if Current_hour < 12:
        
        
        
        
        alternative_speak("Good Morning vivek sir")
        alternative_speak("How are you vivek")
        alternative_speak("I hope you are doing well")
        alternative_speak(f"Today's Date and Day is {now.toString(Qt.DefaultLocaleLongDate)}")
        alternative_speak(f"Current Time is {datetime.datetime.now().hour} hour and {datetime.datetime.now().minute} minute {datetime.datetime.now().strftime('%p')}"  )
        
     
        #notification.notify("Good Morning Vivek sir","jarvis_icon.ico","Good Morning Vivek sir",timeout=10)    
        notification.notify(
            title="Hello sir",
            message="Good Morning Vivek sir",
            #app_icon="jarvis_icon.ico",
            timeout=10
            
            )
        
        
        
    
    elif  12<= Current_hour <18:
        alternative_speak("Good Afternoon Vivek Sir")
       # alternative_speak(f"Today's Date is{datetime.date.today()}")
       
        alternative_speak(" I hope It is nice time to work on laptop.")
        
       
        alternative_speak(f"Today's Date and Day is {now.toString(Qt.DefaultLocaleLongDate)}")
         
        alternative_speak(f"Current Time is {datetime.datetime.now().hour} hour and {datetime.datetime.now().minute} minute"  )
     
        notification.notify(
            title="Hello sir",
            message="Good Afternoon Vivek sir",
            #app_icon="jarvis_icon.ico",
            timeout=10
            
            )
         
    else:
        alternative_speak("Good Evening Vivek Sir")
        alternative_speak("I hope it is wonderful evening for you")
        alternative_speak(f"Current Time is {datetime.datetime.now().hour} hour and {datetime.datetime.now().minute} minute"  )
     
        notification.notify(
            title="Hello sir",
            message="Good Evening Vivek sir",
           # app_icon=r"F:/pythonPractice/jarvis_icon.ico",
            timeout=10
                        
            )
        

def thisIsaTestforJokes():
        data=requests.get("http://www.laughfactory.com/jokes/sex-jokes")
        soup=bs4.BeautifulSoup(data.content,'html.parser');
        for data in soup.findAll('div',{"class": "joke-text"}):
            print(data.text)
            jokes=data.text
            alternative_speak(jokes)
           
            





        

def tellMeAjoke():
    try:
        data=requests.get("http://www.laughfactory.com/jokes/clean-jokes")
        soup=bs4.BeautifulSoup(data.content,'html.parser');
        for data in soup.findAll('div',{"class": "joke-text"}):
            print(data.text)
            jokes=data.text
            alternative_speak(jokes)
            alternative_speak("Do you want to listen next joke vivek,press yes to continue or no to skip this functionality")
            choice=input("Do you want to listen next joke sir").strip()
            
            if choice=="no":
                break
            else:
                continue
            
    except requests.exceptions.ConnectionError as err:
      alternative_speak("Sorry !No internet connection is presend")  
      tryToStartInterentConnection()  

        
            
    
        
#Normal SearchSong Feature        
def searchSong():
   
        driver=webdriver.Chrome(r"C:\Windows\System32\chromedriver_win32\chromedriver.exe")
        
        alternative_speak("Which Song May Please You Sir?")
        songToPlay=input("Enter Which Song to Please You Sir").strip()
        
        alternative_speak(f"{songToPlay} is coming up for you")
        
        driver.get(f"https://www.youtube.com/results?search_query={songToPlay}")

        # driver.implicitly_wait(5)
        
        # searchButton =driver.find_element_by_class_name('gNO89b')
        # searchButton.click()
        
        driver.implicitly_wait(5)
        videoToPlay=driver.find_element_by_xpath('//*[@id="video-title"]/yt-formatted-string')
        
        videoToPlay.click()
        
    
        
        input()
        
#Getting Weather
def getWeather():
     
     try:
      data=requests.get("https://www.worldweatheronline.com/lang/en-in/bhusawal-weather/maharashtra/in.aspx")   
      soup=bs4.BeautifulSoup(data.content,'html.parser')
    
      for i in soup.select('.display-4'):
         print(i.text)
         data=i.text
  
    
      alternative_speak(f"Today's weather in Bhusaawal is {data}")
    
     except requests.exceptions.ConnectionError as err:
      alternative_speak("Sorry !No internet connection is presend")  
      tryToStartInterentConnection()  


#NetConnection Checking
def checkNetConnectionAndThenPlaySong():
    try:
         Manager=PoolManager(10)
         res=Manager.request('GET','https://www.google.com')
         if res:
            alternative_speak("Please Type the song name as you have not responsded")
            searchSong()
         else:
             alternative_speak("Sorry!No Internet is present")        
     
    except urllib3.exceptions.MaxRetryError as err:
         alternative_speak("Internet is not present")   
         tryToStartInterentConnection()
         exit()

#Opening Code Feature     
def openCode():
    
    alternative_speak("Opening Code for You!!!Enjoy")
    subprocess.Popen("C://Program Files//Microsoft VS Code//Code.exe")
    
    

def playNews():
   try: 
    driver=webdriver.Chrome(r"C:\Windows\System32\chromedriver_win32\chromedriver.exe")
        
    NewsToSearch="news headlines today india"
    driver.get(f"https://www.youtube.com/results?search_query={NewsToSearch}")
    driver.implicitly_wait(5)  
    videoToClick=driver.find_element_by_xpath('//*[@id="video-title"]/yt-formatted-string').click()
    
    
    
   except requests.exceptions.ConnectionError as err:
       alternative_speak("Net connection is not presend")
       tryToStartInterentConnection()
   
   
   input()
   
   
def deletingAlltheTrash():
    alternative_speak("Cleaning All the Temporary files for you Sir")
    shutil.rmtree(r"C:/Users/vivek/AppData/Local/Temp/",ignore_errors=True)   
    


def tryToStartInterentConnection():
    alternative_speak("Trying to switch on your internet connection sir")
    # while True: 
    #   postion=pyautogui.position()    
    #   print(postion)
     
 #    this will click to Hidden Icon   
    pyautogui.click(1709,1048)
    
    time.sleep(1)    

#     #this will click on internet browser icon
    pyautogui.click(1656,953)

    time.sleep(1)
    
#     #this will click goto control panel internet option
    pyautogui.click(1630,430)

    time.sleep(2)
#     #this will bring connect button  
    pyautogui.click(631,180)  

    time.sleep(1)
#    #to click on connect button click
    pyautogui.click(597,253)



def shutdownSystem():
    alternative_speak("Please wait..I need to check the Identity and Authority of user")
    time.sleep(1)
    alternative_speak("See u soon sir")
    os.system("shutdown -s -t 0")
    exit()




def fetchMotivationalQuoteDaily():
    try:
    
     data=requests.get("https://www.briantracy.com/blog/personal-success/26-motivational-quotes-for-success/")
     soup=bs4.BeautifulSoup(data.content,'html.parser')
    
     for i in soup.findAll('h3'):
        print(i.text)
        Quotes=i.text
        
    
          
    except requests.exceptions.ConnectionError as err:
      alternative_speak("Sorry !No internet connection is presend")
        
      tryToStartInterentConnection()  
        
if __name__ == "__main__":
    
    
    
    
       
    
    urllib3.disable_warnings();
    
    
    
    GreetVivekSir()
    
    alternative_speak("would you like to listen some jokes sir,Let me know your choice sir")
    
    
    
    choice= input("would you like to listen jokes").strip()
    
    if choice == "yes":
        alternative_speak("Your smile means lot to me")
        alternative_speak("Let me tell you a joke")
        tellMeAjoke()
    
    
    
    #deletingAlltheTrash()
    #getWeather()
    
    
    
    alternative_speak("Please Give me once chance to serve you sir")

    
    # choice=input("Would You like to open code sir,Enter yes to continue").strip()
    
    # if choice == "yes":
    #   alternative_speak("Opening Code for you sir")
    #   openCode()     
     
    # else:
    #    alternative_speak("Maybe we should listen some songs then")  
    
    
   
    
    # alternative_speak("Would Like to play News Sir")
      
      
    # choice=input("Would You Play News,Enter yes to continue").strip()
    # if choice == "yes":
    #   alternative_speak("Playing News for you")
    #   playNews()     
     
    # else:
    #    alternative_speak("Maybe we should listen some songs then")  
    
        
    audio=listen()
    print(audio)
    convertedToAudio=str(audio)
    
    if audio==None:
      audio=listen() 
    
 
    if 'song' in audio:
     searchSongandPlaySongUsingVoice()
    
    elif 'Wikipedia' in audio:
      print(True)
      serachonWikipedUsingVoice()
    
    elif "Google" in audio:
      googleSerchUsingVoice()
    
    elif "Open Code" in audio:
       openCode() 
    
    elif 'play news' in audio:
       playNews()
       
    elif 'Amazon' in audio:
        alternative_speak("Happy Shopping Sir")
        webbrowser.open(f"www.amazon.in/s?k={str(audio)}")
    
    elif 'shutdown' in audio:
        shutdownSystem()        
    
    elif 'activate advanced system' in audio:
        alternative_speak("Verifying Identity")
        time.sleep(2)
        
        alternative_speak("Identity Accepted")
        alternative_speak("User Vivek ")
        alternative_speak("admin level detected ...Granting full access to system")    
        alternative_speak("Running Diagnostics")
        alternative_speak("Getting Resources")
        time.sleep(1)
        alternative_speak("Activating Advanced Shell sir")
        os.system(r"C:\Users\vivek\AppData\Local\Programs\edex-ui\eDEX-UI.exe")
        
    
    
    else:
        
        alternative_speak("functionality is not implemented yet")
        alternative_speak("anything else sir")
        audio=listen()
        
        if None in audio:
            alternative_speak("No response has been specified")
            
        
        elif "yes" in audio:
            audio=listen()
        
        else:
            alternative_speak("Good Bye sir") 
         
    
    # #checkNetConnectionAndThenPlaySong()
    
     
     
    #  openCode()
     
     
         
         
     

     
    #  alternative_speak("would you like to play songs,Enter YES to play songs and NO to quit")
       
     
    #  choice=input("Would You like to listen songs,Enter yes to continue").strip()
    #  if choice == "yes":
    #      checkNetConnectionAndThenPlaySong()      
     
    #  else:
    #      alternative_speak("Have a nice day sir")       
    #      alternative_speak("bye bye vivek sir")   


         