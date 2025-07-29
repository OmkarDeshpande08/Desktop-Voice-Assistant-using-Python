from JarvisAI import JarvisAssistant
import re
import os
import webbrowser
import speedtest
import datetime
import requests
import sys
import socket 
import pyjokes
import time
import pyautogui
import pywhatkit
import wolframalpha
import psutil
import cv2
import subprocess
import platform
from PIL import Image
from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtCore import QTimer, QTime, QDate, Qt
from PyQt5.QtGui import QMovie
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.uic import loadUiType
from Jarvis.features.gui import Ui_MainWindow
from Jarvis.config import config

obj = JarvisAssistant()




def speak(text):
    obj.tts(text)


app_id = config.wolframalpha_id




def search_and_open_folder(folder_name, base_paths):
    """Search and open folder with cross-platform compatibility"""
    for base_path in base_paths:
        try:
            for root, dirs, files in os.walk(base_path):
                for dir_name in dirs:
                    if folder_name.lower() in dir_name.lower():
                        folder_path = os.path.join(root, dir_name)
                        
                        # Cross-platform folder opening
                        if platform.system() == "Windows":
                            os.startfile(folder_path)
                        elif platform.system() == "Darwin":  # macOS
                            subprocess.run(["open", folder_path], check=True)
                        else:  # Linux
                            subprocess.run(["xdg-open", folder_path], check=True)
                        
                        speak(f"Opening {dir_name} folder.")
                        return True
        except Exception as e:
            print(f"Error searching in {base_path}: {e}")
            continue
    return False

def get_ip_address():
    try:
        hostname = socket.gethostname()
        ip_address = socket.gethostbyname(hostname)
        return ip_address
    except Exception as e:
        print(f"Error fetching IP address: {e}")
        return "unknown"


def open_office_application(application):
    """Open office applications with cross-platform support"""
    try:
        app_lower = application.lower()
        
        if platform.system() == "Windows":
            # Windows - Microsoft Office
            office_path = "C:\\Program Files\\Microsoft Office\\root\\Office16\\"
            if app_lower == 'word':
                subprocess.run([os.path.join(office_path, 'WINWORD.EXE')], check=True)
                speak("Opening Microsoft Word.")
            elif app_lower == 'powerpoint':
                subprocess.run([os.path.join(office_path, 'POWERPNT.EXE')], check=True)
                speak("Opening Microsoft PowerPoint.")
            elif app_lower == 'excel':
                subprocess.run([os.path.join(office_path, 'EXCEL.EXE')], check=True)
                speak("Opening Microsoft Excel.")
        else:
            # Linux/macOS - LibreOffice
            if app_lower == 'word':
                subprocess.run(['libreoffice', '--writer'], check=True)
                speak("Opening LibreOffice Writer.")
            elif app_lower == 'powerpoint':
                subprocess.run(['libreoffice', '--impress'], check=True)
                speak("Opening LibreOffice Impress.")
            elif app_lower == 'excel':
                subprocess.run(['libreoffice', '--calc'], check=True)
                speak("Opening LibreOffice Calc.")
        
    except subprocess.CalledProcessError:
        speak(f"Sorry, I couldn't open {application}. Please make sure it's installed.")
    except FileNotFoundError:
        speak(f"Sorry, {application} is not installed on this system.")
    except Exception as e:
        speak(f"Sorry, I encountered an error opening {application}.")
        print(f"Error opening {application}: {e}")
        
def get_battery_percentage():
    battery = psutil.sensors_battery()
    percent = battery.percent if battery else None
    return percent
    

    

def wish():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<=12:
        speak("Good Morning")
    elif hour>12 and hour<18:
        speak("Good afternoon")
    else:
        speak("Good evening")
    #c_time = obj.tell_time()
    #speak(f"Currently it is {c_time}")
    speak("I am your pc assistant. Online and ready sir. Please tell me how may I help you")
# if __name__ == "__main__":


class MainThread(QThread):
    def __init__(self):
        super(MainThread, self).__init__()

    def run(self):
        self.TaskExecution()

    def TaskExecution(self):
        #startup()
        wish()

        while True:
            command = obj.mic_input()

            if re.search('date', command):
                date = obj.tell_me_date()
                print(date)
                speak(date)
                
            elif 'time' in command:
                current_time = datetime.datetime.now().strftime("%I:%M %p")

                speak(f"Sir, the time is {current_time}")
                print(f"Sir, the time is {current_time}")
            
            elif 'day ' in command:
                x = datetime.datetime.now()
                d=x.strftime("%A")
                speak(f"sir,today is{d}")
                
                
            elif "search on chrome" in command or "search on google" in command:
                try:
                    speak("What should I search?")
                    print("What should I search?")
                    search_query = obj.mic_input()
                    speak(f"sure sir! here are the results for {search_query}")
                    webbrowser.open(f"https://www.google.com/search?q={search_query}")
                    print(search_query)
                
                except Exception as e:
                    speak("Can't open now, please try again later.")
                    print("Can't open now, please try again later.")

            elif re.search('weather', command):
                city = command.split(' ')[-1]
                weather_res = obj.weather(city=city)
                print(weather_res)
                speak(weather_res)
                
            elif 'battery' in command:
                battery_percent = get_battery_percentage()
                if battery_percent is not None:
                    speak(f"Sir, your battery percentage is {battery_percent} percent.")
                    print(f"Sir, your battery percentage is {battery_percent} percent.")
                else:
                    speak("Sorry, I couldn't retrieve the battery information at the moment.")
                    
                    
            elif "open notepad" in command:
                try:
                    speak("sure sir! opening text editor")
                    if platform.system() == "Windows":
                        subprocess.run(["notepad.exe"], check=True)
                    else:
                        # Try common Linux text editors
                        try:
                            subprocess.run(["gedit"], check=True)
                        except FileNotFoundError:
                            try:
                                subprocess.run(["kate"], check=True)
                            except FileNotFoundError:
                                subprocess.run(["nano"], check=True)
                except Exception as e:
                    speak("Sorry, I couldn't open the text editor.")
                    print(f"Error opening text editor: {e}")
            
            elif "close notepad" in command:
                try:
                    speak("closing text editor")
                    if platform.system() == "Windows":
                        subprocess.run(["taskkill", "/im", "notepad.exe"], check=True)
                    else:
                        # For Linux, we'll use pkill
                        subprocess.run(["pkill", "-f", "gedit|kate|nano"], check=True)
                except Exception as e:
                    speak("Sorry, I couldn't close the text editor.")
                    print(f"Error closing text editor: {e}")
                
            elif "resume" in command:
                try:
                    speak("sure sir! here is your resume")
                    # Use cross-platform path
                    resume_path = config.resume_path
                    if os.path.exists(resume_path):
                        if platform.system() == "Windows":
                            os.startfile(resume_path)
                        elif platform.system() == "Darwin":
                            subprocess.run(["open", resume_path], check=True)
                        else:
                            subprocess.run(["xdg-open", resume_path], check=True)
                    else:
                        speak("Sorry, I couldn't find your resume file.")
                except Exception as e:
                    speak("Sorry, I couldn't open your resume.")
                    print(f"Error opening resume: {e}")
                
            elif 'email' in command:
                speak('sure sir! opening gmail')
                webbrowser.open("www.gmail.com")
                
            elif "college" in command:
                speak("Sure sir! here is your college website")
                webbrowser.open("www.git.edu")
                
            elif "increase" in command:
                pyautogui.press("volumeup")
                pyautogui.press("volumeup")
                pyautogui.press("volumeup")
                pyautogui.press("volumeup")
                pyautogui.press("volumeup")
                pyautogui.press("volumeup")
                speak("Is this ok sir?")

            elif "decrease" in command:
                pyautogui.press("volumedown")
                pyautogui.press("volumedown")
                pyautogui.press("volumedown")
                pyautogui.press("volumedown")
                pyautogui.press("volumedown")
                speak("Is this ok sir?")
            
            elif "camera" in command:
                try:
                    speak("Opening camera. Press ESC to close.")
                    cap = cv2.VideoCapture(0)
                    
                    if not cap.isOpened():
                        speak("Sorry, I couldn't access the camera.")
                        continue
                    
                    while True:
                        ret, img = cap.read()
                        if not ret:
                            speak("Camera feed lost.")
                            break
                            
                        cv2.imshow('Jarvis Webcam - Press ESC to close', img)
                        k = cv2.waitKey(50)
                        if k == 27:  # ESC key
                            break
                    
                    cap.release()
                    cv2.destroyAllWindows()
                    speak("Camera closed.")
                    
                except Exception as e:
                    speak("Sorry, I encountered an error with the camera.")
                    print(f"Camera error: {e}")
                    try:
                        cap.release()
                        cv2.destroyAllWindows()
                    except:
                        pass
                

            elif "open command prompt" in command or "open terminal" in command:
                try:
                    speak("sure sir! opening terminal")
                    if platform.system() == "Windows":
                        subprocess.run(["cmd"], check=True)
                    else:
                        # Try common Linux terminals
                        try:
                            subprocess.run(["gnome-terminal"], check=True)
                        except FileNotFoundError:
                            try:
                                subprocess.run(["konsole"], check=True)
                            except FileNotFoundError:
                                subprocess.run(["xterm"], check=True)
                except Exception as e:
                    speak("Sorry, I couldn't open the terminal.")
                    print(f"Error opening terminal: {e}")
                
            elif "close command prompt" in command or "close terminal" in command:
                try:
                    speak("closing terminal")
                    if platform.system() == "Windows":
                        subprocess.run(["taskkill", "/im", "cmd.exe"], check=True)
                    else:
                        subprocess.run(["pkill", "-f", "gnome-terminal|konsole|xterm"], check=True)
                except Exception as e:
                    speak("Sorry, I couldn't close the terminal.")
                    print(f"Error closing terminal: {e}")
                
            elif "internet speed" in command:
                speak("Sure sir! let me check please wait")
                wifi = speedtest.Speedtest()
                upload_net = wifi.upload() / 1048576
                download_net = wifi.download() / 1048576
                upload_net = round(upload_net, 2)
                download_net = round(download_net, 2)

                print("Wifi Upload Speed is", upload_net)
                print("Wifi download speed is ", download_net)

                
                speak(f"Wifi download speed is {download_net} gb per second")
                speak(f"Wifi Upload speed is {upload_net} gb per second")

            elif "open calculator" in command:
                try:
                    speak("sure sir! opening calculator")
                    if platform.system() == "Windows":
                        subprocess.run(["calc.exe"], check=True)
                    else:
                        # Try common Linux calculators
                        try:
                            subprocess.run(["gnome-calculator"], check=True)
                        except FileNotFoundError:
                            try:
                                subprocess.run(["kcalc"], check=True)
                            except FileNotFoundError:
                                subprocess.run(["xcalc"], check=True)
                except Exception as e:
                    speak("Sorry, I couldn't open the calculator.")
                    print(f"Error opening calculator: {e}")

            elif "close calculator" in command:
                try:
                    speak("closing calculator")
                    if platform.system() == "Windows":
                        subprocess.run(["taskkill", "/f", "/im", "calc.exe"], check=True)
                    else:
                        subprocess.run(["pkill", "-f", "gnome-calculator|kcalc|xcalc"], check=True)
                except Exception as e:
                    speak("Sorry, I couldn't close the calculator.")
                    print(f"Error closing calculator: {e}")
                
            
            elif 'open folder' in command:
                speak("Sure sir! What's the name of the folder?")
                folder_query = obj.mic_input().lower()
                
                if folder_query != 'none':
                    # Use cross-platform search paths from config
                    base_paths_to_search = config.folder_search_paths
                    if not search_and_open_folder(folder_query, base_paths_to_search):
                        speak(f"Sorry, I couldn't find any folder with the name {folder_query}.")
                else:
                    speak("Sorry, I didn't recognize the folder name.")
            

            elif 'youtube' in command:
                video = command.split(' ')[1]
                speak(f"Okay sir, playing {video} on youtube")
                pywhatkit.playonyt(video)
                
            if 'word' in command:
                open_office_application('word')

            elif 'excel' in command:
                open_office_application('excel')

            elif 'powerpoint' in command:
                open_office_application('powerpoint')

            if "joke" in command:
                joke = pyjokes.get_joke()
                print(joke)
                speak(joke)

            elif "system" in command:
                sys_info = obj.system_info()
                print(sys_info)
                speak(sys_info)

            
            elif "ip address" in command:
                ip_address = get_ip_address()
                print(f"Your IP address is {ip_address}")
                speak(f"Your IP address is {ip_address}")

            elif "switch the window" in command or "switch window" in command:
                speak("Okay sir, Switching the window")
                pyautogui.keyDown("alt")
                pyautogui.press("tab")
                time.sleep(1)
                pyautogui.keyUp("alt")


            elif "take screenshot" in command or "take a screenshot" in command or "capture the screen" in command:
                try:
                    speak("By what name do you want to save the screenshot?")
                    name = obj.mic_input()
                    if name != 'none':
                        speak("Alright sir, taking the screenshot")
                        
                        # Create screenshots directory if it doesn't exist
                        os.makedirs("screenshots", exist_ok=True)
                        
                        img = pyautogui.screenshot()
                        screenshot_name = f"screenshots/{name}.png"
                        img.save(screenshot_name)
                        speak("The screenshot has been successfully captured")
                        
                        # Store the filename for later use
                        global last_screenshot
                        last_screenshot = screenshot_name
                    else:
                        speak("Sorry, I didn't get the name for the screenshot.")
                except Exception as e:
                    speak("Sorry, I couldn't take the screenshot.")
                    print(f"Screenshot error: {e}")

            elif "show me the screenshot" in command:
                try:
                    if 'last_screenshot' in globals() and os.path.exists(last_screenshot):
                        if platform.system() == "Windows":
                            os.startfile(last_screenshot)
                        elif platform.system() == "Darwin":
                            subprocess.run(["open", last_screenshot], check=True)
                        else:
                            subprocess.run(["xdg-open", last_screenshot], check=True)
                        speak("Here it is sir")
                    else:
                        speak("Sorry sir, I couldn't find any recent screenshot")
                except Exception as e:
                    speak("Sorry sir, I am unable to display the screenshot")
                    print(f"Error showing screenshot: {e}")

            

            elif "goodbye" in command or "offline" in command or "bye" in command:
                speak("Alright sir, going offline. It was nice working with you")
                sys.exit()


startExecution = MainThread()


class Main(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.pushButton.clicked.connect(self.startTask)
        self.ui.pushButton_2.clicked.connect(self.close)

    def __del__(self):
        sys.stdout = sys.__stdout__

    # def run(self):
    #     self.TaskExection
    def startTask(self):
        self.ui.movie = QtGui.QMovie("Jarvis/utils/images/live_wallpaper.gif")
        self.ui.label.setMovie(self.ui.movie)
        self.ui.movie.start()
        self.ui.movie = QtGui.QMovie("Jarvis/utils/images/initiating.gif")
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


app = QApplication(sys.argv)
jarvis = Main()
jarvis.show()
exit(app.exec_())