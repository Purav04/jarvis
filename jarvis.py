import os
import sys
import csv
import datetime
# pip install SpeechRecognition
import speech_recognition as sr
from win32com.client import Dispatch
speak = Dispatch("SAPI.SpVoice")

def recognizer():#return
    try:
        r = sr.Recognizer()
        with sr.Microphone() as source:
            audio_data = r.record(source, offset=0,duration=4)
            # speak.Speak("Recognizing...")
            print("recognizing...")
            text = r.recognize_google(audio_data)
            print(text)
        return text
    except:return "None"

def process(name):#return
    try:
        import psutil
        for proc in psutil.process_iter():
                if proc.name()[:len(name)].lower() == name:
                    return True
        return False
    except:return False

def application(name):
    list_of_applications = ['Acer Jumpstart', 'Acer Portal', 'Acer Quick Access', 'Care Center',
                             'python', 'abFiles', 'abPhoto', 'pycharm',
                            'chrome', 'nox', 'notepad', "Internet Download Manager 6.35.17","idm",
                            "Snip & Sketch", "VLC media player"
                            "WinRAR 5.80 (64-bit)"]
    list_of_exe = ["C:\\Program Files (x86)\Acer\\Acer Jumpstart\\wall.exe","C:\\Program Files (x86)\\Acer\\Acer Portal\\AcerPortal.exe","C:\\Program Files\\Acer\\Acer Quick Access\\QuickAccess.exe",
                   "C:\\Program Files (x86)\\Acer\\Care Center\\CareCenter.exe","C:\\Python\\pythonw.exe","C:\\Program Files (x86)\\Acer\\abFiles\\abFilesTrayIcon.exe","C:\\Program Files (x86)\\Acer\\abPhoto\\abPhoto.exe",
                   "C:\\Program Files\\JetBrains\\PyCharm Community Edition 2019.3.1\\bin\\pycharm64.exe","C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe","C:\\Program Files (x86)\\Nox\\bin\\Nox.exe",
                   "C:\\Windows\\system32\\notepad.exe","C:\\Program Files (x86)\\Internet Download Manager\\IDMan.exe","C:\\Program Files (x86)\\Internet Download Manager\\IDMan.exe",
                   "C:\\Windows\\system32\\slui.exe","C:\\Program Files\\VideoLAN\VLC\\vlc.exe","C:\\Program Files\\WinRAR\\WinRAR.exe"]
    try :
        for i in list_of_applications:
            if name.lower() == i[:len(name)].lower():
                if process(name):
                    speak.Speak("process is already running")
                    speak.Speak("do you want to start it again(yes or no)")
                    run = recognizer()
                    if run.lower() == "y":#("yes" or "yep" or "y" or "yeah")
                        os.startfile(list_of_exe[list_of_applications.index(name)])
                else:
                    os.startfile(list_of_exe[list_of_applications.index(name)])
        else:
            speak.Speak("there is no application with this name....do you want to search on chrome(yes or no)")
            conti = recognizer()
            if conti.lower() == "yes":search_on_chrome(name)
            else:sys.exit()
    except:speak.Speak("errr in application")

def folder(name):
    list_of_mains = ['C:/', 'D:/', 'E:/', 'C:/Users/PAKHI/Desktop/']
    try:
        for i in list_of_mains:
            entity = os.listdir(i)
            if name in entity:
                name = i + name
        os.startfile(name)
    except:
        speak.Speak("there is no application with this name....do you want to search on chrome(yes or no)")
        conti = recognizer()
        if conti.lower() == "yes":
            search_on_chrome(name)
        else:
            sys.exit()

def search_on_chrome(name):
        try:
            import webbrowser
            chrome_path = 'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s'
            url = "https://www.google.com.tr/search?q={}".format(name)
            webbrowser.get(chrome_path).open(url)
            return
        except :speak.Speak("errr in search_on_chrome")

def kill_process(name):
    try:
        if name.lower() == "this program": sys.exit()
        list_of_applications = ['Acer Jumpstart', 'Acer Portal', 'Acer Quick Access', 'Care Center','python', 'abFiles', 'abPhoto', 'pycharm','chrome', 'nox', 'notepad',
                                "Internet Download Manager 6.35.17", "idm","Snip & Sketch", "VLC media player","WinRAR 5.80 (64-bit)"]
        list_of_exe = ["C:\\Program Files (x86)\Acer\\Acer Jumpstart\\wall.exe","C:\\Program Files (x86)\\Acer\\Acer Portal\\AcerPortal.exe",
                       "C:\\Program Files\\Acer\\Acer Quick Access\\QuickAccess.exe","C:\\Program Files (x86)\\Acer\\Care Center\\CareCenter.exe", "C:\\Python\\pythonw.exe",
                       "C:\\Program Files (x86)\\Acer\\abFiles\\abFilesTrayIcon.exe","C:\\Program Files (x86)\\Acer\\abPhoto\\abPhoto.exe",
                       "C:\\Program Files\\JetBrains\\PyCharm Community Edition 2019.3.1\\bin\\pycharm64.exe","C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe",
                       "C:\\Program Files (x86)\\Nox\\bin\\Nox.exe","C:\\Windows\\system32\\notepad.exe","C:\\Program Files (x86)\\Internet Download Manager\\IDMan.exe",
                       "C:\\Program Files (x86)\\Internet Download Manager\\IDMan.exe","C:\\Windows\\system32\\slui.exe", "C:\\Program Files\\VideoLAN\VLC\\vlc.exe",
                       "C:\\Program Files\\WinRAR\\WinRAR.exe"]
        if process(name):
            exe_file =  list_of_exe[list_of_applications.index(name)]
            exe = ""
            for i in range(len(exe_file) - 1, -1, -1):
                if exe_file[i] == '\\':
                    break
                exe += exe_file[i]
            exe = exe[::-1]

            os.system(f"TASKKILL /IM {exe}")
            print(f"successfully close application")
        else:
            speak.Speak("process already dead")
    except:print("errr in kill_process")

def output(name):
    try:
        name = name.lower()
        try:
            with open(r"C:\Users\PAKHI\Desktop\data_of_sec.csv", "a", newline='') as csvfile:
                writer = csv.writer(csvfile)
                x = datetime.datetime.now()
                writer.writerow([x.time(),x.date(),name])
                csvfile.close()
        except:pass

        if name=="close right now":sys.exit()
        elif name[:5].lower() == "close":kill_process(name[6:])
        elif name[:6].lower() == "search":search_on_chrome(name[7:])
        elif name[:11].lower() == "open folder":folder(name[12:])
        elif name[:5].lower() == "start":application(name[6:])
        else:
            speak.Speak("do you want to search in google(yes or no)")
            conti = recognizer()
            if conti.lower() == "yes":search_on_chrome(name)
            # else:
            #     speak.Speak("\nif you try another time then start with this keys:\nclose for close the application\nsearch for search on google\n"
            #        "open folder for open the folder\nstart for start application\n")
        speak.Speak("thanks for using")
        speak.Speak("do you want me to do something (yes or no)?")
        conti = recognizer()
        if conti.lower() == "yes":
            speak.Speak("what can i do for you?")
            open__ = recognizer()
            output(open__) # input("enter Folder/Application name::")
    except:
        print("error:")

if __name__ == "__main__":
    speak.Speak("hey purav...what can i do for you?")#
    open = recognizer()
    output(open)
