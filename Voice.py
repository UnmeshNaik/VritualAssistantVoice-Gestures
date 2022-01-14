import speech_recognition as sr
import pyttsx3
import random
import time
import subprocess
import pyautogui

t = time.localtime()
current_time = time.strftime("%H:%M:%S", t)
hour = t.tm_hour

# initializing input and outputs
engine = pyttsx3.init()
listener = sr.Recognizer()
# voices = engine.getProperty('voices')

Labels = {'verb': '', 'target': ''}
once = 1
fd = 0


def speak(content):
    """for speaking the strings passed as arguments"""
    engine.say(content)
    engine.runAndWait()


def greet():
    """Greets user according to current time"""
    if 00 < hour <= 12:
        greetings = ['Hello sir', 'At your service sir', 'Good Morning Sir']
        speak(str(random.choice(greetings))+",it is "+str(t.tm_hour)+" "+str(t.tm_min)+" In the morning,sir")
        # timing adding
    elif 12 < hour <= 6:
        greetings = ['Hello sir', 'At your service sir',
                      'Good Afternoon Sir']
        #speak(str(random.choice(greetings)))
        speak(str(random.choice(greetings))+",it is "+str(t.tm_hour)+" "+str(t.tm_min)+" In the afternoon,sir")
    else:
        greetings = ['Hello sir', 'At your service sir', 'Good Evening Sir']
        #speak(str(random.choice(greetings)))
        speak(str(random.choice(greetings))+",it is "+str(t.tm_hour)+" "+str(t.tm_min)+" In the evening,sir")


def respond():
    reply = ['sir?', 'Yes sir!']
    speak(str(random.choice(reply)))


def remove_stopwords(command):
    """This function finds and removes the stopwords from command"""
    stopwords = ['an', 'the', 'ok', 'hey', 'baymax', 'hello']
    for i in command:
        if i in stopwords:
            command.remove(i)
    # print(command)


def labelize(command):
    """This function is used for tokenizing the command,
    removing stopwords and labeling the command for further processing"""

    verbs = ['open', 'stop', 'close','start']

    command_tokens = (command.lower()).split()
    remove_stopwords(command_tokens)

    for i in command_tokens:
        if i in verbs:
            Labels['verb'] = i
            command_tokens.remove(i)

    Labels['target'] = ' '.join(command_tokens)

def start_typing():
    
    """This function is used for typing direct text in files"""
    
    speak("speak 'Stop typing' to stop sir")
    prev = ""
    while True:
        try:
            with sr.Microphone() as source:
                print("running")
                voice = listener.listen(source, timeout=0.5,
                                        phrase_time_limit=1.5)  # listens from the source and creates audio
                command = listener.recognize_google(voice)
                command = command.lower()
                print(command)
                if 'stop typing' in command:
                    speak("typing stopped")
                    break
                elif 'full stop' in command:
                    command = '.'
                    pyautogui.write(". ")
                elif 'coma' in command:
                    command = ','
                    pyautogui.write(", ")
                elif 'back' in command or 'backspace' in command:
                    for i in range(len(prev)+1):
                        pyautogui.press('backspace')
                elif 'enter' == command or 'search' == command:
                    pyautogui.press("enter")
                else:
                    pyautogui.write(command + " ")
                prev = command
        except:
            pass


def take_input():
    """Takes user input and process commands"""
    speak("Yes sir?")
    while True:
        try:
            with sr.Microphone() as source:
                # listening and getting input in string format
                print("Listening:")
                voice = listener.listen(source, timeout=0.5, phrase_time_limit=2)  # listens from the source and creates audio
                command1 = listener.recognize_google(voice)  # returns text format of audio
                print(command1)
                labelize(command1)

                # respective system calls
                if Labels['verb'] == 'open':
                    # command = command.replace('open', '')
                    print("In open")
                    if ('word' or 'microsoft word' or 'word 2016') in Labels['target']:
                        speak('Opening' + Labels['target'] + ',sir')
                        word = subprocess.Popen(r'"C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE" /w')
                        time.sleep(3)
                        start_typing()

                    elif ('powerpoint' or 'presentation') in Labels['target']:
                        speak('Opening' + Labels['target'] + ',sir')
                        ppt = subprocess.Popen(r'C:\Program Files (x86)\Microsoft Office\Office16\POWERPNT.EXE /w')

                    elif ('chrome' or 'browser') in Labels['target']:
                        speak('opening ' + Labels['target'] + ',sir')
                        chrome = subprocess.Popen('"C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe"')
                        print("Opening incognito")
                        start_typing()

                    elif 'youtube' in Labels['target']:
                        speak('opening ' + Labels['target'] + ',sir')
                        yt = subprocess.Popen(
                            '"C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe" youtube.com')

                    elif 'wikipedia' in Labels['target']:
                        speak('opening ' + Labels['target'] + ',sir')
                        yt = subprocess.Popen(
                            '"C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe" -incognito wikipedia.com')
                        start_typing()

                    elif 'notepad' in Labels['target']:
                        speak('opening ' + Labels['target'] + ',sir')
                        notepad = subprocess.Popen(r'"C:\Windows\System32\notepad.exe"')
                        time.sleep(2)
                        start_typing()

                    else:
                        global app,once
                        with open(r"C:\Users\iumah\PycharmProjects\CWPProject\venv\Paths.txt") as file:
                            print("Searching in files")
                            for line in file:
                                sep_index = line.find('-')
                                if Labels['target'] in line[:sep_index]:
                                    open_app = line.split('-')[1]
                                    print(r"{}".format(open_app[-1]))
                                    app = subprocess.Popen('"{}"'.format(open_app.replace("\n", "")))
                                    break
                            else:
                                print("Not found in files")
                                file.close()
                                if once == 1:
                                    speak("I beg your pardon sir")
                                    once = 0
                                else:
                                    speak("I can't find the application sir. Do you want to add the application? Type 1 if you want to add the application")
                                    if int(input("1.Yes\t2.No\n")) == 1:
                                        path = input("Application path: ")
                                        with open(r"C:\Users\iumah\PycharmProjects\CWPProject\venv\Paths.txt", 'a') as file:
                                            file.write(Labels['target'] + '-' + r"{}".format(path) + '\n')
                                            speak('Opening' + Labels['target'] + ',sir')
                                            app = subprocess.Popen(r'"{}"'.format(path))
                                    once = 1
                        file.close()

                elif Labels['verb'] == 'start':
                    if Labels['target'] == 'typing':
                        start_typing()

                elif Labels['verb'] == "close":
                    if ('word' or 'microsoft word' or 'word 2016') in Labels['target']:
                        speak('Closing' + Labels['target'] + ',sir')
                        word.terminate()

                    elif ('powerpoint' or 'presentation') in Labels['target']:
                        speak('Closing' + Labels['target'] + ',sir')
                        ppt.terminate()

                    elif ('chrome' or 'browser') in Labels['target']:
                        print("Closing chrome")
                        speak('Closing ' + Labels['target'] + ',sir')
                        chrome.terminate()

                    elif 'notepad' in Labels['target']:
                        speak('Closing ' + Labels['target'] + 'sir')
                        notepad.terminate()

                    elif 'youtube' in Labels['target']:
                        print('Closing ' + Labels['target'] + ' sir')
                        speak('Closing ' + Labels['target'] + ' sir')
                        yt.terminate()

                    elif 'application' in Labels['target']:
                        pyautogui.hotkey('alt','f4')

                    else:
                        if once == 1:
                            speak('Sorry ,sir, i can\'t find the application')
                            once = 0
                        else:
                            speak('I beg your pardon')
                            once = 1

                elif 'stop' in Labels['verb']:
                    speak('Thank you sir!')
                    break

                else:
                    speak('I beg your pardon')
        except Exception as e:
            print("Exception",e)
            pass
        Labels['verb'] = Labels['target'] = ''




greet()
while True:
    try:
        with sr.Microphone() as source:
            print("running")
            voice = listener.listen(source, timeout=0.5,
                                    phrase_time_limit=1.5)  # listens from the source and creates audio
            command = listener.recognize_google(voice)
            print(command)
            if 'hello' in command.lower() or 'jarvis' in command.lower():
                take_input()
            elif 'stop' in command.lower():
                speak('Stopping sir')
                break
    except:
        pass
