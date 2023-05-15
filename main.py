import speech_recognition as sr
from docx import Document
from pywikihow import search_wikihow
import pyttsx3
import datetime
import subprocess
import wikipedia
import webbrowser
import os

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)

def speak(*audio):
    engine.say(*audio)
    engine.runAndWait()

def wishme():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good Morning")
        print("Good Morning")
    elif hour >= 12 and hour < 18:
        speak("Good Afternoon")
        print("Good Afternoon")
    else:
        speak("Good Evening")
        print("Good Evening")

    speak("I'am mandle sir, how may i help you")
    print("I'am mandle sir, how may i help you")

def note(text):
    date = datetime.datetime.now()
    file_name = str(date).replace(':', '-') + 'note.txt'
    with open(file_name, 'w') as f:
        f.write(text)

    subprocess.Popen(['notepad.exe', file_name])

def word(text):
    date = datetime.datetime.now()
    file_name = str(date).replace(':', '-') + 'doc.docx'
    with open(file_name, 'w') as f:
        f.write(text)

def takeCommand():
    global query
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening.........")
        r.pause_threshold = 0.5
        audio = r.listen(source)
        try:
            print("Recognizing......")
            query = r.recognize_google(audio, language='en-in')
            print("User said:", query)

        except Exception as e:
            print("Say that again")
            query = "Nothing"
        return query

if __name__ == '__main__':
    speak("Hi")
    wishme()
    while True:
        query = takeCommand().lower()
        if 'wikipedia' in query:
            speak("Searching wikipedia")
            print('Searching wikipedia')
            query = query.replace("wikipedia", " ")
            results = wikipedia.summary(query, sentences=2)
            speak("According to wikipedia")
            print(results)
            speak(results)

        elif 'search' in query:
            speak('Searching')
            print('Searching.....................')
            query = query.replace("search", "")
            webbrowser.open_new_tab(query)

        elif 'activate how to do mode' in query:
            speak("How to do mode is activated")
            print("How to do mode is activated")
            while True:
                speak("please tell me what do you want to know")
                print("please tell me what do you want to know")
                how = takeCommand()
                try:
                    if 'exit' in how or 'close' in how or 'quit' in how:
                        speak("how to do mode is closed")
                        print("how to do mode is closed")
                        break
                    else:
                        max_results = 1
                        how_to = search_wikihow(how, max_results)
                        assert len(how_to) == 1
                        how_to[0].print()
                        speak(how_to[0].summary)
                except Exception as e:
                    speak("sorry sir i am not able to find this")
                    print("sorry sir i am not able to find this")

        elif 'hello' in query:
            speak('hello sir, how are you')
            print('hello sir, how are you')

        elif 'fine' in query:
            speak('well, good to hear that sir')
            print('well, good to hear that sir')

        elif 'open google' in query:
            webbrowser.open("chrome.exe")
            speak("google is open now")
            print("google is open now")

        elif 'open youtube' in query:
            webbrowser.open("youtube.com")
            speak("Youtube is open now")
            print("Youtube is open now")

        elif 'open stack overflow' in query:
            webbrowser.open("stackoverflow.com")

        elif 'open whatsapp' in query or 'whatsapp' in query:
            whatsapp_path = 'C:\\Users\\Manohar\\AppData\\Local\\WhatsApp\\WhatsApp.exe'
            speak("Opening Whats App")
            print('Opening Whats App')
            os.startfile(whatsapp_path)

        elif 'open gmail' in query or 'gmail' in query:
            webbrowser.open_new_tab("gmail.com")
            speak("GMail is open now")

        elif 'open photos' in query or 'open pictures' in query or 'pictures' in query:
            pic_dir = 'D:\\Pictures'
            photos = os.listdir(pic_dir)
            speak('Opening pictures folder')
            print('Opening pictures folder')
            os.startfile(pic_dir)

        elif 'open movies' in query or 'movies' in query or 'movie' in query:
            movie_dir = 'F:\\Movies'
            movies = os.listdir(movie_dir)
            speak('Opening movies folder')
            print('Opening movies folder')
            os.startfile(movie_dir)

        elif 'time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            print()
            speak(f"Sir the time is:{strTime}")
            print(f"Sir the time is:{strTime}")

        elif 'open code' in query:
            code_path = 'D:\\Manohar\\Users\\Manohar\\Program Files\\Microsoft VS Code\\Code.exe'
            speak("Opening Visual Studio Code")
            print('Opening Visual Studio Code')
            os.startfile(code_path)

        elif 'make a note' in query:
            speak('what would you like me to write down')
            print('what would you like me to write down')
            note_text = takeCommand().lower()
            note(note_text)

        elif 'create document file' in query:
            document = Document()
            speak('what would you like me to write down')
            print('what would you like me to write down')
            word_text = takeCommand().lower()
            word(word_text)

        elif '***' in query:
            speak('sir please be polite to me')
            print('sir please be polite to me')

        elif 'open word' in query:
            word_path = 'C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE'
            speak("Opening Microsoft Word")
            print('Opening Microsoft Word')
            os.startfile(word_path)

        elif 'open zoom' in query:
            zoom_path = 'C:\\Users\\Manohar\\AppData\\Roaming\\Zoom\\bin_00\\Zoom.exe'
            speak("Opening zoom")
            print('Opening zoom')
            os.startfile(zoom_path)

        elif 'open ppt' in query:
            ppt_path = 'C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE'
            speak("Opening PowerPoint")
            print('Opening PowerPoint')
            os.startfile(ppt_path)

        elif 'news' in query:
            news = webbrowser.open_new_tab("https://timesofindia.indiatimes.com/home/headlines")
            speak('Here are some headlines from the Times of India,Happy reading')

        elif 'type' in query:
            file = open('file.txt', 'w')
            file.write('query')

        elif 'open excel' in query:
            excel_path = 'C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE'
            speak('opening Excel')
            print('opening Excel')
            os.startfile(excel_path)

        elif 'who are you' in query:
            speak('Iam mandle your personal assistant.')
            print('Iam mandle your personal assistant')

        elif 'what can you do' in query or 'help' in query:
            speak('I am programmed to minor tasks like'
                  'opening youtube,google chrome,gmail,whatsapp and stackoverflow ,read out time,search wikipedia' 
                  'get top headline news from times of india, and also open windows applications like'
                  'MSWord, MSPowerPoint, MSExcel and open few folders and can create and type in a notepad file')

        elif 'creator' in query or 'created' in query:
            speak('I have been created and named as mandle by Manohar')
            print('I have been created and named as mandle by Manohar')

        elif 'exit' in query or 'close' in query or 'shutdown' in query or 'sleep' in query or 'logout' in query:
            speak("have a good day sir")
            print("have a good day sir")
            exit()




