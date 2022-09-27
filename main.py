from tkinter import *
from tkinter import Button, Label, Entry, Text, messagebox
from tkinter import Toplevel
import webbrowser
import keyboard
import speech_recognition as sr
import pyttsx3
import datetime
import subprocess
import pyjokes
import os
import wikipedia

chrome = 'C:/Program Files/Google/Chrome/Application/chrome.exe %s'

engine = pyttsx3.init('sapi5')
rate = engine.getProperty('rate')
engine.setProperty('rate', 170)
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)


def speak(text):
    engine.say(text)
    engine.runAndWait()

try:
    import pywhatkit as kit
    speak("internet Connection is available")
except:
    speak("No internet Connection")
    speak("Please Connect internet for better experience")
    pass

from email.message import EmailMessage
import smtplib

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good Morning!")
    elif hour >= 12 and hour < 18:
        speak("Good Afternoon!")
    else:
        speak("Good Evening!")

    speak("I am your Voice Assistant. Please tell me how may I help you")


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 1
        audio = r.listen(source, phrase_time_limit = 3)

    try:
        query = r.recognize_google(audio, language='en-in')
        speak(f"User said: {query}\n")

    except Exception as e:
        speak("Say that again please...")
        return "None"
    return query


def note(text):
    date = datetime.datetime.now()
    file_name = str(date).replace(":", "-") + "-note.txt"

    with open(file_name, "w") as f:
        f.write(text)

    subprocess.Popen(["notepad.exe", file_name])


def date():
    now = datetime.datetime.now()
    month_name = now.month
    day_name = now.day
    month_names = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
    ordinalnames = [ '1st', '2nd', '3rd', ' 4th', '5th', '6th', '7th', '8th', '9th', '10th', '11th', '12th', '13th', '14th', '15th', '16th', '17th', '18th', '19th', '20th', '21st', '22nd', '23rd','24rd', '25th', '26th', '27th', '28th', '29th', '30th', '31st'] 

    global yearday
    yearday ="Today is "+ month_names[month_name-1] +" " + ordinalnames[day_name-1] + '.'


def Stop():
    window.destroy()


def Home_page():
    Voice_window.destroy()
    window.deiconify()


def Start():
    window.withdraw()

    global Voice_window
    if not any(isinstance(x, Toplevel) for x in window.winfo_children()):

        Voice_window = Toplevel(window)
        Voice_window.title("Info")
        Voice_window.geometry("960x666")
        Voice_window.resizable(0, 0)
        Voice_window.iconbitmap('images\logo.ico')
        Voice_window.configure(bg="#454545")

        name_label = Label(Voice_window, text = name_assistant,width = 300, height=1, bg = "black", fg="white", font = ("Calibri", 92))
        name_label.pack()

        goback = Button(Voice_window, text="<<< Go back", command=Home_page, font=("Calibri",18), fg="white", bg="#454545", borderwidth=0, activebackground="#454545", cursor="hand2")
        goback.place(x=0, y=156)

        exit = Button(Voice_window, text="Exit", command=Stop, font=("Calibri",18), fg="white", bg="#454545", borderwidth=0, activebackground="#454545", cursor="hand2")
        exit.place(x=910, y=156)

        # query = input("Enter: ")
        run = 1
        if __name__=='__main__':

            app_string = ["open word", "open powerpoint", "open excel", "open edge", "open chrome"]
            app_link = [r'\Word 2016.lnk',r'\PowerPoint 2016.lnk', r'\Excel 2016.lnk', r'\Microsoft Edge.lnk', r'\Google Chrome.lnk']
            app_dest = r'C:\ProgramData\Microsoft\Windows\Start Menu\Programs'
            office_dest = r'C:\ProgramData\Microsoft\Windows\Start Menu\Programs'

            query = takeCommand().lower()

            user_said = Text(Voice_window, height=8, width=44, padx=0, pady=0, font=('Arial', 11, 'bold'))
            uasked = 'User said: '+ query
            user_said.insert(1.0, uasked)
            user_said.tag_configure("center", justify="center")
            user_said.tag_add("center", 1.0, "end")
            user_said.place(x=6, y=320)

            if 'search' in query:
                results = 'Searching...'
                speak(results)
                answer = Text(Voice_window, height=8, width=44,  padx=0, pady=0, font=('Arial', 11, 'bold'))
                answer.insert(1.0, results)
                answer.tag_configure("center", justify="center")
                answer.tag_add("center", 1.0, "end")
                answer.place(x=601, y=465)
                query = query.replace("search", "")
                web = 'https://www.google.com/search?q=' + query
                webbrowser.get(chrome).open(web)

            if "hello" in query or "hey" in query:
                results = 'Hello, Sir!'
                speak(results)
                answer = Text(Voice_window, height=8, width=44,  padx=0, pady=0, font=('Arial', 11, 'bold'))
                answer.insert(1.0, results)
                answer.tag_configure("center", justify="center")
                answer.tag_add("center", 1.0, "end")
                answer.place(x=601, y=465)
                wishMe()

            if "good bye" in query or "ok bye" in query or "stop" in query:
                results = 'Your personal assistant is shutting down, Good bye'
                speak(results)
                answer = Text(Voice_window, height=8, width=44,  padx=0, pady=0, font=('Arial', 11, 'bold'))
                answer.insert(1.0, results)
                answer.tag_configure("center", justify="center")
                answer.tag_add("center", 1.0, "end")
                answer.place(x=601, y=465)
                window.destroy()

            if 'joke' in query:
                results = pyjokes.get_joke()
                speak(results)
                answer = Text(Voice_window, height=8, width=44,  padx=0, pady=0, font=('Arial', 11, 'bold'))
                answer.insert(1.0, results)
                answer.tag_configure("center", justify="center")
                answer.tag_add("center", 1.0, "end")
                answer.place(x=601, y=465)

            if "open whatsapp" in query:
                results = "whatsapp is opening now"
                speak(results)
                answer = Text(Voice_window, height=8, width=44,  padx=0, pady=0, font=('Arial', 11, 'bold'))
                answer.insert(1.0, results)
                answer.tag_configure("center", justify="center")
                answer.tag_add("center", 1.0, "end")
                answer.place(x=601, y=465)
                webbrowser.get(chrome).open("web.whatsapp.com")


            if 'open youtube' in query:
                results = "youtube is opening now"
                speak(results)
                answer = Text(Voice_window, height=8, width=44,  padx=0, pady=0, font=('Arial', 11, 'bold'))
                answer.insert(1.0, results)
                answer.tag_configure("center", justify="center")
                answer.tag_add("center", 1.0, "end")
                answer.place(x=601, y=465)
                webbrowser.open_new_tab("https://www.youtube.com")

            if 'open google' in query:
                results = "Google chrome is open now"
                speak(results)
                answer = Text(Voice_window, height=8, width=44,  padx=0, pady=0, font=('Arial', 11, 'bold'))
                answer.insert(1.0, results)
                answer.tag_configure("center", justify="center")
                answer.tag_add("center", 1.0, "end")
                answer.place(x=601, y=465)
                webbrowser.open_new_tab("https://www.google.com")

            if 'open gmail' in query:
                results = "Google Mail open now"
                speak(results)
                answer = Text(Voice_window, height=8, width=44,  padx=0, pady=0, font=('Arial', 11, 'bold'))
                answer.insert(1.0, results)
                answer.tag_configure("center", justify="center")
                answer.tag_add("center", 1.0, "end")
                answer.place(x=601, y=465)
                webbrowser.open_new_tab("mail.google.com")

            if 'open netflix' in query:
                results = "Netflix open now"
                speak(results)
                answer = Text(Voice_window, height=8, width=44,  padx=0, pady=0, font=('Arial', 11, 'bold'))
                answer.insert(1.0, results)
                answer.tag_configure("center", justify="center")
                answer.tag_add("center", 1.0, "end")
                answer.place(x=601, y=465)
                webbrowser.open_new_tab("netflix.com/browse") 

            if 'on youtube' in query:
                try:
                    query = query.replace("play", "")
                    query = query.replace("on youtube", "")
                    results = f"playing {query} on youtube"
                    speak(results)
                    answer = Text(Voice_window, height=8, width=44,  padx=0, pady=0, font=('Arial', 11, 'bold'))
                    answer.insert(1.0, results)
                    answer.tag_configure("center", justify="center")
                    answer.tag_add("center", 1.0, "end")
                    answer.place(x=601, y=465)
                    kit.playonyt(query)
                except:
                    results = f"try again..."
                    speak(results)
                    answer = Text(Voice_window, height=8, width=44,  padx=0, pady=0, font=('Arial', 11, 'bold'))
                    answer.insert(1.0, results)
                    answer.tag_configure("center", justify="center")
                    answer.tag_add("center", 1.0, "end")
                    answer.place(x=601, y=465)
                    Home_page()

            if app_string[0] in query:
                results = "Microsoft office Word is opening now"
                speak(results)
                answer = Text(Voice_window, height=8, width=44,  padx=0, pady=0, font=('Arial', 11, 'bold'))
                answer.insert(1.0, results)
                answer.tag_configure("center", justify="center")
                answer.tag_add("center", 1.0, "end")
                answer.place(x=601, y=465)
                os.startfile(office_dest + app_link[0])

            if app_string[1] in query:
                results = "Microsoft office PowerPoint is opening now"
                speak(results)
                answer = Text(Voice_window, height=8, width=44,  padx=0, pady=0, font=('Arial', 11, 'bold'))
                answer.insert(1.0, results)
                answer.tag_configure("center", justify="center")
                answer.tag_add("center", 1.0, "end")
                answer.place(x=601, y=465)
                os.startfile(office_dest + app_link[1])

            if app_string[2] in query:
                results = "Microsoft office Excel is opening now"
                speak(results)
                answer = Text(Voice_window, height=8, width=44,  padx=0, pady=0, font=('Arial', 11, 'bold'))
                answer.insert(1.0, results)
                answer.tag_configure("center", justify="center")
                answer.tag_add("center", 1.0, "end")
                answer.place(x=601, y=465)
                os.startfile(office_dest + app_link[2])

            if app_string[3] in query:
                results = "Edge is opening now"
                speak(results)
                answer = Text(Voice_window, height=8, width=44,  padx=0, pady=0, font=('Arial', 11, 'bold'))
                answer.insert(1.0, results)
                answer.tag_configure("center", justify="center")
                answer.tag_add("center", 1.0, "end")
                answer.place(x=601, y=465)
                os.startfile(app_dest + app_link[3])

            if app_string[4] in query:
                results = "Chrome is opening now"
                speak(results)
                answer = Text(Voice_window, height=8, width=44,  padx=0, pady=0, font=('Arial', 11, 'bold'))
                answer.insert(1.0, results)
                answer.tag_configure("center", justify="center")
                answer.tag_add("center", 1.0, "end")
                answer.place(x=601, y=465)
                os.startfile(app_dest + app_link[4])

            if 'send whatsapp message' in query:
                query = query.replace("send whatsapp message", "")
                kit.sendwhatmsg_instantly(f"+918580578946", query)

            elif 'send mail' in query:
                msg = "this message is send by voice assistant"
                emailid = "akashdeepsingh4703@gmail.com"
                try:
                    email = EmailMessage()
                    email['To'] = emailid
                    email["Subject"] = msg
                    email['From'] = "Voice Assistant"
                    email.set_content(msg)
                    s = smtplib.SMTP("smtp.gmail.com", 587)
                    s.starttls()
                    s.login('otpforpythondeep@gmail,com', 'singhakash')
                    s.send_message(email)
                    s.close()
                except Exception as e:
                    speak(e)
                    Home_page()

            if 'news' in query:
                results = "Here are some headlines from the Times of India"
                speak(results)
                answer = Text(Voice_window, height=8, width=44,  padx=0, pady=0, font=('Arial', 11, 'bold'))
                answer.insert(1.0, results)
                answer.tag_configure("center", justify="center")
                answer.tag_add("center", 1.0, "end")
                answer.place(x=601, y=465)
                webbrowser.open_new_tab("https://timesofindia.indiatimes.com/city/shimla")

            if 'cricket' in query:
                results = "This is live news from cricbuzz"
                speak(results)
                answer = Text(Voice_window, height=8, width=44,  padx=0, pady=0, font=('Arial', 11, 'bold'))
                answer.insert(1.0, results)
                answer.tag_configure("center", justify="center")
                answer.tag_add("center", 1.0, "end")
                answer.place(x=601, y=465)
                webbrowser.get(chrome).open_new_tab("cricbuzz.com")

            if 'time' in query:
                strTime=datetime.datetime.now().strftime("%H:%M:%S")
                results = f"the time is {strTime}"
                speak(results)
                answer = Text(Voice_window, height=8, width=44,  padx=0, pady=0, font=('Arial', 11, 'bold'))
                answer.insert(1.0, results)
                answer.tag_configure("center", justify="center")
                answer.tag_add("center", 1.0, "end")
                answer.place(x=601, y=465)

            if 'date' in query:
                date()
                speak(yearday)
                answer = Text(Voice_window, height=8, width=44,  padx=0, pady=0, font=('Arial', 11, 'bold'))
                answer.insert(1.0, yearday)
                answer.tag_configure("center", justify="center")
                answer.tag_add("center", 1.0, "end")
                answer.place(x=601, y=465)

            if 'make a note' in query or 'note this' in query:
                results = "your note has been created!"
                speak(results)
                answer = Text(Voice_window, height=8, width=44,  padx=0, pady=0, font=('Arial', 11, 'bold'))
                answer.insert(1.0, results)
                answer.tag_configure("center", justify="center")
                answer.tag_add("center", 1.0, "end")
                answer.place(x=601, y=465)
                query = query.replace("make a note", "")
                note(query)
            
            if 'wikipedia' in query:
                try:
                    speak('Searching Wikipedia...')
                    query = query.replace("wikipedia", "")
                    results = wikipedia.summary(query, sentences=2)
                    speak("According to Wikipedia")
                    speak(results)
                    answer = Text(Voice_window, height=8, width=44,  padx=0, pady=0, font=('Arial', 11, 'bold'))
                    answer.insert(1.0, results)
                    answer.tag_configure("center", justify="center")
                    answer.tag_add("center", 1.0, "end")
                    answer.place(x=601, y=465)
                except:
                    results = "Try Again..."
                    speak(results)
                    answer = Text(Voice_window, height=8, width=44,  padx=0, pady=0, font=('Arial', 11, 'bold'))
                    answer.insert(1.0, results)
                    answer.tag_configure("center", justify="center")
                    answer.tag_add("center", 1.0, "end")
                    answer.place(x=601, y=465)
                    Home_page()


            else:
                Home_page()

    else:
        messagebox.showinfo("showinfo", "Top level already exists")

keyboard.add_hotkey("F3", Home_page)
keyboard.add_hotkey("F4", Start)
keyboard.add_hotkey("Esc", Stop)


def change_name():
    name_info = name.get()
    file=open("Assistant_name.txt", "w")
    file.write(name_info)
    file.close()
    settings_window.destroy()
    window.destroy()


def change_name_window():
    
    global settings_window
    global name


    settings_window = Toplevel(window)
    settings_window.title("Settings")
    settings_window.geometry("300x300")
    settings_window.iconbitmap('images\logo.ico')

    
    name = StringVar()

    current_label = Label(settings_window, text = "Current name: "+ name_assistant)
    current_label.pack()

    enter_label = Label(settings_window, text = "Please enter your Virtual Assistant's name below") 
    enter_label.pack(pady=10)   
    

    Name_label = Label(settings_window, text = "Name")
    Name_label.pack(pady=10)
    
    name_entry = Entry(settings_window, textvariable = name)
    name_entry.pack()


    change_name_button = Button(settings_window, text = "Ok", width = 10, height = 1, command = change_name)
    change_name_button.pack(pady=10)


def info():

    global info_window
    info_window = Toplevel(window)
    info_window.title("Info")
    info_window.geometry("160x80")
    info_window.iconbitmap('images\logo.ico')

    creator_label = Label(info_window,text = "Created by Akash Deep")
    creator_label.pack(pady=5)

    for_label = Label(info_window, text = "For Project")
    for_label.pack(pady=5)


def main_window():
    global window
    window = Tk()
    window.title("Voice Assistant")
    window.geometry("960x500")
    window.resizable(0, 0)
    window.iconbitmap('images\logo.ico')
    window.configure(bg="#454545")


    name_label = Label(text = name_assistant,width = 300, bg = "black", fg="white", font = ("Calibri", 13))
    name_label.pack()

    Start_photo = PhotoImage(file = "images\\start.png", width=250, height=60)
    microphone_button = Button(image=Start_photo, command = Start, borderwidth=0, activebackground="#454545", cursor="hand2")
    microphone_button.pack(pady=12)

    Stop_photo = PhotoImage(file = "images\\stop.png", width=250, height=60)
    microphone_button = Button(image=Stop_photo, command = Stop, borderwidth=0, activebackground="#454545", cursor="hand2")
    microphone_button.pack(pady=12)

    settings_photo = PhotoImage(file = "images\\settings.png", width=250, height=60)
    settings_button = Button(image=settings_photo, command = change_name_window, borderwidth=0, activebackground="#454545", cursor="hand2")
    settings_button.pack(pady=12)

    info_button = Button(text ="Info", command = info, font=("Calibri",18), fg="white", bg="#454545", borderwidth=0, activebackground="#454545", cursor="hand2")
    info_button.pack(pady=12)

    end_label = Label(text ="" ,width = 300, bg = "black", fg="white", font = ("Calibri", 13))
    end_label.pack(side='bottom')

    window.mainloop()


try:
    name_file = open("Assistant_name.txt", "r")
    name_assistant = name_file.read()
except:
    file_name = "Assistant_name.txt"

    with open(file_name, "w") as f:
        f.write("VI")
    
    name_file = open("Assistant_name.txt", "r")
    name_assistant = name_file.read()


main_window()