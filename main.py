import speech_recognition as sr
import win32com.client

import webbrowser
speaker = win32com.client.Dispatch("SAPI.SpVoice")


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone () as source:
        r.pause_threshold = 1
        audio = r.listen(source)
    try:
        query = r.recognize_google(audio, language="en-in")
        print(f"User said:{query}")
        return query
    except Exception as e:
        return "Some Error Occurred. Sorry from Jarvis"
while True:
    print("Listening")
    query = takeCommand()
    sites = [["Youtube", "https://youtube.com"], ["Wikipedia", "https://wikipedia.com"], ["Facebook", "https://facebook.com"], ["Google", "https://google.com"], ["Danson Marketing", "https://dansonmarketing.co.uk"]]
    for site in sites:
        if f"open {site[0]}".lower() in query.lower():
            speaker.Speak(f"Opening {site[0]} Dani ")
            webbrowser.open(site[1])
        

