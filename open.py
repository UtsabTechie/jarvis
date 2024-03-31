import win32com.client
import  speech_recognition as sr
import webbrowser
import datetime

def say(text):
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    speaker.Speak(text)

time09 = int(datetime.datetime.now().strftime("%H"))
if 12 > time09 >= 6 :
    print("Good Morning, how can i help you")
    say("good morning sir, how can i help you")
elif 15> time09 >= 12 :
    print("good noon sir, how can i help you")
    say("good noon sir, how can i help you")
elif 18> time09 >= 15:
    print("good afternoon sir, how can i help you")
    say("good afternoon sir, how can i help you")
elif time09 >= 18 :
    print("good evening sir, how can i help you")
    say("good evening sir, how can i help you")


def take_command():
    r = sr.Recognizer()
    r.pause_threshold = 1
    with sr.Microphone() as source:
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language="en-in")
            print(query)
            return query
        except sr.UnknownValueError:
            print("Sorry, I could not understand what you said.")
            say("Sorry, I could not understand what you said.")
            return ""
        except sr.RequestError as e:
            print(f"Could not request results from Google Web Speech API; {e}")
            say(f"Could not request results from Google Web Speech API; {e}")
            return ""

while True:

    print("Listening...")
    x = take_command()
    if x:
        sites = [["youtube", "https://www.youtube.com/"],["wikipedia", "https://www.wikipedia.com/"],["google", "https://www.google.com/"]]
        for site in sites:
            if f'open {site[0]}'.lower() in x.lower():
                webbrowser.open(site[1])
                say(f"opening {site[0]} sir...")
        if "exit" in x:
            break
        else:
            say(x)
