import os
import datetime
import subprocess
import psutil
import webbrowser
import pyttsx3
import speech_recognition as sr
import pyautogui
import requests

# Text-to-Speech
engine = pyttsx3.init()
engine.setProperty('rate', 175)
engine.setProperty('volume', 1)

def speak(text):
    print(f"\n JARVIS: {text}\n")
    engine.say(text)
    engine.runAndWait()

#  Speech Recognition
def listen():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        recognizer.adjust_for_ambient_noise(source, duration=1)
        try:
            audio = recognizer.listen(source, timeout=5, phrase_time_limit=7)
            command = recognizer.recognize_google(audio).lower()
            print(f" Heard: {command}")
            return command
        except sr.UnknownValueError:
            print(" Speech error: could not understand")
            return ""
        except sr.RequestError:
            speak("Network error with speech recognition.")
            return ""
        except Exception:
            return ""

#  Greeting based on time
def greet_user():
    hour = datetime.datetime.now().hour
    if hour < 12:
        greeting = "Good morning"
    elif hour < 18:
        greeting = "Good afternoon"
    else:
        greeting = "Good evening"
    speak(f"{greeting}, Ayush! How can I assist you today?")

#  Remember tasks
TASKS_FILE = "tasks.txt"

def remember_task(task):
    task = task.replace("that", "", 1).strip()  # Remove leading 'that'
    if os.path.exists(TASKS_FILE):
        with open(TASKS_FILE, "r") as f:
            existing = [line.strip() for line in f.readlines()]
        if task in existing:
            speak(f"You already asked me to remember {task}")
            return
    with open(TASKS_FILE, "a") as f:
        f.write(task + "\n")
    speak(f"Got it! I will remember {task}")

def read_tasks():
    if not os.path.exists(TASKS_FILE):
        speak("You have no tasks saved for tomorrow.")
    else:
        with open(TASKS_FILE, "r") as f:
            tasks = [t.strip() for t in f.readlines()]
        if tasks:
            speak("Here are your tasks for tomorrow:")
            for t in tasks:
                speak(t)
        else:
            speak("You have no tasks for tomorrow.")

#  Brightness and Volume Helpers
def set_brightness(level):
    try:
        import wmi
        level = max(0, min(100, level))
        wmi_interface = wmi.WMI(namespace='wmi')
        methods = wmi_interface.WmiMonitorBrightnessMethods()[0]
        methods.WmiSetBrightness(level, 0)
        speak(f"Brightness set to {level} percent")
    except Exception as e:
        speak("Sorry, I couldn’t change the brightness.")

def set_volume(level):
    try:
        from ctypes import cast, POINTER
        from comtypes import CLSCTX_ALL
        from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
        level = max(0, min(100, level))
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        volume.SetMasterVolumeLevelScalar(level / 100, None)
        speak(f"Volume set to {level} percent")
    except Exception as e:
        speak("Sorry, I couldn’t change the volume.")

# DuckDuckGo Search 
def search_duckduckgo(query):
    try:
        url = f"https://api.duckduckgo.com/?q={query}&format=json&no_redirect=1&skip_disambig=1"
        response = requests.get(url).json()
        answer = response.get("AbstractText", "").strip()
        if answer:
            # Limit to 1 sentences
            sentences = answer.split(". ")
            answer_short = ". ".join(sentences[:1])
            if not answer_short.endswith("."):
                answer_short += "."
            return answer_short
        else:
            # Fallback: open browser
            webbrowser.open(f"https://duckduckgo.com/?q={query}")
            return f"Opened DuckDuckGo for {query}."
    except Exception as e:
        print(e)
        return "Sorry, an error occurred during search."

# Main Command Execution
def execute_command(command):
    command = command.lower()

    if "time" in command:
        speak("The time is " + datetime.datetime.now().strftime("%I:%M %p"))

    elif "open notepad" in command or "open notebook" in command:
        speak("Opening Notepad")
        subprocess.Popen(["notepad.exe"])

    elif "close notepad" in command or "close notebook" in command:
        speak("Closing Notepad")
        os.system("taskkill /f /im notepad.exe")

    elif "type" in command:
        text = command.replace("type", "").strip()
        pyautogui.typewrite(text)
        speak(f"Typed {text} in Notepad")

    elif "open youtube" in command:
        speak("Opening YouTube")
        webbrowser.open("https://www.youtube.com")

    elif "search" in command or "search for" in command:
        query = command.replace("search", "").replace("for", "").strip()
        speak(search_duckduckgo(query))

    elif "set brightness" in command:
        try:
            level = int(command.split()[-1].replace("%", ""))
            set_brightness(level)
        except:
            speak("Brightness command format not recognized.")

    elif "set volume" in command:
        try:
            level = int(command.split()[-1].replace("%", ""))
            set_volume(level)
        except:
            speak("Volume command format not recognized.")

    elif "battery" in command:
        battery = psutil.sensors_battery()
        percent = battery.percent
        plugged = "plugged in" if battery.power_plugged else "not plugged in"
        speak(f"Battery at {percent} percent, {plugged}")

    elif "turn off wifi" in command:
        speak("Turning off Wi-Fi")
        os.system("netsh interface set interface name=\"Wi-Fi\" admin=disable")

    elif "turn on wifi" in command:
        speak("Turning on Wi-Fi")
        os.system("netsh interface set interface name=\"Wi-Fi\" admin=enable")

    elif "joke" in command:
        speak("If you put a million monkeys at a million keyboards, one of them will eventually write a Java program. The rest of them will write Perl.")

    elif "remember" in command:
        task = command.replace("remember", "").strip()
        remember_task(task)

    elif "what i have to do tomorrow" in command:
        read_tasks()

    elif "exit" in command or "quit" in command:
        speak("Shutting down. Goodbye!")
        exit()

    else:
        speak(search_duckduckgo(command))

#  Main Loop
if __name__ == "__main__":
    greet_user()
    while True:
        user_command = listen()
        if user_command:
            execute_command(user_command)
