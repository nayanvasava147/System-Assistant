from flask import Flask, render_template_string, jsonify, request
from comtypes import CoInitialize, CoUninitialize
import os, webbrowser, threading, datetime, pyttsx3, speech_recognition as sr, subprocess, pyautogui, time

app = Flask(__name__)

class VoiceAssistant:
    def _init_(self):
        self.engine = pyttsx3.init()
        voices = self.engine.getProperty('voices')
        self.engine.setProperty('voice', voices[1].id if len(voices) > 1 else voices[0].id)
        self.silent_mode = False
        self.recording = False
        self.output_dir = os.getcwd()
        os.makedirs(self.output_dir, exist_ok=True)
        self.app_keywords = {
            "chrome": "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
            "notepad": "notepad.exe",
            "calculator": "calc.exe"
        }
        self.app_names = ["chrome", "notepad", "calculator"]

    def speak(self, audio):
        if not self.silent_mode:
            self.engine.say(audio)
            self.engine.runAndWait()

    def take_command(self):
        r = sr.Recognizer()
        with sr.Microphone() as source:
            print("Listening...")
            r.pause_threshold = 1
            try:
                audio = r.listen(source, timeout=4, phrase_time_limit=3)
            except sr.WaitTimeoutError:
                self.speak("I didn't hear anything.")
                return "none"
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language='en-in')
            print(f"User said: {query}")
            return query.lower()
        except sr.UnknownValueError:
            self.speak("Sorry, I didn't catch that.")
            return "none"
        except sr.RequestError:
            self.speak("Couldn't connect to the recognition service.")
            return "none"

    def handle_silent_mode(self, query):
        silent_commands = ["chup raho", "shut up", "bandh", "chup", "stay silent", "silent"]
        wake_commands = ["system", "hello system", "hello assistant"]
        if any(word in query for word in silent_commands):
            self.speak("Sorry to interrupt, going silent.")
            self.silent_mode = True
            return True
        elif any(word in query for word in wake_commands):
            self.silent_mode = False
            self.speak("Yeah, I'm here. How can I help you?")
            time.sleep(2)
            return False
        return False

    def take_screenshot(self):
        screenshot = pyautogui.screenshot()
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filepath = os.path.join(self.output_dir, f"screenshot_{timestamp}.png")
        screenshot.save(filepath)
        self.speak(f"Screenshot saved as {filepath}")

    def open_app_or_website(self, query):
        for app_name, path in self.app_keywords.items():
            if app_name in query:
                if "close" in query:
                    os.system(f"taskkill /f /im {app_name}.exe")
                    self.silent_mode = False
                    self.speak(f"{app_name.capitalize()} closed.")
                else:
                    os.startfile(path)
                    self.silent_mode = True
                    self.speak(f"{app_name.capitalize()} opened.")
                return

        site_name = query.replace("open ", "").strip()
        default_url = f"http://www.{site_name}.com"
        self.speak(f"Opening {site_name}.com.")
        webbrowser.open(default_url)

    def youtube_search(self):
        self.speak("YouTube is opened. What would you like to search?")
        for attempt in range(3):
            search_query = self.take_command()
            if search_query and search_query != "none":
                youtube_search_url = f"https://www.youtube.com/results?search_query={search_query.replace(' ', '+')}"
                self.speak(f"Searching YouTube for {search_query}.")
                webbrowser.open(youtube_search_url)
                return
            elif attempt < 2:
                self.speak("I didn't catch that. Please repeat.")
        self.speak("I couldn't hear your search query. Please try again later.")

    def wikipedia_search(self, query):
        topic = query.replace("search", "").strip()
        wikipedia_url = f"https://en.wikipedia.org/wiki/{topic.replace(' ', '_')}"
        self.speak(f"Searching Wikipedia for {topic}.")
        webbrowser.open(wikipedia_url)

    def anime_search(self, query):
        topic = query.replace("anime", "").strip()
        if topic:
            search_url = f"https://hianime.to/search?keyword={topic}"
            self.speak(f"Searching HiAnime for {topic}.")
            webbrowser.open(search_url)
        else:
            self.speak("Please specify the anime title you want to search for.")

    def open_whatsapp(self):
        # Open WhatsApp Desktop or Web
        os.system("start whatsapp:")
        time.sleep(8)

    def open_chat(self, name):
        pyautogui.hotkey("ctrl", "f")
        time.sleep(1)
        pyautogui.write(name)
        time.sleep(2)
        pyautogui.press('esc')
        time.sleep(0.5)
        pyautogui.click(300, 250)  # Adjust based on your screen size for correct click location
        time.sleep(2)

    def send_whatsapp_message(self, name, message):
        self.open_whatsapp()
        self.open_chat(name)
        pyautogui.write(message)
        pyautogui.press("enter")
        self.speak(f"Message sent to {name}")

    def whatsapp_call(self, name, call_type="voice"):
        self.open_whatsapp()
        self.open_chat(name)
        time.sleep(1)
        if call_type == "voice":
            pyautogui.click(1800, 100)  # Adjust these values based on screen size
            self.speak(f"Voice calling {name}")
        elif call_type == "video":
            pyautogui.click(1740, 100)  # Adjust these values based on screen size
            self.speak(f"Video calling {name}")

    def process_whatsapp_command(self, query):
        if "send" in query and "to" in query:
            parts = query.split("to")
            message = parts[0].replace("send", "").strip()
            contact = parts[1].strip()
            self.send_whatsapp_message(contact, message)
        elif "video call" in query:
            contact = query.replace("video call", "").strip()
            self.whatsapp_call(contact, call_type="video")
        elif "call" in query:
            contact = query.replace("call", "").strip()
            self.whatsapp_call(contact, call_type="voice")
        else:
            self.speak("Sorry, I didn't understand the WhatsApp command.")

    def execute_query(self, query):
        if self.handle_silent_mode(query): return
        if "anime" in query: self.anime_search(query)
        elif "youtube" in query: self.youtube_search()
        elif "search" in query: self.wikipedia_search(query)
        elif "open" in query: self.open_app_or_website(query)
        elif "screenshot" in query: self.take_screenshot()
        elif "send" in query or "call" in query or "video call" in query: self.process_whatsapp_command(query)
        elif "stop" in query or "exit" in query:
            self.speak("Shutting down. Have a great day!")
            exit()
        elif "close" in query:
            self.speak("Please specify which app to close.")

    def run(self):
        CoInitialize()
        self.speak("Starting voice assistant. How can I assist you?")
        while True:
            query = self.take_command()
            if query != "none":
                self.execute_query(query)
        CoUninitialize()

HTML_CONTENT = '''
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Voice Recognition</title>
    <style>
        body {
            margin: 0;
            padding: 0;
            font-family: 'Segoe UI', sans-serif;
            background-color: #f4f4f4;
            display: flex;
            justify-content: center;
            align-items: center;
            height: 100vh;
        }

        .container {
            text-align: center;
            background: #ffffff;
            padding: 40px;
            border-radius: 20px;
            box-shadow: 0 0 30px rgba(0, 0, 0, 0.1);
            width: 80%;
            max-width: 600px;
        }

        .mic-icon img {
            width: 60px;
            margin-bottom: 20px;
        }

        .title {
            font-size: 28px;
            font-weight: bold;
            color: #333;
            margin-bottom: 30px;
            letter-spacing: 2px;
        }

        .waveform {
            display: flex;
            justify-content: center;
            gap: 10px;
            margin-bottom: 30px;
        }

        .wave {
            width: 10px;
            height: 60px;
            background: linear-gradient(to bottom, #00c3ff, #ff55c3);
            border-radius: 10px;
            animation: wave-animation 1.2s infinite ease-in-out;
        }

        .wave:nth-child(2) { animation-delay: 0.2s; }
        .wave:nth-child(3) { animation-delay: 0.4s; }
        .wave:nth-child(4) { animation-delay: 0.6s; }
        .wave:nth-child(5) { animation-delay: 0.8s; }

        @keyframes wave-animation {
            0%, 100% { height: 60px; }
            50% { height: 20px; }
        }

        .start-btn {
            background-color: #00c3ff;
            color: white;
            padding: 12px 30px;
            border-radius: 30px;
            font-size: 18px;
            cursor: pointer;
            transition: background 0.3s ease;
        }

        .start-btn:hover {
            background-color: #ff55c3;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="mic-icon">
            <img src="{{ url_for('static', filename='mic.png') }}" alt="Mic Icon">
        </div>
        <div class="title">VOICE RECOGNITION</div>
        <div class="waveform">
            <div class="wave"></div>
            <div class="wave"></div>
            <div class="wave"></div>
            <div class="wave"></div>
            <div class="wave"></div>
        </div>
        <div class="start-btn" onclick="startAssistant()">Start Listening</div>
    </div>

    <script>
        function startAssistant() {
            fetch('/start_assistant', { method: 'POST' })
                .then(res => res.json())
                .then(data => alert(data.response));
        }
    </script>
</body>
</html>
'''

@app.route('/')
def index():
    return render_template_string(HTML_CONTENT)

@app.route('/start_assistant', methods=['POST'])
def start_assistant():
    def run_assistant():
        assistant = VoiceAssistant()
        assistant.run()

    thread = threading.Thread(target=run_assistant)
    thread.daemon = True
    thread.start()
    return jsonify({'response': 'Voice Assistant started'})

if __name__ == '_main_':
    app.run(debug=True)