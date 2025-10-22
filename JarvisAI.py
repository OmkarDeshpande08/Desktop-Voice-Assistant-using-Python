"""
JarvisAI - Core module for the voice assistant
"""

import pyttsx3
import speech_recognition as sr
import datetime
import platform
import psutil
import requests
import json
import os


class JarvisAssistant:
    def __init__(self):
        # Initialize text-to-speech engine
        self.engine = pyttsx3.init()
        self.setup_voice()
        
        # Initialize speech recognition
        self.recognizer = sr.Recognizer()
        self.microphone = sr.Microphone()
        
        # Adjust for ambient noise
        with self.microphone as source:
            self.recognizer.adjust_for_ambient_noise(source, duration=1)
    
    def setup_voice(self):
        """Configure the text-to-speech voice settings"""
        voices = self.engine.getProperty('voices')
        
        # Try to set a male voice (usually index 0)
        if voices:
            self.engine.setProperty('voice', voices[0].id)
        
        # Set speech rate and volume
        self.engine.setProperty('rate', 200)
        self.engine.setProperty('volume', 0.9)
    
    def tts(self, text):
        """Text-to-speech function"""
        try:
            print(f"Jarvis: {text}")
            self.engine.say(text)
            self.engine.runAndWait()
        except Exception as e:
            print(f"TTS Error: {e}")
    
    def mic_input(self):
        """Listen to microphone input and convert to text"""
        try:
            with self.microphone as source:
                print("Listening...")
                # Listen for audio with timeout
                audio = self.recognizer.listen(source, timeout=5, phrase_time_limit=5)
            
            print("Recognizing...")
            # Use Google's speech recognition
            command = self.recognizer.recognize_google(audio, language='en-US')
            print(f"User said: {command}")
            return command.lower()
            
        except sr.WaitTimeoutError:
            print("Listening timeout")
            return "none"
        except sr.UnknownValueError:
            print("Could not understand audio")
            return "none"
        except sr.RequestError as e:
            print(f"Could not request results; {e}")
            return "none"
        except Exception as e:
            print(f"Error in mic_input: {e}")
            return "none"
    
    def tell_me_date(self):
        """Get current date"""
        try:
            today = datetime.date.today()
            return today.strftime("Today is %A, %B %d, %Y")
        except Exception as e:
            return f"Error getting date: {e}"
    
    def tell_time(self):
        """Get current time"""
        try:
            current_time = datetime.datetime.now()
            return current_time.strftime("%I:%M %p")
        except Exception as e:
            return f"Error getting time: {e}"
    
    def weather(self, city="London"):
        """Get weather information for a city"""
        try:
            # Using OpenWeatherMap API (free tier)
            # You need to get an API key from openweathermap.org
            api_key = "YOUR_OPENWEATHER_API_KEY"  # Replace with actual API key
            
            if api_key == "YOUR_OPENWEATHER_API_KEY":
                return f"Weather service not configured. Please add your OpenWeatherMap API key."
            
            base_url = "http://api.openweathermap.org/data/2.5/weather"
            params = {
                "q": city,
                "appid": api_key,
                "units": "metric"
            }
            
            response = requests.get(base_url, params=params, timeout=10)
            
            if response.status_code == 200:
                data = response.json()
                temp = data['main']['temp']
                description = data['weather'][0]['description']
                return f"The weather in {city} is {description} with a temperature of {temp} degrees Celsius"
            else:
                return f"Could not get weather data for {city}"
                
        except requests.RequestException:
            return "Network error: Could not fetch weather data"
        except Exception as e:
            return f"Weather error: {e}"
    
    def system_info(self):
        """Get system information"""
        try:
            # Get system information
            system = platform.system()
            node = platform.node()
            release = platform.release()
            machine = platform.machine()
            processor = platform.processor()
            
            # Get memory info
            memory = psutil.virtual_memory()
            memory_total = round(memory.total / (1024**3), 2)  # GB
            memory_available = round(memory.available / (1024**3), 2)  # GB
            memory_percent = memory.percent
            
            # Get CPU info
            cpu_percent = psutil.cpu_percent(interval=1)
            cpu_count = psutil.cpu_count()
            
            info = f"""System Information:
Operating System: {system} {release}
Computer: {node}
Architecture: {machine}
Processor: {processor}
CPU Cores: {cpu_count}
CPU Usage: {cpu_percent}%
Total Memory: {memory_total} GB
Available Memory: {memory_available} GB
Memory Usage: {memory_percent}%"""
            
            return info
            
        except Exception as e:
            return f"Error getting system info: {e}"
    
    def get_battery_info(self):
        """Get battery information"""
        try:
            battery = psutil.sensors_battery()
            if battery:
                percent = battery.percent
                plugged = battery.power_plugged
                status = "Plugged in" if plugged else "On battery"
                return f"Battery: {percent}% ({status})"
            else:
                return "Battery information not available"
        except Exception as e:
            return f"Error getting battery info: {e}"
    
    def create_directories(self):
        """Create necessary directories"""
        directories = [
            "screenshots",
            "downloads",
            "Jarvis/utils/images"
        ]
        
        for directory in directories:
            try:
                os.makedirs(directory, exist_ok=True)
            except Exception as e:
                print(f"Error creating directory {directory}: {e}")


# Initialize the assistant when module is imported
if __name__ == "__main__":
    assistant = JarvisAssistant()
    assistant.tts("Jarvis Assistant module loaded successfully")