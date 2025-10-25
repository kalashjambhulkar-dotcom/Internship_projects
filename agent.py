
import pyautogui
import subprocess
import os
import time
import speech_recognition as sr  
import pyttsx3  
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.keys import Keys
from translate import Translator  
from datetime import datetime
from webdriver_manager.chrome import ChromeDriverManager  
from fuzzywuzzy import process  
import win32com.client  
import pythoncom


BRAVE_PATH = r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe"
CHROME_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"


recognizer = sr.Recognizer()
microphone = sr.Microphone()


tts_engine = pyttsx3.init()
tts_engine.setProperty('volume', 1.0)  
tts_engine.setProperty('rate', 160)  
voices = tts_engine.getProperty('voices')
if voices:
    tts_engine.setProperty('voice', voices[0].id) 


context = {'browser': None, 'site': None, 'driver': None}


ALLOWED_LANGUAGES = {
    'af': 'Afrikaans', 'sq': 'Albanian', 'am': 'Amharic', 'eu': 'Basque',
    'bn': 'Bengali', 'ca': 'Catalan', 'ceb': 'Cebuano', 'co': 'Corsican',
    'nl': 'Dutch', 'eo': 'Esperanto', 'et': 'Estonian', 'fi': 'Finnish',
    'gl': 'Galician', 'ka': 'Georgian', 'gu': 'Gujarati', 'ht': 'Haitian Creole',
    'ha': 'Hausa', 'haw': 'Hawaiian', 'is': 'Icelandic', 'ig': 'Igbo',
    'ga': 'Irish', 'jw': 'Javanese', 'kn': 'Kannada', 'kk': 'Kazakh',
    'km': 'Khmer', 'ku': 'Kurdish', 'ky': 'Kyrgyz', 'lo': 'Lao',
    'lb': 'Luxembourgish', 'mg': 'Malagasy', 'ms': 'Malay', 'ml': 'Malayalam',
    'mt': 'Maltese', 'mi': 'Maori', 'mr': 'Marathi', 'mn': 'Mongolian',
    'ne': 'Nepali', 'no': 'Norwegian', 'ny': 'Nyanja', 'ps': 'Pashto',
    'sm': 'Samoan', 'gd': 'Scots Gaelic', 'sn': 'Shona', 'sd': 'Sindhi',
    'si': 'Sinhala', 'so': 'Somali', 'su': 'Sundanese', 'sw': 'Swahili',
    'tg': 'Tajik', 'ta': 'Tamil', 'te': 'Telugu', 'th': 'Thai',
    'tr': 'Turkish', 'uk': 'Ukrainian', 'ur': 'Urdu', 'uz': 'Uzbek',
    'vi': 'Vietnamese', 'xh': 'Xhosa', 'yi': 'Yiddish', 'yo': 'Yoruba',
    'zu': 'Zulu'
}


def speak_text(text):
    """
    Speaks the given text loudly and enthusiastically using pyttsx3.
    """
    try:
        print(f"Jarvis: {text}")
        tts_engine.say(text)
        tts_engine.runAndWait()
    except Exception as e:
        print(f"Jarvis: Error in TTS: {str(e)}")


def listen_command(max_retries=3):
    """
    Listens to the microphone and returns the recognized text, with retries.
    """
    retries = 0
    while retries < max_retries:
        try:
            with microphone as source:
                recognizer.adjust_for_ambient_noise(source, duration=1) 
                print("Listening...")
                audio = recognizer.listen(source, timeout=5, phrase_time_limit=10)
            
            command = recognizer.recognize_google(audio).lower()
            print(f"You said: {command}")
            return command
        except sr.WaitTimeoutError:
            speak_text("No command detected. Please try again!")
            retries += 1
        except sr.UnknownValueError:
            speak_text("Sorry, I didn't catch that. Please repeat!")
            retries += 1
        except sr.RequestError as e:
            speak_text(f"Speech recognition error: {str(e)}. Try again!")
            retries += 1
        except Exception as e:
            speak_text(f"Error listening: {str(e)}. Please try again!")
            retries += 1
    speak_text("Too many failed attempts. Please type your command.")
    command = input("Type your command: ").lower()
    print(f"You typed: {command}")
    return command


def get_installed_apps():
    """
    Queries Windows registry for installed applications.
    """
    try:
        pythoncom.CoInitialize() 
        shell = win32com.client.Dispatch("WScript.Shell")
        apps = {}
        reg_paths = [
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
            r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
        ]
        for path in reg_paths:
            try:
                reg = win32com.client.Dispatch("WScript.Registry")
                key = reg.OpenKey(win32com.client.constants.HKEY_LOCAL_MACHINE, path)
                for i in range(reg.QueryInfoKey(key)[0]):
                    subkey_name = reg.EnumKey(key, i)
                    subkey = reg.OpenKey(key, subkey_name)
                    try:
                        app_name = reg.QueryValueEx(subkey, "DisplayName")[0]
                        app_path = reg.QueryValueEx(subkey, "InstallLocation")[0]
                        if app_name and app_path:
                            apps[app_name.lower()] = app_path
                    except:
                        continue
                    reg.CloseKey(subkey)
                reg.CloseKey(key)
            except:
                continue
        pythoncom.CoUninitialize()
        return apps
    except Exception as e:
        speak_text(f"Error discovering apps: {str(e)}!")
        return {}


def get_driver(browser="brave", headless=True):
    """
    Returns a Selenium WebDriver for Brave or Chrome using Service.
    """
    try:
        options = Options()
        if browser.lower() == "brave":
            options.binary_location = BRAVE_PATH
        elif browser.lower() == "chrome":
            options.binary_location = CHROME_PATH
        else:
            speak_text(f"Unsupported browser: {browser}!")
            return None
        if headless:
            options.add_argument('--headless=new')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-gpu')
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        return driver
    except Exception as e:
        speak_text(f"Error initializing {browser}: {str(e)}!")
        return None


def close_browser():
    """
    Closes the current browser if open.
    """
    global context
    if context['driver']:
        try:
            context['driver'].quit()
            speak_text("Closing browser!")
        except Exception as e:
            speak_text(f"Error closing browser: {str(e)}!")
        finally:
            context['driver'] = None
            context['browser'] = None
            context['site'] = None
    else:
        speak_text("No browser is open!")


def brave_search(query, headless=True):
    """
    Performs a Brave Search with Copilot-like intelligence and returns summarized results.
    """
    try:
        
        if "weather" in query and " in " in query:
            city = query.split(" in ")[1].strip()
            query = f"current weather in {city}"
        elif query.startswith("weather "):
            city = query.split("weather ")[1].strip()
            query = f"current weather in {city}"
        
        driver = get_driver(browser="brave", headless=headless)
        if not driver:
            return
        try:
            driver.get("https://search.brave.com/")
            wait = WebDriverWait(driver, 10)
            search_box = wait.until(EC.presence_of_element_located((By.NAME, "q")))
            search_box.clear()
            search_box.send_keys(query)
            search_box.send_keys(Keys.RETURN)
            time.sleep(3)
            results = driver.find_elements(By.CSS_SELECTOR, "div.card")
            summary = f"Results for '{query}': "
            for result in results[:3]:
                try:
                    title = result.find_element(By.CSS_SELECTOR, "a > div > div:nth-child(1) > span").text
                    snippet = result.find_element(By.CSS_SELECTOR, "div.snippet").text
                    summary += f"{title}: {snippet[:100]}... "
                except NoSuchElementException:
                    continue
            if summary == f"Results for '{query}': ":
                driver.get(f"https://www.google.com/search?q={query}")
                time.sleep(3)
                results = driver.find_elements(By.CSS_SELECTOR, "div.g")
                for result in results[:3]:
                    try:
                        title = result.find_element(By.CSS_SELECTOR, "h3").text
                        snippet = result.find_element(By.CSS_SELECTOR, "div.VwiC3b").text
                        summary += f"{title}: {snippet[:100]}... "
                    except NoSuchElementException:
                        continue
                if summary == f"Results for '{query}': ":
                    summary = f"No clear results found for '{query}'. Try a different query!"
            speak_text(summary)
        except TimeoutException:
            speak_text("Timeout during search. Please try again!")
        except Exception as e:
            speak_text(f"Error in search: {str(e)}!")
        finally:
            if headless:
                driver.quit()
            else:
                global context
                context['driver'] = driver
                context['browser'] = 'brave'
    except Exception as e:
        speak_text(f"Error processing query: {str(e)}. Try again!")


def browser_site_action(browser, site, query=None, action="search"):
    """
    Opens Brave/Chrome, navigates to site, searches, and performs action (search/play).
    """
    global context
    if not query and action in ["search", "play"]:
        speak_text(f"What do you want to {action} on {site}?")
        query = listen_command()
        if not query:
            return
    driver = context['driver'] if context['driver'] and context['browser'] == browser.lower() else get_driver(browser=browser, headless=(action != "play"))
    if not driver:
        return
    context['driver'] = driver
    context['browser'] = browser.lower()
    context['site'] = site.lower()
    try:
        if site.lower() == "youtube":
            url = f"https://www.youtube.com/results?search_query={query}" if query else "https://www.youtube.com"
            driver.get(url)
            wait = WebDriverWait(driver, 10)
            if action == "play" and query:
                first_video = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#video-title')))
                first_video.click()
                time.sleep(3)
                speak_text(f"Playing video for '{query}' on YouTube in {browser}!")
            elif action == "search" and query:
                results = driver.find_elements(By.CSS_SELECTOR, '#video-title')
                summary = f"YouTube search results for '{query}' in {browser}: "
                for result in results[:3]:
                    summary += f"{result.text} "
                speak_text(summary)
                if context['driver']:
                    context['driver'].quit()
                    context['driver'] = None
                    context['browser'] = None
                    context['site'] = None
            else:
                speak_text(f"Opening YouTube in {browser}!")
        elif site.lower() == "vimeo":
            url = f"https://vimeo.com/search?q={query}" if query else "https://vimeo.com"
            driver.get(url)
            wait = WebDriverWait(driver, 10)
            if action == "play" and query:
                first_video = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[href*="/"]')))
                first_video.click()
                time.sleep(3)
                try:
                    play_button = driver.find_element(By.CSS_SELECTOR, '.player_wrapper .vp-video-wrapper button')
                    play_button.click()
                except NoSuchElementException:
                    pass
                speak_text(f"Playing video for '{query}' on Vimeo in {browser}!")
            elif action == "search" and query:
                results = driver.find_elements(By.CSS_SELECTOR, 'a[href*="/"]')
                summary = f"Vimeo search results for '{query}' in {browser}: "
                for result in results[:3]:
                    summary += f"{result.text} "
                speak_text(summary)
                if context['driver']:
                    context['driver'].quit()
                    context['driver'] = None
                    context['browser'] = None
                    context['site'] = None
            else:
                speak_text(f"Opening Vimeo in {browser}!")
        elif site.lower() == "dailymotion":
            url = f"https://www.dailymotion.com/search/video/{query}" if query else "https://www.dailymotion.com"
            driver.get(url)
            wait = WebDriverWait(driver, 10)
            if action == "play" and query:
                first_video = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.dm-search-result a')))
                first_video.click()
                time.sleep(3)
                try:
                    play_button = driver.find_element(By.CSS_SELECTOR, 'button[aria-label="Play"]')
                    play_button.click()
                except NoSuchElementException:
                    pass
                speak_text(f"Playing video for '{query}' on Dailymotion in {browser}!")
            elif action == "search" and query:
                results = driver.find_elements(By.CSS_SELECTOR, '.dm-search-result a')
                summary = f"Dailymotion search results for '{query}' in {browser}: "
                for result in results[:3]:
                    summary += f"{result.text} "
                speak_text(summary)
                if context['driver']:
                    context['driver'].quit()
                    context['driver'] = None
                    context['browser'] = None
                    context['site'] = None
            else:
                speak_text(f"Opening Dailymotion in {browser}!")
        else:
            speak_text(f"Unsupported site: {site}!")
            if context['driver']:
                context['driver'].quit()
                context['driver'] = None
                context['browser'] = None
                context['site'] = None
    except TimeoutException:
        speak_text(f"No results found on {site}. Try again!")
        if context['driver']:
            context['driver'].quit()
            context['driver'] = None
            context['browser'] = None
            context['site'] = None
    except Exception as e:
        speak_text(f"Error on {site}: {str(e)}!")
        if context['driver']:
            context['driver'].quit()
            context['driver'] = None
            context['browser'] = None
            context['site'] = None


def open_application(app_name):
    """
    Opens any installed Windows application with fuzzy matching and registry lookup.
    """
    try:
        
        app_map = {
            'brave': BRAVE_PATH,
            'chrome': CHROME_PATH,
            'notepad': 'notepad.exe',
            'calculator': 'calc.exe',
            'file explorer': 'explorer.exe',
            'file manager': 'explorer.exe',
            'vlc': r"C:\Program Files\VideoLAN\VLC\vlc.exe",  # Update path
            'spotify': r"C:\Users\<YourUsername>\AppData\Roaming\Spotify\Spotify.exe",  # Update path
            'word': r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",  # Update Office version
            'excel': r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",  # Update Office version
            'powerpoint': r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",  # Update Office version
            'microsoft office': r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",  # Default to Word
            'office': r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"  # Default to Word
        }
        app_name = app_name.lower().strip()
        if not app_name:
            speak_text("Which application would you like to open?")
            app_name = listen_command()
            if not app_name:
                return
        
        
        app_names = list(app_map.keys())
        match, score = process.extractOne(app_name, app_names)
        if score > 80:
            app_path = app_map[match]
            if os.name == 'nt':
                os.startfile(app_path)
                speak_text(f"Opening {match.replace('file explorer', 'File Explorer').replace('file manager', 'File Explorer')}!")
            else:
                speak_text(f"Application '{match}' not supported on non-Windows!")
            return
        
        # Registry lookup for other installed applications
        installed_apps = get_installed_apps()
        if installed_apps:
            app_names = list(installed_apps.keys())
            match, score = process.extractOne(app_name, app_names)
            if score > 80:
                app_path = installed_apps[match]
                executable = next((f for f in os.listdir(app_path) if f.endswith('.exe')), None)
                if executable:
                    os.startfile(os.path.join(app_path, executable))
                    speak_text(f"Opening {match}!")
                    return
        
        # System path search as fallback
        if os.name == 'nt':
            result = subprocess.run(['where', app_name], capture_output=True, text=True)
            if result.returncode == 0:
                app_path = result.stdout.split('\n')[0].strip()
                os.startfile(app_path)
                speak_text(f"Opening {app_name}!")
                return
            else:
                speak_text(f"Application '{app_name}' not found. Please check the name or install it!")
        else:
            try:
                subprocess.run(['open', f'/Applications/{app_name}.app'], check=True)
                speak_text(f"Opening {app_name}!")
            except subprocess.CalledProcessError:
                speak_text(f"Application '{app_name}' not found on non-Windows!")
    except Exception as e:
        speak_text(f"Error opening application: {str(e)}!")

# Function to create a new file
def create_file(file_name):
    """
    Creates a new empty file in the current directory.
    """
    try:
        with open(file_name, 'w') as f:
            pass
        speak_text(f"Created file '{file_name}'!")
    except Exception as e:
        speak_text(f"Error creating file: {str(e)}!")

# Function to translate text using local translate library
def translate_text(text, target_language):
    """
    Translates text into a specified minor language using translate library.
    """
    try:
        if target_language.lower() not in ALLOWED_LANGUAGES:
            return f"Language '{target_language}' not supported or is a major language!"
        
        lang_code = target_language.lower()
        translator = Translator(to_lang=lang_code)
        translation = translator.translate(text)
        lang_name = ALLOWED_LANGUAGES[lang_code]
        return f"Translation to {lang_name}: {translation}"
    except Exception as e:
        return f"Error translating text: {str(e)}!"

# Function to get today's date
def get_today_date():
    """
    Returns today's date in a spoken format.
    """
    today = datetime.now().strftime("%B %d, %Y")
    return f"Today is {today}!"

# Main function to interact with the user via voice
def main():
    """
    Main function to run the Jarvis virtual assistant with voice input/output.
    """
    speak_text("Hi, I'm Jarvis! Try saying 'open File Explorer', 'open Spotify', 'play YouTube funny cats', 'weather Nagpur', 'today's date', 'translate hello to Swahili', 'close browser', or 'exit'!")
    
    while True:
        try:
            command = listen_command()
            
            if not command:
                continue
                
            # Parse command for intent using simple keyword matching
            intent = "general"
            if "play" in command:
                intent = "play"
            elif "open" in command:
                intent = "open"
            elif "weather" in command or "what is" in command and "weather" in command:
                intent = "weather"
            elif "translate" in command:
                intent = "translate"
            elif "create" in command:
                intent = "create"
            elif "close" in command:
                intent = "close"
            elif "date" in command or "day" in command:
                intent = "date"
            
            if 'exit' in command:
                if context['driver']:
                    context['driver'].quit()
                    context['driver'] = None
                    context['browser'] = None
                    context['site'] = None
                speak_text("Goodbye!")
                break
            elif 'help' in command:
                help_text = "Try saying: open app (e.g., File Explorer, Spotify), create file, translate text to language, play YouTube query, weather city, today's date, close browser!"
                speak_text(help_text)
            elif intent == "close":
                close_browser()
            elif intent == "open":
                app_name = command.split('open ')[1].strip() if 'open ' in command else ""
                if 'search' in command or 'play' in command:
                    parts = command.split('open ')[1].split(' ', 2)
                    if len(parts) >= 3:
                        browser = parts[0].strip()
                        action = parts[1].strip()
                        site_query = parts[2].strip().split(' on ', 1)
                        if len(site_query) == 2:
                            site, query = site_query[0].strip(), site_query[1].strip()
                            browser_site_action(browser, site, query, action)
                        else:
                            site = parts[2].strip()
                            browser_site_action(browser, site, action=action)
                    else:
                        speak_text("Please say: open browser search/play site query!")
                else:
                    open_application(app_name)
            elif intent == "create":
                file_name = command.split('create ')[1].strip()
                create_file(file_name)
            elif intent == "translate":
                parts = command.split('translate ')[1].split(' to ')
                if len(parts) == 2:
                    text, lang = parts[0].strip(), parts[1].strip()
                    translation = translate_text(text, lang)
                    speak_text(translation)
                else:
                    speak_text("Please say: translate text to language_code!")
            elif intent == "weather":
                if ' in ' in command:
                    city = command.split(' in ')[1].strip()
                elif command.startswith('weather '):
                    city = command.split('weather ')[1].strip()
                else:
                    speak_text("Please specify a city for the weather!")
                    city = listen_command()
                    if not city:
                        continue
                brave_search(f"current weather in {city}", headless=True)
            elif intent == "date":
                date_info = get_today_date()
                speak_text(date_info)
            elif intent == "play" and context['site']:
                query = command.split('play ')[1].strip()
                browser_site_action(context['browser'], context['site'], query, action="play")
            elif intent == "play":
                query = command.split('play ')[1].strip()
                browser_site_action("brave", "youtube", query, action="play")
            else:
                speak_text("Could you clarify what you want to do? Try 'open app', 'play YouTube query', or 'weather city'!")
                command = listen_command()
                if command:
                    brave_search(command, headless=True)
                
            # Small delay between interactions
            time.sleep(0.5)
            
        except KeyboardInterrupt:
            if context['driver']:
                context['driver'].quit()
            speak_text("Interrupted. Goodbye!")
            break
        except Exception as e:
            speak_text(f"An error occurred: {str(e)}. Please try again!")

if __name__ == "__main__":
    # Configure pyautogui for safety
    pyautogui.FAILSAFE = True
    pyautogui.PAUSE = 0.5  # Add a small pause between actions
    
    main()