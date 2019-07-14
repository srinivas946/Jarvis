import speech_recognition as sr
import win32com.client as wincl
import webbrowser, os, wikipedia, sys
os.environ['PYGAME_HIDE_SUPPORT_PROMPT'] = "hide"
import pygame, time
from num2words import num2words
from datetime import datetime
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By


# ==========================
#   DIFFERENT ACTIONS
# ==========================
class Action:

    def __init__(self):
        self.chrome_driver = os.getcwd()+'chromedriver.exe'

    def opengoogle(self):
        webbrowser.open('www.google.com')

    def files(self):
        os.system('explorer')

    def search(self, text):
        webbrowser.open(text)

    def info(self, text, st):
        res = wikipedia.summary(text, sentences=1)
        print(res)
        st.speak(res)

    def anonymous(self, text, st):
        res = wikipedia.summary(text, sentences=1)
        print(res)
        st.speak(res)

    def excel(self):
        os.system('start excel')

    def word(self):
        os.system('start winword')

    def notepad(self):
        os.system('start notepad')

    def give_text(self, lt, source, st):
        while True:
            lstn = lt.listen(source)
            try:
                text = lt.recognize_google(lstn, language='en-in')
                if text.__contains__('exit') or text.__contains__('Exit'):
                    self.exit()
                break
            except: st.speak('Unable to Recognise. Please Enter Again')
        return text

    def youtube(self, st, lt, source):
        st.speak('You Tube is ready to open')
        driver = webdriver.Chrome(executable_path=self.chrome_driver)
        driver.get('https://www.youtube.com')
        search = driver.find_element_by_id('search')
        st.speak('Enter your search sir')
        text = self.give_text(lt, source, st)

        search.send_keys(text)
        driver.find_element_by_id('search-form').submit()
        try:
            myElem = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '//*[@id="video-title"]')))
            link_dict, obj_dict, count = {}, {}, 1
            for obj in myElem:
                link_dict[num2words(count)] = obj.text
                obj_dict[num2words(count)] = obj
                count += 1
            for k, v in link_dict.items():
                print(f'{k} - {v}')
            st.speak(f'Select Video in range one to {num2words(len(link_dict))}')
            text = num2words(self.give_text(lt, source, st))
            obj_dict[text].click()
        except TimeoutException: st.speak('Sorry ! Time Out')
        time.sleep(20)
        driver.quit()

    def maps(self, st, lt, source):
        driver = webdriver.Chrome(executable_path=os.getcwd()+'chromedriver.exe')
        driver.get('https://www.google.co.in/maps/')
        text = self.give_text(lt, source, st)
        print('Your Text : '+text)
        if text == 'nearby':
            search = driver.find_element_by_id('searchboxinput')
            search.send_keys(self.give_text(lt, source, st))
            driver.find_element_by_id('searchbox-searchbutton').click()
            try:
                myElem = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'section-result-title-container')))
                link_dict, obj_dict, count = {}, {}, 1
                for val in myElem:
                    link_dict[num2words(count)] = val.text
                    obj_dict[num2words(count)] = val
                    count += 1
                for k, v in link_dict.items():
                    print(f'{k} - {v}')
                text = input(f'\nSelect Result between one to {num2words(len(link_dict))} : ')
                obj_dict[text].click()
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="pane"]/div/div[1]/div/div/div[4]/div[1]/div/button'))).click()
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="sb_ifc51"]/input'))).send_keys(input('Enter Your Location : '))
                driver.find_element_by_xpath('//*[@id="directions-searchbox-0"]/button[1]').click()
            except TimeoutException:
                pass
        elif text == 'directions' or text == 'Directions':
            driver.find_element_by_xpath('//*[@id="searchbox-directions"]').click()
            st.speak('Enter Starting Point')
            text = self.give_text(lt, source, st)
            driver.find_element_by_xpath('//*[@id="sb_ifc51"]/input').send_keys(text)
            st.speak('Enter Destination Point')
            text = self.give_text(lt, source, st)
            driver.find_element_by_xpath('//*[@id="sb_ifc52"]/input').send_keys(text)
            driver.find_element_by_xpath('//*[@id="directions-searchbox-1"]/button[1]').click()
        else:
            print('Your Entered mode in unavailable')
        time.sleep(20)
        driver.quit()

    def write_file(self, st, lt, source):
        f_w = open(os.getcwd()+'\\lesson.txt', 'w')
        st.speak('Please Enter the text Sir ')
        lstn, text = lt.listen(source), ''
        while True:
            try:
                text = lt.recognize_google(lstn, language='en-in')
                break
            except: st.speak('Unable to Recognise. Please Enter Again')
        f_w.write(text)
        f_w.close()
        st.speak('File Created Successfully')
        st.speak('Shall I open the file sir')
        while True:
            try:
                res = lt.recognize_google(lt.listen(source))
                break
            except: st.speak('Unable to Recognise. Please say again')
        if res.__contains__('open') or res.__contains__('Open'):
            os.system('start '+os.getcwd()+'\\lesson.txt')
        elif res.__contains__('no') or res.__contains__('No'):
            pass


    def music(self):
        pygame.mixer.init()
        for path in os.listdir(os.getcwd()):
            if path.endswith('.mp3'):
                pygame.mixer.music.load(os.getcwd()+'\\'+path)
                pygame.mixer.music.play()
                while pygame.mixer.music.get_busy(): pygame.time.Clock().tick(1)

    def exit(self):
        sys.exit()

# =========================
#   LISTEN TO USER VOCALS
# =========================
class Listen:

    def __init__(self):
        self.st = Speak_Text()
        self.at = Action()

    def commands(self, r, source):
        while True:
            self.st.speak('Enter Your Command Sir')
            audio = r.listen(source)
            try:
                text = r.recognize_google(audio)
                print(f'You Entered {text}')
                if text.__contains__('google') or text.__contains__('Google'): self.at.opengoogle()
                elif text.__contains__('search') or text.__contains__('Search'): self.at.search(text)
                elif text.__contains__('file') or text.__contains__('File'): self.at.files()
                elif text.__contains__('about') or text.__contains__('about'): self.at.info(text, self.st)
                elif text.__contains__('exit') or text.__contains__('Exit'):
                    self.st.speak('Thank u Sir see u again bye bye')
                    self.at.exit()
                elif text.__contains__('excel') or text.__contains__('Excel'): self.at.excel()
                elif text.__contains__('Document') or text.__contains__('document'): self.at.word()
                elif text.__contains__('notepad') or text.__contains__('Notepad'): self.at.notepad()
                elif text.__contains__('music') or text.__contains__('Music'): self.at.music()
                elif text.__contains__('Create') or text.__contains__('create'): self.at.write_file(self.st, r, source)
                elif text.__contains__('youtube') or text.__contains__('Youtube') or text.__contains__('YouTube'): self.at.youtube(self.st, r, source)
                elif text.__contains__('maps') or text.__contains__('Maps'): self.at.maps(self.st, r, source)
                else: self.at.anonymous(text, self.st)
            except Exception: self.st.speak('Sorry! Unable to Recognize')


    def listen_vocal(self):
        r = sr.Recognizer()
        with sr.Microphone() as source:
            self.st.speak('Hello Sir  I am Bheem man how can i help u')
            audio = r.listen(source)
            try:
                text = r.recognize_google(audio, language='en-in')
                if text.__contains__('Bheem') or text.__contains__('bheem'):
                    current_time = datetime.now().hour
                    if current_time in range(1, 12):
                        self.st.speak('Hi Sir  Good Morning')
                    elif current_time in range(13, 16):
                        self.st.speak('Hi Sir Good Afternoon')
                    elif current_time in range(17, 19):
                        self.st.speak('Hi Sir Good Evening')
                    elif current_time in range(20, 24):
                        self.st.speak('Hi Sir Good Night')
                    self.commands(r, source)
                else:
                    self.commands(r, source)
            except: print('Unable to Recognize')


# ========================================
#   TEXT TO SPEECH USING PYWIN32 PACKAGE
# ========================================
class Speak_Text:

    def speak(self, text):
        speak = wincl.Dispatch('SAPI.SpVoice')
        speak.Speak(text)
        


if __name__ == '__main__':
    ls = Listen()
    ls.listen_vocal()