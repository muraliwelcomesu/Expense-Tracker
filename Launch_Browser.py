#! python3
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.alert import Alert
 
import time
import User
import webbrowser,pyperclip
from tkinter import messagebox
from selenium.webdriver.support.wait import WebDriverWait
 

class OpenURL():
 
    def launch_url():
        url = 'give urlname'
        webbrowser.open(url)
        
      
if __name__ == '__main__':
    o = OpenURL()
    o.launch_url()
 