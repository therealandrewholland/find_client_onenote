from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException

import time
import os
import json

class getITGlueIDs():
    def __init__(self):
        chrome_options = Options()
        chrome_options.add_argument("--headless=new")
        self.driver = webdriver.Chrome(options=chrome_options)

        self.client_ids_data = {}
        self.progress = 0
        self.old_percent = 0
        with open('get_itglue_ids.js', 'r') as file:
            self.get_client_id_script = file.read()

    def __del__(self):
        self.driver.quit()

    def login(self, username=None, password=None, mfa_code=None):
        LOGIN_URL = "https://aunalytics.itglue.com/login"
        ALL_CLIENTS_URL = "https://aunalytics.itglue.com/organizations#sortBy=name:asc&filters=%5B%20%5D&partial="

        self.progress += 1
        self.driver.get(LOGIN_URL)
        for i in range(2):
            self.progress += 1
            time.sleep(1)
 
        self.driver.find_element(By.NAME, 'username').send_keys(username)
        self.driver.find_element(By.NAME, 'password').send_keys(password)
        self.driver.find_element(By.NAME, 'password').send_keys(Keys.ENTER)
        time.sleep(1)
        self.progress += 1
        
        try:
            self.driver.find_element(By.NAME, 'mfa')
        except NoSuchElementException:
            self.progress = 0
            return False, "Login credentials failed. Please try again."

        self.driver.find_element(By.NAME, 'mfa').send_keys(mfa_code)
        self.driver.find_element(By.NAME, 'mfa').send_keys(Keys.ENTER)

        for i in range(5):
            self.progress += 1
            time.sleep(1)
        # Wait for redirection or for login to fail

        try:
            self.driver.find_element(By.NAME, 'username')
            self.progress = 0
            return False, "MFA verification failed. Please try again."
        except NoSuchElementException:

            self.driver.get(ALL_CLIENTS_URL)
            for i in range(10):
                self.progress += 1
                time.sleep(1)
            # Wait for search results to load 
            
            return True, "Login successful."  

    def getClientIDs(self, client_list):
        client_list_json = json.dumps(client_list)
        last_height = self.driver.execute_script("return window.pageYOffset")
        total_height = self.driver.execute_script("return document.documentElement.scrollHeight")

        # Scrolls down in increments
        self.driver.execute_script("window.scrollBy(0, 500);")

        # Execute the client ID collection Javascript file
        full_script = self.get_client_id_script + f"\nreturn getClientData({json.dumps(client_list_json)});"
        new_clients = self.driver.execute_script(full_script)
        for client in new_clients:
            self.client_ids_data[client['name']] = client['id']

        # Calculate new scroll height and compare with last scroll height. Also calculates a progress % based on scroll height
        new_height = self.driver.execute_script("return window.pageYOffset")

        new_percent = round((new_height / total_height) * 81, 2)
        if new_percent != self.old_percent:
            self.progress += (new_percent - self.old_percent)
        self.old_percent = new_percent
        
        if new_height == last_height:
            self.progress = 100
        
        last_height = new_height


    def updateJSON(self, data):
        CLIENT_IDS_JSON = {}

        #If JSON already exists, only update with new clients (if any)
        if os.path.exists("it_glue_client_ids.json"):
            with open("it_glue_client_ids.json", "r") as file:
                CLIENT_IDS_JSON = json.load(file)

        CLIENT_IDS_JSON.update(data)

        with open("it_glue_client_ids.json", "w") as json_file:
            json.dump(CLIENT_IDS_JSON, json_file, indent=4)    
