import pandas as pd
import random
import schedule
import threading
import time
import csv
import os

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementNotInteractableException, ElementClickInterceptedException
from webdriver_manager.chrome import ChromeDriverManager

class Constants:
    PEOPLE_SEARCH_TIMEOUT = 30
    KEYWORD_SEARCH_TIMEOUT = 30
    CONNECT_MODAL_TIMEOUT = 30
    CUSTOM_MESSAGE_TIMEOUT = 30
    BUTTON_CLICK_DELAY = random.randint(5,10)
    LOG_FILE = "connection_log.csv"
    PROCESSED_PROFILES_FILE = "processed_profiles.csv"
    CONNECTION_INTERVAL = 10
    SCHEDULE_TIME = "11:54"
    CHROME_PROFILE_PATH = "C:\\Users\\Brackets\\AppData\\Local\\Google\\Chrome\\User Data\\Profile 8"
    CONNECTION_LIMIT = 1  # Max number of connections per run

def initialize_driver():
    options = Options()
    options.add_argument(f"user-data-dir={Constants.CHROME_PROFILE_PATH}")
    options.add_argument("profile-directory=Default")
    options.add_experimental_option("detach", True)
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    return driver

def login_to_linkedin(driver, email, password):
    driver.get("https://www.linkedin.com")
    wait = WebDriverWait(driver, 30)

    try:
        if "feed" in driver.current_url:
            print("Already logged in to LinkedIn")
            return

        sign_in_button = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Sign in")))
        sign_in_button.click()

        email_field = wait.until(EC.presence_of_element_located((By.ID, "username")))
        email_field.send_keys(email)

        password_field = wait.until(EC.presence_of_element_located((By.ID, "password")))
        password_field.send_keys(password)

        login_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[type='submit']")))
        login_button.click()

        wait.until(EC.url_contains("feed"))
        print("Successfully logged in to LinkedIn")

    except TimeoutException:
        print("Login process timed out.")
    except NoSuchElementException as e:
        print(f"Element not found during login: {e}")
    except Exception as e:
        print(f"Unexpected error during login: {e}")

def read_csv(file_path):
    df = pd.read_csv(file_path)
    return df

def save_processed_profiles(processed_profiles):
    processed_profiles.to_csv(Constants.PROCESSED_PROFILES_FILE, index=False)

def load_processed_profiles():
    if os.path.exists(Constants.PROCESSED_PROFILES_FILE):
        return pd.read_csv(Constants.PROCESSED_PROFILES_FILE)
    else:
        return pd.DataFrame(columns=["LinkedIn Profile URL"])

def log_result(profile_url, status):
    with open(Constants.LOG_FILE, mode='a', newline='') as file:
        writer = csv.writer(file)
        writer.writerow([profile_url, status])

def send_connection_request(driver, profile_data, processed_profiles):
    wait = WebDriverWait(driver, 30)
    connections_sent = 0

    for index, row in profile_data.iterrows():
        if connections_sent >= Constants.CONNECTION_LIMIT:
            print(f"Connection limit of {Constants.CONNECTION_LIMIT} reached.")
            break

        url = row['LinkedIn Profile URL']
        connection_note = row['Custom Note']

        if url in processed_profiles["LinkedIn Profile URL"].values:
            continue

        driver.get(url)
        try:
            # Random scrolling to simulate human behavior
            scroll_pause_time = random.uniform(1, 3)
            for i in range(random.randint(1, 3)):
                driver.execute_script(f"window.scrollTo(0, {random.randint(100, 600)});")
                time.sleep(scroll_pause_time)

            connect_buttons = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//span[text()='Connect']")))
            for button in connect_buttons:
                if "Connect" in button.get_attribute("innerText"):
                    button.click()
                    time.sleep(Constants.BUTTON_CLICK_DELAY)

                    try:
                        add_note_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[aria-label='Add a note']")))
                        add_note_button.click()
                        time.sleep(Constants.BUTTON_CLICK_DELAY)

                        custom_note_field = wait.until(EC.visibility_of_element_located((By.ID, "custom-message")))
                        custom_note_field.send_keys(connection_note)
                        time.sleep(Constants.BUTTON_CLICK_DELAY)

                        send_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[aria-label='Send invitation']")))
                        send_button.click()

                        print(f"Successfully sent connection request with note to {url}")
                        log_result(url, "Success with Note")

                    except (TimeoutException, ElementNotInteractableException, ElementClickInterceptedException):
                        try:
                            # If the "Add a note" button is not found, send the connection request without a note
                            send_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[aria-label='Send now']")))
                            send_button.click()

                            print(f"Successfully sent connection request without note to {url}")
                            log_result(url, "Success without Note")
                            time.sleep(Constants.BUTTON_CLICK_DELAY)

                        except (TimeoutException, ElementNotInteractableException, ElementClickInterceptedException) as e:
                            print(f"Skipping profile: {url} due to error: {str(e)}")
                            log_result(url, "Skipped")

                    # Mark the profile as processed
                    processed_profiles = processed_profiles._append({"LinkedIn Profile URL": url}, ignore_index=True)
                    save_processed_profiles(processed_profiles)
                    time.sleep(Constants.CONNECTION_INTERVAL)

                    connections_sent += 1  # Increment the counter
                    break

        except (TimeoutException, NoSuchElementException) as e:
            print(f"Skipping profile: {url} due to error: {str(e)}")
            log_result(url, "Skipped")


def scheduled_task(driver, profile_data):
    processed_profiles = load_processed_profiles()
    send_connection_request(driver, profile_data, processed_profiles)

def run_schedule(driver, profile_data):
    schedule.every().day.at(Constants.SCHEDULE_TIME).do(scheduled_task, driver, profile_data)

    while True:
        schedule.run_pending()
        time.sleep(1)

if __name__ == "__main__":
    driver = initialize_driver()

    email = "*****"
    password = "********"

    login_to_linkedin(driver, email, password)

    csv_file_path = "LinkedIn_Connections.csv"
    profile_data = read_csv(csv_file_path)

    # Start scheduling in a separate thread
    schedule_thread = threading.Thread(target=run_schedule, args=(driver, profile_data))
    schedule_thread.start()

    input("Press Enter to close the browser...")
    driver.quit()
