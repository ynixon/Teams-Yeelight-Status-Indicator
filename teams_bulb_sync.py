import os
import sys
import time
import yaml
import argparse
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchWindowException, WebDriverException
from yeelight import Bulb
from webdriver_manager.chrome import ChromeDriverManager

# Disable TensorFlow Lite warnings
os.environ["TF_CPP_MIN_LOG_LEVEL"] = "3"


# Function to load the configuration from the YAML file
def load_config(file_path):
    try:
        with open(file_path, "r") as file:
            config = yaml.safe_load(file)
            return config.get("settings", {}), config.get("status_mappings", {})
    except Exception as e:
        print(f"Error loading configuration file: {e}")
        sys.exit(1)


# Function to validate paths and settings
def validate_settings(settings):
    email = settings.get("email", "")
    bulb_ip = settings.get("bulb_ip", "")

    # Debugging outputs
    print(f"Loaded settings:")
    print(f"  Email: {email}")
    print(f"  Bulb IP: {bulb_ip}")

    # Validate required settings
    if not email:
        print("Error: Email address is missing in the configuration.")
        sys.exit(1)
    if not bulb_ip:
        print("Error: Bulb IP is missing in the configuration.")
        sys.exit(1)
    return email, bulb_ip


# Function to initialize the Yeelight bulb
def initialize_bulb(bulb_ip):
    try:
        bulb = Bulb(bulb_ip, effect="smooth", duration=500, auto_on=True)
        bulb.get_properties()  # Test connection
        print(f"Connected to Yeelight bulb at {bulb_ip}")
        return bulb
    except Exception as e:
        print(f"Error initializing Yeelight bulb: {e}")
        sys.exit(1)


# Function to reconnect to the Yeelight bulb
def reconnect_bulb(bulb, bulb_ip, retries=3):
    for attempt in range(retries):
        try:
            print(f"Reconnecting to bulb at IP {bulb_ip}... Attempt {attempt + 1}")
            bulb.__init__(bulb_ip, effect="smooth", duration=500, auto_on=True)
            bulb.get_properties()  # Test connection
            print("Reconnection successful.")
            return bulb
        except Exception as e:
            print(f"Reconnection failed: {e}")
            time.sleep(2 ** attempt)  # Exponential backoff
    print("Failed to reconnect to the bulb after multiple attempts.")
    sys.exit(1)


# Function to handle Teams login
def login_to_teams(driver, email):
    try:
        print("Filling in the email address...")
        email_input = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located(
                (By.XPATH, "//input[@type='email' and @name='loginfmt']")
            )
        )
        email_input.send_keys(email)

        print("Clicking the submit button...")
        submit_button = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located(
                (By.XPATH, "//input[@type='submit' and @id='idSIButton9']")
            )
        )
        submit_button.click()
        print("Email submitted successfully.")
    except Exception as e:
        print(f"Error during Teams login: {e}")
        sys.exit(1)


# Function to get Teams status
def get_teams_status(driver, status_mappings):
    try:
        status_button = WebDriverWait(driver, 60).until(
            EC.presence_of_element_located(
                (By.XPATH, "//*[contains(@aria-label, 'status') and @role='button']")
            )
        )
        aria_label = status_button.get_attribute("aria-label").lower()
        print(f"{aria_label}")

        for key, mapping in status_mappings.items():
            if key in aria_label:
                return mapping  # Return the entire mapping dictionary
        return {"status": "Unknown", "color": "128,128,128"}  # Default mapping
    except Exception as e:
        print(f"Error retrieving status: {e}")
        return {"status": "Unknown", "color": "128,128,128"}


# Function to update the Yeelight bulb color based on Teams status
def update_bulb_color(bulb, status_mapping, bulb_ip):
    try:
        status = status_mapping.get("status", "Unknown")
        rgb_color = status_mapping.get("color", "128,128,128")  # Default to Gray
        r, g, b = map(int, rgb_color.split(","))  # Convert RGB string to integers

        # Update bulb color
        bulb.set_rgb(r, g, b)

        # Map RGB to ANSI escape sequence for console color
        console_color = f"\033[38;2;{r};{g};{b}m"

        # Print status in the corresponding color
        print(f"{console_color}Updated bulb color for status: {status}\033[0m")
    except Exception as e:
        print(f"Failed to update Yeelight bulb color: {e}")
        reconnect_bulb(bulb, bulb_ip)


# Main function
def main():
    parser = argparse.ArgumentParser(
        description="Sync Microsoft Teams status with a Yeelight bulb."
    )
    parser.add_argument(
        "--config",
        type=str,
        default="config.yaml",
        help="Path to the configuration file (default: config.yaml).",
    )
    args = parser.parse_args()

    # Load configuration
    print(f"Loading configuration from: {args.config}")
    settings, status_mappings = load_config(args.config)

    # Display loaded mappings for debugging
    print(f"Loaded status mappings:")
    for status, mapping in status_mappings.items():
        print(f"  {status}: {mapping}")

    # Validate settings and extract necessary parameters
    email, bulb_ip = validate_settings(settings)

    # Initialize the Yeelight bulb
    print(f"Attempting to connect to the Yeelight bulb at IP: {bulb_ip}")
    bulb = initialize_bulb(bulb_ip)

    # Set up ChromeDriver using webdriver-manager
    options = webdriver.ChromeOptions()
    options.add_argument("--disable-gpu")
    options.add_argument("--ignore-certificate-errors")
    options.add_argument("--allow-insecure-localhost")
    options.add_argument("--log-level=3")
    options.add_argument('--ignore-ssl-errors')
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)


    # Open Microsoft Teams Web
    driver.get("https://teams.microsoft.com/")
    WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.XPATH, "//body"))
    )
    print("Microsoft Teams page loaded.")

    # Login to Teams
    print(f"Logging in with email: {email}")
    login_to_teams(driver, email)

    # Synchronize Teams status with the Yeelight bulb
    try:
        while True:
            status = get_teams_status(driver, status_mappings)  # Only retrieve status
            print("Teams Status:", status)
            update_bulb_color(bulb, status, bulb_ip)
            time.sleep(15)  # Check every 15 seconds
    except KeyboardInterrupt:
        print("Exiting...")
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
