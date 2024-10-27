import os
import sys
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchWindowException
from yeelight import Bulb

# Function to determine the path to ChromeDriver
def get_chromedriver_path():
    # Priority 1: Command line argument
    if len(sys.argv) > 1:
        driver_path = sys.argv[1]
        print(f"Using ChromeDriver path from command line argument: {driver_path}")
    # Priority 2: Environment variable
    elif "CHROMEDRIVER_PATH" in os.environ:
        driver_path = os.environ["CHROMEDRIVER_PATH"]
        print(f"Using ChromeDriver path from environment variable: {driver_path}")
    # Priority 3: Default to chromedriver in the current directory
    else:
        driver_path = os.path.join(os.getcwd(), "chromedriver.exe")
        print(f"Using default ChromeDriver path: {driver_path}")
    
    # Verify that the path exists
    if not os.path.exists(driver_path):
        print(
            "ChromeDriver not found at specified path.\n"
            "Usage:\n"
            "  python teams_yeelight_status.py <path_to_chromedriver> <bulb_ip>\n"
            "or set the environment variable CHROMEDRIVER_PATH and BULB_IP.\n"
        )
        sys.exit(1)
    
    return driver_path

# Function to determine the IP of the Yeelight bulb
def get_bulb_ip():
    if len(sys.argv) > 2:
        bulb_ip = sys.argv[2]
        print(f"Using bulb IP from command line argument: {bulb_ip}")
    elif "BULB_IP" in os.environ:
        bulb_ip = os.environ["BULB_IP"]
        print(f"Using bulb IP from environment variable: {bulb_ip}")
    else:
        bulb_ip = "192.168.1.100"  # Default IP if not specified
        print(f"Using default bulb IP: {bulb_ip}")
    
    return bulb_ip

# Initialize ChromeDriver path and bulb IP
driver_path = get_chromedriver_path()
bulb_ip = get_bulb_ip()

# Initialize the bulb
bulb = Bulb(bulb_ip, effect="smooth", duration=500, auto_on=True)

# Set up ChromeDriver service
service = Service(driver_path)

# Set up Chrome options to disable GPU acceleration and run headless if desired
options = webdriver.ChromeOptions()
options.add_argument("--disable-gpu")

# Initialize the WebDriver
driver = webdriver.Chrome(service=service, options=options)

# Open Microsoft Teams Web
driver.get("https://teams.microsoft.com/")
print("Opened Microsoft Teams... waiting for page to load")

# Wait for Teams page to fully load
time.sleep(10)

# Function to get Teams status
def get_teams_status():
    try:
        print("Attempting to locate the status element...")
        # Locate the button with the aria-label containing "status"
        status_button = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, "//*[contains(@aria-label, 'status')]"))
        )
        
        # Extract the status from the aria-label
        aria_label = status_button.get_attribute("aria-label").lower()
        print(f"Aria-label found: {aria_label}")

        # Determine the status based on the aria-label content
        if "available" in aria_label:
            return "Available"
        elif "busy" in aria_label:
            return "Busy"
        elif "away" in aria_label:
            return "Away"
        else:
            return "Unknown"
    except NoSuchWindowException:
        print("Browser window was closed. Exiting the application.")
        sys.exit(0)
    except Exception as e:
        print("Error retrieving status:", e)
        return "Unknown"

# Function to update bulb color based on Teams status
def update_bulb_color(status):
    try:
        if status == "Busy":
            bulb.set_rgb(255, 0, 0)  # Red for busy
        elif status == "Available":
            bulb.set_rgb(0, 255, 0)  # Green for available
        elif status == "Away":
            bulb.set_rgb(255, 255, 0)  # Yellow for away
        else:
            bulb.set_rgb(128, 128, 128)  # Gray for unknown status
        print(f"Updated bulb color for status: {status}")
    except Exception as e:
        print("Failed to update Yeelight bulb color:", e)
        
        # Attempt to reconnect
        try:
            print("Attempting to reconnect to the bulb...")
            bulb.__init__(bulb_ip, effect="smooth", duration=500, auto_on=True)
            # Test connection with a simple command
            bulb.get_properties()
            print("Reconnection successful.")
        except Exception as reconnect_error:
            print("Reconnection failed:", reconnect_error)

# Main loop to check status periodically
try:
    while True:
        status = get_teams_status()
        print("Teams Status:", status)
        update_bulb_color(status)
        time.sleep(15)  # Check every 15 seconds
except KeyboardInterrupt:
    print("Exiting...")
finally:
    driver.quit()
