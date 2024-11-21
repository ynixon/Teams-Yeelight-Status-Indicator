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
from whatsapp_api_client_python import API

# Disable TensorFlow Lite warnings
os.environ["TF_CPP_MIN_LOG_LEVEL"] = "3"


# Function to extract MFA number
def extract_mfa_number(driver, retries=3):
    for attempt in range(retries):
        try:
            print("Waiting for MFA number to appear...")
            mfa_element = WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.ID, "idRichContext_DisplaySign"))
            )
            mfa_number = mfa_element.text.strip()
            print(f"Extracted MFA number: {mfa_number}")
            return mfa_number
        except Exception as e:
            print(f"Attempt {attempt + 1} failed: {e}")
            time.sleep(2 ** attempt)  # Exponential backoff
    print("Failed to extract MFA number after multiple attempts.")
    driver.save_screenshot("mfa_error_page.png")
    print("Screenshot saved as mfa_error_page.png.")
    print("Page source at the time of failure:")
    print(driver.page_source)
    return None


# Function to send MFA number via WhatsApp
def send_mfa_via_whatsapp(mfa_number, settings):
    try:
        green_api_instance = settings.get("GREEN_API_INSTANCE")
        green_api_token = settings.get("GREEN_API_TOKEN")
        whatsapp_number = settings.get("WHATSAPP_NUMBER")

        if not (green_api_instance and green_api_token and whatsapp_number):
            print(
                "Error: Green API credentials or WhatsApp number missing in configuration."
            )
            return

        greenAPI = API.GreenAPI(green_api_instance, green_api_token)
        # Convert the MFA number to a string
        message = f"Approve for Walert using this code: {str(mfa_number)}"
        response = greenAPI.sending.sendMessage(whatsapp_number, message)
        if response:
            print("MFA number sent successfully via WhatsApp.")
        else:
            print("Failed to send MFA number via WhatsApp.")
    except Exception as e:
        print(f"Error sending MFA number via WhatsApp: {e}")


# Function to load the configuration from the YAML file
def load_config(file_path):
    # Search paths in order
    search_paths = [
        os.path.abspath(file_path),  # Provided path (e.g., "config.yaml")
        os.path.join(os.getcwd(), "config.yaml"),  # Current working directory
        os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "config.yaml"
        ),  # Same folder as the script
    ]

    for path in search_paths:
        if os.path.exists(path):
            print(f"Loading configuration from: {path}")
            try:
                with open(path, "r") as file:
                    config = yaml.safe_load(file)
                    return config.get("settings", {}), config.get("status_mappings", {})
            except Exception as e:
                print(f"Error loading configuration file: {e}")
                sys.exit(1)

    print(
        "Error: Configuration file 'config.yaml' not found in any of the expected paths."
    )
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
            time.sleep(2**attempt)  # Exponential backoff
    print("Failed to reconnect to the bulb after multiple attempts.")
    sys.exit(1)


# Function to handle Teams login
def login_to_teams(driver, email, password):
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
        driver.save_screenshot("email_submitted_page.png")
        print("Screenshot saved as email_submitted_page.png.")

        # Check for the password field
        print("Checking for password field...")
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "i0118"))
        )

        print("Filling in the password...")
        # Re-locate the password field after page transition
        password_field = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "i0118"))
        )
        password_field.clear()  # Clear any pre-filled data
        password_field.send_keys(password)

        print("Submitting the password...")
        # Re-locate the submit button after password entry
        password_submit = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "idSIButton9"))
        )
        password_submit.click()
        driver.save_screenshot("password_submitted_page.png")
        print("Screenshot saved as password_submitted_page.png.")
        print("Password submitted successfully. Waiting for MFA screen...")

    except Exception as e:
        print(f"Error during Teams login: {e}")
        # Capture the page source for debugging
        print("Page source at the time of failure:")
        print(driver.page_source)
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

    try:
        # Load configuration
        print(f"Loading configuration from: {args.config}")
        settings, status_mappings = load_config(args.config)

        # Validate settings and extract necessary parameters
        email, bulb_ip = validate_settings(settings)
        password = settings.get("password", "")

        if not password:
            print("Error: Password is missing in the configuration.")
            sys.exit(1)

        # Initialize the Yeelight bulb
        print(f"Attempting to connect to the Yeelight bulb at IP: {bulb_ip}")
        bulb = initialize_bulb(bulb_ip)

        # Set up ChromeDriver using webdriver-manager
        options = webdriver.ChromeOptions()
        options.add_argument("--disable-gpu")
        options.add_argument("--ignore-certificate-errors")
        options.add_argument("--allow-insecure-localhost")
        options.add_argument("--log-level=3")
        options.add_argument("--ignore-ssl-errors")

        options.add_argument("--headless")  # Run Chrome in headless mode
        options.add_argument("--no-sandbox")  # Required for running in headless environments
        options.add_argument("--disable-dev-shm-usage")  # Avoid shared memory issues
        options.add_argument("--window-size=1920,1080")  # Optional: Specify window size for screenshots

        driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()), options=options
        )

        # Open Microsoft Teams Web
        driver.get("https://teams.microsoft.com/")
        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, "//body"))
        )
        driver.save_screenshot("teams_login_page.png")
        print("Screenshot saved as teams_login_page.png.")
        print("Microsoft Teams page loaded.")

        # Login to Teams
        print(f"Logging in with email: {email}")
        login_to_teams(driver, email, password)

        # Handle MFA
        mfa_number = extract_mfa_number(driver)
        if mfa_number:
            send_mfa_via_whatsapp(mfa_number, settings)

        # Synchronize Teams status with the Yeelight bulb
        try:
            while True:
                status = get_teams_status(driver, status_mappings)  # Only retrieve status
                print("Teams Status:", status)
                update_bulb_color(bulb, status, bulb_ip)
                time.sleep(15)  # Check every 15 seconds
        except KeyboardInterrupt:
            print("Exiting gracefully due to keyboard interrupt.")
        except Exception as e:
            print(f"Unexpected error during status sync: {e}")
    except KeyboardInterrupt:
        print("Exiting gracefully due to keyboard interrupt.")
    except Exception as e:
        print(f"Unexpected error: {e}")
    finally:
        try:
            driver.quit()
        except Exception as e:
            print(f"Error closing the WebDriver: {e}")


if __name__ == "__main__":
    main()
