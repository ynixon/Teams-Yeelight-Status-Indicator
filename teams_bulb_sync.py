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
from colorama import init, Fore, Style
init(autoreset=True)

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
            print(f"Extracted MFA number: \033[97m{mfa_number}\033[0m")
            return mfa_number
        except Exception as e:
            print(f"Attempt {attempt + 1} failed: {e}")
            max_delay = 60  # Set maximum delay in seconds
            delay = min(2 ** attempt, max_delay)
            time.sleep(delay)  # Exponential backoff with maximum delay
    print("Failed to extract MFA number after multiple attempts.")
    driver.save_screenshot("mfa_error_page.png")
    print("Screenshot saved as mfa_error_page.png.")
    print("Page source at the time of failure:")
    print(driver.page_source)
    return None


def create_driver(headless=True):
    options = webdriver.ChromeOptions()
    if headless:
        options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--ignore-certificate-errors")
    options.add_argument("--allow-insecure-localhost")
    options.add_argument("--log-level=3")
    options.add_argument("--ignore-ssl-errors")
    options.add_argument("--no-sandbox")  # Required for running in headless environments
    options.add_argument("--disable-dev-shm-usage")  # Avoid shared memory issues
    options.add_argument("--window-size=1920,1080")  # Optional: Specify window size for screenshots
    options.add_argument("--disable-webgl")
    return webdriver.Chrome(
        service=Service(ChromeDriverManager().install()), options=options
    )


# Function to restart the driver and navigate to Microsoft Teams
def restart_driver(settings, old_driver=None):
    try:
        print(Fore.CYAN + "Restarting WebDriver..." + Style.RESET_ALL)
        headless_mode = settings.get("headless", True)
        cookies = []

        # Save cookies if an old driver exists
        if old_driver:
            try:
                cookies = old_driver.get_cookies()
                print(Fore.GREEN + f"Saved {len(cookies)} cookies from the old driver." + Style.RESET_ALL)
            except Exception as e:
                print(Fore.YELLOW + f"Error retrieving cookies: {e}" + Style.RESET_ALL)
            finally:
                try:
                    old_driver.quit()
                    print(Fore.GREEN + "Old WebDriver session closed successfully." + Style.RESET_ALL)
                except Exception as e:
                    print(Fore.RED + f"Error closing old WebDriver: {e}" + Style.RESET_ALL)

        # Create a new driver
        driver = create_driver(headless=headless_mode)
        print(Fore.CYAN + "Navigating to Microsoft Teams..." + Style.RESET_ALL)
        driver.get("https://teams.microsoft.com/")

        # Wait for the page to load
        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, "//title[contains(text(), 'Microsoft Teams')]"))
        )
        print(Fore.GREEN + "Teams page loaded successfully." + Style.RESET_ALL)

        # Restore cookies if any were saved
        if cookies:
            try:
                for cookie in cookies:
                    driver.add_cookie(cookie)
                driver.refresh()  # Refresh after restoring cookies
                print(Fore.GREEN + "Cookies restored successfully." + Style.RESET_ALL)
            except Exception as e:
                print(Fore.YELLOW + f"Error restoring cookies: {e}" + Style.RESET_ALL)

        return driver
    except Exception as e:
        print(Fore.RED + f"Error restarting WebDriver: {e}" + Style.RESET_ALL)
        return None


# Update keep_session_alive to use restart_driver
def keep_session_alive(driver, settings):
    try:
        # Check if driver is active
        driver.execute_script("return window.location.href;")
    except WebDriverException:
        print("WebDriver session disconnected. Restarting...")
        return restart_driver(settings, driver)
    return driver


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
    print("Loaded settings:")
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
def initialize_bulb(bulb_ip, retries=3):
    for attempt in range(retries):
        try:
            print(f"Attempting to connect to Yeelight bulb at IP: {bulb_ip} (Attempt {attempt + 1})")
            bulb = Bulb(bulb_ip, effect="smooth", duration=500, auto_on=True)
            bulb.get_properties()  # Test connection
            print(f"Connected to Yeelight bulb at {bulb_ip}")
            return bulb
        except Exception as e:
            print(Fore.YELLOW + f"Error initializing Yeelight bulb: {e}" + Style.RESET_ALL)
            max_delay = 60  # Set maximum delay in seconds
            delay = min(2 ** attempt, max_delay)
            time.sleep(delay)  # Exponential backoff with maximum delay

    print(Fore.RED + "Failed to connect to the Yeelight bulb after multiple attempts. Exiting..." + Style.RESET_ALL)
    sys.exit(1)  # Exit if the bulb fails to connect


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
        print("Navigating to Microsoft Teams...")
        driver.get("https://teams.microsoft.com")
        WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.ID, "i0116")))

        print("Teams page loaded successfully.")
        print(f"Logging in with email: {email}")
        email_field = driver.find_element(By.ID, "i0116")
        email_field.send_keys(email)
        driver.find_element(By.ID, "idSIButton9").click()
        print("Filling in the email address...")
        time.sleep(2)  # Wait for the next page to load

        print("Clicking the submit button...")
        WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.ID, "i0118")))
        driver.save_screenshot("email_submitted_page.png")
        print("Screenshot saved as email_submitted_page.png.")

        print("Checking for password field...")
        password_field = driver.find_element(By.ID, "i0118")
        password_field.send_keys(password)
        print("Filling in the password...")
        driver.find_element(By.ID, "idSIButton9").click()
        time.sleep(2)  # Wait for the next page to load

        # Add more steps if there are additional login steps (e.g., MFA)
    except Exception as e:
        print(f"Error during Teams login: {e}")
        driver.save_screenshot("login_error_page.png")
        print("Screenshot saved as login_error_page.png.")
        print("Page source at the time of failure:")
        print(driver.page_source)
        raise


# Function to get Teams status
def get_teams_status(driver, status_mappings):
    """Retrieve Teams status and map it to a predefined status."""
    try:
        print("Locating status button...")
        status_button = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//*[contains(@aria-label, 'status') and @role='button']")
            )
        )
        aria_label = status_button.get_attribute("aria-label").lower()

        # Extract the status from the aria-label
        status_key = aria_label.replace("your profile, status ", "").strip()

        print(f"Profile status: \033[97m{status_key}\033[0m")  # Debugging information

        # Match extracted status to predefined mappings
        for key, mapping in status_mappings.items():
            if key in status_key:
                return mapping  # Return the matching status mapping

        # Default mapping for unknown statuses
        return {"status": "Unknown", "color": "255,255,255"}  # Bright white
    except Exception as e:
        print(f"Error retrieving status: {e}")
        driver.save_screenshot("error_retrieving_status.png")
        print("Screenshot saved as error_retrieving_status.png.")
        print("Page source at the time of failure:")
        print(driver.page_source)
        return {"status": "Unknown", "color": "255,255,255"}  # Bright white fallback


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
        print(f"Updated bulb color for status: {console_color}{status}\033[0m")
    except Exception as e:
        print(f"Failed to update Yeelight bulb color: {e}")
        reconnect_bulb(bulb, bulb_ip)


# Main function refactored to avoid duplication
def main():
    driver = None
    bulb = None
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
        bulb = initialize_bulb(bulb_ip)

        password = settings.get("password", "")
        refresh_interval = settings.get("refresh_interval", 15)  # Default to 15 seconds
        if not isinstance(refresh_interval, int) or refresh_interval <= 0:
            print("Invalid refresh_interval in configuration. Defaulting to 15 seconds.")
            refresh_interval = 15

        if not password:
            print("Error: Password is missing in the configuration.")
            sys.exit(1)

        # Initialize the Yeelight bulb
        print(f"Attempting to connect to the Yeelight bulb at IP: {bulb_ip}")
        bulb = initialize_bulb(bulb_ip)

        # Start WebDriver
        driver = restart_driver(settings)
        if not driver:
            print("Failed to initialize WebDriver. Exiting.")
            sys.exit(1)

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
                driver = keep_session_alive(driver, settings)  # Ensure the session is active
                if driver:
                    # Check if Teams is still logged in
                    driver.execute_script("return window.location.href;")
                    status = get_teams_status(driver, status_mappings)  # Retrieve status
                    update_bulb_color(bulb, status, bulb_ip)  # Update bulb color
                time.sleep(refresh_interval)  # Wait before the next check
        except KeyboardInterrupt:
            print("Exiting gracefully due to keyboard interrupt.")
        except (WebDriverException, NoSuchWindowException) as e:
            print(f"Driver error caught: {e}. Restarting driver...")
            driver = restart_driver(settings)  # Reuse restart logic
        except Exception as e:
            print(f"Unexpected error during status sync: {e}")
    except KeyboardInterrupt:
        print("Exiting gracefully due to keyboard interrupt.")
    except Exception as e:
        print(f"Unexpected error: {e}")
    finally:
        if bulb:  # Only attempt to reset the bulb if it's not None
            try:
                # Reset bulb color to bright white
                print("Resetting bulb to bright \033[97mwhite\033[0m...")
                bulb.set_rgb(255, 255, 255)  # Bright white
                print("Bulb color reset to bright \033[97mwhite\033[0m.")
            except Exception as e:
                print(f"Error resetting bulb color: {e}")
        else:
            print("Skipping bulb reset as the bulb was not initialized.")

        if driver:
            try:
                driver.quit()
                print("WebDriver closed successfully.")
            except Exception as e:
                print(f"Error closing the WebDriver: {e}")


if __name__ == "__main__":
    main()
