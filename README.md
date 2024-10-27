# Teams-Yeelight-Status-Indicator

**Teams-Yeelight-Status-Indicator** is a Python script that synchronizes your Microsoft Teams status with a Yeelight smart bulb. The bulb color reflects your Teams status, providing a quick visual indicator of your availability.

## Features

- Changes the Yeelight bulb color based on your Teams status:
  - **Green** for Available
  - **Red** for Busy
  - **Yellow** for Away
  - **Gray** for Unknown
- Automatically attempts to reconnect to the bulb if a connection is lost.
- Uses Selenium to fetch Teams status from the Teams web app.

## Requirements

- **Python 3.8+**
- **Yeelight smart bulb**
- **Selenium** and **ChromeDriver**
- **Microsoft Teams account**

## Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/ynixon/Teams-Yeelight-Status-Indicator.git
   cd Teams-Yeelight-Status-Indicator
   ```

2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Download ChromeDriver:**
   - Ensure your ChromeDriver version matches your installed Chrome browser version.
   - Place `chromedriver.exe` in the project directory or specify its location using either:
     - A command-line argument
     - The environment variable `CHROMEDRIVER_PATH`

4. **Set the Bulb IP Address:**
   - Specify the Yeelight bulb’s IP address using either:
     - A command-line argument
     - The environment variable `BULB_IP`

## Usage

Run the script with the following options:

1. **Basic Usage**:
   ```bash
   python teams_yeelight_status.py
   ```

2. **Using Command-Line Arguments**:
   ```bash
   python teams_yeelight_status.py <path_to_chromedriver> <bulb_ip>
   ```
   Example:
   ```bash
   python teams_yeelight_status.py C:\path\to\chromedriver.exe 192.168.1.100
   ```

3. **Using Environment Variables**:
   ```bash
   set CHROMEDRIVER_PATH=<path_to_chromedriver>
   set BULB_IP=<bulb_ip>
   python teams_yeelight_status.py
   ```

The script will open Microsoft Teams in a browser window, detect your current status, and update the Yeelight bulb color accordingly. It checks the status every 15 seconds and adjusts the color as needed.

## Troubleshooting

- **"Failed to update Yeelight bulb color: Bulb closed the connection"**: This error usually occurs if the bulb is powered off. Ensure that the bulb is on and connected to the same network as your computer.
- **"Error retrieving status"**: If Selenium can’t locate the Teams status, ensure that Teams is fully loaded, or try increasing the wait time in the script.

## Customization

You can customize the bulb colors for each Teams status by modifying the `update_bulb_color` function in `teams_yeelight_status.py`.

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you’d like to change.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
