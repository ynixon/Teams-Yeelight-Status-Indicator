# Teams-Yeelight-Status-Indicator

**Teams-Yeelight-Status-Indicator** is a Python script that synchronizes your Microsoft Teams status with a Yeelight smart bulb. The bulb color changes dynamically based on your Teams availability, offering a quick and visual way to indicate your status.

---

## Features

- **Visual Status Indication**:
  - **Green** for Available
  - **Red** for Busy
  - **Yellow** for Away
  - **Gray** for Unknown
- **Automatic Bulb Reconnection**: Reconnects to the Yeelight bulb if the connection is lost.
- **Selenium Integration**: Fetches Teams status via the Microsoft Teams web app.

---

## Requirements

- **Python 3.8+**
- **Yeelight smart bulb** (connected to the same network as your PC)
- **Selenium** and **ChromeDriver**
- **Microsoft Teams account**

---

## Installation

1. **Clone the repository**:
   ```bash
   git clone https://github.com/ynixon/Teams-Yeelight-Status-Indicator.git
   cd Teams-Yeelight-Status-Indicator
   ```

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Download ChromeDriver**:
   - Ensure that your ChromeDriver version matches your installed Chrome browser version.
   - Place the `chromedriver.exe` in the project directory or specify its location in the `config.yaml` file.

4. **Configure the Script**:
   - Edit the `config.yaml` file to include:
     - The path to `chromedriver.exe`.
     - Your Microsoft Teams email address.
     - The Yeelight bulb’s IP address.
   Example:
   ```yaml
   settings:
     chromedriver_path: "c:/tools/chromedriver.exe"
     email: "your_email@domain.com"
     bulb_ip: "192.168.1.100"

   status_mappings:
     available: "Available"
     busy: "Busy"
     away: "Away"
     do not disturb: "Busy"
     be right back: "Away"
     offline: "Unknown"
     in a call: "Busy"
     in a meeting: "Available"
     presenting: "Busy"
   ```

---

## Usage

1. **Basic Usage**:
   Run the script with the default configuration:
   ```bash
   python teams_bulb_sync.py
   ```

2. **Specifying a Custom Configuration File**:
   Provide a custom configuration file using the `--config` argument:
   ```bash
   python teams_bulb_sync.py --config custom_config.yaml
   ```

The script will:
- Launch Microsoft Teams in a visible Chrome browser.
- Fetch your current status.
- Update the Yeelight bulb color accordingly.
- Check your Teams status every 15 seconds.

---

## Troubleshooting

- **Bulb Connection Issues**:
  - Ensure the bulb is powered on and connected to the same network as your PC.
  - Verify the correct bulb IP address in the `config.yaml` file.
  - Restart the bulb if it fails to respond.

- **ChromeDriver Errors**:
  - Make sure ChromeDriver is installed and its version matches your Chrome browser.
  - Update the `chromedriver_path` in `config.yaml` if necessary.

- **Teams Status Not Detected**:
  - Confirm that Teams is fully loaded before running the script.
  - Increase Selenium's wait time for status detection in the script if necessary.

---

## Customization

- **Change Status Colors**:
  Modify the `status_mappings` section in `config.yaml` to adjust the bulb color for specific statuses.

- **Adjust Polling Interval**:
  Change the frequency of status checks by editing the `time.sleep(15)` line in the script.

---

## Contributing

Pull requests and issues are welcome! For major changes, open an issue first to discuss your ideas.

---

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
