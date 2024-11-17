# Teams-Yeelight-Status-Indicator

**Teams-Yeelight-Status-Indicator** is a Python script that synchronizes your Microsoft Teams status with a Yeelight smart bulb. The bulb color changes dynamically based on your Teams availability, offering a quick and visual way to indicate your status.

---

## Features

- **Visual Status Indication**:
  - Dynamic RGB color mapping based on Teams status.
  - Colors are configurable via `config.yaml`.
- **Automatic Bulb Reconnection**: Reconnects to the Yeelight bulb if the connection is lost.
- **Selenium Integration**: Fetches Teams status via the Microsoft Teams web app.

---

## Requirements

- **Python 3.8+**
- **Yeelight smart bulb** (connected to the same network as your PC)
- **Selenium** and **webdriver-manager**
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

3. **Configure the Script**:
   - Edit the `config.yaml` file to include:
     - Your Microsoft Teams email address.
     - The Yeelight bulb’s IP address.
   Example:
   ```yaml
   settings:
     email: "your_email@domain.com"
     bulb_ip: "192.168.1.100"

   status_mappings:
     available:
       status: Available
       color: "0,255,0"  # Green
     busy:
       status: Busy
       color: "255,0,0"  # Red
     away:
       status: Away
       color: "255,255,0"  # Yellow
     do not disturb:
       status: Busy
       color: "255,0,0"  # Red
     be right back:
       status: Away
       color: "255,255,0"  # Yellow
     offline:
       status: Unknown
       color: "128,128,128"  # Gray
     in a call:
       status: Busy
       color: "255,0,0"  # Red
     in a meeting:
       status: Available
       color: "0,255,0"  # Green
     presenting:
       status: Busy
       color: "255,0,0"  # Red
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
  - `webdriver-manager` is used, so manual installation of ChromeDriver is not required.
  - Ensure your Chrome browser is up to date.

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
