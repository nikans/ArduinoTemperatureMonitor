# Arduino Temperature Monitor

A Windows Python application for monitoring temperature data from an Arduino Uno with MAX6675 thermocouple module. 100% vibe-coded.

![screen](https://github.com/user-attachments/assets/dfcd1aff-c1a5-4979-8554-2ba155cf2051)

## Features

- Real-time temperature monitoring via USB serial connection
- CSV data logging with timestamps and temperature change calculations
- Live temperature vs time graph
- Automatic Arduino port detection
- Error handling and status logging
- Multi-language support (English, Russian, French, Italian, German, Chinese)
- Origin integration for real-time data analysis

## Requirements

- Python 3.7 or higher (doesn't work with 3.13): https://www.python.org/downloads/windows/
- Arduino drivers: https://www.arduino.cc/en/software/ https://sparks.gogo.co.nz/ch340.html
- Arduino Uno with MAX6675 thermocouple module https://www.analog.com/media/en/technical-documentation/data-sheets/max6675.pdf
- USB cable to connect Arduino to computer

## Installation

### Option 1: Install from Source (Development)

1. Install Python dependencies:
```bash
pip install -r requirements.txt
```

2. Upload the Arduino code to your Arduino Uno (see Arduino Code section below)

### Option 2: Install Pre-built Executable (Recommended)

1. **Build the executable** (if not already built):
   ```bash
   build.bat
   ```

2. **Extract the ZIP file** and choose installation method:

   **Method A: Simple Installation (No Admin Required)**
   - Double-click `install_simple.bat`
   - Installs to your user directory: `%USERPROFILE%\TemperatureMonitor`

   **Method B: System Installation (Admin Required)**
   - Right-click on `install.bat`
   - Select "Run as administrator"
   - Installs to: `C:\Program Files\TemperatureMonitor`

3. **Upload the Arduino code** to your Arduino Uno (see Arduino Code section below)

### Uninstallation

To uninstall the application:
- **For user installations** (installed with `install_simple.bat`):
  - Double-click `uninstall.bat` (no admin privileges needed)
- **For system installations** (installed with `install.bat`):
  - Right-click on `uninstall.bat`
  - Select "Run as administrator"

## Usage

1. Connect your Arduino Uno to the computer via USB
2. Run the application:
```bash
python temperature_monitor.py
```

3. Click "Start Measurement" to begin recording
4. Click "Stop" to end recording and open the measurements folder

## Arduino Code

Your existing Arduino code in `temperature_monitor.ino` is already configured correctly:

```cpp
#include "max6675.h"

int thermoDO  = 12;   // MISO (Data Out от датчика → в Arduino)
int thermoCS  = 10;   // CS (Chip Select, "включение" датчика)
int thermoCLK = 13;   // SCK (Serial Clock, синхронизация)

MAX6675 thermocouple(thermoCLK, thermoCS, thermoDO);

void setup() {
  Serial.begin(9600);
  delay(500);
}

void loop() {
  double tempC = thermocouple.readCelsius();
  Serial.println(tempC);
  delay(1000); // каждую секунду
}
```

The Python application will work with your existing Arduino setup.

## File Structure

- `temperature_monitor.py` - Main application
- `requirements.txt` - Python dependencies
- `measurements/` - Folder containing CSV files (created automatically)
- `README.md` - This file

## CSV Output Format

The application creates CSV files with the following columns:
- **Time (ms)**: Milliseconds elapsed since measurement start (excluding Arduino initialization delay)
- **Temperature**: Temperature reading in Celsius
- **Change**: First derivative (temperature change per second)

## Language Configuration

The application supports multiple languages with automatic detection:

### Supported Languages
- **English** (en) - Default
- **Russian** (ru) - Русский
- **French** (fr) - Français
- **Italian** (it) - Italiano
- **German** (de) - Deutsch
- **Chinese Simplified** (zh) - 中文

### Changing Language

**Method 1: Edit config.ini**
```ini
[General]
language = fr  # Change to desired language code
```

**Method 2: Use language configuration utility**
```bash
python language_config.py fr
```

**Method 3: Interactive configuration**
```bash
python language_config.py
```

### Automatic Detection
When set to `auto`, the application detects the system language:
- Russian systems → Russian interface
- French systems → French interface
- Italian systems → Italian interface
- German systems → German interface
- Chinese systems → Chinese interface
- All others → English interface

## Troubleshooting

- **Arduino not found**: Make sure the Arduino is connected via USB and the correct drivers are installed
- **Serial communication errors**: Check that the Arduino is sending data at 9600 baud rate
- **Permission errors**: Run the application as administrator if needed
- **Language not changing**: Restart the application after changing language settings

## Notes

- The application automatically detects Arduino ports
- CSV files are saved in the `measurements` folder with timestamps
- The graph shows the last 100 data points for performance
- Temperature change is calculated as the first derivative between consecutive readings
