# Mouth-Operated Mouse V2

An open-source, affordable assistive technology project that enables computer control through mouth movements and sip/puff actions. This is an updated version of the original Mouth-Operated Mouse, now utilizing an Arduino Leonardo for direct HID (Human Interface Device) emulation and an enhanced Python application for calibration and training.

## Overview

This project creates a mouth-operated mouse combining a pressure sensor for sip/puff actions with a joystick for cursor movement, allowing users with limited hand mobility to control a computer. The device functions as follows:

*   **Joystick**: Controls cursor movement (operated by mouth/chin)
*   **Sip/Puff Actions**:
    *   Hard sip → Right click
    *   Soft sip → Scroll down
    *   Neutral → No action
    *   Soft puff → Scroll up
    *   Hard puff → Left click

## Key Features of V2

*   **Arduino Leonardo Integration**: Direct USB HID emulation, eliminating the need for a separate Python script for mouse control on the computer side. The Arduino Leonardo acts directly as a mouse.
*   **Enhanced Python Application (`App.py`)**: A comprehensive GUI application built with `customtkinter` for:
    *   **Tuning & Profiles**: Adjust pressure thresholds and joystick sensitivity, save and load custom profiles.
    *   **Trainer**: Interactive exercises to improve sip/puff and joystick control accuracy.
    *   **Calibrate Sensor**: Visualize real-time pressure data and assist in setting optimal thresholds.
    *   **Stick Control**: Visual representation of joystick input and deadzone.
*   **Improved Responsiveness**: Direct HID communication from Arduino Leonardo offers lower latency.
*   **Modular Design**: Easy customization and integration of components.

## Components

1.  **Arduino Leonardo** (~$5 USD)
    *   Replaces Arduino Uno for direct HID capabilities.
2.  **MPXV7002DP** (~$20 USD)
    *    - Specifically designed for sip/puff applications
3.  **Joystick Module**
    *   Analog Thumb Joystick Module (~$2 USD)
4.  **Tubing and Mouthpiece:**
    *   Food-grade silicone tubing (~$5 USD) - 1/8" inner diameter
    *   Plastic mouthpiece (custom-made or adapted)
5.  **Additional Components:**
    *   Breadboard for prototyping (~$5 USD)
    *   Jumper wires (~$3-5 USD)
    *   USB cable for Arduino (~$3-5 USD)
    *   Small project box/enclosure (~$5-10 USD)

**Estimated Total Cost:** Approximately $40 USD

## Setup Instructions

### 1. Hardware Assembly

Follow these basic circuit connections:

*   **Pressure Sensor (MPXV7002DP) Connections:**
    *   GND pin → Arduino GND
    *   +5V pin → Arduino 5V
    *   Analog output pin → Arduino A0
*   **Joystick Module Connections:**
    *   GND pin → Arduino GND
    *   +5V pin → Arduino 5V
    *   VRx (X-axis) → Arduino A1
    *   VRy (Y-axis) → Arduino A2
    *   SW (Switch, optional) → Not used in V2.ino, but can be connected to a digital pin if desired.
*   **Tubing Connection:**
    *   Connect silicon food-grade tubing to the pressure port on the sensor.
    *   The other end of the tubing connects to a mouthpiece.

### 2. Arduino IDE Setup

1.  **Install Arduino IDE**: Download and install the Arduino IDE from the [official website](https://www.arduino.cc/en/software).
2.  **Install Arduino Leonardo Board**: Go to `Tools > Board > Boards Manager...` and search for "Arduino AVR Boards". Install the package that includes the Arduino Leonardo.
3.  **Install Libraries**: The `V2.ino` sketch uses the built-in `Mouse.h` library. No additional library installations are required for the Arduino sketch.
4.  **Upload Sketch**: Open `V2.ino` in the Arduino IDE, select `Tools > Board > Arduino Leonardo`, and choose the correct `Port`. Then, click `Upload`.

### 3. Python Application Setup

1.  **Install Python**: Ensure you have Python 3.x installed. You can download it from [python.org](https://www.python.org/downloads/).
2.  **Install Dependencies**: Open a terminal or command prompt and navigate to the directory where `App.py` is located. Install the required Python libraries using pip:
    ```bash
    pip install pyserial customtkinter pyautogui
    ```
3.  **Run Application**: Execute the Python application:
    ```bash
    python App.py
    ```

## Usage

Once the Arduino sketch is uploaded and the Python application is running, you can use the `App.py` interface to:

*   **Connect to Arduino**: Select the serial port connected to your Arduino Leonardo and click "Connect".
*   **Tune Parameters**: Adjust the pressure thresholds (Hard Sip, Neutral Min/Max, Soft Puff, Hard Puff) and joystick deadzone/cursor speed. Apply settings to the Arduino.
*   **Calibrate Sensor**: Use the "Calibrate Sensor" tab to visualize real-time pressure readings and fine-tune your thresholds for optimal performance.
*   **Train**: Utilize the "Trainer" tab to practice and improve your control.
*   **Manage Profiles**: Save and load different configurations as profiles.

## Troubleshooting

*   **Arduino Not Detected**: Ensure the Arduino Leonardo drivers are correctly installed and the correct port is selected in the Arduino IDE and `App.py`.
*   **Serial Communication Issues**: Verify that no other application is using the serial port. Restarting the Arduino IDE or `App.py` might help.
*   **Mouse Not Moving/Clicking**: Check the pressure sensor and joystick connections. Ensure the `V2.ino` sketch is successfully uploaded to the Arduino Leonardo.
*   **Calibration**: The pressure thresholds are highly dependent on your specific sensor and lung capacity. Use the "Calibrate Sensor" tab in `App.py` to find your optimal settings.

## Contributing

Contributions are welcome! Please feel free to fork the repository, make improvements, and submit pull requests. For major changes, please open an issue first to discuss what you would like to change.

## License

This project is open-source and available under the [MIT License](https://opensource.org/licenses/MIT).

