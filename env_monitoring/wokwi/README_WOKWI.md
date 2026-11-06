# PRAYAS Sensor Node — Wokwi Simulation

Interactive ESP32-S3 sensor-node simulation. Paste these 3 files into a Wokwi
project and press Play — no hardware, no installs.

## 1. Open it (2 minutes)

1. Go to https://wokwi.com/projects/new/esp32-s3 (free account)
2. Replace the contents of `diagram.json` with the file in this folder
3. Replace `sketch.ino` with the file in this folder
4. Add the libraries: click `libraries.txt` (under the `+` menu) and paste
   the file in this folder (or add each library via the Libraries tab)
5. Press Play, open the Serial Monitor (115200 baud)

## 2. Drive a disaster (demo script)

| Action in simulation | What it proves |
|---|---|
| Drag HC-SR04 distance slider down fast | Water-level rise rate → flood WATCH → WARNING → EMERGENCY, siren + LED |
| Drag DHT22 temperature up past 45 °C, humidity down | Temp + humidity confirm → fire alert (needs 2 channels, like the real correlator) |
| Drag MQ gas `ppm` slider up | Gas spike confirms fire / industrial alert |
| Turn soil potentiometer knob up | VWC crosses 45 % (saturated) → 55 % (critical), feeds flood + landslide scores |
| Click the blue rain button repeatedly | Tip counting → mm/hr → landslide-lite score |

Serial output is one `PRAYAS ...` telemetry line per 10 s sense cycle, with
`alert=` / `level=` / `conf=` fields matching `detection_engine.py`
(INFO < 0.30 < WATCH < 0.60 < WARNING < 0.85 < EMERGENCY).

## 3. What is real vs stand-in (use this caption in the PPT)

- Simulated for real: ESP32-S3 GPIO/I2C/ADC reads, threshold + correlation
  scoring, siren/LED actuation, 10 s cadence.
- Functional stand-ins, labelled on the diagram: HC-SR04 → JSN-SR04T,
  DHT22 → BME680 (temp/humidity only), MPU6050 → LIS3DH, potentiometer →
  soil probe, pushbutton → tipping bucket.
- NOT simulated: SX1262 LoRa RF link and PMS5003 laser scattering (no Wokwi
  models exist). Their MCU pins are reserved in firmware and on the diagram.
  The simulation proves firmware logic + sensor interfacing, not RF range,
  power autonomy, or AQI optics.

## 4. Real-hardware wiring deltas (before building the physical node)

- JSN-SR04T ECHO is 5 V → voltage divider (e.g. 1k/2k) to GPIO6. Never 5 V
  straight into the ESP32.
- MQ-135/MQ-2 AO swing to ~5 V → divider or 3.3 V-compatible module before
  GPIO1/GPIO2 (ADC1 only — ADC2 clashes with WiFi).
- Power path: 6 V solar → CN3722 MPPT → 3.2 V LiFePO4 → 5 V boost → DevKit
  VIN; 3.3 V rail for BME680/LIS3DH; 5 V rail for PMS5003/MQ/JSN-SR04T.
- Avoid strapping pins 0/3/45/46 for outputs, USB pins 19/20, 35/36/37
  (octal flash on some variants), 43/44 (USB serial), 48 (onboard RGB LED).
