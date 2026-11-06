/*
 * PRAYAS sensor node firmware (Wokwi simulation build)
 * Board: ESP32-S3-DevKitC-1 | Framework: Arduino
 *
 * Simulates ONE village sensor node: flood (ultrasonic) + fire (temp/humidity/gas)
 * + soil moisture + rain gauge + siren/LED alert output.
 * Detection thresholds mirror env_monitoring/detection_engine.py.
 *
 * ---- Wokwi part stand-ins (see diagram.json) ----
 *  HC-SR04  -> JSN-SR04T water-level probe
 *  DHT22    -> BME680 (temp + humidity channels only)
 *  MPU6050  -> LIS3DH accelerometer (tilt / vibration RMS)
 *  pot knob -> capacitive soil-moisture probe
 *  pushbtn  -> tipping-bucket reed switch (press = 1 tip = 0.2 mm)
 *
 * ---- NOT simulated (pins reserved, see diagram labels) ----
 *  SX1262 LoRa : SPI SCK12/MISO13/MOSI11/NSS10 + DIO1-14/RST-15/BUSY-16
 *  PMS5003 AQI : UART1 RX-17
 *
 * Required libraries (Wokwi: add via Libraries tab, or use libraries.txt):
 *  - "DHT sensor library for ESPx"  (DHTesp.h)
 *  - "Adafruit MPU6050" + "Adafruit Unified Sensor"
 */

#include <Arduino.h>
#include <Wire.h>
#include "DHTesp.h"
#include <Adafruit_MPU6050.h>
#include <Adafruit_Sensor.h>

// ---------------- Pin map (ESP32-S3-DevKitC-1) ----------------
// ADC1-only analog pins (safe with WiFi): 1, 2, 4
#define PIN_MQ135_AO   1    // MQ-135 analog (ADC1_CH0)
#define PIN_MQ2_AO     2    // MQ-2 analog (ADC1_CH1)
#define PIN_SOIL_AO    4    // soil probe analog (ADC1_CH3)
#define PIN_SONIC_TRIG 5
#define PIN_SONIC_ECHO 6
#define PIN_DHT        7    // DHT22 one-wire data
#define PIN_I2C_SDA    8
#define PIN_I2C_SCL    9
// 10,11,12,13 reserved: LoRa SPI (NSS, MOSI, SCK, MISO) — no Wokwi model
// 14,15,16 reserved: LoRa DIO1 / RESET / BUSY
// 17 reserved: PMS5003 UART1 RX
#define PIN_SIREN      21
#define PIN_ALERT_LED  38
#define PIN_MQ135_DO   39
#define PIN_MQ2_DO     40
#define PIN_RAIN       42   // reed switch to GND, INPUT_PULLUP

// ---------------- Thresholds (mirror detection_engine.py) ----------------
#define WL_WARN_RATE    2.0f    // cm/min
#define WL_CRIT_RATE    5.0f
#define WL_FLASH_RATE   10.0f
#define FIRE_TEMP_RATE  2.0f    // degC/min
#define FIRE_TEMP_ABS   45.0f
#define FIRE_HUM_RATE   5.0f    // %/min drop
#define FIRE_HUM_ABS    20.0f
#define SOIL_SAT        45.0f   // % VWC
#define SOIL_CRIT       55.0f
#define RAIN_WARN       30.0f   // mm/hr (1-hr intensity)
#define RAIN_CRIT       50.0f
#define SENSE_MS        10000UL // 10 s PRAYAS cadence
#define SONIC_MOUNT_CM  300.0f  // sensor height above riverbed
#define MM_PER_TIP      0.2f

DHTesp dht;
Adafruit_MPU6050 mpu;
bool imuOk = false;

volatile uint32_t rainTips = 0;
volatile uint32_t lastTipMs = 0;
void IRAM_ATTR onRainTip() {
  uint32_t now = millis();
  if (now - lastTipMs > 120) { rainTips++; lastTipMs = now; } // debounce
}

// Flood history (10-min sliding window for rise-rate)
#define HIST_N 60
float wlHist[HIST_N]; uint32_t wlTime[HIST_N]; uint8_t wlIdx = 0, wlCount = 0;
float lastTemp = NAN, lastHum = NAN; uint32_t lastTempMs = 0;

float readWaterLevelCm() {
  digitalWrite(PIN_SONIC_TRIG, LOW); delayMicroseconds(4);
  digitalWrite(PIN_SONIC_TRIG, HIGH); delayMicroseconds(10);
  digitalWrite(PIN_SONIC_TRIG, LOW);
  long us = pulseIn(PIN_SONIC_ECHO, HIGH, 30000UL); // 30 ms timeout
  if (us == 0) return NAN;
  float dist = us / 58.0f;                 // cm, sensor -> water surface
  return SONIC_MOUNT_CM - dist;            // cm of water depth
}

void pushWl(float wl, uint32_t now) {
  wlHist[wlIdx] = wl; wlTime[wlIdx] = now;
  wlIdx = (wlIdx + 1) % HIST_N; if (wlCount < HIST_N) wlCount++;
}

float waterRiseRate() { // cm/min over up to 10 min window
  if (wlCount < 2) return 0;
  uint8_t oldest = (wlIdx + HIST_N - wlCount) % HIST_N;
  uint8_t newest = (wlIdx + HIST_N - 1) % HIST_N;
  float dtMin = (wlTime[newest] - wlTime[oldest]) / 60000.0f;
  if (dtMin < 0.05f) return 0;
  return (wlHist[newest] - wlHist[oldest]) / dtMin;
}

float adcToVwc(int raw) { return raw * 100.0f / 4095.0f; }       // soil %
float adcToPpm(int raw) { return raw * 1000.0f / 4095.0f; }      // gas, indicative

const char* levelName(float s) {
  if (s >= 0.85f) return "EMERGENCY";
  if (s >= 0.60f) return "WARNING";
  if (s >= 0.30f) return "WATCH";
  return "INFO";
}

void alertOutput(float score) {
  if (score >= 0.60f) {
    digitalWrite(PIN_ALERT_LED, HIGH);
    tone(PIN_SIREN, score >= 0.85f ? 1200 : 880, score >= 0.85f ? 800 : 300);
  } else {
    digitalWrite(PIN_ALERT_LED, LOW);
    noTone(PIN_SIREN);
  }
}

void setup() {
  Serial.begin(115200);
  pinMode(PIN_SONIC_TRIG, OUTPUT);
  pinMode(PIN_SONIC_ECHO, INPUT);
  pinMode(PIN_SIREN, OUTPUT);
  pinMode(PIN_ALERT_LED, OUTPUT);
  pinMode(PIN_MQ135_DO, INPUT);
  pinMode(PIN_MQ2_DO, INPUT);
  pinMode(PIN_RAIN, INPUT_PULLUP);
  attachInterrupt(digitalPinToInterrupt(PIN_RAIN), onRainTip, FALLING);

  Wire.begin(PIN_I2C_SDA, PIN_I2C_SCL);
  dht.setup(PIN_DHT, DHTesp::DHT22);
  imuOk = mpu.begin();                       // addr 0x68
  if (imuOk) { mpu.setAccelerometerRange(MPU6050_RANGE_8_G); }

  Serial.println(F("PRAYAS node boot: ESP32-S3 | 10s cadence | LoRa+PMS pins reserved"));
  if (!imuOk) Serial.println(F("WARN: MPU6050 not found (ilich check wiring SDA8/SCL9)"));
}

void loop() {
  static uint32_t lastSense = 0;
  uint32_t now = millis();
  if (now - lastSense < SENSE_MS) { delay(50); return; }
  lastSense = now;

  // ---- Sense ----
  float wl = readWaterLevelCm();
  if (!isnan(wl)) pushWl(wl, now);
  float riseRate = waterRiseRate();

  TempAndHumidity th = dht.getTempAndHumidity();
  float temp = th.temperature, hum = th.humidity;
  float tempRate = 0, humDrop = 0;
  if (!isnan(temp) && !isnan(lastTemp) && lastTempMs > 0) {
    float dtMin = (now - lastTempMs) / 60000.0f;
    if (dtMin > 0.05f) {
      tempRate = (temp - lastTemp) / dtMin;
      humDrop  = (lastHum - hum) / dtMin;
    }
  }
  if (!isnan(temp)) { lastTemp = temp; lastHum = hum; lastTempMs = now; }

  float ax = 0, ay = 0, az = 0;
  if (imuOk) {
    sensors_event_t a, g, t;
    mpu.getEvent(&a, &g, &t);
    ax = a.acceleration.x; ay = a.acceleration.y; az = a.acceleration.z;
  }
  float tiltDeg = degrees(acos(constrain(az / 9.81f, -1.0f, 1.0f))); // tilt from vertical

  int mq135Raw = analogRead(PIN_MQ135_AO), mq2Raw = analogRead(PIN_MQ2_AO);
  bool mq135Do = digitalRead(PIN_MQ135_DO) == LOW; // DO pulls LOW over threshold
  bool mq2Do   = digitalRead(PIN_MQ2_DO) == LOW;
  float soilVwc = adcToVwc(analogRead(PIN_SOIL_AO));

  static uint32_t tipsAtWindowStart = 0; static uint32_t windowStart = 0;
  if (windowStart == 0) windowStart = now;
  float winHrs = (now - windowStart) / 3600000.0f;
  if (winHrs >= 1.0f) { tipsAtWindowStart = rainTips; windowStart = now; winHrs = 1.0f; }
  float rainRate = (winHrs > 0.001f) ? (rainTips - tipsAtWindowStart) * MM_PER_TIP / winHrs : 0;

  // ---- Decide: flood score (mirrors FloodDetector._compute_score) ----
  float flood = 0;
  if (riseRate > WL_FLASH_RATE) flood += 0.5f;
  else if (riseRate > WL_CRIT_RATE) flood += 0.3f;
  else if (riseRate > WL_WARN_RATE) flood += 0.15f;
  if (rainRate > 124.5f) flood += 0.3f;
  else if (rainRate > 64.5f) flood += 0.15f;
  if (soilVwc > SOIL_CRIT) flood += 0.2f;
  else if (soilVwc > SOIL_SAT) flood += 0.1f;
  flood = min(flood, 1.0f);

  // ---- Decide: fire score (2+ confirming channels, mirrors ForestFireDetector) ----
  float sTemp = 0, sHum = 0, sGas = 0;
  if (tempRate > FIRE_TEMP_RATE) sTemp += 0.6f;
  if (!isnan(temp) && temp > FIRE_TEMP_ABS) sTemp += 0.4f;
  sTemp = min(sTemp, 1.0f);
  if (humDrop > FIRE_HUM_RATE) sHum += 0.6f;
  if (!isnan(hum) && hum < FIRE_HUM_ABS) sHum += 0.4f;
  sHum = min(sHum, 1.0f);
  if (mq135Do || mq2Do) sGas = 1.0f;
  else if (adcToPpm(mq135Raw) > 400 || adcToPpm(mq2Raw) > 400) sGas = 0.5f;
  int confirming = (sTemp > 0.3f) + (sHum > 0.3f) + (sGas > 0.3f);
  float fire = 0;
  if (confirming >= 2) {
    fire = 0.15f * sTemp + 0.10f * sHum + 0.10f * sGas
         + min(confirming - 1, 3) * 0.05f;   // agreement bonus
    // NOTE: vision_smoke/vision_fire weights (0.30/0.35) need a camera —
    // real node adds them; sim caps fire score accordingly.
    fire = min(fire * 2.0f, 1.0f);
  }

  // ---- Decide: landslide-lite (rain intensity + tilt anomaly) ----
  float slide = 0;
  if (rainRate > RAIN_CRIT) slide += 0.4f;
  else if (rainRate > RAIN_WARN) slide += 0.25f;
  if (soilVwc > SOIL_CRIT) slide += 0.25f;
  else if (soilVwc > SOIL_SAT) slide += 0.15f;
  if (tiltDeg > 5.0f) slide += 0.2f;         // slope movement proxy
  slide = min(slide, 1.0f);

  float score = max(flood, max(fire, slide));
  const char* lvl = levelName(score);
  const char* hazard = (score == flood && score > 0) ? "flood"
                     : (score == fire && score > 0) ? "fire"
                     : (score == slide && score > 0) ? "landslide" : "none";
  alertOutput(score);

  // ---- Telemetry (matches MQTT payload fields in implementation_plan §6.1) ----
  Serial.print(F("PRAYAS t=")); Serial.print(now / 1000);
  Serial.print(F(" wl_cm=")); Serial.print(isnan(wl) ? -1 : wl, 1);
  Serial.print(F(" rise_cm_min=")); Serial.print(riseRate, 2);
  Serial.print(F(" temp_C=")); Serial.print(isnan(temp) ? -99 : temp, 1);
  Serial.print(F(" hum_pct=")); Serial.print(isnan(hum) ? -1 : hum, 1);
  Serial.print(F(" gas_ppm=")); Serial.print(max(adcToPpm(mq135Raw), adcToPpm(mq2Raw)), 0);
  Serial.print(F(" soil_vwc=")); Serial.print(soilVwc, 1);
  Serial.print(F(" rain_mm_hr=")); Serial.print(rainRate, 1);
  Serial.print(F(" tilt_deg=")); Serial.print(tiltDeg, 1);
  Serial.print(F(" alert=")); Serial.print(hazard);
  Serial.print(F(" level=")); Serial.print(lvl);
  Serial.print(F(" conf=")); Serial.println(score, 2);
}
