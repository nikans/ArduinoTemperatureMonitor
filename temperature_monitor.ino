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
  delay(500);
}