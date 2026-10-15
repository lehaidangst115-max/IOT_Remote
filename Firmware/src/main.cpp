#include <Arduino.h>
#include <WiFi.h>
#include <WiFiClientSecure.h>
#include <PubSubClient.h>

// ================= CẤU HÌNH – SỬA Ở ĐÂY =================
const char* WIFI_SSID = "123";   // điền tên WiFi
const char* WIFI_PASS = "12345678";   // điền mật khẩu WiFi

const char*    MQTT_BROKER  = "ha8969e1.ala.asia-southeast1.emqxsl.com";
const uint16_t MQTT_PORT    = 8883;
const char*    MQTT_USER    = "dang";
const char*    MQTT_PASS    = "12345678";   // điền mật khẩu MQTT
const char*    TOPIC_ANSWER = "class/quiz/answer";

const int BLOCK_ID_DEFAULT = 33;   // block id lúc bật nguồn (0..99), đổi bằng chế độ A+OK
const int STT_DEFAULT      = 1;   // stt lúc bật nguồn (0..99), đổi bằng chế độ D+OK

const uint32_t CONFIRM_LED_MS   = 1000;  // thời gian LED xác nhận sáng sau khi gửi đáp án
const uint32_t CONNECTED_LED_MS = 2000;  // thời gian LED xác nhận sáng khi vừa kết nối xong
const uint32_t BLINK_MS         = 300;   // chu kỳ nháy LED A khi đang kết nối
const uint32_t MODE_BLINK_MS    = 300;   // chu kỳ nháy khi ở chế độ chọn block id / stt
const uint32_t COMBO_HOLD_MS    = 2000;  // giữ tổ hợp bao lâu để vào chế độ chọn
const uint32_t DEBOUNCE_MS      = 30;

// ---- module LED 7 đoạn MAX7219 (1 module 8 chữ số) ----
const uint8_t  PIN_DIN = 5;
const uint8_t  PIN_CLK = 6;
const uint8_t  PIN_CS  = 7;
const uint8_t  DISPLAY_BRIGHTNESS = 8;      // độ sáng 0..15
const bool     DIGIT0_IS_RIGHTMOST = true;  // true: DIG0 của module là chữ số ngoài cùng bên phải
                                            // (đa số module 8 số). Nếu hiển thị bị ngược thì đổi thành false
const uint32_t ANSWER_SHOW_MS = 5000;       // đáp án vừa gửi hiện trên module bao lâu
// ==========================================================

// Nút: A, B, C, D, Xác nhận (nối xuống GND, dùng pull-up nội)
// Lưu ý: GPIO9 là chân BOOT, không nhấn giữ nút ở GPIO9 lúc bật nguồn/reset
const uint8_t BTN_PINS[5] = {2, 3, 9, 4, 0};
// LED: A, B, C, D, Xác nhận (sáng ở mức cao)
// Lưu ý: GPIO20/21 là TX/RX của UART, có thể gây nhiễu nếu Serial đi qua UART
const uint8_t LED_PINS[5] = {1, 8, 10, 20, 21};
const char    ANSWERS[4]  = {'A', 'B', 'C', 'D'};
const uint8_t IDX_A       = 0;
const uint8_t IDX_B       = 1;
const uint8_t IDX_D       = 3;
const uint8_t IDX_CONFIRM = 4;

const uint32_t DISPLAY_BLINK_MS = 500;

// Phím kết hợp với OK để chỉnh Block ID
const uint8_t KEY_BLOCK = IDX_A;

// Phím kết hợp với OK để chỉnh STT
const uint8_t KEY_STT = IDX_D;

WiFiClientSecure secureClient;
PubSubClient     mqtt(secureClient);

int selected = -1;  // -1: chưa chọn, 0..3: A..D
int blockId  = BLOCK_ID_DEFAULT;
int stt      = STT_DEFAULT;

// ---- chế độ hoạt động ----
// NORMAL: chọn và gửi đáp án | BLOCK: chọn block id (giữ A+OK) | STT: chọn stt (giữ D+OK)
enum Mode : uint8_t { MODE_NORMAL = 0, MODE_BLOCK = 1, MODE_STT = 2 };
Mode     mode         = MODE_NORMAL;
Mode     comboKey     = MODE_NORMAL;  // tổ hợp đang được giữ
bool     comboFired   = false;        // tổ hợp đã kích hoạt, chờ nhả hết các nút
bool     okSuppressed = false;        // lần nhấn OK hiện tại không được tính là "xác nhận/lưu"
uint32_t comboStart   = 0;

// ---- trạng thái kết nối (để điều khiển LED báo hiệu) ----
bool     wasOnline         = false;
uint32_t connectedLedUntil = 0;   // 0: không đang đếm thời gian

// ---- đáp án vừa gửi (hiện trên module LED 7 đoạn) ----
char     shownAnswer = 0;   // 0: không hiện
uint32_t answerUntil = 0;

// ---- chống rung nút ----
bool     lastRaw[5]     = {false};
bool     stableSt[5]    = {false};   // true = đang giữ nút
bool     pressEdge[5]   = {false};   // true trong 1 vòng lặp khi vừa nhấn
bool     releaseEdge[5] = {false};   // true trong 1 vòng lặp khi vừa nhả
uint32_t lastChange[5]  = {0};

void updateButtons() {
  for (uint8_t i = 0; i < 5; i++) {
    pressEdge[i] = false;
    releaseEdge[i] = false;
    bool raw = (digitalRead(BTN_PINS[i]) == LOW);
    if (raw != lastRaw[i]) {
      lastRaw[i] = raw;
      lastChange[i] = millis();
    }
    if ((millis() - lastChange[i] > DEBOUNCE_MS) && raw != stableSt[i]) {
      stableSt[i] = raw;
      if (raw) pressEdge[i] = true;
      else     releaseEdge[i] = true;
    }
  }
}

void allLedsOff() {
  for (uint8_t i = 0; i < 5; i++) digitalWrite(LED_PINS[i], LOW);
}

void showSelected() {
  for (uint8_t i = 0; i < 4; i++) {
    digitalWrite(LED_PINS[i], (i == selected) ? HIGH : LOW);
  }
}

// ================= MODULE MAX7219 (tự điều khiển, không cần thư viện) =================
// Mã đoạn: bit7=DP, bit6=A, bit5=B, bit4=C, bit3=D, bit2=E, bit1=F, bit0=G
const uint8_t SEG_DIGIT[10] = {0x7E, 0x30, 0x6D, 0x79, 0x33, 0x5B, 0x5F, 0x70, 0x7F, 0x7B};
const uint8_t SEG_DASH  = 0x01;
const uint8_t SEG_BLANK = 0x00;

uint8_t segLast[8]  = {0};
bool    segFirst    = true;

void max7219Write(uint8_t reg, uint8_t val) {
  digitalWrite(PIN_CS, LOW);
  shiftOut(PIN_DIN, PIN_CLK, MSBFIRST, reg);
  shiftOut(PIN_DIN, PIN_CLK, MSBFIRST, val);
  digitalWrite(PIN_CS, HIGH);
}

void displayInit() {
  pinMode(PIN_DIN, OUTPUT);
  pinMode(PIN_CLK, OUTPUT);
  pinMode(PIN_CS, OUTPUT);
  digitalWrite(PIN_CS, HIGH);
  digitalWrite(PIN_CLK, LOW);

  max7219Write(0x0F, 0x00);               // tắt test hiển thị
  max7219Write(0x0C, 0x01);               // bật hoạt động
  max7219Write(0x09, 0x00);               // không dùng giải mã BCD, tự gửi mã đoạn
  max7219Write(0x0B, 0x07);               // quét cả 8 chữ số
  max7219Write(0x0A, DISPLAY_BRIGHTNESS); // độ sáng
  for (uint8_t r = 1; r <= 8; r++) max7219Write(r, 0x00);
}

uint8_t answerSegments(char c) {
  switch (c) {
    case 'A': return 0x77;   // A
    case 'B': return 0x1F;   // b
    case 'C': return 0x4E;   // C
    case 'D': return 0x3D;   // d
    default:  return SEG_BLANK;
  }
}

// Bố cục (vị trí 0 = ngoài cùng bên trái):
// [0][1] block id | [2] '-' | [3][4] stt | [5] '-' | [6] đáp án vừa gửi | [7] trống
void updateDisplay() {

  // Hết thời gian hiển thị đáp án
  if (shownAnswer != 0 &&
      (int32_t)(millis() - answerUntil) >= 0) {
    shownAnswer = 0;
  }

  uint8_t seg[8];

  // ================= HIỂN THỊ MẶC ĐỊNH =================

  // Block ID
  seg[0] = SEG_DIGIT[(blockId / 10) % 10];
  seg[1] = SEG_DIGIT[blockId % 10];

  seg[2] = SEG_DASH;

  // STT
  seg[3] = SEG_DIGIT[(stt / 10) % 10];
  seg[4] = SEG_DIGIT[stt % 10];

  seg[5] = SEG_DASH;

  // Đáp án
  seg[6] = answerSegments(shownAnswer);
  seg[7] = SEG_BLANK;


  // ================= NHẤP NHÁY =================

  bool blinkOn = ((millis() / DISPLAY_BLINK_MS) % 2 == 0);

  if (!blinkOn) {

    // Đang chỉnh Block ID
    if (mode == MODE_BLOCK) {
      seg[0] = SEG_BLANK;
      seg[1] = SEG_BLANK;
    }

    // Đang chỉnh STT
    else if (mode == MODE_STT) {
      seg[3] = SEG_BLANK;
      seg[4] = SEG_BLANK;
    }
  }


  // ================= KIỂM TRA THAY ĐỔI =================

  // Không cập nhật nếu dữ liệu không thay đổi
  if (!segFirst &&
      memcmp(seg, segLast, sizeof(seg)) == 0) {
    return;
  }

  segFirst = false;
  memcpy(segLast, seg, sizeof(seg));


  // ================= CẬP NHẬT MAX7219 =================

  for (uint8_t p = 0; p < 8; p++) {

    uint8_t reg = DIGIT0_IS_RIGHTMOST ? (8 - p) : (1 + p);

    max7219Write(reg, seg[p]);
  }
}
// =======================================================================================

// ---- kết nối WiFi / MQTT (không chặn vòng lặp quá lâu) ----
void maintainConnection() {
  static uint32_t lastTry = 0;

  if (WiFi.status() != WL_CONNECTED) {
    if (millis() - lastTry > 10000) {
      lastTry = millis();
      Serial.printf("[WiFi] dang ket noi... status=%d\n", WiFi.status());
      WiFi.begin(WIFI_SSID, WIFI_PASS);
    }
    return;
  }

  if (!mqtt.connected()) {
    if (millis() - lastTry > 5000) {
      lastTry = millis();
      String clientId = "quiz-esp32c3-" + String((uint32_t)ESP.getEfuseMac(), HEX);
      Serial.print("[MQTT] dang ket noi...");
      if (mqtt.connect(clientId.c_str(), MQTT_USER, MQTT_PASS)) {
        Serial.println(" OK");
      } else {
        Serial.printf(" that bai, rc=%d\n", mqtt.state());
      }
    }
    return;
  }

  mqtt.loop();
}

// ---- LED báo trạng thái kết nối ----
// Đang kết nối (chưa đủ WiFi + MQTT): LED A nháy (nếu đang ở chế độ thường và chưa chọn đáp án)
// Vừa kết nối xong: LED xác nhận sáng CONNECTED_LED_MS rồi tắt
void updateStatusLeds() {
  bool online = (WiFi.status() == WL_CONNECTED) && mqtt.connected();

  if (online && !wasOnline) {
    Serial.println("[TRANG THAI] da ket noi WiFi + MQTT");
    if (selected >= 0) showSelected();
    else digitalWrite(LED_PINS[IDX_A], LOW);    // dừng nháy LED A
    digitalWrite(LED_PINS[IDX_CONFIRM], HIGH);  // LED xác nhận sáng
    connectedLedUntil = millis() + CONNECTED_LED_MS;
  } else if (!online && wasOnline) {
    Serial.println("[TRANG THAI] mat ket noi");
  }
  wasOnline = online;

  if (!online) {
    if (selected < 0 && mode == MODE_NORMAL) {
      digitalWrite(LED_PINS[IDX_A], ((millis() / BLINK_MS) % 2) ? HIGH : LOW);
    }
  } else if (connectedLedUntil != 0 && (int32_t)(millis() - connectedLedUntil) >= 0) {
    digitalWrite(LED_PINS[IDX_CONFIRM], LOW);
    connectedLedUntil = 0;
  }
}

// ---- chế độ chọn block id / stt ----
int modeLed() { return (mode == MODE_BLOCK) ? IDX_A : IDX_D; }  // LED nháy báo đang ở chế độ

void printValue() {
  if (mode == MODE_BLOCK) {
    Serial.printf("[BLOCK] hang chuc=%d | hang don vi=%d -> block_id=%d\n",
                  blockId / 10, blockId % 10, blockId);
  } else {
    Serial.printf("[STT] hang chuc=%d | hang don vi=%d -> stt=%d\n",
                  stt / 10, stt % 10, stt);
  }
}

void enterMode(Mode m) {
  mode = m;
  selected = -1;      // bỏ đáp án đang chọn dở
  allLedsOff();
  Serial.printf("[%s] vao che do chon (A: +hang chuc, B: +hang don vi, OK: luu & thoat)\n",
                m == MODE_BLOCK ? "BLOCK" : "STT");
  printValue();
}

void exitMode() {
  digitalWrite(LED_PINS[modeLed()], LOW);
  if (mode == MODE_BLOCK) Serial.printf("[BLOCK] da luu, thoat che do. block_id=%d\n", blockId);
  else                    Serial.printf("[STT] da luu, thoat che do. stt=%d\n", stt);
  mode = MODE_NORMAL;
}

void sendAnswer() {
  if (selected < 0) {
    Serial.println("[GUI] chua chon dap an, bo qua");
    return;
  }
  if (!mqtt.connected()) {
    Serial.println("[GUI] chua ket noi MQTT, khong gui duoc");
    return;
  }

  char payload[128];
  snprintf(payload, sizeof(payload),
           "{\"block_id\":\"%d\",\"stt\":%d,\"answer\":\"%c\"}",
           blockId, stt, ANSWERS[selected]);

  if (mqtt.publish(TOPIC_ANSWER, payload)) {
    Serial.printf("[GUI] %s -> %s\n", TOPIC_ANSWER, payload);

    // hiện đáp án vừa gửi trên module LED 7 đoạn
    shownAnswer = ANSWERS[selected];
    answerUntil = millis() + CONFIRM_LED_MS + ANSWER_SHOW_MS;
    updateDisplay();

    digitalWrite(LED_PINS[IDX_CONFIRM], HIGH);  // LED xác nhận sáng
    delay(CONFIRM_LED_MS);
    allLedsOff();                               // tắt hết LED
    selected = -1;                              // cho phép chọn và gửi lại
  } else {
    Serial.println("[GUI] publish that bai");
  }
}

void handleButtons() {
  // 1. Kiểm tra tổ hợp phím
  bool comboBlock = stableSt[KEY_BLOCK] && !stableSt[KEY_STT] && stableSt[IDX_CONFIRM];
  bool comboSTT = stableSt[KEY_STT] && !stableSt[KEY_BLOCK] && stableSt[IDX_CONFIRM];
  Mode target = comboBlock ? MODE_BLOCK : (comboSTT ? MODE_STT : MODE_NORMAL);

  if (target == MODE_NORMAL) {
    comboStart = 0;
  } else if (mode == MODE_NORMAL && !comboFired) {
    if (comboStart == 0 || comboKey != target) {
      comboStart = millis();
      comboKey = target;
    } else if (millis() - comboStart >= COMBO_HOLD_MS) {
      comboFired = true;
      okSuppressed = true;
      enterMode(target);
    }
  }

  if (!stableSt[KEY_BLOCK] && !stableSt[KEY_STT] && !stableSt[IDX_CONFIRM])
    comboFired = false;

  // 2. Xử lý nút A, B, C, D
  for (uint8_t i = 0; i < 4; i++) {
    if (!pressEdge[i]) continue;

    if (mode != MODE_NORMAL) {
      int* value = (mode == MODE_BLOCK) ? &blockId : &stt;

      if (i == IDX_A) {
        *value = (((*value / 10) + 1) % 10) * 10 + (*value % 10);
        printValue();
      } else if (i == IDX_B) {
        *value = (*value / 10) * 10 + (((*value % 10) + 1) % 10);
        printValue();
      }
    } else if ((i == KEY_BLOCK || i == KEY_STT) && stableSt[IDX_CONFIRM]) {
      okSuppressed = true;
    } else {
      selected = i;
      showSelected();
      Serial.printf("[CHON] %c\n", ANSWERS[i]);
    }
  }

  // 3. Xử lý nút OK
  if (pressEdge[IDX_CONFIRM] && mode == MODE_NORMAL &&
      (stableSt[KEY_BLOCK] || stableSt[KEY_STT])) {
    okSuppressed = true;
  }

  if (releaseEdge[IDX_CONFIRM]) {
    if (!okSuppressed) {
      if (mode != MODE_NORMAL) exitMode();
      else sendAnswer();
    }
    okSuppressed = false;
  }
}

void setup() {
  Serial.begin(115200);

  for (uint8_t i = 0; i < 5; i++) {
    pinMode(BTN_PINS[i], INPUT_PULLUP);
    pinMode(LED_PINS[i], OUTPUT);
    digitalWrite(LED_PINS[i], LOW);
  }

  displayInit();
  updateDisplay();

  WiFi.mode(WIFI_STA);
  WiFi.begin(WIFI_SSID, WIFI_PASS);
  WiFi.setTxPower(WIFI_POWER_8_5dBm);

  secureClient.setInsecure();  // TLS nhưng không kiểm tra chứng chỉ
  mqtt.setServer(MQTT_BROKER, MQTT_PORT);
}

void loop() {
  maintainConnection();
  updateStatusLeds();

  // Chế độ chọn block id: LED A nháy | chế độ chọn stt: LED D nháy
  if (mode != MODE_NORMAL) {
    digitalWrite(LED_PINS[modeLed()], ((millis() / MODE_BLINK_MS) % 2) ? HIGH : LOW);
  }

  updateButtons();
  handleButtons();
  updateDisplay();
}