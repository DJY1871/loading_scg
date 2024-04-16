# SgsApp Script Function
## Introduction
用於自動操控Escort Console軟體對sgs溫溼度感測器進行資料下載與配置的腳本。
## Script Introduce
### Loading Script
功能: 自動下載Sensor內儲存的資料並將其寫入控制台附帶Senosr屬性，異常狀態時會回報錯誤訊息並跳出。
* 異常狀態回報訊息說明 :
  * 無法正確啟動軟體: Can't activate ESCORT Console Pro app. Please check the path where your app executable file is saved.
  * 沒有正確進入對應視窗: Can't find XXX window.
  * 沒有搜尋到Sensor的屬性:
    * Error: Unable to find the window or read text from it.
    * Please check Battery,COM connect,Sensor connect.
  * Sensor內沒有儲存的資料:
    * Can't find any data in sensor.
    * Error with empty text.
    * Error with no any text found.
### Configuration Script
功能: 用Command Line的方式將資料傳輸給腳本對Sensor做配置。
* 異常狀態回報說明: 同上
* Cmdline資料規格:
   * 1.description
   * 2.mintemper
   * 3.maxtemper
   * 4.increment
   * 5.AlChoose: EX~ "T_F_T_F_T_T_T"
   * 6.Aloutofspec1
   * 7.Aloutofspec2
   * 8.Duration
   * 9.readinterval: EX~ 03days10h21m06s
   * 10.startlogchoose
   * 11.starttime: EX~ 03h02m    or  03days10h21m    or  "2045/3/4 上 03:22"
   * 12.finishlogchoose
   * 13.aftertime: EX~ 10   or  03days10h21m    or  "2045/3/4 上 03:22"
   * 14.enable beep: EX~ T
   * 15.new battery fitted: EX~ F
   * 16.old_description
* 不更動cmdline寫法
  * description輸入de則f維持原值
  * mintemper輸入def，2,3,4維持原值
  * AlChoose輸入def則維持原值
* EX~ configuration.exe def -20 70 5 F_F_F_F_F_F_F 0 0 U 00days00h00m06s 2 00days00h01m 1 10 T F 39-78
### Get Property Script
功能: 回傳Sensor附帶屬性
* 屬性內容:
  * Serial Number
  * Battery Status
  * Description
  * Product Code
* 異常狀態回報說明: 同上
## 注意事項
當腳本在搜尋Sensor時預設是超過10秒未找到會自動跳脫， 如果突然提前關閉頁面可能會有異常，建議如果要跳脫可以使用熱鍵ESC。
