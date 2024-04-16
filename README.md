# SgsApp Script Function
## Introduction
用於自動操控Escort Console軟體對sgs溫溼度感測器進行資料下載與配置的腳本
## Script Introduce
### Loading Script
功能: 自動下載Sensor內儲存的資料並將其寫入控制台附帶Senosr屬性，異常狀態時會回報錯誤訊息並跳出。
* 異常狀態回報訊息說明:
  * 無法正確啟動軟體 : Can't activate ESCORT Console Pro app. Please check the path where your app executable file is saved.
  * 沒有正確進入對應視窗 : Can't find XXX window.
  * 沒有搜尋到Sensor的屬性 :
    * Error: Unable to find the window or read text from it.
    * Please check Battery,COM connect,Sensor connect.
  * Sensor內沒有儲存的資料 :
    * Can't find any data in sensor.
    * Error with empty text.
    * Error with no any text found.
### Configuration Script

