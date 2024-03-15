#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>

$windowTitle = "ESCORT Console Pro 2.12.07"

;Run 預設路徑執行檔
If  Not WinExists($windowTitle) Then
	Run("C:\Program Files (x86)\Escort DLS\Escort Console\Escort.exe")
EndIf
WinWait($windowTitle, "", 15)

If WinExists("[CLASS:#32770; TITLE:ESCORT Console]") Then
    ControlClick("ESCORT Console", "", "[CLASS:Button; INSTANCE:2]")
EndIf

if WinExists($windowtitle) Then
	ControlSend($windowtitle,"","","{F6}")
	Sleep(1000)
	$configWindowTitle = "Program and Configure"
	