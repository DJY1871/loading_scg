#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>

$windowTitle = "ESCORT Console Pro 2.12.07"

;Run 預設路徑執行檔
If  Not WinExists($windowTitle) Then
	Run("C:\Program Files (x86)\Escort DLS\Escort Console\Escort.exe")
EndIf
WinWait($windowTitle, "", 15)

#cs
;Creat a Warning GUI
$hGUI = GUICreate("Warning", 300, 100, -1, -1, $WS_POPUPWINDOW, $WS_EX_TOPMOST + $WS_EX_TOOLWINDOW)
GUISetState(@SW_SHOW)
GUICtrlCreateLabel("Automation Script is running now, please don't move!!!", 10, 40, 280, 20)
#ce

If WinExists("[CLASS:#32770; TITLE:ESCORT Console]") Then
    ControlClick("ESCORT Console", "", "[CLASS:Button; INSTANCE:2]")
EndIf

If WinExists($windowTitle) Then
    ControlSend($windowTitle, "", "", "{F7}")
    Sleep(1000)
    $downloadWindowTitle = "Download Readings"
	ControlClick($downloadWindowTitle, "", "[CLASS:Button; INSTANCE:1]")
    WinWait($downloadWindowTitle, "", 15)
    Local $i = 1
    While $i <= 100
        Sleep(1000)
		ControlClick($downloadWindowTitle, "", "[CLASS:AfxWnd70s; INSTANCE:1]")
        ControlSend($windowTitle, "", "", "{ENTER}")
        Sleep(500)
        If WinExists("ESCORT iLog/Qualaire Properties") Then
            ExitLoop
        EndIf
        $i += 1
    WEnd
	
    $propertiesWindowTitle = "ESCORT iLog/Qualaire Properties"
    Sleep(500)
    $visibleText = WinGetText($propertiesWindowTitle)
	;print property of device
    If $visibleText <> "" Then
        $serialNumber = StringRegExpReplace($visibleText, "(?s).*Serial Number:\s*([^\r\n]*).*", "\1")
        ConsoleWrite("Serial Number: " & $serialNumber & @CRLF)
        
        $batteryStatus = StringRegExpReplace($visibleText, "(?s).*Battery Status:\s*([^\r\n]*).*", "\1")
        $description = StringRegExpReplace($visibleText, "(?s).*Description:\s*([^\r\n]*).*", "\1")
        $productcode = StringRegExpReplace($visibleText, "(?s).*Product Code:\s*([^\r\n]*).*", "\1")
        ConsoleWrite("Battery Status: " & $batteryStatus & @CRLF)
        ConsoleWrite("Description: " & $description & @CRLF)
        ConsoleWrite("Product Code: " & $productcode & @CRLF)
        ; MsgBox(0, "Status", "Serial Number: " & $serialNumber & " Battery Status: " & $batteryStatus)
        ControlSend($propertiesWindowTitle, "", "", "{ENTER}")
    Else
        ConsoleWrite("Error: Unable to find the window or read text from it." & @CRLF)
    EndIf
EndIf