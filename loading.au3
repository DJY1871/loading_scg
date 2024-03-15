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
    WinWait($downloadWindowTitle, "", 10)
    ControlClick($downloadWindowTitle, "", "[CLASS:Button; INSTANCE:1]")
    Local $i = 1
    While $i <= 100
        Sleep(1000)
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

    If WinExists($downloadWindowTitle) Then
        $controlID = 1004
        If ControlCommand($downloadWindowTitle, "", $controlID, "IsVisible", "") Then
            
            $buttonText = ControlGetText($downloadWindowTitle, "", $controlID)
			
			;click download button
            If StringInStr($buttonText, "Download", 2) Then
                ControlClick($downloadWindowTitle, "", "[CLASS:Button; INSTANCE:1]")
            EndIf
			
			;click Next button
            While $i <= 14000
                $buttonText = ControlGetText($downloadWindowTitle, "", $controlID)
                Sleep(500)
                If StringInStr($buttonText, "Next", 2) Then
                    $isEnabled = ControlCommand($downloadWindowTitle, "", $controlID, "IsEnabled", "")
    
                    If $isEnabled Then
                        ControlClick($downloadWindowTitle, "", "[CLASS:Button; INSTANCE:1]")
                        ExitLoop
                    EndIf 
                EndIf
                $i += 1
            WEnd

            $buttonText = ControlGetText($downloadWindowTitle, "", $controlID)
            Sleep(500)
            If StringInStr($buttonText, "Finish", 2) Then
                $isEnabled = ControlCommand($downloadWindowTitle, "", $controlID, "IsEnabled", "")
                If $isEnabled Then
                    ControlClick($downloadWindowTitle, "", "[CLASS:Button; INSTANCE:1]")
                    Sleep(500)
                EndIf 
            EndIf
        EndIf
    EndIf

    $windowTitlePattern = "ESCORT Console Pro 2.12.07 - " & $serialNumber & "-"
    $windows = WinList()
    $found = False
    Sleep(1000)
    For $i = 1 To $windows[0][0]
        If StringInStr($windows[$i][0], $windowTitlePattern) > 0 Then
            
            $found = True
            ControlSend($windows[$i][0], "", "", "^c")
            Sleep(500)
            $clipContent = ClipGet()

            If $clipContent <> "" Then
                ConsoleWrite($clipContent & @CRLF)
            Else
                ConsoleWriteError("Error with empty text" & @CRLF)
            EndIf
            WinClose($windowTitlePattern)

            WinWait("ESCORT Console", "", 10)
            If WinExists("[CLASS:#32770; TITLE:ESCORT Console]") Then
                ControlClick("ESCORT Console", "否(&N)", "[CLASS:Button; INSTANCE:2]")
            EndIf

            ExitLoop
        EndIf
    Next

    If Not $found Then
        ConsoleWriteError("Error with no any text found" & @CRLF)
    EndIf
    ; MsgBox(0, "","")
    WinClose($windowTitle)

EndIf
