#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>

$windowTitle = "ESCORT Console Pro 2.12.07"
$Runerrro = 0

;Run 預設路徑執行檔
If  Not WinExists($windowTitle) Then
	$Runerrro = Run("C:\Program Files (x86)\Escort DLS\Escort Console\Escort.exe")
	if $Runerrro = 0 Then
		ConsoleWrite("Can't activate ESCORT Console Pro app. Please check the path where your app executable file is saved." & @CRLF)
		Exit
    EndIf
EndIf
;WinWait($windowTitle,"",20)
WinSetState(WinWait($windowTitle,"",20), "", @SW_HIDE)

;Creat a Warning GUI
$hGUI = GUICreate("Warning", 300, 100, -1, -1, $WS_POPUPWINDOW, $WS_EX_TOPMOST + $WS_EX_TOOLWINDOW)
GUISetState(@SW_SHOW)
GUICtrlCreateLabel("Automation Script is running now, please don't move!!!", 10, 40, 280, 20)

Sleep(2000)
GUIDelete()

If WinExists("[CLASS:#32770; TITLE:Tip of the Day]") Then
    ControlClick("Tip of the Day", "", "[CLASS:Button; INSTANCE:4]")
EndIf

If WinExists("[CLASS:#32770; TITLE:ESCORT Console]") Then
    ControlClick("ESCORT Console", "", "[CLASS:Button; INSTANCE:2]")
EndIf

If WinExists($windowTitle) Then
    ;call downloadWindow
    ControlSend($windowTitle, "", "", "{F7}")
    Sleep(1000)
    $downloadWindowTitle = "Download Readings"
    ;WinWait($downloadWindowTitle, "", 10)
	WinSetState(WinWait($downloadWindowTitle, "", 10), "", @SW_HIDE)
    if WinExists($downloadWindowTitle) Then
        ControlClick($downloadWindowTitle, "", "[CLASS:Button; INSTANCE:1]")
        ;call properties window
        Local $i = 1
        While $i <= 10
            Sleep(500)
            ControlClick($downloadWindowTitle,"","CLASS:AfxWnd70s; INSTANCE:1]")
            ControlSend($downloadWindowTitle, "", "", "{ENTER}")
            Sleep(500)
            If WinExists("ESCORT iLog/Qualaire Properties") Then
				WinSetState(WinExists("ESCORT iLog/Qualaire Properties"), "", @SW_HIDE)
                ExitLoop
            EndIf
            $i += 1
        WEnd
    Else
        ConsoleWrite("Can't find " & $downloadWindowTitle & " window"& @CRLF)
        CloseAppWindows($windowTitle)
        Exit
	EndIf

    $propertiesWindowTitle = "ESCORT iLog/Qualaire Properties"
    Sleep(500)
    $visibleText = WinGetText($propertiesWindowTitle)
	
	;print property of device
    If $visibleText <> "" Then
        ;Parse the property string
        $serialNumber = StringRegExpReplace($visibleText, "(?s).*Serial Number:\s*([^\r\n]*).*", "\1")
        $batteryStatus = StringRegExpReplace($visibleText, "(?s).*Battery Status:\s*([^\r\n]*).*", "\1")
        $description = StringRegExpReplace($visibleText, "(?s).*Description:\s*([^\r\n]*).*", "\1")
        $productcode = StringRegExpReplace($visibleText, "(?s).*Product Code:\s*([^\r\n]*).*", "\1")

        ;Write property in console
        ConsoleWrite("Serial Number: " & $serialNumber & @CRLF)
        ConsoleWrite("Battery Status: " & $batteryStatus & @CRLF)
        ConsoleWrite("Description: " & $description & @CRLF)
        ConsoleWrite("Product Code: " & $productcode & @CRLF)
        ; MsgBox(0, "Status", "Serial Number: " & $serialNumber & " Battery Status: " & $batteryStatus)
        ControlSend($propertiesWindowTitle, "", "", "{ENTER}")
    Else
        ConsoleWrite("Error: Unable to find the window or read text from it." & @CRLF)
        ConsoleWrite("Please check Battery,COM connect,Sensor connect." & @CRLF)
        CloseAppWindows($windowTitle)
		Exit
    EndIf

    ;go to download page
    If WinExists($downloadWindowTitle) Then
        $controlID = 1004
        If ControlCommand($downloadWindowTitle, "", $controlID, "IsEnabled", "") Then
            
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
        Else
            ConsoleWrite("Can't find any data in sensor" & @CRLF)
            CloseAppWindows($windowTitle)
            Exit
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
    CloseAppWindows($windowTitle)
Else
    ConsoleWrite("Can't find " & $windowTitle & "window"& @CRLF)
    CloseAppWindows($windowTitle)
    Exit
EndIf


Func CloseAppWindows($appWindowTitle)
    ; 获取当前所有窗口的列表
    Local $aWinList = WinList()

    ; 遍历窗口列表，关闭属于目标应用程序的所有窗口
    For $i = 1 To $aWinList[0][0]
        If StringInStr(WinGetTitle($aWinList[$i][1]), $appWindowTitle) Then
            WinClose($aWinList[$i][1])
        EndIf
    Next
EndFunc