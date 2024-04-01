#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <GuiListView.au3>

$windowTitle = "ESCORT Console Pro 2.12.07"

;Run 預設路徑執行檔
If  Not WinExists($windowTitle) Then
	$Runerrro = Run("C:\Program Files (x86)\Escort DLS\Escort Console\Escort.exe")
	if $Runerrro = 0 Then
		ConsoleWrite("Can't activate ESCORT Console Pro app. Please check the path where your app executable file is saved." & @CRLF)
		Exit
    EndIf
EndIf
WinWait($windowTitle, "", 15)
WinSetState($windowTitle,"",@SW_MINIMIZE)

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

if WinExists($windowTitle) Then
    ;maxsize window and lock 
    WinSetState($windowTitle,"",@SW_MAXIMIZE)
    WinSetState($windowTitle,"",@SW_LOCK)
    Sleep(200)
    ;open configwindow
	ControlSend($windowTitle,"","","{F6}")
	Sleep(1000)
	$configWindowTitle = "Program and Configure"
	WinWait($configWindowTitle,"",10)
	if WinExists($configWindowTitle) Then
        If ControlCommand($configWindowTitle, "", 1004, "IsEnabled", "") Then
            ControlClick($configWindowTitle, "", "[CLASS:Button; INSTANCE:1]")
            Sleep(200)
        EndIf
        ;call properties window
        Local $i = 1
        While $i <= 10
            Sleep(300)
            Sleep(200)
            ;double click list view
            HotKeySet("{Esc}","hotkeyuse")
            Local $hListView = ControlGetHandle($configWindowTitle, "", "[CLASS:SysListView32]")
			_GUICtrlListView_ClickItem($hListView, 0)
			_GUICtrlListView_ClickItem($hListView, 0)
            Sleep(500)
            If WinExists("ESCORT iLog/Qualaire Properties") Then
                ExitLoop
            EndIf
            $i += 1
        WEnd
        ;unlock window exit maxsize
        WinSetState($windowTitle,"",@SW_UNLOCK)
        WinSetState($windowTitle,"",@SW_RESTORE)
    Else
        ConsoleWrite("Can't find " & $configWindowTitle & " window"& @CRLF)
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
        for $i  = 1 to 10
            ControlClick($propertiesWindowTitle,"OK",1)
            if WinExists($propertiesWindowTitle) = 0 Then
                ExitLoop
            EndIf
        Next
        if WinExists($propertiesWindowTitle) Then
            ConsoleWrite("Can't close " & $propertiesWindowTitle & " window")
        EndIf
    Else
        ConsoleWrite("Error: Unable to find the window or read text from it." & @CRLF)
        ConsoleWrite("Please check Battery,COM connect,Sensor connect." & @CRLF)
        CloseAppWindows($windowTitle)
		Exit
    EndIf

    if $description <> $cmdline[16] Then
        ConsoleWrite("description not fit old is " & $cmdline[16] & ",new is " & $description & @CRLF)
        CloseAppWindows($windowTitle)
        Exit
    EndIf
    ;set description
    IF WinExists($configWindowTitle) Then
        ;click next to details
        ControlClick($configWindowTitle,"","[CLASS:Button; INSTANCE:1]")
        Sleep(500)
        $description = $cmdline[1]
        if $description <> "def" Then
            ControlSetText($configWindowTitle,"","[CLASS:Edit; INSTANCE:1]",$description)
        EndIf
        ;click next to sensor display
        ControlClick($configWindowTitle,"","[CLASS:Button; INSTANCE:1]")
        Sleep(500)
    Else
        ConsoleWrite("Can't find " & $configWindowTitle & "window"& @CRLF)
        CloseAppWindows($windowTitle)
        Exit
    EndIf
    
    ;set sensors display
    $buttonid = 1004
    $configsnesor = "Configure Sensor"
    If WinExists($configWindowTitle) Then
        for $i = 1 to 5
            ControlClick($configWindowTitle,"",5182)
            If WinExists($configsnesor) Then
                ExitLoop
            EndIf
        Next
        If WinExists($configsnesor) Then
            ControlClick($configsnesor,"",$buttonid)
            Sleep(500)
            ;set temperature
            $mintemper = $cmdline[2]
            $maxtemper = $cmdline[3]
            $increment = $cmdline[4]
            if $mintemper <> "def" Then
                ControlSetText($configsnesor,"", "[CLASS:Edit; INSTANCE:2]", $mintemper)
                Sleep(200)
                ControlSetText($configsnesor,"", "[CLASS:Edit; INSTANCE:3]", $maxtemper)
                Sleep(200)
                ControlSetText($configsnesor,"", "[CLASS:Edit; INSTANCE:4]", $increment)
                Sleep(200)
            EndIf
            ControlClick($configsnesor,"",$buttonid)
            Sleep(500)
            ;set alarm
            local $Alchose[8]
            $Alchose = StringSplit($cmdline[5],"_")
            if $Alchose[1] <> "def" Then
                For $i = 5 To 11
                    If $Alchose[$i - 4] = "T" Then
                        ControlClick($configsnesor, "", "[CLASS:Button; INSTANCE:" & $i & "]")
                        Sleep(200)
                    EndIf
                Next
                ControlClick($configsnesor, "","[CLASS:Button; INSTANCE:7]")
                ControlSend($configsnesor,"","[CLASS:Edit; INSTANCE:1]",$cmdline[6])
                ControlSend($configsnesor,"","[CLASS:Edit; INSTANCE:2]",$cmdline[7])
            EndIf
            ;exit configsensor window
            for $i = 1 to 10
                ControlClick($configsnesor,"",$buttonid)
                Sleep(200)
                if WinExists($configsnesor) = 0 Then
                    ExitLoop
                EndIf
            Next
            Sleep(500)
            if WinExists($configsnesor) Then
                ConsoleWrite("Can't close " & $configsnesor & " window")
                CloseAppWindows($windowTitle)
                Exit
            EndIf
        Else
            ConsoleWrite("Can't find " & $configsnesor & "window"& @CRLF)
            CloseAppWindows($windowTitle)
        EndIf
    Else
        ConsoleWrite("Can't find " & $configWindowTitle & "window"& @CRLF)
        CloseAppWindows($windowTitle)
        Exit
    EndIf

    ;to sampling time setting window
    ControlClick($configWindowTitle,"",$buttonid)
    ControlClick($configWindowTitle,"",$buttonid)

    ;set sampling time
    If WinExists($configWindowTitle) Then
        ;Duration time
        ControlCommand($configWindowTitle,"",5244,"SelectString",$cmdline[8])
        Sleep(50)
        ;set reading interval days
        Local $readinterval[2]
        $temp = StringReplace($cmdline[9],"days"," ")
        $readinterval = StringSplit($temp," ")
        ControlSend($configWindowTitle,"",5221,$readinterval[1])
        Sleep(50)
        ControlSend($configWindowTitle,"",5223,$readinterval[2])
        Sleep(50)
        ;start log time
        Switch  $cmdline[10]
            ;set wait hr min
            Case 1
                ControlClick($configWindowTitle,"","[CLASS:Button; INSTANCE:6]")
                Sleep(50)
                ControlSend($configWindowTitle,"",5227,$cmdline[11])
                Sleep(50)
            ;set in time day hr min
            Case 2
                ;devide xxdaysxxhxxm
                Local $starttime[2]
                $temp = StringReplace($cmdline[11],"days"," ")
                $starttime = StringSplit($temp," ")
                ControlClick($configWindowTitle,"","[CLASS:Button; INSTANCE:7]")
                Sleep(50)
                ControlSetText($configWindowTitle,"",5229,$starttime[1])
                Sleep(50)
                ControlSend($configWindowTitle,"",5231,$starttime[2])
                Sleep(50)
            ;set log at precise time
            Case 3
                $day = $cmdline[11]
                ControlClick($configWindowTitle,"","[CLASS:Button; INSTANCE:8]")
                If StringInStr($day, "上") Then
                    ControlSend($configWindowTitle,"",5232,$day)
                    For $i = 1 To 2
                        ControlSend("Program and Configure","",5232,"{left}")
                        Sleep(50)
                    Next
                    ControlSend($configWindowTitle,"",5232,"上")
                Else
                    ControlSend($configWindowTitle,"",5232,$day)
                    For $i = 1 To 2
                        ControlSend("Program and Configure","",5232,"{left}")
                        Sleep(50)
                    Next
                    ControlSend($configWindowTitle,"",5232,"下")
                EndIf
            Case Else
                ConsoleWrite("Error")
        EndSwitch

        ;finish log time
        Switch  $cmdline[12]
            ;set after reading time
            Case 1
                ControlClick($configWindowTitle,"","[CLASS:Button; INSTANCE:9]")
                Sleep(50)
                ControlSetText($configWindowTitle,"",5237,$cmdline[13])
                Sleep(50)
            ;set after day hr/min
            Case 2
                ;devide xxdaysxxhxxm
                Local $aftertime[2]
                $temp = StringReplace($cmdline[13],"days"," ")
                $aftertime = StringSplit($temp," ")
                ControlClick($configWindowTitle,"","[CLASS:Button; INSTANCE:10]")
                Sleep(50)
                ControlSend($configWindowTitle,"",5239,$aftertime[1])
                Sleep(50)
                ControlSend($configWindowTitle,"",5241,$aftertime[2])
                Sleep(50)
            ;set log at precise time
            Case 3
                $day = $cmdline[13]
                ControlClick($configWindowTitle,"","[CLASS:Button; INSTANCE:11]")
                If StringInStr($day, "上") Then
                    ControlSend($configWindowTitle,"",5242,$day)
                    For $i = 1 To 2
                        ControlSend("Program and Configure","",5232,"{left}")
                        Sleep(50)
                    Next
                    ControlSend($configWindowTitle,"",5242,"上")
                Else
                    ControlSend($configWindowTitle,"",5242,$day)
                    For $i = 1 To 2
                        ControlSend("Program and Configure","",5242,"{left}")
                        Sleep(50)
                    Next
                    ControlSend($configWindowTitle,"",5242,"下")
                EndIf
            Case 4
                ControlClick($configWindowTitle,"",5236)
            Case Else
                ConsoleWrite("Can't find " & $configWindowTitle & "window"& @CRLF)
        EndSwitch
    Else
        ConsoleWrite("Can't find " & $windowTitle & "window"& @CRLF)
        CloseAppWindows($windowTitle)
        Exit
    EndIf

    ;enable beeper in logger
    If $cmdline[14] = "F" Then 
        ControlClick($configWindowTitle,"",5243)
        Sleep(50)
    EndIf
    ;battary fitted
    If $cmdline[15] = "T" Then 
        ControlClick($configWindowTitle,"",5245)
        Sleep(50)
    EndIf

    ;programing
    If WinExists($configWindowTitle) Then
        for $i = 1 to 200
            Sleep(500)
            ControlClick($configWindowTitle,"",$buttonid)
            if WinExists($configWindowTitle) = 0 Then
                ExitLoop
            EndIf
        Next
    Else
        ConsoleWrite("Can't find " & $windowTitle & "window"& @CRLF)
        CloseAppWindows($windowTitle)
        Exit
    EndIf
    Sleep(500)
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

Func hotkeyuse()
    ConsoleWrite("Exit searching window" & @CRLF)
    CloseAppWindows($windowTitle)
    Exit
EndFunc

;how to use exe cmdline
#cs 
    configuration.exe
    1.description
    2.mintemper                       不變的話第一個改def
    3.maxtemper
    4.increment
    5.AlChoose EX~ "T F T F T T T"    不變的話第一個改def
    6.Aloutofspec1
    7.Aloutofspec2
    8.Duration
    9.readinterval    EX~ 03days10h21m06s 
    10.startlogchoose EX~ 1
    11.starttime      EX~ 03h02m    or  03days10h21m    or  "2045/3/4 上 03:22"
    12.finishlogchoose EX~ 3
    13.aftertime      EX~ 10   or  03days10h21m    or  "2045/3/4 上 03:22"
    14.enable beep EX~ "T
    15.new battery fitted EX~ "T
    16.old_description
#ce
;EX~ configuration.exe def -20 70 5 F_F_F_F_F_F_F 0 0 U 00days00h00m06s 2 00days00h01m 1 10 T F 39-78
;"C:\Program Files (x86)\AutoIt3\Aut2Exe\Aut2exe_x64.exe" /In "configuration.au3" /x64 /console