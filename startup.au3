#include <TrayConstants.au3>
#include <Date.au3>
#include <Array.au3>

; See Exit's _SingleScript.au3 --> https://www.autoitscript.com/forum/topic/178681-_singlescript-assure-that-only-one-script-with-the-same-name-is-running/
; Based on my comment at the end of the thread, I'm sure I modified my copy to implement the changes I wanted, so I've included that function now instead
; #include <_SingleScript.au3>

; The default tray menu items will not be shown and items are not checked when selected. These are options 1 and 2 for TrayMenuMode. (Yay, magic numbers!)
AutoItSetOption("TrayMenuMode", 1+2)

; IDK, some title... I'm a programmer, not a namer
Global Const $__g_sTitle = "Startup"

; The config file
Global Const $__g_sConfig = @ScriptDir & "\startup.ini"

; The 2D 0-based array of programs that is loaded from the config
Global $__g_asPrograms[0][0]
; Enum that describes the columns of the program array above
Global Enum $__Name, _
			$__Program, _
			$__Parameters, _
			$__Startup, _
			$__LastRun, _
			$__RunFrequency, _
			$__Directory, _
			$__SIZE

; Set this to true to view data printed in the log and not start programs
Global $__g_bDebugging = True

Main()

Func Main()

	; Set up the tray
	TrayCreateItem("Startup")
	TraySetToolTip("Startup2")
	TrayItemSetState(-1, $TRAY_DISABLE)
	TrayCreateItem("")
	Local $idOpenConfig = TrayCreateItem("Open Config")
	Local $idReloadConfig = TrayCreateItem("Reload Config")
	; Local $idCreateItem = TrayCreateItem("Add Task")
	TrayCreateItem("")
	Local $idStartup = TrayCreateMenu("Startup Tasks")
	Local $idTasks = TrayCreateMenu("Tasks")

	; Get the list of programs to use
	LoadPrograms()

	; Holds the list of IDs to watch. Lets us know which program was clicked
	Local $aTrayItems[UBound($__g_asPrograms)]
	Local $idParent
	; For each program
	For $i=0 To UBound($__g_asPrograms) - 1
		; If it should be launched on startup
		If $__g_asPrograms[$i][$__Startup] Then
			; We're starting, aren't we? Run it
			StartProgram($i)
			$idParent = $idStartup
		Else
			; Other program, goes in Tasks list
			$idParent = $idTasks
		EndIf
		; Create the tray item
		$aTrayItems[$i] = TrayCreateItem($__g_asPrograms[$i][0], $idParent)
	Next

	TrayCreateItem("")
	Local $idExit = TrayCreateItem("Exit")

	; Main Tray Loop
	Local $vMsg
	While True
		$vMsg = TrayGetMsg()
		Switch $vMsg
			Case $TRAY_EVENT_NONE, $TRAY_EVENT_PRIMARYDOWN, $TRAY_EVENT_PRIMARYUP, $TRAY_EVENT_SECONDARYDOWN, $TRAY_EVENT_SECONDARYUP, $TRAY_EVENT_PRIMARYDOUBLE, $TRAY_EVENT_SECONDARYDOUBLE
				; Skip
			Case $idOpenConfig
				ShellExecute($__g_sConfig)
			Case $idReloadConfig
				LoadPrograms()
			Case $idExit
				ExitLoop
			Case Else
				; Okay, hopefully only program events are left at this point
				For $i=0 To UBound($aTrayItems) - 1
					If $aTrayItems[$i] = $vMsg Then
						; Start the program requested
						StartProgram($i, True)
						ExitLoop
					EndIf
				Next
		EndSwitch
	WEnd

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: StartProgram
; Description ...: Starts a program from the global array by index, optionally ignoring if it's already running
; Syntax ........: StartProgram($iIndex[,  $bOverride = False])
; Parameters ....: $iIndex              - 
;                  $bOverride           - [optional] Ignore if program is already running. Default is false.
; Return values .: None
; Author ........: Seadoggie
; Modified ......: June 30, 2023
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func StartProgram($iIndex, $bOverride = Default)

	If IsKeyword($bOverride) Or $bOverride = False Then
		; Don't check if the program is already running
	Else
		; if it's an AutoIt script (simple check, but good enough for me)
		If $__g_asPrograms[$iIndex][$__Program] = @AutoItExe Then
			; If the script is already running, then don't run it
			If _SingleScript(3, $__g_asPrograms[$iIndex][$__Parameters]) Then Return Debug("Already Running: " & $__g_asPrograms[$iIndex][$__Parameters])
		Else
			; If the process is running, then don't run it
			If ProcessExists(FileName($__g_asPrograms[$iIndex][$__Program])) Then Return Debug("Already Running: " & $__g_asPrograms[$iIndex][$__Program])
		EndIf
		
		; If the program has a run freqency And it hasn't been too long
		If ($__g_asPrograms[$iIndex][$__RunFrequency] <> 0) _
		And (_DateDiff("h", $__g_asPrograms[$iIndex][$__LastRun], _NowCalc()) < $__g_asPrograms[$iIndex][$__RunFrequency]) Then
			Return Debug("Run Frequency too close: " & $__g_asPrograms[$iIndex][$__Program] _ 
				& @CRLF & @TAB & "Parameters: " & $__g_asPrograms[$iIndex][$__Parameters] _
				& @CRLF & @TAB & "DateDiff: " & _DateDiff("D", $__g_asPrograms[$iIndex][$__LastRun], _NowCalc()) _
				& @CRLF & @TAB & "RunFrequency: " & $__g_asPrograms[$iIndex][$__RunFrequency] _
				& @CRLF & @TAB & "Last Run: " & $__g_asPrograms[$iIndex][$__LastRun])
		EndIf
	EndIf

	; If we're in debug mode, don't run the program, just debug it and quit
	If $__g_bDebugging Then 
		Return Debug('NAME:' & $__g_asPrograms[$iIndex][$__Name] & ' FILENAME:"' & $__g_asPrograms[$iIndex][$__Program] & '" PARAMETERS:"' & $__g_asPrograms[$iIndex][$__Parameters] & '"')
	EndIf

	; Launch the program
	ShellExecute('"' & $__g_asPrograms[$iIndex][$__Program] & '"', $__g_asPrograms[$iIndex][$__Parameters], $__g_asPrograms[$iIndex][$__Directory])
	If ErrMsg(ShellExecute) Then
		; Let the user know it failed (and don't count it as "running")
		MsgBox($MB_ICONERROR, $__g_sTitle, "Failed to Launch: " + $__g_asPrograms[$iIndex][0])
	Else
		; Record the last run
		IniWrite($__g_sConfig, $__g_asPrograms[$iIndex][$__Name], "LastRun", _NowCalc())
	EndIf

EndFunc

Func FileName($sFile)
	Return StringTrimLeft($sFile, StringInStr($sFile, "\", 0, -1))
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: LoadPrograms
; Description ...: Loads programs from config file into global program array
; Syntax ........: LoadPrograms()
; Parameters ....: None
; Return values .: None
; Author ........: Seadoggie
; Modified ......: June 30, 2023
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func LoadPrograms()

	Local $aProgramNames = IniReadSection($__g_sConfig, "Programs")

	ReDim $__g_asPrograms[UBound($aProgramNames)-1][$__SIZE]

	For $i=1 To UBound($aProgramNames) - 1
		$__g_asPrograms[$i-1][$__Name] = $aProgramNames[$i][0]
		$__g_asPrograms[$i-1][$__Program] = IniRead($__g_sConfig, $aProgramNames[$i][0], "Program", "")
		$__g_asPrograms[$i-1][$__Parameters] = IniRead($__g_sConfig, $aProgramNames[$i][0], "Parameters", "")
		$__g_asPrograms[$i-1][$__Startup] = IniRead($__g_sConfig, $aProgramNames[$i][0], "Startup", "False") = "True" ; Force a boolean variable type
		$__g_asPrograms[$i-1][$__LastRun] = IniRead($__g_sConfig, $aProgramNames[$i][0], "LastRun", "")
		$__g_asPrograms[$i-1][$__RunFrequency] = Number(IniRead($__g_sConfig, $aProgramNames[$i][0], "RunFrequency", ""))
		$__g_asPrograms[$i-1][$__Directory] = IniRead($__g_sConfig, $aProgramNames[$i][0], "Directory", "")
	Next

EndFunc

Func Debug($sText, $sStart = "+ ")

	ConsoleWrite($sStart & $sText & @CRLF)

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: ErrMsg
; Description ...: Prints the error and message to the log. Returns the error.
; Syntax ........: ErrMsg([$sMsg = ""[, $iError = @error[, $iExtended = @extended]]])
; Parameters ....: $sMsg                - [optional] a string value. Default is "".
;                  $iError              - [optional] an integer value. Default is @error.
;                  $iExtended           - [optional] an integer value. Default is @extended.
;                  $iScriptLineNum      - [optional] an integer value. Default is @ScriptLineNumber.
; Return values .: $iError (the one passed in) and preserves @error and @extended
; Author ........: Seadoggie01
; Modified ......: September 30, 2020
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: 
; ===============================================================================================================================
Func ErrMsg($sMsg = "", $iError = @error, $iExtended = @extended, $iScriptLineNum = @ScriptLineNumber)

	; If we've used a Function as the message (it looks really nice when I do) then print the name of the function
	If IsFunc($sMsg) Then $sMsg = FuncName($sMsg)

	; If there is an error, then write the error message
	If $iError Then
		Local $sText = ""
		If StringInStr(@ScriptName, ".au3") Then $sText = '"' & @ScriptFullPath & '" (' & $iScriptLineNum & ',5) : (See below for message)' & @CRLF
		If $iExtended <> 0 Then $sMsg = "Extended: " & $iExtended & " - " & $sMsg
		$sText &= "! Error: " & $iError & " - " & $sMsg
		Debug($sText, "")
	EndIf
	; Preserve the error
	Return SetError($iError, $iExtended, $iError)

EndFunc

;==============================================================================================================
; UDF Name:         SingleScript.au3
; Description:      iMode=0  Close all executing scripts with the same name and continue.
;                   iMode=1  Wait for completion of predecessor scripts with the same name.
;                   iMode=2  Exit if other scripts with the same name are executing.
;                   iMode=3  Test, if other scripts with the same name are executing.
;
; Syntax:           _SingleScript([iMode=0, [$sScriptName = @ScriptName]])
;                   Default:  iMode=0
; Parameter(s):     iMode:     0/1/2/3    see above
;                   sScriptName: The name of the script to test - defaults to this script
; Requirement(s):   none
; Return Value(s): -1= error      @error=-1   invalid iMode
;                   0= no other script executing @error=0 @extended=0
;                   1= other script executing @error=0 @extended=1 (only iMode=3)
; Example:
;               #include <SingleScript.au3>  ; http://www.autoitscript.com/forum/index.php?showtopic=178681
;               _SingleScript() ; Close mode ( iMode defaults to 0 )
;               MsgBox(Default, Default, "No other script with name " & StringTrimRight(@ScriptName, 4) & " is executing.", 0)
;               ; see other example at end of this UDF
;
; Author:       Exit   ( http://www.autoitscript.com/forum/user/45639-exit )
; Modified:     Seadoggie - added optional $sScriptName to check if other scripts are running. Modified around 2021.10.13 Based on version 2021.04.14
; SourceCode:   http://www.autoitscript.com/forum/index.php?showtopic=178681   Version: 2021.04.14
; COPYLEFT:     � 2013 Freeware by "Exit"
;               ALL WRONGS RESERVED
;==============================================================================================================
Func _SingleScript($iMode = 0, $sScriptName = @ScriptName)
    Local $oWMI, $oProcess, $oProcesses, $aHandle, $aError
    Local $sPrefix = StringLeft($sScriptName, StringInStr($sScriptName, ".") - 1)
    Local $sMutexName = "_SingleScript " & $sPrefix
    If $iMode < 0 Or $iMode > 3 Then Return SetError(-1, -1, -1)
    If $iMode = 0 Or $iMode = 3 Then ; (iMode = 0) close all other scripts with the same name.  (iMode = 3) check, if others are running.
        $oWMI = ObjGet("winmgmts:\\" & @ComputerName & "\root\CIMV2")
        If @error Then
            RunWait(@ComSpec & ' /c net start winmgmt  ', '', @SW_HIDE)
            RunWait(@ComSpec & ' /c net continue winmgmt  ', '', @SW_HIDE)
            $oWMI = ObjGet("winmgmts:\\" & @ComputerName & "\root\CIMV2")
        EndIf
        $oProcesses = $oWMI.ExecQuery("SELECT * FROM Win32_Process", "WQL", 0x30)
        For $oProcess In $oProcesses
            If $oProcess.ProcessId = @AutoItPID Then ContinueLoop
            If ($oProcess.Name = "AutoIt3.exe" And StringInStr($oProcess.CommandLine, "AutoIt3Wrapper")) Then ContinueLoop
            If Not (StringInStr($oProcess.Name & $oProcess.CommandLine, $sPrefix)) Then ContinueLoop
            If $iMode = 3 Then Return SetError(0, 1, 1) ; indicate other script is running. Return value and @extended set to 1.
            Sleep(1000) ; allow previous process to terminate
            If ProcessClose($oProcess.ProcessId) Then ContinueLoop
            MsgBox(262144, "Debug " & $sScriptName, "Error: " & @error & " Extended: " & @extended & @LF & "SingleScript Processclose error: " & $oProcess.Name & @LF & "******", 5)
        Next
    EndIf
    $aHandle = DllCall("kernel32.dll", "handle", "CreateMutexW", "struct*", 0, "bool", 1, "wstr", $sMutexName) ; try to create Mutex
    $aError = DllCall("kernel32.dll", "dword", "GetLastError") ; retrieve last error
    If Not $aError[0] Then Return SetError(0, 0, 0)
    If $iMode = "2" Then Exit 1
    If $iMode = "0" Then Return SetError(1, 0, 1) ; should not occur
    DllCall("kernel32.dll", "dword", "WaitForSingleObject", "handle", $aHandle[0], "dword", -1) ; infinite wait for lock
    Return SetError(0, 0, 0)
EndFunc   ;==>_SingleScript