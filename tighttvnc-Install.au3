#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseX64=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <AutoItConstants.au3>
#include <Date.au3>
#include <FileConstants.au3>

Local $user = ""
Local $password = ""
Local $domain = ""

main()

Func main()
	;IF statement to check if TightVNC is installed or not
	;CASE TRUE: Breaks execution as intended use as been reached
	;CASE FALSE: Starts script execution
	If Not isVNCInstalled() Then
		validateOS()
	Else
		Break
	EndIf
EndFunc

;Install VNC on a 64bit machine
Func installVNC64()
	;COPIA DE FICHEIROS
	DirCopy("\\hdsdom.root\NETLOGON\tightvnc\script", "c:\temp\script",1)
	;DESINSTALAR UltraVNC
	RunAsWait($user, $domain, $password, 0, "C:\temp\script\UninstallVNC.bat", "", @SW_HIDE)
	;INSTALAR TightVNC
	RunAswait($user, $domain, $password, 0, "msiexec.exe /i c:\temp\script\tightvnc64.msi ADDLOCAL=Server /qn /norestart", "", @SW_HIDE)
	;REGISTO COM CONFIGURAÇÕES E REINICO DE SERVIÇOS
	RunAsWait($user, $domain, $password, 0, "C:\temp\script\regbatch.bat","", @SW_HIDE)
	;REMOVER PASTA DE SCRIPTS
	DirRemove("C:\temp\script",1)
EndFunc

;Inbstall VNC on a 32bit machine
Func installVNC32()
	;COPIA DE FICHEIROS
	DirCopy("\\hdsdom.root\NETLOGON\tightvnc\script", "c:\temp\script",1)
	;DESINSTALAR UltraVNC
	RunAsWait($user, $domain, $password, 0, "C:\temp\script\UninstallVNC.bat", "", @SW_HIDE)
	;INSTALAR TightVNC
	RunAswait($user, $domain, $password, 0, "msiexec.exe /i c:\temp\script\tightvnc32.msi ADDLOCAL=Server /qn /norestart", "", @SW_HIDE)
	;REGISTO COM CONFIGURAÇÕES E REINICO DE SERVIÇOS
	RunAsWait($user, $domain, $password, 0, "C:\temp\script\regbatch.bat","", @SW_HIDE)
	;REMOVER PASTA DE SCRIPTS
	DirRemove("C:\temp\script",1)
EndFunc

;Validate Os Architecture and install corresponding VNC version
Func installVNC()
	Dim $Arch = @OSArch
	If $Arch = "X64" Then
		installVNC64()
	ElseIf $Arch ="X86" Then
		installVNC32()
	EndIf
	createFlag()
EndFunc

;Valide OS Version and Architecture to decide which install function to use
Func validateOS()
	Dim $OS = @OSVersion
;In case of a Windows XP machine: script skips execution
;In case of a Windows 7/8/10 machine: script executes
	Switch $OS
		Case "WIN_10"
			installVNC()
		Case "WIN_7"
			installVNC()
		Case "WIN_8"
			installVNC()
		Case Else
			writeFile()
			Break
	EndSwitch
EndFunc

;Write to file located in a shared folder accessible by any user authenticated in the network with informetion of when script was executed and machine info
Func writeFile()
	;Variable holding the name of the computer where script is executed
	Local $name = @ComputerName
	;Variable defining the path where file with Ip and name info will be created
	Local $path = "\\172.16.110.201\tightVNC\" & $name & "-" & _NowDate() & ".txt"
	;Variable to append all posiible IPs of the machine where the script is executed
	Local $ip = @IPAddress1 & " - " & @IPAddress2 & " - " & @IPAddress3 & " - " & @IPAddress4 & " - "
	;Variable to hold date and time of execution of script
	Local $timedate = _NowDate() & "-" & _NowTime() & "->"

	;Write to the file time and date of script execution
	If Not FileWrite($path, $timedate) Then
		Return False
	EndIf

	;Create a file handler variable to control file execution in append mode
	Local $filehandler = FileOpen($path, $FO_APPEND)
	If $filehandler = -1 Then
		Return False
	EndIf

	;Write IP information to the file
	FileWrite($path, $ip & @CRLF)

	;Close file
	FileClose($filehandler)
EndFunc

;Create the flag that is going to be checked in case of a seconf execution in the same machine
Func createFlag()
	;Variable to be writetn into flag file to display date and time of execution
	Local $timedate = "DONE AT ->" & _NowDate() & "-" & _NowTime()
	;Variable holding info of path to flag
	Local $flagpath = "c:\temp\tightOK.txt"
	;Create file handler variable and file at desired location
	Local $flaghandler = FileOpen($flagpath, $FO_APPEND)
	If $flaghandler = -1 Then
		Return False
	EndIf

	;Write to the file the date and time of execution
	FileWrite($flagpath, $timedate & @CRLF)

	;Close file
	FileClose($flaghandler)
EndFunc

;Check if TightVNC is already installed on the machine
Func isVNCInstalled()
	;Variable to hold array of all installed software on the machine
	Local $installedPrograms = getListInstalledSoftware()

	;IF statement to validate if the list was obtained successfuly or not. IF met halts execution and returns false
	If $installedPrograms == -1 Then
		Return False
	EndIf

	;FOR LOOP to run trough every instance of the list of installed software and compare to the set string holding the name of the software to check
	For $program in $installedPrograms
		;IF the string comparission condition is met LOOP halts execution and returns true
		If StringLower($program.Caption) == StringLower("TightVNC") Then
			Return True
		EndIf
	Next

	;FOR LOOP did not did not find any instance where software was installed and returns false
	Return False
EndFunc

;Function to store all softaware installed on machine and store in array $listOfPrograms
Func getListInstalledSoftware()
	$wbemFlagReturnImmediately = 0x10
	$wbemflagForwardOnly = 0x20
	$listOfPrograms = ""

	$objWMIService = ObjGet("winmgmts:\\localhost\root\CIMV2")
	$listOfPrograms = $objWMIService.ExecQuery("SELECT * FROM Win32_Product", "WQL", $wbemFlagReturnImmediately + $wbemflagForwardOnly)

	;IF array $listOfPrograms is an object returns it to the instance when it was called
	If IsObj($listOfPrograms) Then
		Return $listOfPrograms
	EndIf

	;IF statement was not met and returns -1 to be handeled by an IF satement in another function
	Return -1
EndFunc
