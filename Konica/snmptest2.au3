#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=KM-icon.ico
#AutoIt3Wrapper_Outfile=Konica-Totals.exe
#AutoIt3Wrapper_Res_Description=Grabs page totals from Konica Minolta's
#AutoIt3Wrapper_Res_Fileversion=1.0.0.1
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_LegalCopyright=Richard Easton 2014
#AutoIt3Wrapper_Run_Tidy=y
#AutoIt3Wrapper_Tidy_Stop_OnError=n
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <array.au3>
#include <date.au3>
#include <String.au3>
#include 'SNMPUDF.au3'


Global $Port = 161 ; UDP 161 = SNMP port
Global $SNMP_Version = 2 ; SNMP v2c (1 for SNMP v1)
Global $SNMP_Community = "public" ; SNMPString(Community) (change it)
Global $SNMP_ReqID = 1
Global $SNMP_Command
Global $Start = 1
Global $result
Global $Timeout_msec = 2500


If Not @Compiled Then Opt("trayicondebug", 1)

$ini = @ScriptDir & "\devices.ini"
$sect = "printer"

$aPrinter = IniReadSection($ini, $sect)
$result = "Meter reading of the Konica printer from: " & _NowCalc() & @CRLF


UDPStartup()

For $i = 1 To $aPrinter[0][0]
	$IP = $aPrinter[$i][0]
	$var = Ping($IP, 250)
	$Socket = UDPOpen($aPrinter[$i][0], $Port)
	ConsoleWrite(@LF & $IP & " " & $aPrinter[$i][1] & @LF)
	$result &= $aPrinter[$i][1]
	$OIDsect = $aPrinter[$i][1]
	$OIDsect = StringTrimLeft($OIDsect, (StringInStr($OIDsect, " ", 0, -1)))
	$aOID = IniReadSection($ini, $OIDsect)
	FileWriteLine(@DesktopDir & "\results-" & @MDAY & @MON & "-" & @HOUR & @MIN & ".txt", @CRLF & $aPrinter[$i][1])
	For $x = 1 To $aOID[0][0]
		$SNMP_ReqID += 1
		$SNMP_Command = _SNMPBuildPacket($aOID[$x][1], $SNMP_Community, $SNMP_Version, $SNMP_ReqID, "A0")
		UDPSend($Socket, $SNMP_Command)
		_StartListener()
		Sleep(200)
		$result &= @TAB & $aOID[$x][0] & @TAB & $SNMP_Util[1][0] & @CRLF
		;_ArrayDisplay($SNMP_Util, $aPrinter[$i][1])
		ConsoleWrite($aOID[$x][0] & " = " & $SNMP_Util[1][1] & @LF)
		FileWriteLine(@DesktopDir & "\results-" & @MDAY & @MON & "-" & @HOUR & @MIN & ".txt", $aOID[$x][0] & " = " & $SNMP_Util[1][1])
	Next
	UDPCloseSocket($Socket)
Next
MsgBox(64, "Konica Totals", "All Devices Processed", 10)

UDPShutdown()
Sleep(500)


ConsoleWrite($result & @LF)





Func _StartListener()
	If $Start = 1 Then
		$Timeout = TimerInit()
		While (1)
			$srcv = UDPRecv($Socket, 2048)
			If ($srcv <> "") Then
				$result = _ShowSNMPReceived($srcv)
				;ConsoleWrite(stringmid(_HexToString($srcv), 47) & @CRLF)
				;_ArrayDisplay($result)
				ExitLoop
			EndIf
			Sleep(200)
			If TimerDiff($Timeout) > $Timeout_msec Then
				ExitLoop
			EndIf
		WEnd
	EndIf
EndFunc   ;==>_StartListener

Func OnAutoItExit()
	UDPCloseSocket($Socket)
	UDPShutdown()
EndFunc   ;==>OnAutoItExit
