#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=KM-icon.ico
#AutoIt3Wrapper_Outfile=Konica-Totals.exe
#AutoIt3Wrapper_Res_Description=Grabs page totals from SMNP devices
#AutoIt3Wrapper_Res_Fileversion=1.0.2.5
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_LegalCopyright=Richard Easton 2014
#AutoIt3Wrapper_Res_requestedExecutionLevel=None
#AutoIt3Wrapper_Run_Tidy=y
#AutoIt3Wrapper_Tidy_Stop_OnError=n
#EndRegion

#include <array.au3>
#include <date.au3>
#include <String.au3>
#include 'SNMPUDF.au3'
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <ListViewConstants.au3>
#include <StaticConstants.au3>
#include <TabConstants.au3>
#include <WindowsConstants.au3>
#include <Excel.au3>
#include <FileConstants.au3>
#include <GuiListView.au3>


Global $Port = 161 ; UDP 161 = SNMP port
Global $SNMP_Version = 2 ; SNMP v2c (1 for SNMP v1)
Global $SNMP_Community = "public" ; SNMPString(Community) (change it)
Global $SNMP_ReqID = 1
Global $SNMP_Command
Global $Start = 1
Global $result
Global $Timeout_msec = 10000


If Not @Compiled Then Opt("trayicondebug", 1)

$ini = @ScriptDir & "\devices.ini"
$sect = "printer"

$aPrinter = IniReadSection($ini, $sect)


;gui
$KT_V2 = GUICreate("Konica Totals", 615, 600, -1, -1)
GUISetIcon("J:\Extra User Areas\RIE\Programming\Konica Totals\KM-icon.ico", -1)
$Tab1 = GUICtrlCreateTab(0, 0, 617, 601)
$TAB_KT = GUICtrlCreateTabItem("Konica Totals")

GUICtrlCreateLabel("Information", 16, 32, 56, 17)
$Info = GUICtrlCreateListView("Printer Name|IP Address|Description|Serial Number|Total Pages", 8, 48, 601, 521, $LVS_EX_GRIDLINES + $LVS_EX_FULLROWSELECT + $LVS_SORTDESCENDING)
$Process = GUICtrlCreateButton("Process", 8, 570, 105, 25)
GUICtrlSetBkColor(-1, 0x00ff00)
$Export = GUICtrlCreateButton("Export", 120, 570, 105, 25)
GUICtrlSetBkColor(-1, 0x3073BD)
GUICtrlSetColor(-1, 0xffffff)
GUICtrlSetState($Export, $GUI_DISABLE)


$TAB_settings = GUICtrlCreateTabItem("Settings")

GUICtrlCreateLabel("Current Devices", 8, 32, 80, 17)
$Cur_Kons = GUICtrlCreateList("", 8, 48, 161, 321)
GUICtrlSetColor(-1, 0x000000)

For $i = 1 To $aPrinter[0][0]
	$IP = $aPrinter[$i][0]
	$Konica = StringSplit($aPrinter[$i][1], " ")
	GUICtrlSetData($Cur_Kons, $Konica[1] & @CRLF, $Cur_Kons)
	GUICtrlSetColor(-1, 0x000000)
Next


$Add_Device = GUICtrlCreateButton("Add New Device", 8, 370, 163, 25)
GUICtrlCreateTabItem("")

GUISetState(@SW_SHOW)

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit

		Case $Process
			_GUICtrlListView_DeleteAllItems($Info)
			UDPStartup()

			For $i = 1 To $aPrinter[0][0]
				$IP = $aPrinter[$i][0]
				$var = Ping($IP, 500)
				Sleep(500)
				If Not $var Then
					$name = StringSplit($aPrinter[$i][1], " ")
					$Kitem = GUICtrlCreateListViewItem($name[1] & "|" & $IP & "|Not responding", $Info)
					GUICtrlSetColor($Kitem, 0xff0f00)
				Else
					$Socket = UDPOpen($aPrinter[$i][0], $Port)
					$result &= $aPrinter[$i][1]
					$OIDsect = $aPrinter[$i][1]
					$OIDsect = StringTrimLeft($OIDsect, (StringInStr($OIDsect, " ", 0, -1)))
					$aOID = IniReadSection($ini, $OIDsect)
					$name = StringSplit($aPrinter[$i][1], " ")
					;$Kitem = GUICtrlCreateListViewItem($name[1], $Info)
					$output = $name[1] & "|" & $IP & "|"
					For $x = 1 To $aOID[0][0]
						$SNMP_ReqID += 1
						$SNMP_Command = _SNMPBuildPacket($aOID[$x][1], $SNMP_Community, $SNMP_Version, $SNMP_ReqID, "A0")
						UDPSend($Socket, $SNMP_Command)
						_StartListener()
						Sleep(100)
						$output &= $SNMP_Util[1][1] & "|"
					Next
					$Kitem = GUICtrlCreateListViewItem($output, $Info)
					GUICtrlSetColor($Kitem, 0x000000)
					UDPCloseSocket($Socket)
				EndIf
			Next
			MsgBox(64, "Konica Totals", "All Devices Processed", 10)
			UDPShutdown()
			Sleep(1000)
			GUICtrlSetState($Export, 64)

		Case $Export
			$filename = @DesktopDir & "\results - " & @MDAY & @MON & " - " & @HOUR & @MIN & " .csv"
			FileWriteLine($filename, "Printer Name,IP Address,Description,Serial Number,Total Pages")
			$listcount = _GUICtrlListView_GetItemCount($Info)

			For $x = 0 To $listcount - 1
				$stext = ""
				$item = _GUICtrlListView_GetItemTextArray($Info, $x)
				For $i = 1 To $item[0]
					$stext &= $item[$i] & ","
				Next
				FileWriteLine($filename, $stext)
			Next
			MsgBox(64, "Export Completed", "Exported results to " & @CRLF & @CRLF & $filename, 5)
	EndSwitch
WEnd



Func _StartListener()
	If $Start = 1 Then
		$Timeout = TimerInit()
		While (1)
			$srcv = UDPRecv($Socket, 2048)
			If ($srcv <> "") Then
				$result = _ShowSNMPReceived($srcv)
				ExitLoop
			EndIf
			Sleep(1000)
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


