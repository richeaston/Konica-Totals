#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.2.4.9
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

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