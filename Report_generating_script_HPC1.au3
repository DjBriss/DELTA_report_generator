#include <Date.au3>
#include <Constants.au3>
#include <WindowsConstants.au3>
#include <Excel.au3>

#cs
This script will open the Delta software and generate the report.
It changes the report name and file location as needed, and then closes the Delta software
#ce
Global $day, $month, $year, $time = 0
Global $path = "C:\Documents and Settings\Maintenance\My Documents\DELTA_reports\"
GetDateInfo()
foldercheck()
Delta()
Move_report()
If FileExists($path & $month & $year & "\" & $day & ".xls") Then
;MsgBox($MB_OK, "TITLE", "BEFORE REPORT GENERSATOR")
report_generator()
Sleep(1000)
time_in_report()
EndIf 

Func Delta()
Run("C:\Program Files\Delta Controls\3.33\System\ORCAview.exe")
WinWait("Logon")
Login()
WinWait("ORCAview - WHPC")
WinActivate("ORCAview - WHPC")
Gotoreport()
Savereport($month, $day, $year)
Sleep(1000)
WinClose("ORCAview - WHPC")
EndFunc

Func foldercheck()
   If Not FileExists($path & $month & $year) Then
	  DirCreate($path & $month & $year)
	  FileCopy($path & "HPC_monthly_reportmaster.xls", $path & $month & $year & "\monthly_report.xls")
   EndIf
   If Not FileExists($path & $month & $year & "\monthly_report.xls") Then
	  FileCopy($path & "HPC_monthly_reportmaster.xls", $path & $month & $year & "\monthly_report.xls")
   EndIf
EndFunc

Func GetDateInfo()
Local $SYSTEMTIME = _Date_Time_GetSystemTime()
Local $Datestring = _Date_Time_SystemTimeToDateTimeStr($SYSTEMTIME)
Local $Date = StringSplit($Datestring, "/")
Local $monthname = _DateToMonth($Date[1])
Local $temp_year = StringSplit($Date[3], " ")
$day = $Date[2]
$month = $monthname
$year = $temp_year[1]
EndFunc

Func Login()
ControlClick("Logon", "", "[CLASS:Edit; INSTANCE:1]", "left", 1)
ControlSend("Logon", "", "[CLASS:Edit; INSTANCE:1]", "DELTA")
ControlClick("Logon", "", "[CLASS:Edit; INSTANCE:2]","left" , 1)
ControlSend("Logon", "", "[CLASS:Edit; INSTANCE:2]", "LOGIN")
Sleep(2000)
ControlClick("Logon", "", "[CLASS:Button; INSTANCE:1]", "left", 1)
EndFunc

Func Gotoreport()
   WinActivate("ORCAview - WHPC")
   WinMenuSelectItem("ORCAview - WHPC","","&Tools","&Navigator")
   Sleep(2000)
   ControlClick("Navigator - Network","","[CLASS:SysListView32; INSTANCE:1]", "left",2,11,64)
   ;ControlClick("Navigator - Reports","","[CLASS:SysListView32; INSTANCE:1]", "left",2,46, 28)
EndFunc

Func Savereport($monthname, $day, $year)
   Sleep(1500)
   ControlClick("Navigator - Reports","","[CLASS:SysListView32; INSTANCE:1]", "right",1,47, 28)
   Sleep(250)
   ControlSend("Navigator - Reports","","[CLASS:SysListView32; INSTANCE:1]", "e")
   WinWait("Exporting Records")
EndFunc   

Func Move_report()
   FileMove($path & "report.xls",$path & $month & $year & "\" & $day & ".xls",1)
EndFunc

Func report_generator()
ControlClick("Program Manager","","[CLASS:SysListView32; INSTANCE:1]", "left",2, 33, 399)
WinWait("Form1")
WinActivate("Form1")
;select the location
ControlClick("Form1", "","[NAME:ListBox1]","left",2,45, 35)
WinWait("WindowsApplication2")
ControlSend("WindowsApplication2","","","{ENTER}")
Sleep(1000)
;trick the program into triggering the date
ControlClick("Form1", "","[NAME:DateTimePicker1]","left",2,60,10)
ControlSend("Form1", "","[NAME:DateTimePicker1]", "{DOWN}")
ControlSend("Form1", "","[NAME:DateTimePicker1]", "{UP}")
Sleep(1000)
;select the extension
ControlClick("Form1", "","[NAME:ListBox2]","left",2,49, 7)
WinWait("WindowsApplication2")
ControlSend("WindowsApplication2","","","{ENTER}")
Sleep(1000)
ControlClick("Form1", "","[NAME:Button1]","left",2)
WinWait("WindowsApplication2")
ControlSend("WindowsApplication2","","","{ENTER}")
WinWait("WindowsApplication2")
ControlSend("WindowsApplication2","","","{ENTER}")
WinWait("WindowsApplication2")
;ControlSend("WindowsApplication2","","","{ENTER}")
Winclose("Form1")
EndFunc

Func time_in_report()
   Local $workbook = _ExcelBookOpen($path & $month & $year & "\monthly_report.xls")
   _ExcelWriteCell( $workbook, @Hour & ":" & @Min, 2, ($day + 3)) 
   Sleep(1000)
   _ExcelBookClose($workbook)
EndFunc
