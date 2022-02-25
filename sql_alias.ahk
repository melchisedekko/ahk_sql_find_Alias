#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance force

#include <RemoteSQL> ; se devi fare cose con sql
IniRead, ConnectionString, config.ini, SensitiveInformation, ConnectionString
;DebugWindow(ConnectionString)
global DB := New RemoteSQL(ConnectionString)
DB.Connect()

Send {Ctrl Down}c{Ctrl Up}
Sleep, 200
Clipboardstring := Clipboard
;DebugWindow(Clipboardstring)
if (RegExMatch(Clipboardstring, "P)[A-Z1-9&]{3,7}") && IsUpper(Clipboardstring)) {
;DebugWindow("Alias")
rs := DB.Query("SELECT name, base_object_name from sys.synonyms WHERE name = '" . Clipboardstring . "'")
if ((rs.Length()>0 || rs.Count()>0) && rs.Empty != "No Results") {
HideTrayTip()
vText := JEE_StrReplaceChars(rs[1].base_object_name, "[]", "", vCount)
TableData_array := StrSplit(vText, ".")
Sleep 300 
TrayTip, % "Found Table for " . rs[1].Code , % "TableName -> " . TableData_array[2] . "`n Schema -> " . TableData_array[1] , 10, 1
Clipboard := TableData_array[1] . "." . TableData_array[2]
}
else{
HideTrayTip()
Sleep 500 
TrayTip, % "Table not Found for " . Clipboardstring , % "TableName -> ...`n Schema -> ...", 10, 2
Clipboard := "not found"
}
}
else if (StrLen(Clipboardstring)>0 && StrLen(Clipboardstring)<100) {
;DebugWindow("TableName")	
rs := DB.Query("SELECT name, base_object_name from sys.synonyms WHERE base_object_name like '%" . Clipboardstring . "%'")
;DebugWindow(rs)
if ((rs.Length()>0 || rs.Count()>0) && rs.Empty != "No Results") {
HideTrayTip()
vText := JEE_StrReplaceChars(rs[1].base_object_name, "[]", "", vCount)
TableData_array := StrSplit(vText, ".")
Sleep 300  
TrayTip, % "Found Alias for " . TableData_array[2] , % "Alias -> " . rs[1].name . "`n Schema -> " . TableData_array[1] , 10, 1
Clipboard := TableData_array[1] . "." . rs[1].name
}
else{
HideTrayTip()
Sleep 500 
TrayTip, % "Table not Found for " . Clipboardstring , % "TableName -> ...`n Schema -> ...", 10, 2	
Clipboard := "not found"
}
}
try {
	DB.Exit()
}
DB := ""
Sleep, 3000 ; Delay needed because if the script dies before the TrayTip is displayed, it will not diplay it 
ExitApp

HideTrayTip() {
    TrayTip  ; Attempt to hide it the normal way.
    if SubStr(A_OSVersion,1,3) = "10." {
        Menu Tray, NoIcon
        Sleep 300  ; It may be necessary to adjust this sleep.
        Menu Tray, Icon
    }
}

JEE_StrReplaceChars(vText, vNeedles, vReplaceText:="", ByRef vCount:="")
{
	vCount := StrLen(vText)
	;Loop, Parse, vNeedles ;change it to this for older versions of AHK v1
	Loop, Parse, % vNeedles
		vText := StrReplace(vText, A_LoopField, vReplaceText)
	vCount := vCount-StrLen(vText)
	return vText
}

IsUpper(String) {
   Return (String == Format("{:U}", String))
}