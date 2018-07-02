;
; AutoHotkey Version: 1.x
; Language:       English
; Platform:       Win9x/NT
; Author:         A.N.Other <myemail@nowhere.com>
;
; Script Function:
;	Template script (you can customize this template by editing "ShellNew\Template.ahk" in your Windows folder)
;

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

sleep, 50
stdout := FileOpen("*", "r")  
Text := stdout.read()
;Text := Clipboard
Text_array := StrSplit(Text, "IamSep")
if Text_array[1]{
	ControlSetText, TMemo1, % Text_array[1], UniReport
}
if Text_array[2]{
	ControlSetText, Edit3, % Text_array[2], UniReport
}
if Text_array[3]{
	ControlSetText, Edit2, % Text_array[3], UniReport
}

