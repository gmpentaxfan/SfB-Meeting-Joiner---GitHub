#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseX64=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
AutoItSetOption("WinTitleMatchMode", 4)
Func _WinWaitActivate($title,$text,$timeout=0)
	WinWait($title,$text,$timeout)
	If Not WinActive($title,$text) Then WinActivate($title,$text)
	WinWaitActive($title,$text,$timeout)
EndFunc
;Run("SfB Meeting Joiner.exe")
;_WinWaitActivate("Testing Tuesday (1 Participant)","")

;$handle = WinGetHandle("[CLASS:LyncConversationWindowClass]","")
;MsgBox (4096+1,"Error1", $handle)
;ControlClick($handle, "", "[CLASS:NetUIHWND; INSTANCE:1]","","",380,155)

_WinWaitActivate("[CLASS:NUIDialog]","")
$handle = WinGetHandle("[CLASS:NUIDialog]","")
;MsgBox (4096+1,"Error2", $handle)
ControlClick($handle, "", "[CLASS:NetUIHWND; INSTANCE:1]","","",380,157)
Send("{ALTDOWN}n{ALTUP}")
ControlClick($handle, "", "[CLASS:NetUIHWND; INSTANCE:1]")
$handle = WinGetHandle("[CLASS:NetUIHWND]","")
;MsgBox (4096+1,"Error3",$handle)
ControlClick($handle, "", "[CLASS:NetUIHWND; INSTANCE:1]","","",380,155)
