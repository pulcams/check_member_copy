#include <Date.au3>

;=================================================================
; Quick check of member copy in Voyager Cataloging module for:
; 1. 'completeness' (full call no., single call no., 6xx or lit.)
; 2. elvl is full (not 3,5,8, or M)
; 3. RelDes (7xx $e, 1xx $e)
; Outputs csv logfile and MsgBox prompts.
; Using Voyager 9.0
; From 20150106
; pmg
;=================================================================

Local $action = "" ; action for MsgBox
Local $bib = "none" ; bib id default
Local $bibid = "" ; bib id from $bib (split window title)
Local $elvl = "" ; encoding level
Local $f050a_count = 0
Local $flag = 0 ; for MsgBox
Local $fullcallno = False
Local $header = "date, bib, track, full_call_no, mult_call_no, subj_or_lit, elvl" ; for log file
Local $lit = ""
Local $log = @ScriptDir & "\log.csv"
Local $multcallno = False
Local $msg = "" ; log entry
Local $online = False
Local $reldes = True ; relationship des. are OK by default
Local $suba = 0 ; to count occurrences of 050$a
Local $sublit = False ; 6xx or call no. P or 008/33
Local $track = "" ; slow, med. or fast track -- goes into the log file

;====================================
; log file test
;====================================
If FileGetSize($log) == 0 Then
	$hFile = FileOpen($log, 1)
	FileWriteLine($hFile, $header) ; insert column names if log file is blank
	FileClose($hFile)
EndIf

WinActivate("Voyager Cataloging")
Send("!m")
;====================================
; LDR - get ldr/17 elvl
;====================================
Send("{TAB}")
Send("{RIGHT}")
Send("{HOME}+{END}")
Send("^c")
Sleep(100)
$ldr = ClipGet()
$elvl = StringMid($ldr, 18, 1)
If $elvl == " " Then
	$elvl = "full"
EndIf
;====================================
; 008 - get literary form 008/33
;====================================
; NOTE: also testing for call no. starting w/ P in 050 and 090 check
Send("{TAB 2}")
Send("{RIGHT}")
Send("{HOME}+{END}")
Send("^c")
Sleep(100)
$f008 = ClipGet()
$lit = StringMid($f008, 34, 1)
If $lit <> "0" And $lit <> "" Then
	$sublit = True
EndIf
Send("!m")
;====================================
; enter record
;====================================
Send("{shiftdown}{tab 2}{shiftup}")
Send("{home}^{home}")
Send("{f8}")
Send("^c")
Sleep(100)
$thisfield = ClipGet()

While $thisfield <> "" And $thisfield < "800" ; loop through 7xx or to the end of the record
	;===============================
	; 050
	;===============================
	If $thisfield == "050" Then
		Send("{TAB 3}{HOME}+{END}")
		Send("^c")
		Sleep(100)
		$f050 = ClipGet()
		If StringInStr($f050, "‡b") And StringRegExp($f050, '[0-9]') Then
			$fullcallno = True
		Else
			$fullcallno = False ; no numbers in 050$b
		EndIf
		StringReplace($f050, "‡a", "")
		$suba = @extended ; the count of $a in this 050 field
		$f050a_count = $f050a_count + $suba ; the total 050$a count in this record
		If $f050a_count > 1 Then
			$multcallno = True ; has more than one ‡a
		EndIf
		If StringInStr($f050, "‡a P") Then ; if call no. starts with P
			$sublit = True
		EndIf
		Send("+{TAB 3}")
	EndIf
	;===============================
	; 090
	;===============================
	If $thisfield == "090" And $fullcallno == False Then
		Send("{TAB 3}{HOME}+{END}")
		Send("^c")
		Sleep(100)
		$f090 = ClipGet()
		If StringInStr($f090, "‡b") Then
			$fullcallno = True
		EndIf
		If StringInStr($f090, "‡a P") Then ; if call no. starts with P
			$sublit = True
		EndIf
		Send("+{TAB 3}")
	EndIf
	;===============================
	; 1xx
	;===============================
	If StringMid($thisfield, 1, 1) == "1" Then
		Send("{TAB 3}{HOME}+{END}")
		Send("^c")
		Sleep(100)
		$f1xx = ClipGet()
		If Not StringInStr($f1xx, "‡e") Then
			$reldes = False
		EndIf
		Send("+{TAB 3}")
	EndIf
	;===============================
	; 300
	;===============================
	If $thisfield == "300" Then
		Send("{TAB 3}{HOME}+{END}")
		Send("^c")
		Sleep(100)
		$f300 = ClipGet()
		If StringInStr($f300, "online") Then
			$online = True
		EndIf
		Send("+{TAB 3}")
	EndIf
	;===============================
	; 6xx
	;===============================
	If $sublit == False Then
		If StringMid($thisfield, 1, 1) == "6" Then
			Send("{TAB 2}")
			Send("^c")
			Sleep(100)
			$ind2 = ClipGet()
			If $ind2 == "0" Then
				$sublit = True
			EndIf
			Send("+{TAB 2}")
		EndIf
	EndIf
	;===============================
	; 7xx
	;===============================
	If StringMid($thisfield, 1, 1) == "7" Then
		Send("{TAB 3}{HOME}+{END}")
		Send("^c")
		Sleep(100)
		$f7xx = ClipGet()
		If Not StringInStr($f7xx, "‡e") Then
			$reldes = False
		EndIf
		Send("+{TAB 3}")
	EndIf

	ClipPut("") ; clear clipboard

	Send("{down}{f8}") ; keep moving

	Send("^c")
	Sleep(100) ; this may need to be tweaked

	$thisfield = ClipGet()

	;MsgBox(0, "", $thisfield) ; for debugging
WEnd

;====================================
; get bib id
;====================================
$bib = StringSplit(WinGetTitle("Voyager Cataloging"), " ")
Sleep(100)
$bibid = $bib[5]

;====================================
; report
;====================================
If $online == True Then ; online resource - stop the bus!
	$flag = 16 ; stop-sign icon
	$track = "online"
	$action = $action & @CRLF & "Online resource"
ElseIf $fullcallno == True And $multcallno == False And $sublit == True Then ;  record is 'complete'
	If StringInStr("358M", $elvl) Or $reldes == False Then
		$flag = 48 ; exclamation-point icon
		If StringInStr("358M", $elvl) Then
			$track = "slow"
			$action = $action & @CRLF & "ENCODING LEVEL (" & $elvl & ")"
		EndIf
		If $reldes == False Then
			$track = "medium"
			$action = $action & @CRLF & "relationship designators"
		EndIf
		If $fullcallno == False Or $multcallno == True Then
			If $fullcallno == False And $multcallno == True Then
				$action = $action & @CRLF & "call number"
			ElseIf $fullcallno == False Then
				$action = $action & @CRLF & "call  no. (not full)"
			ElseIf $multcallno == True Then
				$action = $action & @CRLF & "call no. (multiple)"
			EndIf
		EndIf
	Else
		$flag = 64 ; information-sign icon
		$track = "fast"
		$action = "OK to process"
	EndIf
Else ; record is not 'complete'
	;------------------------------------
	; debug: why going to hold
	;------------------------------------
;~ 	If $sublit == False Then
;~ 		$action = $action & @CRLF & "To the hold!"
;~ 	EndIf
;~ 	If $fullcallno == False Or $multcallno == True Then
;~ 		If $fullcallno == False And $multcallno == True Then
;~ 			$action = $action & @CRLF & "Check the call number"
;~ 		ElseIf $fullcallno == False Then
;~ 			$action = $action & @CRLF & "Check the call  no. (not full)"
;~ 		ElseIf $multcallno == True Then
;~ 			$action = $action & @CRLF & "Check the call no. (multiple)"
;~ 		EndIf
;~ 	EndIf
	$flag = 48
	$action = "To the hold!"
	$track = "hold"
EndIf

; write to log
$msg = _Now() & "," & $bibid & "," & $track & "," & $fullcallno & "," & $multcallno & "," & $sublit & "," & $elvl
Local $hFile = FileOpen($log, 1)
FileWriteLine($hFile, $msg)
FileClose($hFile)

Send("!m") ; return to top of Voyager record display

MsgBox($flag, "", $action) ; prompt the human