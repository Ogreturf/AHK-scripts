setkeydelay,200
setwindelay, 50
#Singleinstance
Inputbox, Fileloc, Needs absolute File Path
;Fileloc := ""
;msgbox, %Fileloc%
Loop, read, %Fileloc%
{
    LineNumber = %A_Index%
    Loop, parse, A_LoopReadLine, CSV
    {
	If (LineNumber=1) 
	Continue
	If(A_Index = 1)
{
	If(LineNumber>1)
	{
     		; MsgBox, 4, , Field %LineNumber%-%A_Index% is:`n%A_LoopField%`n`nContinue?
			 Distr := A_Loopfield
			 
			 
	} 
}
	 
	
	DC:
   	If(A_Index = 2)
{
	If(LineNumber>1)
	{
     		 ;MsgBox, 4, , Field %LineNumber%-%A_Index% is:`n%A_LoopField%`n`nContinue?
			 DC := A_Loopfield
			 
			 
	}
}
	
	Mod:
	If(A_Index = 6)
{
	If(LineNumber>1)
	{
       		;MsgBox, 4, , Field %LineNumber%-%A_Index% is:`n%A_LoopField%`n`nContinue?
			Mod := A_Loopfield
	}
}
	
	If(A_Index = 7)
{
	If(LineNumber>1)
	{
       		;MsgBox, 4, , Field %LineNumber%-%A_Index% is:`n%A_LoopField%`n`nContinue?
			QTY := A_Loopfield
	}

}
      IfMsgBox, No
            return
			
   }
   If (LineNumber=1)
   Continue
	Else
	Gosub, Action
}

Action:
;Distr := Clipboard

;Winmenuselectitem,  5250 Terminal - ADR7 (Session A ), , Edit, Paste

WinWait, 5250 Terminal - ADR7 (Session A ), 
IfWinNotActive, 5250 Terminal - ADR7 (Session A ), , WinActivate, 5250 Terminal - ADR7 (Session A ), 
WinWaitActive, 5250 Terminal - ADR7 (Session A ), 
Sendinput,%Distr%%DC%


;DC := Clipboard

;Winmenuselectitem,  5250 Terminal - ADR7 (Session A ), , Edit, Paste

WinWait, 5250 Terminal - ADR7 (Session A ), 
IfWinNotActive, 5250 Terminal - ADR7 (Session A ), , WinActivate, 5250 Terminal - ADR7 (Session A ), 
WinWaitActive, 5250 Terminal - ADR7 (Session A ), 
Send,{ENTER}
Sendinput,{F6}{tab}{Tab}

ifinstring, Mod, N-R to D
{
;Msgbox, RtoD `n DEQ N
Send,R{tab}D%Distr%{Tab}%QTY%{Enter}{pause}{F3}Y{f3}x{enter}
}
ifinstring, Mod, N-D to R
{
;Msgbox, DtoR `n DEQ N
Send,D%Distr%{tab}R{tab}}%QTY%{Enter}{pause}{F3}Y{f3}x{enter}
}
ifinstring, Mod, Y-R to D
{
;Msgbox, RtoD `n DEQ Y
Send,R{tab}D%Distr%{Tab}%QTY%{Enter}{pause}{F3}Y{f3}x{enter}
}
ifinstring, Mod, Y-R to D
{
;Msgbox, DtoR `n DEQ N
Send,D%Distr%{tab}R{tab}}%QTY%{Enter}{pause}{F3}Y{f3}x{enter}
}

return
