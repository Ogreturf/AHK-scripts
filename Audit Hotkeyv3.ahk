setkeydelay,200
setwindelay, 50
#Singleinstance
#persistent


;^1::
;Inputbox, FileLoc
Fileloc := "K:\MV\General Office\Q's folder\Hotkeys\New method\Reference Files\audits.csv"
;msgbox, %Fileloc%
;Return

^1::
inputbox, start
OnExit, ExitSub
loop, read, %Fileloc%
{
    LineNumber = %A_Index%
    Loop, parse, A_LoopReadLine, CSV
    {
	If (LineNumber=>1) 
{
	;msgbox, %LineNumber%
	Continue
}
	If(A_Index = 1)
{
	If(LineNumber>=start)
	{
     	 Keywait, Space,D
		  sleep 300
		 ;MsgBox, 4, , Field %LineNumber%-%A_Index% is:`n%A_LoopField%`n`nContinue?
		 Sendinput, {Backspace}%A_LoopField%{enter} 
		 FileAppend,%A_LoopField%`,`, Complete`n, C:\Temp\Temp.txt
		Audit :=  A_LoopField	
		
	} 
}

}

}
FileMove, C:\Temp\Temp.txt, %FileLoc%, 1
Exitapp

ExitSub:
Msgbox, Your last completed audit is %Audit%. `n a reminder has been saved at C:\Temp\Audit.txt
FileDelete C:\Temp\Audit.txt
FileDelete C:\Temp\Temp.txt
FileAppend, %Audit% row number %LineNumber%.`n, C:\Temp\Audit.txt
Exitapp 
