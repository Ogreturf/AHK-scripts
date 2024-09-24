setkeydelay,200
setwindelay, 50
#Singleinstance
#persistent
^+h::
msgbox, this hotkey is meant to assist with doing WIC/Material Conversions.`n Press Control+Shift+A to input the wic, wait a few moments and a message should appear prompting whether you want to save the item or not.`n If you select yes a copy is stored in your temp files under the name material history. `n The text file contains your most recent search, a CSV file contains a record of all previous searches. If you wish to clear your search history delete the .csv


;Set file path

;msgbox, %Fileloc%
VeryStart:
return
^+a::
Fileloc := "K:\MV\General Office\Q's folder\Hotkeys\New method\Material Lookup.csv"
Start:


Inputbox, Searched,Insert Wic #
Loop{
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
			 Legacy := A_LoopField
			 
		} 
}
	If(A_Index = 2)
	{
		If(LineNumber>1)
		{
     		; MsgBox, 4, , Field %LineNumber%-%A_Index% is:`n%A_LoopField%`n`nContinue?
			Material := A_LoopField
			 	
			 
		} 
		;msgbox, If index 2 `n%Legacy%`n%Material% 
			
		;If Legacy != %searched%
		;{
		;continue
		;msgbox, success
		;}
	}
		If(A_Index = 3)
{
		If(LineNumber>1)
		{
     		; MsgBox, 4, , Field %LineNumber%-%A_Index% is:`n%A_LoopField%`n`nContinue?
			 UOM := A_LoopField
			 
		} 
}
		If(A_Index = 4)
{
		If(LineNumber>1)
		{
     		; MsgBox, 4, , Field %LineNumber%-%A_Index% is:`n%A_LoopField%`n`nContinue?
			 UPC := A_LoopField
			 
		} 
		if Legacy = %searched%
		{
		;msgbox, Conversion is %Legacy%`n%Material%
		Gosub, End
		}	
}
}

}
if Material = null
break
End:
Msgbox,4,, WIC is %Legacy%`nMaterial is %Material%`nUnit of Measure is %UOM%`nUPC is %UPC%`n do you wish to save this item?
;
	
IfMsgBox, No
         reload

FileAppend,%Legacy%`,%Material%`,%UOM%`,%UPC%`n,C:\Temp\Material_History.csv
FileDelete, C:\Temp\Material_History_Shortlist.txt
FileAppend,%Legacy%`,%Material%`,%UOM%`,%UPC%`n,C:\Temp\Material_History_Shortlist.txt
FileMove, C:\Temp\Material_History_Shortlist.txt, C:\Temp\Material_History_Shortlist.txt

If winexist("Material_History_Shortlist.txt - Notepad")
{
WinWait Material_History_Shortlist.txt ahk_class Notepad,, .250
 WinClose, Material_History_Shortlist.txt ahk_class Notepad
 WinWait Material_History_Shortlist.txt ahk_class Notepad,, .250
 WinClose, *Material_History_Shortlist.txt ahk_class Notepad

Run, C:\Temp\Material_History_Shortlist.txt
}


else
{
Run, C:\Temp\Material_History_Shortlist.txt
}
return
}