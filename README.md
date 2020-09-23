<div align="center">

## Hex/String Conversion


</div>

### Description

These functions allow you to easily convert a string to its hex value and back again.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Lewis E\. Moten III](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/lewis-e-moten-iii.md)
**Level**          |Beginner
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Strings](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/strings__4-26.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/lewis-e-moten-iii-hex-string-conversion__4-6804/archive/master.zip)

### API Declarations

Copyright (c) 2001, Lewis Moten. All rights reserved.


### Source Code

email = "lewis@moten.com"
Response.Write "Original value: " & email & "<BR>"
email = StringToHex(email)
Response.Write "Hex value: " & email & "<BR>"
email = HexToString(email)
Response.Write "Back to string value: " & email & "<BR>"
Function StringToHex(ByRef pstrString)
	Dim llngIndex
	Dim llngMaxIndex
	Dim lstrHex
	llngMaxIndex = Len(pstrString)
	For llngIndex = 1 To llngMaxIndex
		lstrHex = lstrHex & Right("0" & Hex(Asc(Mid(pstrString, llngIndex, 1))), 2)
	Next
	StringToHex = lstrHex
End Function
Function HexToString(ByRef pstrHex)
	Dim llngIndex
	Dim llngMaxIndex
	Dim lstrString
	llngMaxIndex = Len(pstrHex)
	For llngIndex = 1 To llngMaxIndex Step 2
		lstrString = lstrString & Chr("&h" & Mid(pstrHex, llngIndex, 2))
	Next
	HexToString = lstrString
End Function

