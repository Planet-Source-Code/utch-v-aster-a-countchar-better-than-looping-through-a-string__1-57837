<div align="center">

## A CountChar \- Better than looping through a string\!\!\!


</div>

### Description

It used to be, when I wanted to count the number of times a charector appeared in a string, I would loop through the string letter by letter and keep a count. I have replace my old methodolgies with this one, which also allows you to search for substrings (more than 1-digit long)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[\[\]\)utch\[\]v\[\]aster](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/utch-v-aster.md)
**Level**          |Beginner
**User Rating**    |4.2 (25 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/utch-v-aster-a-countchar-better-than-looping-through-a-string__1-57837/archive/master.zip)





### Source Code

<BR><BR>Public Function CountChar(vText as String, vChar as String, Optional IgnoreCase as Boolean) as Integer<BR>
&nbsp;&nbsp;If IgnoreCase Then <BR>
&nbsp;&nbsp;&nbsp;&nbsp;vText = LCase$(vText) <BR>
&nbsp;&nbsp;&nbsp;&nbsp;vChar = LCase$(vChar) <BR> &nbsp;&nbsp;End If <BR>
&nbsp;&nbsp;Dim L as Integer <BR>
&nbsp;&nbsp;L = Len(vText) <BR><BR>
&nbsp;&nbsp;vText = Replace$(vText, vChar, "") <BR>
&nbsp;&nbsp;CountChar = (L - Len(vText)) / Len(vChar) <BR>
End Function

