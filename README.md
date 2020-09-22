<div align="center">

## Add Horizontal Scrollbar to Listbox


</div>

### Description

Automatically add a horizontal scrollbar to a listbox, quickly and completely. I wasnt satisfied with the ones on planet so I created my own.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[N/A](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/empty.md)
**Level**          |Beginner
**User Rating**    |5.0 (80 globes from 16 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/add-horizontal-scrollbar-to-listbox__1-10898/archive/master.zip)

### API Declarations

```
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const LB_SETHORIZONTALEXTENT = &H194
```


### Source Code

```
Public Sub AddScroll(List As ListBox)
 Dim i As Integer, intGreatestLen As Integer, lngGreatestWidth As Long
 'Find Longest Text in Listbox
 For i = 0 To List.ListCount - 1
 If Len(List.List(i)) > Len(List.List(intGreatestLen)) Then
  intGreatestLen = i
 End If
 Next i
 'Get Twips
 lngGreatestWidth = List.Parent.TextWidth(List.List(intGreatestLen) + Space(1))
'Space(1) is used to prevent the last Character from being cut off
 'Convert to Pixels
 lngGreatestWidth = lngGreatestWidth \ Screen.TwipsPerPixelX
 'Use api to add scrollbar
 SendMessage List.hwnd, LB_SETHORIZONTALEXTENT, lngGreatestWidth, 0
End Sub
```

