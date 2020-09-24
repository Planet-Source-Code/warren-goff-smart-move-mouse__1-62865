<div align="center">

## Smart Move Mouse


</div>

### Description

I wanted to have the cursor to move to my focused command buttons to direct the action and be more user friendly. This little thing does the trick. I found it in some old code of mine but I think it was borrowed back then.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Warren Goff](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/warren-goff.md)
**Level**          |Beginner
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/warren-goff-smart-move-mouse__1-62865/archive/master.zip)





### Source Code

```
Public Type POINTAPI
  X As Long
  Y As Long
End Type
Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32.dll" (ByVal X As Long, ByVal Y As Long) As Long
Sub MoveMouse(X As Single, Y As Single)
Dim pt As POINTAPI
  pt.X = X
  pt.Y = Y
  ClientToScreen Form1.hwnd, pt
  SetCursorPos pt.X, pt.Y
End Sub
Private Sub Form_Activate()
  Command1.SetFocus
  MoveMouse 97, 72
End Sub
```

