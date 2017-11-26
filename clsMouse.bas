Option Explicit
'Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
'Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
'Declare Function SendInput Lib "user32" (ByVal nCommands As Long, iCommand As Any, ByVal cSize As Long) As Long

'Public Const MOUSEEVENTF_MOVE = &H1 '  mouse move
'Public Const MOUSEEVENTF_LEFTDOWN = &H2 '  left button down
'Public Const MOUSEEVENTF_LEFTUP = &H4 '  left button up
'Public Const MOUSEEVENTF_RIGHTDOWN = &H8 '  right button down
'Public Const MOUSEEVENTF_RIGHTUP = &H10 '  right button up
'Public Const MOUSEEVENTF_MIDDLEDOWN = &H20 '  middle button down
'Public Const MOUSEEVENTF_MIDDLEUP = &H40 '  middle button up
'Public Const MOUSEEVENTF_ABSOLUTE = &H8000& '  absolute move
Const INPUT_MOUSE = 0
Const INPUT_KEYBOARD = 1
Const INPUT_HARDWARE = 2
Dim gi As GENERALINPUT
Dim mi As MOUSEINPUT

Sub click()
    mouseDown
    mouseUp
End Sub
Private Sub rightClick()
  mouseAction MOUSEEVENTF_RIGHTDOWN
  mouseAction MOUSEEVENTF_RIGHTUP
End Sub
'alias
Sub contextClick()
    rightClick
End Sub
Private Sub doubleClick()
  click
  click
End Sub
Sub mouseDown()
   mouseAction MOUSEEVENTF_LEFTDOWN
End Sub
Sub mouseMove(ByVal x As Long, ByVal y As Long)
   mouseAction MOUSEEVENTF_ABSOLUTE Or MOUSEEVENTF_MOVE, x, y
End Sub
Sub mouseUp()
   mouseAction MOUSEEVENTF_LEFTUP
End Sub
'alias
Sub move(ByVal x As Long, ByVal y As Long)
   mouseMove x, y
End Sub
Sub moveRelative(ByVal x As Long, ByVal y As Long)
   mouseAction MOUSEEVENTF_MOVE, x, y
End Sub
Sub moveAndClick(ByVal x As Long, ByVal y As Long)
   move x, y
   click
End Sub
'Low level helper function
Sub mouseAction(iFlags As Long, Optional x As Long = 0, Optional y As Long = 0)
   Dim gi As GENERALINPUT
   Dim mi As MOUSEINPUT
   Dim tX As Long
   Dim tY As Long
   
   If (iFlags = (MOUSEEVENTF_ABSOLUTE Or MOUSEEVENTF_MOVE)) Or (iFlags = MOUSEEVENTF_ABSOLUTE) Then
        tX = x * (65535 / ScreenWidth)
        tY = y * (65535 / ScreenHeight)
    Else
        tX = x
        tY = y
   End If
   
   gi.dwType = INPUT_MOUSE
   mi.dwFlags = iFlags
   mi.dx = tX
   mi.dy = tY
   
   CopyMemory VarPtr(gi.xi(0)), VarPtr(mi), Len(mi)
   SendInput 1&, gi, Len(gi)
End Sub
