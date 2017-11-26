Option Explicit
Public Declare Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As GENERALINPUT, ByVal cbSize As Long) As Long
                                    
Const INPUT_MOUSE = 0
Const INPUT_KEYBOARD = 1
Const INPUT_HARDWARE = 2

Public Type MOUSEINPUT
  dx As Long
  dy As Long
  mouseData As Long
  dwFlags As Long
  time As Long
  dwExtraInfo As Long
End Type

Public Type KEYBDINPUT
  wVk As Integer
  wScan As Integer
  dwFlags As Long
  time As Long
  dwExtraInfo As Long
End Type

Public Type HARDWAREINPUT
  uMsg As Long
  wParamL As Integer
  wParamH As Integer
End Type

Public Type GENERALINPUT
  dwType As Long
  xi(0 To 23) As Byte
End Type
Sub testme()
    Debug.Print GetForegroundWindow()
End Sub

'extern.Declare micLong, "GetForegroundWindow", "user32.dll", "GetForegroundWindow"
'extern.Declare micLong, "AttachThreadInput", "user32.dll", "AttachThreadInput", micLong, micLong, micLong
'extern.Declare micLong, "GetWindowThreadProcessId", "user32.dll", "GetWindowThreadProcessId", micLong, micLong
'extern.Declare micLong, "GetCurrentThreadId", "kernel32.dll", "GetCurrentThreadId"
'extern.Declare micLong, "GetCursor", "user32.dll", "GetCursor"

Function get_cursor()
    Dim hwnd, pid, thread_id
    hwnd = GetForegroundWindow()
    pid = GetWindowThreadProcessId(hwnd, 0)
    thread_id = GetCurrentThreadId()
    AttachThreadInput pid, thread_id, True
    get_cursor = GetCursor()
    AttachThreadInput pid, thread_id, False
End Function
'testit: call drawrectangle(getdc(0),10,10,100,100)
Sub drawRectangle(lDC As Long, x1, y1, x2, y2, Optional blinkCount = 1)
    Dim tp As POINTAPI
    Dim hOldPen As Long, hPen As Long
    Dim logBR As LOGBRUSH
    Dim i As Integer
    
    i = 1
    hPen = CreatePen(0, 1, vbRed)
       
    hOldPen = SelectObject(lDC, hPen)
    MoveToEx lDC, x1, y1, tp
    LineTo lDC, x2, y1
    LineTo lDC, x2, y2
    LineTo lDC, x1, y2
    LineTo lDC, x1, y1
       
    DeleteObject SelectObject(lDC, hOldPen)
End Sub

'Some bit logic in excel 2007 and 2010
' 2013 has bitwise functions BITNOT, BITAND, BITOR and BITNOT
Public Function BITWISE_XOR(x As Long, y As Long) As Long
    BITWISE_XOR = x Xor y
End Function
 
Public Function BITWISE_NOT(x As Long) As Long
    BITWISE_NOT = Not x
End Function
 
Public Function BITWISE_AND(x As Long, y As Long) As Long
    BITWISE_AND = x And y
End Function
 
Public Function BITWISE_OR(ByVal x As Long, ByVal y As Long) As Long
    BITWISE_OR = x Or y
End Function
Public Function ScreenHeight() As Long
    ScreenHeight = GetSystemMetrics(SM_CYSCREEN)
End Function

Public Function ScreenWidth() As Long
    ScreenWidth = GetSystemMetrics(SM_CXSCREEN)
End Function

Public Function ProcIDFromWnd(ByVal hwnd As Long) As Long
   Dim idProc As Long
  
   ' Get PID for this HWnd
   GetWindowThreadProcessId hwnd, idProc
   ProcIDFromWnd = idProc
End Function
      
Public Function GetWinHandle(hInstance As Long) As Long
   Dim tempHwnd As Long
      
   ' Grab the first window handle that Windows finds:
   tempHwnd = FindWindow(vbNullString, vbNullString)
   
   ' Loop until you find a match or there are no more window handles:
   Do Until tempHwnd = 0
      ' Check if no parent for this window
      If GetParent(tempHwnd) = 0 Then
         ' Check for PID match
         If hInstance = ProcIDFromWnd(tempHwnd) Then
            ' Return found handle
            GetWinHandle = tempHwnd
            ' Exit search loop
            Exit Do
         End If
      End If
   
      ' Get the next window handle
      tempHwnd = GetWindow(tempHwnd, GW_HWNDNEXT)
      'Debug.Print tempHwnd
   Loop
End Function
