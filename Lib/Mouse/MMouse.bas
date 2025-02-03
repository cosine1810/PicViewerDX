Attribute VB_Name = "MMouse"
Option Explicit

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const MOUSEEVENTF_MIDDLEDOWN = &H20
Public Const MOUSEEVENTF_MIDDLEUP = &H40
Public Const MOUSEEVENTF_RIGHTDOWN = &H8
Public Const MOUSEEVENTF_RIGHTUP = &H10

'Public Type POINTAPI
'    x As Long
'    y As Long
'
'End Type


'SetCursorPos Cursor(i).Pos.x, Cursor(i).Pos.y                       '设置鼠标指针坐标
    '模拟鼠标事件
 '   If (Not pLB) And (Cursor(i).LButton) Then mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0

Public Function GetCursorByObjectPos(ByVal ob As Object) As POINTAPI
    GetCursorPos GetCursorByObjectPos
    
    On Error Resume Next
    Do
        GetCursorByObjectPos.X = GetCursorByObjectPos.X - ob.left / Screen.TwipsPerPixelX
        GetCursorByObjectPos.Y = GetCursorByObjectPos.Y - ob.top / Screen.TwipsPerPixelY
        Set ob = ob.parent
        
    Loop Until ob.parent Is Nothing
    
    On Error GoTo 0
End Function

Public Function GetFormCursorPos(ByVal ob As Object) As POINTAPI
    Dim p As POINTAPI
    p = GetCursorByObjectPos(ob)
    p.X = p.X - 8
    p.Y = p.Y - 49
    GetFormCursorPos = p
End Function

Public Function IsMouseLeftPressed() As Boolean
    Dim Res As Long
    Res = GetAsyncKeyState(1)
    If Res = -32767 Then
        IsMouseLeftPressed = True
    End If
End Function

Public Function IsMouseLeftDown() As Boolean
    Dim Res As Long
    Res = GetAsyncKeyState(1)
    
    If Res = -32768 Or Res = -32767 Then
        IsMouseLeftDown = True
    End If
End Function
