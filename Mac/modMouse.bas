Attribute VB_Name = "modMouse"
Public Type POINTAPI
        X As Long
        Y As Long
End Type
Public p As POINTAPI
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Function GetX()
Dim p As POINTAPI
GetCursorPos p
GetX = p.X * 15 '- mX
End Function

Public Function GetY()
Dim p As POINTAPI
GetCursorPos p
GetY = p.Y * 15 ' - mY
End Function


