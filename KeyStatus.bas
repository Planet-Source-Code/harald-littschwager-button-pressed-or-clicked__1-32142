Attribute VB_Name = "KeyStatus"
Option Explicit

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public Function ReturnState() As Boolean    'Enterkey status as bool
ReturnState = CBool((GetAsyncKeyState(vbKeyReturn) And &H8000) = &H8000)
End Function
