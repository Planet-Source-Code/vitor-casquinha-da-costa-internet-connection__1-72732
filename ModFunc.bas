Attribute VB_Name = "ModFunc"
Option Explicit

Private Declare Function SetLayeredWindowAttributes Lib "USER32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "USER32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "USER32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000

Public Sub TransWin(ByVal hWnd As Long, ByVal lTrans As Byte)
Dim lRet As Long
lRet = GetWindowLong(hWnd, GWL_EXSTYLE)
lRet = lRet Or WS_EX_LAYERED
Call SetWindowLong(hWnd, GWL_EXSTYLE, lRet)
Call SetLayeredWindowAttributes(hWnd, 0, lTrans, LWA_ALPHA)
End Sub

