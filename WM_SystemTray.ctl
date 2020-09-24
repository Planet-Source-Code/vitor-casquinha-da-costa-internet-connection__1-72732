VERSION 5.00
Begin VB.UserControl SysTray 
   ClientHeight    =   660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   600
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   660
   ScaleWidth      =   600
   Begin VB.Menu mnuTray 
      Caption         =   "TrayMenu"
   End
End
Attribute VB_Name = "SysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'

'Created by: Michael Hollifield
'If used please acknowledge that I did some of it.

'THanks,

'Michael

'Basic Usage is as follows in the parent program:

'systraycontrol.addtosystemtray me,mnuFile,mnuHelp,"This is the toolTip"
    





'***************************************************************
'Windows API/Global Declarations for :Windows 95 System Tray
'***************************************************************






Private Const WM_LBUTTONDBLCLICK = &H203
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_MBUTTONDOWN = &H207

Private Const WM_RBUTTONUP = &H205
Private Const WM_MOUSEMOVE = &H200
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private VBGTray As NOTIFYICONDATA
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long


Public Event LeftButtonClick()
Public Event RightButtonClick()
Public Event LeftButtonDblClick()
Public Event RightButtonDblClick()

'Default Property Values:
Private Const m_def_ToolTipText = ""
Private Const m_def_Menu = 0

'Property Variables:
Private m_ToolTipText As String
Private m_Menu_Left As Object
Private m_Menu_Right As Object

Private m_Parent As Object
Private m_bInit As Boolean
Private m_CurrentImage As Picture


Private Sub Events(sEvent As String)
    'here is where the events are triggered to the parent software. Could have put this
    'in the actualy mouse event but thought it might be cleaner here.
    'Either way works
    'Dont want to do stuff if we are not initialized.
    'Double click events do not happen if you have menus assigned to them.
    
    If m_bInit Then
        
        Select Case sEvent
            Case "LeftButtonDoubleClick"
                RaiseEvent LeftButtonDblClick
            
            Case "LeftButtonClick"
                RaiseEvent LeftButtonClick
                If IsObject(m_Menu_Left) Then
                    result = SetForegroundWindow(m_Parent.hWnd)
                    m_Parent.PopupMenu m_Menu_Left
                End If
            
            Case "RightButtonDoubleClick"
                RaiseEvent RightButtonDblClick
            
            Case "RightButtonClick"
                RaiseEvent RightButtonClick
                If IsObject(m_Menu_Right) Then
                    result = SetForegroundWindow(m_Parent.hWnd)
                    m_Parent.PopupMenu m_Menu_Right
                End If
        
        End Select
    End If

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Static lngMsg As Long
    Static blnFlag As Boolean
    Dim result As Long
        
    'mouse move event on the icon in the system tray
    'Need to have the on error here. When you are closing the software it will still
    'have an event here as the mouse moves. If the parent is gone (closed)
    'windows will have a cow.
    
    'some care here for reentrancy via the boolean busy variable
    
    On Error GoTo Move_Error
    
    If m_bInit Then
        
        lngMsg = X / Screen.TwipsPerPixelX
        If blnFlag = False Then
            blnFlag = True
            
            Select Case lngMsg
                Case WM_LBUTTONUP
                    Call Events("LeftButtonClick")
                Case WM_LBUTTONDBLCLICK
                    'doubleclick
                    Call Events("LeftButtonDoubleClick")
                Case WM_RBUTTONUP
                    Call Events("RightButtonClick")
                Case WM_RBUTTONDBLCLK
                    Call Events("RightButtonDoubleClick")
            End Select
    
            blnFlag = False
    
        End If
    End If
    
    Exit Sub
    
Move_Error:
        
End Sub

Private Sub UserControl_Resize()
    'really doesnt make a difference. due to the control not being seen
    
    UserControl.Width = 555
    UserControl.Height = 555
    
End Sub

Private Sub UserControl_Terminate()
    'when the parent terminates this control it will remove itself from the system tray
    'note. Windows doesnt seem to update the system tray very good. You might
    'see the icon still there but if you move the mouse over the icon it will dissappear.
    'If you know how to fix this please tell me. Thank You.
    
    Call RemoveFromSystemTray
    
End Sub
Public Sub Refresh()

    'need to just update the icon in the tray
    VBGTray.cbSize = Len(VBGTray)
    VBGTray.hWnd = UserControl.hWnd
    
    VBGTray.hIcon = m_CurrentImage
    VBGTray.szTip = m_ToolTipText & vbNullChar
    
    Call Shell_NotifyIcon(NIM_MODIFY, VBGTray)
    
End Sub
Public Property Set Image(ByVal New_Image As Picture)
Attribute Image.VB_Description = "Returns/sets a graphic to be displayed in a control."
    
    Set m_CurrentImage = New_Image
    
    VBGTray.cbSize = Len(VBGTray)
    VBGTray.hWnd = UserControl.hWnd
    
    VBGTray.hIcon = New_Image
    Call Shell_NotifyIcon(NIM_MODIFY, VBGTray)
    
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_ToolTipText = m_def_ToolTipText
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Set Picture = PropBag.ReadProperty("Image", Nothing)
    m_ToolTipText = PropBag.ReadProperty("ToolTipText", m_def_ToolTipText)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Image", Picture, Nothing)
    Call PropBag.WriteProperty("Menu", m_Menu, m_def_Menu)
    Call PropBag.WriteProperty("ToolTipText", m_ToolTipText, m_def_ToolTipText)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
    ToolTipText = m_ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    
    'this is for setting the tooltiptext after adding the icon. So if you want to
    'change the text for something your program is doing then here is the call.
    
    m_ToolTipText = New_ToolTipText
    PropertyChanged "ToolTipText"
    
    VBGTray.cbSize = Len(VBGTray)
    VBGTray.hWnd = UserControl.hWnd
    
    VBGTray.szTip = m_ToolTipText & vbNullChar
    
    Call Shell_NotifyIcon(NIM_MODIFY, VBGTray)

End Property

Public Property Let Parent_Handle(ByVal hwd As Long)

    m_Handle = hwd
    
End Property
Public Property Let Parent_Form(ByRef oParent As Object)

    Set m_Parent = oParent
    
End Property

Public Sub AddToSystemTray(ByRef oParent As Object, ByRef oMenuL As Object, ByRef oMenuR As Object, sToolTip As String)
    
    'Routine to add the icon to the system tray. oMenuL and oMenuR are for left mouse
    'click and right mouse click.
    'Basic Usage is as follows in the parent program:
    
    'systraycontrol.addtosystemtray me,mnuFile,mnuHelp,"This is the toolTip"
    
    m_bInit = True
    Set m_Parent = oParent
    Set m_Menu_Left = oMenuL
    Set m_Menu_Right = oMenuR
    m_ToolTipText = sToolTip
    
    VBGTray.cbSize = Len(VBGTray)
    VBGTray.hWnd = UserControl.hWnd
    VBGTray.uId = vbNull
    VBGTray.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    VBGTray.ucallbackMessage = WM_MOUSEMOVE
    
    VBGTray.hIcon = m_Parent.Icon
    
   Set m_CurrentImage = m_Parent.Icon
    
    'tooltiptext
    VBGTray.szTip = m_ToolTipText & vbNullChar
    Call Shell_NotifyIcon(NIM_ADD, VBGTray)

End Sub

Public Sub RemoveFromSystemTray()

    'you can call this manually or let the control handle this itsself
    If m_bInit Then
        
        On Error Resume Next
        VBGTray.cbSize = Len(VBGTray)
        VBGTray.hWnd = UserControl.hWnd
        VBGTray.uId = vbNull


        Call Shell_NotifyIcon(NIM_DELETE, VBGTray)
        'm_bInit = False
    End If

End Sub
