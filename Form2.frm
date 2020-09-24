VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7350
   LinkTopic       =   "Form2"
   ScaleHeight     =   735
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00DC7334&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tmp As Long
Dim Alpha As Long

Private Sub Form_Load()
Me.Top = 0
Me.Left = Screen.Width - Me.Width

ModFunc.TransWin Me.hWnd, 180

Alpha = 180
tmp = 0
Timer1.Enabled = True
End Sub

Private Sub Label1_Click()
tmp = 6
End Sub

Private Sub Timer1_Timer()
tmp = tmp + 1
If tmp > 5 Then
    Timer1.Interval = 10
    If Alpha > 10 Then Alpha = Alpha - 10: ModFunc.TransWin Me.hWnd, Alpha
    If Alpha <= 10 Then Timer1.Enabled = True: Unload Me
End If
End Sub
