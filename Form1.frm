VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   855
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   3015
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   855
   ScaleWidth      =   3015
   StartUpPosition =   3  'Windows Default
   Begin vb6projectProject1.SysTray SysTray1 
      Left            =   1200
      Top             =   120
      _ExtentX        =   979
      _ExtentY        =   979
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1920
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2400
      Top             =   240
   End
   Begin VB.Menu file 
      Caption         =   "file"
      Begin VB.Menu CheckIP 
         Caption         =   "&Check IP"
      End
      Begin VB.Menu lin 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim voice As SpVoice

Private Sub CheckIP_Click()
temp = NetCon.GetIPAddress
If temp <> "127.0.0.1" Then
    MsgBox "IP address: " & temp
Else
    MsgBox "No internet connection available"
End If
End Sub

Private Sub exit_Click()
If MsgBox("Are you sure do you want to quit from program?", vbYesNo + vbQuestion, "NetDetect") = vbYes Then End
End Sub

Private Sub Form_Load()
If App.PrevInstance Then End

Set voice = New SpVoice

SysTray1.AddToSystemTray Me, Me, file, "NetDetect"

Me.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
SysTray1.RemoveFromSystemTray
End Sub

Private Sub Timer1_Timer()
temp = NetCon.GetIPAddress
If temp <> Text1 Then
    Text1 = temp
    If Text1 <> "127.0.0.1" Then
        voice.Speak "internet connection acquired.", SVSFlagsAsync
        Form2.Label1 = "Internet Acquired"
        Form2.Show
    Else
        voice.Speak "internet connection lost.", SVSFlagsAsync
        Form2.Label1 = "Internet Lost"
        Form2.Show
    End If
End If
End Sub
