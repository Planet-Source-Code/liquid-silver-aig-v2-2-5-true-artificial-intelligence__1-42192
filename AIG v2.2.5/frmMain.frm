VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00FF8000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AIG2"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3405
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "frmMain.frx":151A
   MousePointer    =   99  'Custom
   Picture         =   "frmMain.frx":1DE4
   ScaleHeight     =   2265
   ScaleWidth      =   3405
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrStart 
      Interval        =   10000
      Left            =   1440
      Top             =   1125
   End
   Begin VB.TextBox txtOut 
      Appearance      =   0  'Flat
      BackColor       =   &H00B5BC00&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   1485
      Left            =   150
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   645
      Width           =   3135
   End
   Begin VB.TextBox txtIn 
      Appearance      =   0  'Flat
      BackColor       =   &H00B5BC00&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   135
      TabIndex        =   0
      Top             =   330
      Width           =   2775
   End
   Begin VB.Label cmdExit 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   3180
      MouseIcon       =   "frmMain.frx":1B19A
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   45
      Width           =   195
   End
   Begin VB.Label cmdMinimize 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   3000
      TabIndex        =   4
      Top             =   45
      Width           =   195
   End
   Begin VB.Label Title 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   15
      Width           =   3390
   End
   Begin VB.Label cmdOK 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3000
      TabIndex        =   2
      Top             =   315
      Width           =   300
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Beeping As Boolean
Private Busy As Boolean

Private Sub cmdExit_Click()
    Unload Me
    End
End Sub

Private Sub cmdMinimize_Click()
    Me.WindowState = 1
End Sub

Private Sub cmdOK_Click()
    If Busy Then Exit Sub
    Busy = True
    
    If Beeping = True Then Beep
    
    txtIn.SelStart = 0
    txtIn.SelLength = Len(txtIn.Text)
    txtIn.SetFocus
    
    txtOut = "<THINKING>"
    cmdOK.Enabled = False
    tmrStart.Enabled = False
    
    DoEvents
    
    txtOut = RequestAnswer(txtIn.Text, Joining)
    cmdOK.Enabled = True
    If Joining = False Then tmrStart.Enabled = True
    Beeping = True
    Busy = False
End Sub

Private Sub Form_Load()
    Call Remove_Title
    txtOut = "<LOADING>"
    DoEvents
    Beeping = True
    Busy = False
    
    Initialize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    txtOut = "<UNLOADING>"
    DoEvents
    
    Terminate
    End
End Sub

Private Sub Title_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  MousePointer = 15
  Call ReleaseCapture
  Call SendMessage(hWnd, &HA1, 2, 0&)
  MousePointer = 99
End Sub

Private Sub tmrStart_Timer()
    If Busy Then Exit Sub
    Busy = True
    
    txtOut = "<THINKING>"
    cmdOK.Enabled = False
    tmrStart.Enabled = False
    frmMain.tmrStart.Interval = (Int((11) * Rnd) + 10) * 1000
    
    DoEvents
    
    Beep
    txtOut = Starter()
    cmdOK.Enabled = True
    If Joining = False Then tmrStart.Enabled = True
    Busy = False
End Sub

Private Sub txtIn_Change()
    If Trim(txtIn) = "" And Continue = True Then
        tmrStart.Enabled = False
        tmrStart.Enabled = True
    Else
        tmrStart.Enabled = False
    End If
End Sub

Private Sub txtIn_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Beeping = False: Call cmdOK_Click: Beeping = False
    Beeping = True
End Sub
