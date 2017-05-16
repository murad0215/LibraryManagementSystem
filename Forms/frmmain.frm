VERSION 5.00
Begin VB.Form frmmain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Main Menu"
   ClientHeight    =   10305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12150
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10305
   ScaleWidth      =   12150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   11160
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   11640
      Top             =   0
   End
   Begin VB.Image cmdsrch 
      Height          =   1920
      Left            =   4200
      Picture         =   "frmmain.frx":000C
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   3960
   End
   Begin VB.Image cmdrete 
      Height          =   3105
      Left            =   6840
      Picture         =   "frmmain.frx":73CD0
      Stretch         =   -1  'True
      ToolTipText     =   "Click This Button To Enter The Interface Through Which You Can Return A Book"
      Top             =   6120
      Width           =   2880
   End
   Begin VB.Image cmdret 
      Height          =   3105
      Left            =   8400
      Picture         =   "frmmain.frx":19DEB4
      Stretch         =   -1  'True
      ToolTipText     =   "Click This Button To Enter The Interface Through Which You Can Return A Book"
      Top             =   6840
      Width           =   2880
   End
   Begin VB.Image cmdisue 
      Height          =   3105
      Left            =   3840
      Picture         =   "frmmain.frx":410E38
      Stretch         =   -1  'True
      ToolTipText     =   "Click This Button To Enter The Interface Through Which You Can Issue A Book"
      Top             =   6960
      Width           =   2880
   End
   Begin VB.Image cmdisu 
      Height          =   3105
      Left            =   720
      Picture         =   "frmmain.frx":4FE05C
      Stretch         =   -1  'True
      ToolTipText     =   "Click This Button To Enter The Interface Through Which You Can Issue A Book"
      Top             =   6840
      Width           =   2880
   End
   Begin VB.Image cmdmeme 
      Height          =   3360
      Left            =   6000
      Picture         =   "frmmain.frx":75D3F0
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   3360
   End
   Begin VB.Image cmdbooke 
      Height          =   3360
      Left            =   2760
      Picture         =   "frmmain.frx":917874
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   3360
   End
   Begin VB.Image Image1 
      Height          =   1755
      Left            =   1845
      Picture         =   "frmmain.frx":A55D38
      Stretch         =   -1  'True
      Top             =   120
      Width           =   8460
   End
   Begin VB.Label lblmain 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "MAIN  MENU"
      BeginProperty Font 
         Name            =   "LCD"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   975
      Left            =   4125
      TabIndex        =   1
      Top             =   1800
      Width           =   4215
   End
   Begin VB.Label lbluser 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Administrator"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   10380
      TabIndex        =   0
      ToolTipText     =   "Click This Button To Log Out"
      Top             =   9960
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   1755
      Left            =   1845
      Picture         =   "frmmain.frx":A707D0
      Stretch         =   -1  'True
      Top             =   120
      Width           =   8460
   End
   Begin VB.Image cmdbook 
      Height          =   3360
      Left            =   240
      Picture         =   "frmmain.frx":AB5B54
      Stretch         =   -1  'True
      ToolTipText     =   "Click This Button To Enter In The Book Information Interface"
      Top             =   2880
      Width           =   3360
   End
   Begin VB.Image cmdmem 
      Height          =   3360
      Left            =   8640
      Picture         =   "frmmain.frx":DCBC18
      Stretch         =   -1  'True
      ToolTipText     =   "Click This Button To Enter Into The Member Information Interface"
      Top             =   2880
      Width           =   3360
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdisue_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdisue.Visible = False
End Sub

Private Sub cmdout_Click()
    
End Sub

Private Sub cmdbooke_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdbooke.Visible = False
End Sub

Private Sub cmdmeme_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdmeme.Visible = False
End Sub

Private Sub cmdrete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdrete.Visible = False
End Sub

Private Sub cmdsrche_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdsrche.Visible = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Visible = True
cmdbooke.Visible = True
cmdmeme.Visible = True
cmdisue.Visible = True
cmdrete.Visible = True
'cmdsrche.Visible = True
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Visible = False
End Sub

Private Sub Image3_Click()

End Sub

Private Sub lblmain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Visible = True
End Sub

Private Sub lbluser_Click()

End Sub

Private Sub Timer1_Timer()
'lblmain.ForeColor = &HFF00&
If lblmain.ForeColor = &HFFFF& Then
    lblmain.ForeColor = &HFF00&
End If
End Sub

Private Sub Timer2_Timer()
If lblmain.ForeColor = &HFF00& Then
    lblmain.ForeColor = &HFFFF&
End If
End Sub
