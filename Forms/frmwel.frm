VERSION 5.00
Begin VB.Form frmwel 
   BorderStyle     =   0  'None
   ClientHeight    =   10215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11610
   LinkTopic       =   "Form1"
   Picture         =   "frmwel.frx":0000
   ScaleHeight     =   10215
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer29 
      Interval        =   10200
      Left            =   10920
      Top             =   6960
   End
   Begin VB.Timer Timer28 
      Interval        =   10000
      Left            =   9000
      Top             =   6960
   End
   Begin VB.Timer Timer27 
      Interval        =   9900
      Left            =   8760
      Top             =   6960
   End
   Begin VB.Timer Timer26 
      Interval        =   9800
      Left            =   8520
      Top             =   6960
   End
   Begin VB.Timer Timer25 
      Interval        =   9700
      Left            =   8280
      Top             =   6960
   End
   Begin VB.Timer Timer24 
      Interval        =   9600
      Left            =   8040
      Top             =   6960
   End
   Begin VB.Timer Timer23 
      Interval        =   9500
      Left            =   7560
      Top             =   6960
   End
   Begin VB.Timer Timer22 
      Interval        =   9400
      Left            =   7320
      Top             =   6960
   End
   Begin VB.Timer Timer21 
      Interval        =   9300
      Left            =   6840
      Top             =   6960
   End
   Begin VB.Timer Timer20 
      Interval        =   9200
      Left            =   6600
      Top             =   6960
   End
   Begin VB.Timer Timer19 
      Interval        =   9100
      Left            =   6360
      Top             =   6960
   End
   Begin VB.Timer Timer18 
      Interval        =   9000
      Left            =   6000
      Top             =   6960
   End
   Begin VB.Timer Timer17 
      Interval        =   8900
      Left            =   5760
      Top             =   6960
   End
   Begin VB.Timer Timer16 
      Interval        =   8800
      Left            =   5520
      Top             =   6960
   End
   Begin VB.Timer Timer15 
      Interval        =   8700
      Left            =   5280
      Top             =   6960
   End
   Begin VB.Timer Timer14 
      Interval        =   8600
      Left            =   5040
      Top             =   6960
   End
   Begin VB.Timer Timer13 
      Interval        =   8500
      Left            =   4800
      Top             =   6960
   End
   Begin VB.Timer Timer12 
      Interval        =   7500
      Left            =   4920
      Top             =   6000
   End
   Begin VB.CommandButton Command1 
      Height          =   975
      Left            =   8280
      Picture         =   "frmwel.frx":31C81
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9120
      Width           =   1815
   End
   Begin VB.Timer Timer11 
      Interval        =   1500
      Left            =   5280
      Top             =   3600
   End
   Begin VB.Timer Timer10 
      Interval        =   1000
      Left            =   1560
      Top             =   2520
   End
   Begin VB.Timer Timer9 
      Interval        =   500
      Left            =   840
      Top             =   1800
   End
   Begin VB.Timer Timer8 
      Interval        =   6500
      Left            =   3840
      Top             =   5040
   End
   Begin VB.Timer Timer7 
      Interval        =   6000
      Left            =   3480
      Top             =   5040
   End
   Begin VB.Timer Timer6 
      Interval        =   5500
      Left            =   3120
      Top             =   5040
   End
   Begin VB.Timer Timer5 
      Interval        =   5000
      Left            =   2760
      Top             =   5040
   End
   Begin VB.Timer Timer4 
      Interval        =   4000
      Left            =   3840
      Top             =   4560
   End
   Begin VB.Timer Timer3 
      Interval        =   3500
      Left            =   3480
      Top             =   4560
   End
   Begin VB.Timer Timer2 
      Interval        =   3000
      Left            =   3120
      Top             =   4560
   End
   Begin VB.Timer Timer1 
      Interval        =   2500
      Left            =   2760
      Top             =   4560
   End
   Begin VB.CommandButton Command2 
      Height          =   975
      Left            =   10080
      Picture         =   "frmwel.frx":3C993
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9120
      Width           =   1455
   End
   Begin VB.Label lblmurad 
      BackStyle       =   0  'Transparent
      Caption         =   "MURAD"
      BeginProperty Font 
         Name            =   "Mortal Kombat 4"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   9480
      TabIndex        =   30
      Top             =   6480
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label lblm 
      BackStyle       =   0  'Transparent
      Caption         =   "m"
      BeginProperty Font 
         Name            =   "Mortal Kombat 4"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   8880
      TabIndex        =   29
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lbla3 
      BackStyle       =   0  'Transparent
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Mortal Kombat 4"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   8640
      TabIndex        =   28
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lbll2 
      BackStyle       =   0  'Transparent
      Caption         =   "l"
      BeginProperty Font 
         Name            =   "Mortal Kombat 4"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   8400
      TabIndex        =   27
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lbls2 
      BackStyle       =   0  'Transparent
      Caption         =   "s"
      BeginProperty Font 
         Name            =   "Mortal Kombat 4"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   8160
      TabIndex        =   26
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lbli 
      BackStyle       =   0  'Transparent
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "Mortal Kombat 4"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   8040
      TabIndex        =   25
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lbll1 
      BackStyle       =   0  'Transparent
      Caption         =   "l"
      BeginProperty Font 
         Name            =   "Mortal Kombat 4"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   7560
      TabIndex        =   24
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblu 
      BackStyle       =   0  'Transparent
      Caption         =   "u"
      BeginProperty Font 
         Name            =   "Mortal Kombat 4"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   7320
      TabIndex        =   23
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblz 
      BackStyle       =   0  'Transparent
      Caption         =   "z"
      BeginProperty Font 
         Name            =   "Mortal Kombat 4"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   6840
      TabIndex        =   22
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lbla2 
      BackStyle       =   0  'Transparent
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Mortal Kombat 4"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   6600
      TabIndex        =   21
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblw 
      BackStyle       =   0  'Transparent
      Caption         =   "w"
      BeginProperty Font 
         Name            =   "Mortal Kombat 4"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   6240
      TabIndex        =   20
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lble 
      BackStyle       =   0  'Transparent
      Caption         =   "e"
      BeginProperty Font 
         Name            =   "Mortal Kombat 4"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   6000
      TabIndex        =   19
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lbln 
      BackStyle       =   0  'Transparent
      Caption         =   "n"
      BeginProperty Font 
         Name            =   "Mortal Kombat 4"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   5760
      TabIndex        =   18
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblh2 
      BackStyle       =   0  'Transparent
      Caption         =   "h"
      BeginProperty Font 
         Name            =   "Mortal Kombat 4"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   5520
      TabIndex        =   17
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lbla1 
      BackStyle       =   0  'Transparent
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Mortal Kombat 4"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   5280
      TabIndex        =   16
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblh1 
      BackStyle       =   0  'Transparent
      Caption         =   "h"
      BeginProperty Font 
         Name            =   "Mortal Kombat 4"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   5040
      TabIndex        =   15
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lbls1 
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Mortal Kombat 4"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   4800
      TabIndex        =   14
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblcode 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Front End Coding:"
      BeginProperty Font 
         Name            =   "Mortal Kombat 4"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   240
      TabIndex        =   11
      Top             =   6000
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Label rami2 
      BackStyle       =   0  'Transparent
      Caption         =   "Ashiqur Rahman Rami"
      BeginProperty Font 
         Name            =   "Mortal Kombat 4"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   8160
      TabIndex        =   10
      Top             =   8280
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label rami3 
      BackStyle       =   0  'Transparent
      Caption         =   "Ashiqur Rahman Rami"
      BeginProperty Font 
         Name            =   "Mortal Kombat 4"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   6600
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label rami4 
      BackStyle       =   0  'Transparent
      Caption         =   "Ashiqur Rahman Rami"
      BeginProperty Font 
         Name            =   "Mortal Kombat 4"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   4800
      TabIndex        =   8
      Top             =   5040
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Label rami1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ashiqur Rahman Rami"
      BeginProperty Font 
         Name            =   "Mortal Kombat 4"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   9720
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label rahat2 
      BackStyle       =   0  'Transparent
      Caption         =   "Rafsan Jani Rahat"
      BeginProperty Font 
         Name            =   "Mortal Kombat 4"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   4320
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label rahat3 
      BackStyle       =   0  'Transparent
      Caption         =   "Rafsan Jani Rahat"
      BeginProperty Font 
         Name            =   "Mortal Kombat 4"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   7440
      TabIndex        =   5
      Top             =   3240
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label rahat4 
      BackStyle       =   0  'Transparent
      Caption         =   "Rafsan Jani Rahat"
      BeginProperty Font 
         Name            =   "Mortal Kombat 4"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   4800
      TabIndex        =   4
      Top             =   4320
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label rahat1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rafsan Jani Rahat"
      BeginProperty Font 
         Name            =   "Mortal Kombat 4"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label lbldsgn 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Interface design:"
      BeginProperty Font 
         Name            =   "Mortal Kombat 4"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   3840
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Label lblhd 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "The ultimate software dealing with library"
      BeginProperty Font 
         Name            =   "Mortal Kombat 4"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1215
      Left            =   1410
      TabIndex        =   1
      Top             =   2520
      Visible         =   0   'False
      Width           =   8775
   End
   Begin VB.Label lblmain 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LIBRARY MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "Mortal Kombat 4"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   1440
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   8775
   End
End
Attribute VB_Name = "frmwel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
frmwel.Show
End Sub

Private Sub Command2_Click()
Me.Hide
frmsec.Show
Unload Me
End Sub

Private Sub Label15_Click()

End Sub

Private Sub Label16_Click()

End Sub

Private Sub Label5_Click()

End Sub

Private Sub Label8_Click()

End Sub

Private Sub Timer1_Timer()
If rahat2.Visible = True Or rahat3.Visible = True Or rahat4.Visible = True Then
    Exit Sub
Else
    rahat1.Visible = True
End If
End Sub

Private Sub Timer10_Timer()
lblhd.Visible = True
End Sub

Private Sub Timer11_Timer()
lbldsgn.Visible = True
End Sub

Private Sub Timer12_Timer()
lblcode.Visible = True
End Sub

Private Sub Timer13_Timer()
lbls1.Visible = True
End Sub

Private Sub Timer14_Timer()
lblh1.Visible = True
End Sub

Private Sub Timer15_Timer()
lbla1.Visible = True
End Sub

Private Sub Timer16_Timer()
lblh2.Visible = True
End Sub

Private Sub Timer17_Timer()
lbln.Visible = True
End Sub

Private Sub Timer18_Timer()
lble.Visible = True
End Sub

Private Sub Timer19_Timer()
lblw.Visible = True
End Sub

Private Sub Timer2_Timer()
If rahat3.Visible = True Or rahat4.Visible = True Then
    Exit Sub
Else
    rahat1.Visible = False
    rahat2.Visible = True
End If
End Sub

Private Sub Timer20_Timer()
lbla2.Visible = True
End Sub

Private Sub Timer21_Timer()
lblz.Visible = True
End Sub

Private Sub Timer22_Timer()
lblu.Visible = True
End Sub

Private Sub Timer23_Timer()
lbll1.Visible = True
End Sub

Private Sub Timer24_Timer()
lbli.Visible = True
End Sub

Private Sub Timer25_Timer()
lbls2.Visible = True
End Sub

Private Sub Timer26_Timer()
lbll2.Visible = True
End Sub

Private Sub Timer27_Timer()
lbla3.Visible = True
End Sub

Private Sub Timer28_Timer()
lblm.Visible = True
End Sub

Private Sub Timer29_Timer()
lblmurad.Visible = True
End Sub

Private Sub Timer3_Timer()
If rahat4.Visible = True Then
    Exit Sub
Else
    rahat2.Visible = False
    rahat3.Visible = True
End If

End Sub

Private Sub Timer4_Timer()
    rahat3.Visible = False
    rahat4.Visible = True
End Sub

Private Sub Timer5_Timer()
If rami2.Visible = True Or rami3.Visible = True Or rami4.Visible = True Then
    Exit Sub
Else
    rami1.Visible = True
End If
End Sub

Private Sub Timer6_Timer()
If rami3.Visible = True Or rami4.Visible = True Then
    Exit Sub
Else
    rami1.Visible = False
    rami2.Visible = True
End If
End Sub

Private Sub Timer7_Timer()
If rami4.Visible = True Then
    Exit Sub
Else
    rami2.Visible = False
    rami3.Visible = True
End If

End Sub

Private Sub Timer8_Timer()
    rami3.Visible = False
    rami4.Visible = True
End Sub

Private Sub Timer9_Timer()
lblmain.Visible = True
End Sub
