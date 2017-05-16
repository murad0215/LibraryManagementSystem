VERSION 5.00
Begin VB.Form frmmain2 
   BorderStyle     =   0  'None
   Caption         =   "MAIN  MENU"
   ClientHeight    =   9045
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12045
   LinkTopic       =   "Form2"
   Picture         =   "frmmain2.frx":0000
   ScaleHeight     =   9045
   ScaleWidth      =   12045
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CMDEXIT 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11520
      TabIndex        =   11
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Search and Report"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   10
      Top             =   2880
      Width           =   3495
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Return  Books"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   9
      Top             =   4320
      Width           =   2895
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Issue Books"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   8
      Top             =   6120
      Width           =   2295
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Book Information"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   3480
      Width           =   3375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Member Infomation"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      Top             =   4800
      Width           =   3735
   End
   Begin VB.Label search 
      BackStyle       =   0  'Transparent
      Caption         =   "Search and Report"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label returnbook 
      BackStyle       =   0  'Transparent
      Caption         =   "Return  Books"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   4
      Top             =   4320
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label isuebook 
      BackStyle       =   0  'Transparent
      Caption         =   "Issue Books"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   6120
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label bookinfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Book Information"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   3480
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label memberinfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Member Infomation"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   4800
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label lbluser 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Administrator"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   9120
      TabIndex        =   0
      Top             =   8640
      Width           =   2895
   End
End
Attribute VB_Name = "frmmain2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub bookinfo_Click()
If Not lbluser.Caption = "Administrator" Then
    frmbook.bookdata.RecordsetType = 2 - snapshot
Else
    frmbook.bookdata.RecordsetType = 1 - dynaset
End If

frmbook.Show
End Sub

Private Sub CMDEXIT_Click()
'choice = MsgBox("Are You Sure You Want To Log Out?", vbQuestion + vbYesNo, "Confirmation")
'If choice = vbYes Then
'    Me.Hide
'    lbluser.Caption = ""
'    MsgBox "Thank You For Using THE LIBRARY MANAGEMENT SYSTEM", vbInformation, "Log Out"
'   frmsec.Show
'    Unload Me
'End If

    Unload Me
    'Unload All

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)



bookinfo.Visible = False
Label7.Visible = True
memberinfo.Visible = False
Label6.Visible = True
isuebook.Visible = False
Label8.Visible = True
returnbook.Visible = False
Label9.Visible = True
search.Visible = False
Label10.Visible = True
End Sub

Private Sub isuebook_Click()
If Not lbluser.Caption = "Administrator" Then
    MsgBox "You Do Not Have Any Right To Enter The Issue Form", vbCritical, "No Right To Perform Action"
Else
    frmisu.Show
End If
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
search.Visible = True
Label10.Visible = False
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
memberinfo.Visible = True
Label6.Visible = False
End Sub



Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
bookinfo.Visible = True
Label7.Visible = False
End Sub



Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
isuebook.Visible = True
Label8.Visible = False
End Sub



Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
returnbook.Visible = True
Label9.Visible = False
End Sub

Private Sub lbluser_Click()
choice = MsgBox("Are You Sure You Want To Log Out?", vbQuestion + vbYesNo, "Confirmation")
If choice = vbYes Then
    Me.Hide
    lbluser.Caption = ""
    MsgBox "Thank You For Using THE LIBRARY MANAGEMENT SYSTEM", vbInformation, "Log Out"
    frmsec.Show
    Unload Me
    
End If
End Sub

Private Sub memberinfo_Click()
If Not lbluser.Caption = "Administrator" Then
    frmmem.memdata.RecordsetType = 2 - snapshot
Else
    frmmem.memdata.RecordsetType = 1 - dynaset
End If

frmmem.Show
End Sub

Private Sub returnbook_Click()
If Not lbluser.Caption = "Administrator" Then
    MsgBox "You Do Not Have Any Right To Enter The Return Form", vbCritical, "No Right To Perform Action"
Else
    frmret.Show
End If
End Sub

Private Sub search_Click()
frmsrch.Show

End Sub
