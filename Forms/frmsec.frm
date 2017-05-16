VERSION 5.00
Begin VB.Form frmsec 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login Screen"
   ClientHeight    =   8940
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6735
   Icon            =   "frmsec.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmsec.frx":08CA
   ScaleHeight     =   8940
   ScaleWidth      =   6735
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdent 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Mortal Kombat 3"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4320
      MaskColor       =   &H0000C000&
      Picture         =   "frmsec.frx":BFAF
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6720
      UseMaskColor    =   -1  'True
      Width           =   1935
   End
   Begin VB.CommandButton cmdnew 
      BackColor       =   &H00E0E0E0&
      Caption         =   "New User"
      BeginProperty Font 
         Name            =   "Mortal Kombat 3"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      MaskColor       =   &H0000C000&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7920
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton cmdcan 
      BackColor       =   &H00E0E0E0&
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Mortal Kombat 3"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      MaskColor       =   &H0000C000&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7920
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtpass 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   7200
      Width           =   1935
   End
   Begin VB.TextBox txtuser 
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   6720
      Width           =   1935
   End
   Begin VB.Data secdata 
      Caption         =   "secdata"
      Connect         =   "Access"
      DatabaseName    =   "D:\Library Management System\DB\LMS97DB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select * from tblsec order by User"
      Top             =   9000
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "LIBRARY MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "Mortal Kombat 3"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   180
      TabIndex        =   7
      Top             =   600
      Width           =   6375
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   7200
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Menu mnuchange 
      Caption         =   "Chan&ge Password"
   End
End
Attribute VB_Name = "frmsec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim stat

Private Sub EnableAll()

cmdent.Enabled = False
cmdcan.Visible = True
txtuser.Text = ""
txtpass.Text = ""
mnuchange.Enabled = False

End Sub

Private Sub DisableAll()

cmdent.Enabled = True
cmdcan.Visible = False
txtuser.Text = ""
txtpass.Text = ""
mnuchange.Enabled = True

End Sub

Private Sub cmdcan_Click()
secdata.Recordset.CancelUpdate
secdata.RecordSource = "select * from tblsec"
secdata.Refresh
cmdnew.Caption = "New User"
cmdcan.Visible = False
cmdent.Enabled = True
txtuser.Text = ""
txtpass.Text = ""
mnuchange.Enabled = True
End Sub

Private Sub cmdent_Click()

Dim user1, pass1, user2, pass2
If txtuser.Text = "" Then
    MsgBox "Please Enter Your User Name And Your Password In The Following Text Boxes.", vbExclamation, "Enter Username And Password"
    txtuser.SetFocus
Else
user1 = txtuser.Text
pass1 = txtpass.Text
secdata.RecordSource = "select * from tblsec where User='" & user1 & "' and Password='" & pass1 & "'"
secdata.Refresh

If secdata.Recordset.RecordCount = 0 Then
    MsgBox "Your Login Info Is Incorrect", vbCritical, "Access Denied"
Else
 user2 = secdata.Recordset.Fields("user").Value
 pass2 = secdata.Recordset.Fields("Password").Value
    If user1 = txtuser.Text And pass1 = txtpass.Text Then
        MsgBox "Welcome To THE GENERAL PUBLIC LIBRARY Mr. " & user1, vbInformation, "Accesss Granted"
        frmmain2.lbluser.Caption = txtuser.Text
        Unload Me
        frmmain2.Show

    Else
        MsgBox "The Password Is Incorrect", vbCritical, "Access Denied"
        txtuser.Text = ""
        txtpass.Text = ""
        secdata.RecordSource = "select * from tblsec"
        secdata.Refresh
End If
End If
End If

End Sub

Private Sub cmdnew_Click()

If cmdnew.Caption = "New User" Then
    
    secdata.Recordset.AddNew
    txtuser.SetFocus
    cmdnew.Caption = "Save"
    EnableAll
    
ElseIf cmdnew.Caption = "Save" Then
    
    If txtuser.Text = "" Then
        MsgBox "Please Enter A Valid User Name In The Respective Text Box", vbInformation, "Input Error"
        txtuser.SetFocus
        Exit Sub
    Else

user1 = txtuser.Text

    secdata.RecordSource = "select * from tblsec where User='" & user1 & "'"
    secdata.Refresh
If Not stat = "Edit" Then
    If secdata.Recordset.RecordCount = 0 Then

        With secdata.Recordset
            .AddNew
            ![User] = txtuser.Text
            ![Password] = txtpass.Text
            .Update
        End With
        MsgBox "Congratulations. The Record Of The New Member Has Been Saved", vbInformation, "Update Successful"
    Else
    
        user2 = secdata.Recordset.Fields("user").Value
        
        If user1 = user2 Then
            MsgBox "The User Name You Typed Is Already Present. Please Choose Another Username", vbExclamation, "User Name Already Exists"
            secdata.Refresh
            txtuser.SetFocus
            txtpass.Text = ""
            Exit Sub
        End If
    End If
Else
        With secdata.Recordset
            .Edit
            ![User] = txtuser.Text
            ![Password] = txtpass.Text
            .Update
        End With
End If
    secdata.RecordSource = "select * from tblsec"
    secdata.Refresh
    cmdnew.Caption = "New User"
    DisableAll
End If
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Image1.Visible = True
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Image1.Visible = False
End Sub

Private Sub mnuchange_Click()
Dim username1
stat = "Edit"
username1 = InputBox("Please Enter The UserName You Want To Change Password", "Type In UserName")
secdata.RecordSource = "select * from tblsec where User='" & username1 & "'"
secdata.Refresh
If Not secdata.Recordset.RecordCount = 0 Then
    If username1 = secdata.Recordset.Fields("User").Value Then
        EnableAll
        txtuser.Text = username1
        secdata.Recordset.Edit
        cmdnew.Caption = "Save"
    Else
        MsgBox "No User Found Named: " & username1, vbExclamation, "Error"
        DisableAll
        cmdnew.Caption = "New User"
        secdata.RecordSource = "Select * from tblsec"
        secdata.Refresh
    End If
Else
    MsgBox "No User Found Named: " & username1, vbExclamation, "Error"
End If

End Sub
