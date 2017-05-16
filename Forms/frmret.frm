VERSION 5.00
Begin VB.Form frmret 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Return A Book"
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   12030
   Icon            =   "frmret.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmret.frx":08CA
   ScaleHeight     =   9015
   ScaleWidth      =   12030
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2400
      Top             =   0
   End
   Begin VB.TextBox txtbname 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3360
      TabIndex        =   19
      Top             =   7200
      Width           =   2055
   End
   Begin VB.TextBox txtstat 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3360
      TabIndex        =   18
      Top             =   7800
      Width           =   2055
   End
   Begin VB.TextBox txtmname 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8160
      TabIndex        =   17
      Top             =   7200
      Width           =   2055
   End
   Begin VB.TextBox txtbis 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   9360
      TabIndex        =   16
      Top             =   7800
      Width           =   855
   End
   Begin VB.Data bookdata 
      Caption         =   "bookdata"
      Connect         =   "Access"
      DatabaseName    =   "e:\Library Management System\DB\LMS97DB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tblbook"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data isudata 
      Caption         =   "isudata"
      Connect         =   "Access"
      DatabaseName    =   "e:\Library Management System\DB\LMS97DB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tbliss"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data memdata 
      Caption         =   "memdata"
      Connect         =   "Access"
      DatabaseName    =   "e:\Library Management System\DB\LMS97DB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tblmember"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.ComboBox cmbbid 
      Height          =   315
      Left            =   3600
      TabIndex        =   5
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox txtbid 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2160
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   4
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox txtmid 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   3
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox txtfine 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox txtrdt 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox txtlate 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10035
      TabIndex        =   23
      Top             =   8520
      Width           =   1935
   End
   Begin VB.Label lblclock 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8423
      TabIndex        =   24
      Top             =   8520
      Width           =   1695
   End
   Begin VB.Label lblhelp 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   68
      TabIndex        =   25
      Top             =   8520
      Width           =   8535
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Return Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   22
      Top             =   2640
      Width           =   3015
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Member Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   21
      Top             =   6720
      Width           =   3015
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Book Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   20
      Top             =   6720
      Width           =   3015
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LIBRARY MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   3195
      TabIndex        =   15
      Top             =   840
      Width           =   7095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   2280
      TabIndex        =   14
      Top             =   7800
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   2280
      TabIndex        =   13
      Top             =   7200
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Books Issued"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   6480
      TabIndex        =   12
      Top             =   7800
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   6600
      TabIndex        =   11
      Top             =   7200
      Width           =   735
   End
   Begin VB.Label lblbid 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Days Late"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Member ID"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label lbladd 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Book ID"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label lblcat 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Fine"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Returning Date"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Menu MNUDEFAULTERS 
      Caption         =   "&DEFAULTERS"
   End
   Begin VB.Menu MNUPRINT 
      Caption         =   "&PRINT LATE LETTER"
   End
End
Attribute VB_Name = "frmret"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim b As Integer
Dim c As String
Dim d As Integer
Dim Index As String
Dim a As Integer

Private Sub cmbbid_Click()

isudata.Recordset.MoveFirst
Do While Not isudata.Recordset.EOF
 If isudata.Recordset.B_ID = cmbbid.Text And isudata.Recordset.Fields("Status").Value = "Not Returned" Then
    txtbid.Text = cmbbid.Text
    txtrdt.Text = isudata.Recordset.Ret_date
 t = CDate(Date)
 f = CDate(txtrdt.Text)
  If f < t Then
    e = t - f
    g = CInt(e)
    h = g * 5
    txtlate.Text = g
    txtfine.Text = h
 Else
    txtlate.Text = 0
    txtfine.Text = 0
 End If
   choice = MsgBox("ARE YOU SURE YOU WANT TO RETURN THE BOOK?", vbYesNo, "CONFIRMATION")
If choice = vbYes Then
    With isudata.Recordset
     .Edit
      .Fields("Status") = "Returned"
      ![Ret_date] = Date
      ![Late] = txtlate.Text
      ![fine] = txtfine.Text
     .Update
    End With
    With bookdata.Recordset
        .Edit
        .Fields("Status") = "Available"
        .Update
    End With
    With memdata.Recordset
        .Edit
        a = CInt(.Fields("Issued"))
        a = a - 1
        .Fields("Issued") = a
        .Update
        memdata.Refresh
    End With
    If txtfine.Text = 0 Then
        MsgBox "Thank You For Returning The Book On Time", vbInformation, "Book Returned"
    Else
        MsgBox "Thank You For Returning The Book. Your Late Fine Is Tk. " & h & ".", vbInformation, "Book Returned"
    End If
        txtrdt.Text = ""
        txtfine.Text = ""
        txtlate.Text = ""
Else
    txtrdt.Text = ""
    txtfine.Text = ""
    txtlate.Text = ""
End If
    Exit Do
End If
  isudata.Recordset.MoveNext
Loop
cmbbid.Clear
txtbid.Text = ""
memdata.Refresh
cmbbid.Refresh
isudata.Recordset.MoveFirst
    Do While Not isudata.Recordset.EOF
     If isudata.Recordset.M_ID = txtmid.Text And isudata.Recordset.Status = "Not Returned" Then
      cmbbid.AddItem isudata.Recordset.B_ID
    End If
    isudata.Recordset.MoveNext
    Loop

End Sub

Private Sub Form_Load()
lblclock.Caption = Format(Time, "HH:MM:SS")
lbldate.Caption = Format(Date, "long date")
End Sub

Private Sub MNUDEFAULTERS_Click()

If denvLMS.rscomdefaulters.State = adStateOpen Then
        denvLMS.rscomdefaulters.Close
End If
    
If denvLMS.rscomdefaulters.State = adstateclose Then
    denvLMS.rscomdefaulters.Open
End If

If denvLMS.rscomdefaulters.RecordCount = 0 Then
    MsgBox "There Are No Defaulter Today!!", vbExclamation + vbOKOnly, "No Defaulters!"
    Exit Sub
Else
    rptdefaulters.Show
End If

End Sub

Private Sub MNUPRINT_Click()

memid = InputBox("Please Enter The Defaulter's ID You Want To Send A Letter To", "Enter Member ID")

If memid = "" Then
    MsgBox "Please Type In A Valid Member ID", vbExclamation + vbOKOnly
    Exit Sub
End If

    If denvLMS.rscomletter_Grouping.State = adStateOpen Then
        denvLMS.rscomletter_Grouping.Close
    End If
    
    denvLMS.comletter_Grouping Trim(memid)
    
    If denvLMS.rscomletter_Grouping.RecordCount = 0 Then
        MsgBox "The Member ID You Typed In Is Either Not Valid, Or The Member Is Not A Defaulter. Please Type-In The Member ID Correctly", vbExclamation + vbOKOnly, "Not Valid MEMBER ID"
        Exit Sub
    End If
    
    rptletter.Show

End Sub

Private Sub Timer1_Timer()
lblclock.Caption = Format(Time, "HH:MM:SS")
lbldate.Caption = Format(Date, "long date")
End Sub

Private Sub txtbid_Change()
txtbname.Text = ""
txtstat.Text = ""
If Not txtbid.Text = "" Then
    bookdata.Recordset.MoveFirst
    Do While Not bookdata.Recordset.EOF
    
    If bookdata.Recordset.B_ID = txtbid.Text Then
    txtbname.Text = bookdata.Recordset.B_Name
    txtstat.Text = bookdata.Recordset.Status
    
    Exit Do
End If
bookdata.Recordset.MoveNext
Loop
End If
End Sub

Private Sub txtmid_Change()
cmbbid.Clear
txtmname.Text = ""
txtbis.Text = ""
If Not txtmid.Text = "" Then
    memdata.Recordset.MoveFirst
    Do While Not memdata.Recordset.EOF
     If memdata.Recordset.M_ID = txtmid.Text Then
        txtmname.Text = memdata.Recordset.M_Name
        txtbis.Text = memdata.Recordset.Issued
     Exit Do
    End If
    memdata.Recordset.MoveNext
    Loop
End If
isudata.Recordset.MoveFirst
    Do While Not isudata.Recordset.EOF
     If isudata.Recordset.M_ID = txtmid.Text And isudata.Recordset.Status = "Not Returned" Then
      cmbbid.AddItem isudata.Recordset.B_ID
     End If
    isudata.Recordset.MoveNext
    Loop
End Sub
