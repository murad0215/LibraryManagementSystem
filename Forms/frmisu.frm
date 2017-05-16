VERSION 5.00
Begin VB.Form frmisu 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Issue A Book"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   12030
   Icon            =   "frmisu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   12030
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbbid 
      Height          =   315
      Left            =   10560
      TabIndex        =   33
      Top             =   4440
      Width           =   1335
   End
   Begin VB.TextBox txtrmid 
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
      Left            =   9120
      MaxLength       =   6
      TabIndex        =   31
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1080
      Top             =   840
   End
   Begin VB.TextBox txtstat 
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
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   6720
      Width           =   2055
   End
   Begin VB.TextBox txtbname 
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
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   6240
      Width           =   2055
   End
   Begin VB.TextBox txtbis 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   6600
      Width           =   495
   End
   Begin VB.TextBox txtmname 
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
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   6120
      Width           =   1815
   End
   Begin VB.TextBox txtvalid 
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
      Left            =   8760
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   7200
      Width           =   1455
   End
   Begin VB.TextBox txtiid 
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
      Left            =   2160
      TabIndex        =   6
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox txtrdt 
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
      Left            =   2160
      TabIndex        =   5
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox txtidt 
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
      Left            =   2160
      TabIndex        =   4
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox txtmid 
      BackColor       =   &H00FFFFFF&
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
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   2
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox txtbid 
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
      Height          =   375
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   3
      Top             =   3240
      Width           =   1335
   End
   Begin VB.FileListBox filphotos 
      Height          =   285
      Left            =   3960
      Pattern         =   "*.jpg;*.jpeg;*.bmp;*.gif;*.png"
      TabIndex        =   0
      Top             =   4680
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data memdata 
      Caption         =   "memdata"
      Connect         =   "Access"
      DatabaseName    =   "e:\Library Management System\DB\LMS97DB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tblmember"
      Top             =   6000
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.Data isudata 
      Caption         =   "isudata"
      Connect         =   "Access"
      DatabaseName    =   "e:\Library Management System\DB\LMS97DB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tbliss"
      Top             =   6360
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.Data bookdata 
      Caption         =   "bookdata"
      Connect         =   "Access"
      DatabaseName    =   "e:\Library Management System\DB\LMS97DB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tblbook"
      Top             =   6720
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&CANCEL"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton cmdsave 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "&SAVE"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton cmdadd 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&ADD"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Re-Issue a Book"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   34
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label Label12 
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7800
      TabIndex        =   32
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
      Left            =   9900
      TabIndex        =   28
      Top             =   7800
      Width           =   2055
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
      Left            =   8325
      TabIndex        =   29
      Top             =   7800
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
      Left            =   90
      TabIndex        =   30
      Top             =   7800
      Width           =   8415
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Member Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   25
      Top             =   5760
      Width           =   3015
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Book Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   24
      Top             =   5760
      Width           =   3015
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Issue a Book"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   23
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LIBRARY MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   2228
      TabIndex        =   22
      Top             =   360
      Width           =   7575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2640
      TabIndex        =   21
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2640
      TabIndex        =   20
      Top             =   6720
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   7200
      TabIndex        =   17
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Books Issued"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   7200
      TabIndex        =   16
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Valid Till"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   7200
      TabIndex        =   15
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label lblbid 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Issue ID"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2280
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2760
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label lblcat 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Issue Date"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   3720
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Image imgStudent 
      BorderStyle     =   1  'Fixed Single
      Height          =   2415
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Menu mnumenu 
      Caption         =   "&Menu"
      Begin VB.Menu mnuadd 
         Caption         =   "&ADD"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnusave 
         Caption         =   "&SAVE"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnucancel 
         Caption         =   "&CANCEL"
         Enabled         =   0   'False
         Shortcut        =   ^{INSERT}
      End
   End
   Begin VB.Menu mnusrch 
      Caption         =   "Search A &Book"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "frmisu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim b As Integer
Dim c As String
Dim d As Integer
Dim Index As String
Dim a As Integer

Private Sub dispic()

'Displayed Photo
If IsNull(memdata.Recordset.Fields("Pic-file")) Or memdata.Recordset.Fields("Pic-file") = "" Then
  Photo = ""
  imgStudent.Picture = LoadPicture("")
  filphotos.Refresh
Else
  Photo = memdata.Recordset.Fields("Pic-File")
  
'Match file to listing in file control
  If filphotos.ListCount <> 0 Then
    For N = 0 To filphotos.ListCount
      If Photo = filphotos.List(N) Then
        filphotos.ListIndex = N
        imgStudent.Picture = LoadPicture(filphotos.Path + "\" + Photo)
        Exit Sub
      End If
    Next N
  End If
  imgStudent.Picture = LoadPicture("")
  MsgBox "Photo not in StudentPhotos folder.", vbOKOnly + vbInformation, "File Not Found"
End If
End Sub


Private Sub EnableAll()

mnusave.Enabled = True
mnucancel.Enabled = True
txtmid.Enabled = True
txtbid.Enabled = True
mnuadd.Enabled = False
mnusrch.Enabled = True
cmdadd.Enabled = False
cmdsave.Enabled = True
cmdcancel.Enabled = True

End Sub

Private Sub DisableAll()

mnusave.Enabled = False
mnucancel.Enabled = False
txtmid.Enabled = False
txtbid.Enabled = False
mnuadd.Enabled = True
mnusrch.Enabled = False
cmdadd.Enabled = True
cmdsave.Enabled = False
cmdcancel.Enabled = False

End Sub
Private Sub ClearAll()

txtbid.Text = ""
txtmid.Text = ""
txtiid.Text = ""
txtidt.Text = ""
txtrdt.Text = ""
txtbname.Text = ""
txtstat.Text = ""
txtmname.Text = ""
txtbis.Text = ""

End Sub

Private Sub cbobid_Click()
isudata.Recordset.MoveFirst
    Do While Not isudata.Recordset.EOF
If isudata.Recordset.B_ID = cbobid.Text And isudata.Recordset.Fields("Status").Value = "Not Returned" Then
    txtbid.Text = cbobid.Text
   choice = MsgBox("ARE YOU SURE YOU WANT TO RETURN THE BOOK?", vbYesNo, "CONFIRMATION")
If choice = vbYes Then
    With isudata.Recordset
     .Edit
      .Fields("Status") = "Returned"
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
End If
    Exit Do
End If
   isudata.Recordset.MoveNext
Loop
cbobid.Clear
txtbid.Text = ""
isudata.Recordset.MoveFirst
    Do While Not isudata.Recordset.EOF
        If isudata.Recordset.M_ID = txtmid.Text And isudata.Recordset.Status = "Not Returned" Then
            cbobid.AddItem isudata.Recordset.B_ID
        End If
    isudata.Recordset.MoveNext
Loop

End Sub

Private Sub cmbbid_Click()
isudata.Recordset.MoveFirst
Do While Not isudata.Recordset.EOF

   choice = MsgBox("ARE YOU SURE YOU WANT TO REISSUE THE BOOK?", vbYesNo, "CONFIRMATION")
If choice = vbYes Then
    With isudata.Recordset
     .Edit
        ![Iss_date] = Date
        ![Ret_date] = Date + 7
     .Update
    End With
End If
    Exit Do

  isudata.Recordset.MoveNext
Loop
cmbbid.Clear
memdata.Refresh
cmbbid.Refresh
isudata.Recordset.MoveFirst
    Do While Not isudata.Recordset.EOF
     If isudata.Recordset.M_ID = txtrmid.Text And isudata.Recordset.Status = "Not Returned" Then
      cmbbid.AddItem isudata.Recordset.B_ID
    End If
    isudata.Recordset.MoveNext
    Loop
End Sub

Private Sub cmdadd_Click()
EnableAll
ClearAll
isudata.Recordset.MoveLast

If isudata.Recordset.RecordCount = 0 Then
    txtiid.Text = "I00001"
Else

If isudata.Recordset.RecordCount < 10 Then
    isudata.Recordset.MoveLast
    
    txtiid.Text = isudata.Recordset.Fields("Iss_ID").Value
    iid = Mid(txtiid.Text, 2, 6)
    iid = (iid + 1)
    txtiid.Text = "I0000" + CStr(iid)
ElseIf isudata.Recordset.RecordCount > 9 Then
    isudata.Recordset.MoveLast

    txtiid.Text = isudata.Recordset.Fields("Iss_ID").Value
    iid = Right(txtiid.Text, 2)
    iid = (iid + 1)
    txtiid.Text = "I000" + CStr(iid)
ElseIf isudata.Recordset.RecordCount >= 99 And 999 Then
    isudata.Recordset.MoveLast

    txtiid.Text = isudata.Recordset.Fields("Iss_ID").Value
    iid = Right(txtiid.Text, 2)
    iid = (iid + 1)
    txtiid.Text = "I00" + CStr(iid)
ElseIf isudata.Recordset.RecordCount >= 999 And 9999 Then
    isudata.Recordset.MoveLast

    txtiid.Text = isudata.Recordset.Fields("Iss_ID").Value
    iid = Right(txtiid.Text, 2)
    iid = (iid + 1)
    txtiid.Text = "I0" + CStr(iid)
ElseIf isudata.Recordset.RecordCount >= 9999 And 99999 Then
    isudata.Recordset.MoveLast

    txtiid.Text = isudata.Recordset.Fields("Iss_ID").Value
    iid = Right(txtiid.Text, 2)
    iid = (iid + 1)
    txtiid.Text = "I" + CStr(iid)
End If
    
End If

txtidt.Text = Date
txtrdt.Text = Date + 7


End Sub

Private Sub cmdadd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblhelp.Caption = "Click This Button To Add A New Record"
End Sub

Private Sub cmdcancel_Click()
ClearAll
DisableAll
txtmid.Text = ""
txtbid.Text = ""
txtbname.Text = ""
txtstat.Text = ""
txtmname.Text = ""
txtbis.Text = ""
isudata.Refresh
imgStudent.Refresh
frmisu.Refresh
filphotos.Refresh
End Sub

Private Sub cmdcancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblhelp.Caption = "Click This Button To Cancel The Update Process"
End Sub

Private Sub cmdsave_Click()

If txtmname.Text = "" Then
    MsgBox "The Member ID You Have Typed-In Is Not Valid. Please Enter A Valid Memeber ID", vbCritical + vbOKOnly, "MEMBER ID Not Valid"
    txtmid.SetFocus
    Exit Sub
End If

If txtbname.Text = "" Then
    MsgBox "The Book ID You Have Typed-In Is Not Valid. Please Enter A Valid Book ID", vbCritical + vbOKOnly, "BOOK ID Not Valid"
    txtbid.SetFocus
    Exit Sub
End If

If txtbid.Text = "" Then
    MsgBox "Please Select The Book You Want To Issue.", vbInformation, "Please Select A Book"
    Exit Sub
End If

If txtstat.Text = "Issued" Then
    MsgBox "This Book Has Already Been Issued.  Please Select Another Book", vbInformation, "Please Select Another Book"
    Exit Sub
End If

If txtmid.Text = "" Then
    MsgBox "Please Select The Member Who Wants To Issue A Book.", vbInformation, "Please Select A Member"
    Exit Sub
End If

If txtbis.Text >= "3" Then
    MsgBox "The Member Has Already Issued 3 Books. He/She Cannot Issue Any More Book Until And Unless He/She Returns One", vbInformation, "The Member Cannot Issue Any More Books"
    Exit Sub
End If

Dim date1, date2 As Date
date1 = txtvalid.Text
date2 = Date
If date1 < date2 Then
    MsgBox "This Member Is Not Valid Anymore. Please Ask The Member To Renew His/Her Membership Before He/She Can Issue Any More Book", vbExclamation, "Member Not Valid Anymore"
    Exit Sub
End If

With isudata.Recordset
  
If isudata.Recordset.RecordCount = 0 Then
.AddNew
    ![Iss_ID] = txtiid.Text
    ![B_ID] = txtbid.Text
    ![B_Name] = txtbname.Text
    ![M_ID] = txtmid.Text
    ![M_Name] = txtmname.Text
    ![Iss_date] = txtidt.Text
    ![Ret_date] = txtrdt.Text
    ![Status] = "Not Returned"
    .Update
    
With memdata.Recordset
    .Edit

    b = CInt(txtbis.Text)
    b = b + 1
    .Fields("Issued").Value = b
    .Update
    
End With

With bookdata.Recordset
    .Edit
    c = "Issued"
    .Fields("Status").Value = c
    .Update
End With
    
Else
.AddNew

    ![Iss_ID] = txtiid.Text
    ![B_ID] = txtbid.Text
    ![B_Name] = txtbname.Text
    ![M_ID] = txtmid.Text
    ![M_Name] = txtmname.Text
    ![Iss_date] = txtidt.Text
    ![Ret_date] = txtrdt.Text
    ![Status] = "Not Returned"
   .Update
   
With memdata.Recordset
    .Edit

    b = CInt(txtbis.Text)
    b = b + 1
    memdata.Recordset.Fields("Issued").Value = b
    memdata.Recordset.Update
    
End With

With bookdata.Recordset
    .Edit
    c = "Issued"
    bookdata.Recordset.Fields("Status").Value = c
    bookdata.Recordset.Update
End With

End If

End With
MsgBox "Update Successful.", vbInformation, "Update Successful"
Call ClearAll
Call DisableAll
bookdata.Refresh
txtmid.Text = ""
txtbid.Text = ""
txtbname.Text = ""
txtstat.Text = ""
txtmname.Text = ""
txtbis.Text = ""

End Sub

Private Sub cmdsave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblhelp.Caption = "Click This Button To Complete The Update Process"
End Sub

Private Sub filphotos_Click()
Photo = filphotos.FileName
imgStudent.Picture = LoadPicture(filphotos.Path + "\" + Photo)
End Sub

Private Sub Form_Load()
filphotos.Path = App.Path + "\MEMBER PHOTOS\"
lblclock.Caption = Format(Time, "HH:MM:SS")
lbldate.Caption = Format(Date, "long date")
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblhelp.Caption = ""
End Sub

Private Sub lblclock_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblclock.ToolTipText = Format(Date, "long date")
End Sub

Private Sub mnuadd_Click()
EnableAll
ClearAll

If isudata.Recordset.RecordCount = 0 Then
    txtiid.Text = "I00001"
Else

If isudata.Recordset.RecordCount < 9 Then
    isudata.Recordset.MoveLast

    txtiid.Text = isudata.Recordset.Fields("Iss_ID").Value
    iid = Right(txtiid.Text, 1)
    iid = (iid + 1)
    txtiid.Text = "I0000" + CStr(iid)
ElseIf isudata.Recordset.RecordCount >= 9 And 99 Then
    isudata.Recordset.MoveLast

    txtiid.Text = isudata.Recordset.Fields("Iss_ID").Value
    iid = Right(txtiid.Text, 2)
    iid = (iid + 1)
    txtiid.Text = "I000" + CStr(iid)
ElseIf isudata.Recordset.RecordCount >= 99 And 999 Then
    isudata.Recordset.MoveLast

    txtiid.Text = isudata.Recordset.Fields("Iss_ID").Value
    iid = Right(txtiid.Text, 2)
    iid = (iid + 1)
    txtiid.Text = "I00" + CStr(iid)
ElseIf isudata.Recordset.RecordCount >= 999 And 9999 Then
    isudata.Recordset.MoveLast

    txtiid.Text = isudata.Recordset.Fields("Iss_ID").Value
    iid = Right(txtiid.Text, 2)
    iid = (iid + 1)
    txtiid.Text = "I0" + CStr(iid)
ElseIf isudata.Recordset.RecordCount >= 9999 And 99999 Then
    isudata.Recordset.MoveLast

    txtiid.Text = isudata.Recordset.Fields("Iss_ID").Value
    iid = Right(txtiid.Text, 2)
    iid = (iid + 1)
    txtiid.Text = "I" + CStr(iid)
End If
    
End If

txtidt.Text = Date
txtrdt.Text = Date + 7

End Sub

Private Sub mnucancel_Click()
ClearAll
DisableAll
txtmid.Text = ""
txtbid.Text = ""
txtbname.Text = ""
txtstat.Text = ""
txtmname.Text = ""
txtbis.Text = ""
isudata.Refresh
End Sub

Private Sub mnure_Click()

End Sub

Private Sub mnusave_Click()

If txtmname.Text = "" Then
    MsgBox "The Member ID You Have Typed-In Is Not Valid. Please Enter A Valid Memeber ID", vbCritical + vbOKOnly, "MEMBER ID Not Valid"
    txtmid.SetFocus
    Exit Sub
End If

If txtbname.Text = "" Then
    MsgBox "The Book ID You Have Typed-In Is Not Valid. Please Enter A Valid Book ID", vbCritical + vbOKOnly, "BOOK ID Not Valid"
    txtbid.SetFocus
    Exit Sub
End If

If txtbid.Text = "" Then
    MsgBox "Please Select The Book You Want To Issue.", vbInformation, "Please Select A Book"
    Exit Sub
End If

If txtstat.Text = "Issued" Then
    MsgBox "This Book Has Already Been Issued.  Please Select Another Book", vbInformation, "Please Select Another Book"
    Exit Sub
End If

If txtmid.Text = "" Then
    MsgBox "Please Select The Member Who Wants To Issue A Book.", vbInformation, "Please Select A Member"
    Exit Sub
End If

If txtbis.Text >= "3" Then
    MsgBox "The Member Has Already Issued 3 Books. He/She Cannot Issue Any More Book Until And Unless He/She Returns One", vbInformation, "The Member Cannot Issue Any More Books"
    Exit Sub
End If

Dim date1, date2 As Date
date1 = txtvalid.Text
date2 = Date
If date1 < date2 Then
    MsgBox "This Member Is Not Valid Anymore. Please Ask The Member To Renew His/Her Membership Before He/She Can Issue Any More Book", vbExclamation, "Member Not Valid Anymore"
    Exit Sub
End If

With isudata.Recordset
  
If isudata.Recordset.RecordCount = 0 Then
.AddNew
    ![Iss_ID] = txtiid.Text
    ![B_ID] = txtbid.Text
    ![B_Name] = txtbname.Text
    ![M_ID] = txtmid.Text
    ![M_Name] = txtmname.Text
    ![Iss_date] = txtidt.Text
    ![Ret_date] = txtrdt.Text
    ![Status] = "Not Returned"
    .Update
    
With memdata.Recordset
    .Edit

    b = CInt(txtbis.Text)
    b = b + 1
    .Fields("Issued").Value = b
    .Update
    
End With

With bookdata.Recordset
    .Edit
    c = "Issued"
    .Fields("Status").Value = c
    .Update
End With
    
Else
.AddNew

    ![Iss_ID] = txtiid.Text
    ![B_ID] = txtbid.Text
    ![B_Name] = txtbname.Text
    ![M_ID] = txtmid.Text
    ![M_Name] = txtmname.Text
    ![Iss_date] = txtidt.Text
    ![Ret_date] = txtrdt.Text
    ![Status] = "Not Returned"
   .Update
   
With memdata.Recordset
    .Edit

    b = CInt(txtbis.Text)
    b = b + 1
    memdata.Recordset.Fields("Issued").Value = b
    memdata.Recordset.Update
    
End With

With bookdata.Recordset
    .Edit
    c = "Issued"
    bookdata.Recordset.Fields("Status").Value = c
    bookdata.Recordset.Update
End With

End If

End With
MsgBox "Update Successful.", vbInformation, "Update Successful"
Call ClearAll
Call DisableAll
bookdata.Refresh
txtmid.Text = ""
txtbid.Text = ""
txtbname.Text = ""
txtstat.Text = ""
txtmname.Text = ""
txtbis.Text = ""
End Sub

Private Sub mnusrch_Click()

Dim bname
Dim counter As Integer
bname = InputBox("Please Enter The Book's Name You Want To Search", "Enter Book Name")
bookdata.RecordSource = "select * from tblbook where b_name='" & bname & "' and Status='Available'"
bookdata.Refresh
If bookdata.Recordset.RecordCount = 0 Then
    MsgBox "No Books Found.", vbExclamation, "No Books Found"
Else
    BookID = bookdata.Recordset.Fields("B_ID").Value
    txtbid.Text = BookID
End If

End Sub

Private Sub Timer1_Timer()
lblclock.Caption = Format(Time, "HH:MM:SS")
lbldate.Caption = Format(Date, "long date")
End Sub

Private Sub txtbid_Change()

txtbname.Text = ""
txtstat.Text = ""
If Not txtbid.Text = "" Then
bookdata.RecordSource = "select * from tblbook"
bookdata.Refresh
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

txtmname.Text = ""
txtbis.Text = ""
txtvalid.Text = ""
If Not txtmid.Text = "" Then
    memdata.Recordset.MoveFirst
    Do While Not memdata.Recordset.EOF
    If memdata.Recordset.M_ID = txtmid.Text Then
        dispic
        If memdata.Recordset.Fields("Pic-File").Value = "" Then
            imgStudent.Picture = LoadPicture("")
            filphotos.Refresh
        End If
        txtmname.Text = memdata.Recordset.M_Name
        txtbis.Text = memdata.Recordset.Issued
        txtvalid.Text = memdata.Recordset.Renew_Date
        If txtmname.Text = "" Then
            imgStudent.Picture = LoadPicture("")
            filphotos.Refresh
        End If
        Exit Do
    End If
    memdata.Recordset.MoveNext
    Loop
End If

End Sub

Private Sub txtrmid_Change()
cmbbid.Clear
txtmname.Text = ""
txtbis.Text = ""
txtvalid.Text = ""
If Not txtrmid.Text = "" Then
    memdata.Recordset.MoveFirst
    Do While Not memdata.Recordset.EOF
     If memdata.Recordset.M_ID = txtrmid.Text Then
        txtmname.Text = memdata.Recordset.M_Name
        txtbis.Text = memdata.Recordset.Issued
        txtvalid.Text = memdata.Recordset.Renew_Date
     Exit Do
    End If
    memdata.Recordset.MoveNext
    Loop
End If
isudata.Recordset.MoveFirst
    Do While Not isudata.Recordset.EOF
     If isudata.Recordset.M_ID = txtrmid.Text And isudata.Recordset.Status = "Not Returned" Then
      cmbbid.AddItem isudata.Recordset.B_ID
     End If
    isudata.Recordset.MoveNext
    Loop
End Sub
