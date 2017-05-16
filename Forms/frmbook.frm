VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frmbook 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Book Information Entry & Retrivation Form"
   ClientHeight    =   9105
   ClientLeft      =   2985
   ClientTop       =   1455
   ClientWidth     =   13050
   DrawMode        =   2  'Blackness
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmbook.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9221.384
   ScaleMode       =   0  'User
   ScaleWidth      =   13050
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
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Click This Button To Cancel The Update Process"
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4680
      Top             =   4680
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
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Click This Button To Complete The Update Process"
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton cmddel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&DELETE"
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
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Click This Button To Delete The Currently Selected Record"
      Top             =   2760
      Width           =   1935
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
      Height          =   405
      Left            =   2025
      TabIndex        =   10
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txtpub 
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
      Left            =   2025
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3840
      Width           =   2055
   End
   Begin VB.TextBox txtaut 
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
      Left            =   2025
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   3360
      Width           =   2055
   End
   Begin VB.TextBox txtCat 
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
      Left            =   2025
      TabIndex        =   7
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox txtISBN 
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
      Left            =   2025
      Locked          =   -1  'True
      MaxLength       =   13
      TabIndex        =   6
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox txtname 
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
      Left            =   2025
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1920
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
      Height          =   405
      Left            =   2025
      TabIndex        =   4
      Top             =   4800
      Width           =   2055
   End
   Begin VB.TextBox txtshelf 
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
      Left            =   2025
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   4320
      Width           =   2055
   End
   Begin VB.ComboBox cmbCat 
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
      Left            =   4200
      TabIndex        =   2
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Data bookdata 
      Caption         =   "bookdata"
      Connect         =   "Access"
      DatabaseName    =   "E:\Library Management System\DB\LMS97DB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select * from tblbook order by B_ID"
      Top             =   9840
      Visible         =   0   'False
      Width           =   2340
   End
   Begin MSDBCtls.DBCombo cmbbid 
      Bindings        =   "frmbook.frx":08CA
      Height          =   390
      Left            =   3465
      TabIndex        =   0
      Top             =   1440
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   688
      _Version        =   393216
      ListField       =   "B_ID"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBGrid.DBGrid bgrid 
      Bindings        =   "frmbook.frx":08E1
      Height          =   1935
      Left            =   105
      OleObjectBlob   =   "frmbook.frx":08F8
      TabIndex        =   1
      Top             =   6720
      Width           =   11775
   End
   Begin VB.CommandButton cmdedit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&EDIT"
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
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Click This Button To Edit The Currently Selected Record"
      Top             =   2160
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
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Click This Button To Add A New Record"
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton cmdlast 
      BackColor       =   &H00E0E0E0&
      Caption         =   "LAST"
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
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Click This Button To Jump To The Last Record"
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CommandButton cmdnext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "NEXT"
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
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Click This Button To Jump To The Last Record"
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CommandButton cmdpre 
      BackColor       =   &H00E0E0E0&
      Caption         =   "PREVIOUS"
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
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Click This Button To Move To The Previous Record"
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CommandButton cmdfirst 
      BackColor       =   &H00E0E0E0&
      Caption         =   "FIRST"
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Click This Button To Move To The First Record"
      Top             =   6000
      Width           =   1935
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
      Left            =   9840
      TabIndex        =   30
      Top             =   8640
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
      Left            =   8280
      TabIndex        =   29
      Top             =   8640
      Width           =   1575
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
      Left            =   45
      TabIndex        =   28
      Top             =   8640
      Width           =   8295
   End
   Begin VB.Label Label2 
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
      Height          =   735
      Left            =   1800
      TabIndex        =   19
      Top             =   360
      Width           =   8415
   End
   Begin VB.Label lblbid 
      BackColor       =   &H00C0C0C0&
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
      Height          =   375
      Left            =   705
      TabIndex        =   18
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   705
      TabIndex        =   17
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblISBN 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "ISBN"
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
      Left            =   705
      TabIndex        =   16
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Author"
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
      Left            =   705
      TabIndex        =   15
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label lblcat 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
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
      Left            =   705
      TabIndex        =   14
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   705
      TabIndex        =   13
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Shelf"
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
      Left            =   705
      TabIndex        =   12
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Publisher"
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
      Left            =   705
      TabIndex        =   11
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Menu mnumenu 
      Caption         =   "&Menu"
      Begin VB.Menu mnuadd 
         Caption         =   "&ADD RECORD"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnusave 
         Caption         =   "&SAVE RECORD"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuedit 
         Caption         =   "&EDIT RECORD"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnucancel 
         Caption         =   "&CANCEL UPDATE"
         Enabled         =   0   'False
         Shortcut        =   ^{INSERT}
      End
      Begin VB.Menu mnudel 
         Caption         =   "&DELETE RECORD"
         Shortcut        =   +{DEL}
      End
   End
End
Attribute VB_Name = "frmbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub ClearAll()

txtbid.Text = ""
txtname.Text = ""
txtaut.Text = ""
txtpub.Text = ""
txtISBN.Text = ""
txtCat.Text = ""
txtshelf.Text = ""
txtstat.Text = ""
cmbCat.Text = ""

End Sub

Private Sub EnableAll()

cmdsave.Enabled = True
cmdcancel.Enabled = True
cmdadd.Enabled = False
cmddel.Enabled = False
cmdedit.Enabled = False
txtname.Locked = False
txtISBN.Locked = False
txtaut.Locked = False
txtpub.Locked = False
txtshelf.Locked = False
cmbCat.Enabled = True
cmbbid.Enabled = False
mnuadd.Enabled = False
mnudel.Enabled = False
mnuedit.Enabled = False
mnusave.Enabled = True
mnucancel.Enabled = True
cmdfirst.Enabled = False
cmdnext.Enabled = False
cmdpre.Enabled = False
cmdlast.Enabled = False

End Sub

Private Sub DisableAll()

cmdsave.Enabled = False
cmdcancel.Enabled = False
cmdadd.Enabled = True
cmddel.Enabled = True
cmdedit.Enabled = True
txtname.Locked = True
txtISBN.Locked = True
txtaut.Locked = True
txtpub.Locked = True
txtshelf.Locked = True
cmbCat.Enabled = False
txtCat.Enabled = False
'framenav.Visible = True
cmbbid.Enabled = True
cmbCat.Text = ""
mnuadd.Enabled = True
mnuedit.Enabled = True
mnudel.Enabled = True
mnucancel.Enabled = False
mnusave.Enabled = False
cmdfirst.Enabled = True
cmdnext.Enabled = True
cmdpre.Enabled = True
cmdlast.Enabled = True

End Sub

Private Sub bgrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblhelp.Caption = ""
End Sub

Private Sub cmbbid_Click(Area As Integer)
If bookdata.Recordset.RecordCount = 0 Then
    MsgBox "There are no records!!", vbCritical, "No Records!!!"
Else

    bookdata.Recordset.Bookmark = cmbbid.SelectedItem
    
    With bookdata.Recordset
    
     txtbid.Text = ![B_ID]
     txtname.Text = ![B_Name]
     txtaut.Text = ![author]
     txtpub.Text = ![Publisher]
     txtISBN.Text = ![isbn]
     txtCat.Text = ![Category]
     txtshelf.Text = ![Shelf]
     txtstat.Text = ![Status]
     
    End With

End If
End Sub

Private Sub cmbbid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblhelp.Caption = "Click This Drop-Down Box To Jump To Your Required Record"
End Sub

Private Sub cmbCat_Click()
If cmbCat.Text = "Other" Then
    txtCat.Enabled = True
    cmbCat.Enabled = False
Else
    txtCat.Text = cmbCat.Text
End If
End Sub

Private Sub cmdadd_Click()
If bookdata.RecordsetType = 2 - snapshot Then
    MsgBox "You Do Not Have The Right To Make Any Changes To The Database.", vbCritical, "No Right To Perform Action"
Else
EnableAll
cmbbid.Text = ""
ClearAll

If bookdata.Recordset.RecordCount = 0 Then
    bookdata.Recordset.AddNew
    txtbid.Text = "B00001"
Else
If bookdata.Recordset.RecordCount < 9 Then
    bookdata.Recordset.MoveLast
    txtbid.Text = bookdata.Recordset.Fields("B_ID").Value
    bid = Right(txtbid.Text, 1)
    bid = (bid + 1)
    txtbid.Text = "B0000" + CStr(bid)
    bookdata.Recordset.AddNew
ElseIf bookdata.Recordset.RecordCount >= 9 And 99 Then
    bookdata.Recordset.MoveLast
    txtbid.Text = bookdata.Recordset.Fields("B_ID").Value
    bid = Right(txtbid.Text, 2)
    bid = (bid + 1)
    txtbid.Text = "B000" + CStr(bid)
    bookdata.Recordset.AddNew
ElseIf bookdata.Recordset.RecordCount >= 99 And 999 Then
    bookdata.Recordset.MoveLast
    txtbid.Text = bookdata.Recordset.Fields("B_ID").Value
    bid = Right(txtbid.Text, 2)
    bid = (bid + 1)
    txtbid.Text = "B00" + CStr(bid)
    bookdata.Recordset.AddNew
ElseIf bookdata.Recordset.RecordCount >= 999 And 9999 Then
    bookdata.Recordset.MoveLast
    txtbid.Text = bookdata.Recordset.Fields("B_ID").Value
    bid = Right(txtbid.Text, 2)
    bid = (bid + 1)
    txtbid.Text = "B0" + CStr(bid)
    bookdata.Recordset.AddNew
ElseIf bookdata.Recordset.RecordCount >= 9999 And 99999 Then
    bookdata.Recordset.MoveLast
    txtbid.Text = bookdata.Recordset.Fields("B_ID").Value
    bid = Right(txtbid.Text, 2)
    bid = (bid + 1)
    txtbid.Text = "B" + CStr(bid)
    bookdata.Recordset.AddNew
End If
    
End If
txtname.SetFocus
txtstat.Text = "Available"
End If

End Sub

Private Sub cmdcan_Click()
ClearAll
DisableAll
cmbbid.Text = ""
bookdata.Refresh
End Sub

Private Sub cmdadd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblhelp.Caption = "Click This Button To Add A New Record"
End Sub

Private Sub cmdcancel_Click()

DisableAll
bookdata.Recordset.CancelUpdate
bookdata.Refresh
txtbid.Text = ""
cmbbid.Text = ""

End Sub

Private Sub cmdcancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblhelp.Caption = "Click This Button To Cancel The Update Process"
End Sub

Private Sub cmddel_Click()
If bookdata.RecordsetType = 2 - snapshot Then
    MsgBox "You Do Not Have The Right To Make Any Changes To The Database.", vbCritical, "No Right To Perform Action"
Else
If bookdata.Recordset.RecordCount = 0 Then
    MsgBox "There are no records to be deleted!", vbCritical, "No Records!!!"
Else
    If txtbid.Text = "" Then
        MsgBox "Please Select The Record You Want To Delete", vbInformation, "Please Select A Record To Delete"
        Exit Sub
    End If
    choice = MsgBox("Are you sure want to delete the record?", vbYesNo, "Delete Confirmation")
    If choice = vbYes Then
        bookdata.Recordset.Delete
        bookdata.Refresh
        MsgBox "The Record Has Been Deleted From The Data Successfully", vbInformation, "Delete Successful"
    End If
End If

cmbbid.Text = ""
ClearAll
End If
End Sub

Private Sub cmddel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblhelp.Caption = "Click This Button To Delete The Currently Selected Record"
End Sub

Private Sub cmdedit_Click()
If bookdata.RecordsetType = 2 - snapshot Then
    MsgBox "You Do Not Have The Right To Make Any Changes To The Database.", vbCritical, "No Right To Perform Action"
Else
If bookdata.Recordset.RecordCount = 0 Then
    MsgBox "There are no records to be edited!", vbCritical, "No Records!!!"
Else
    If txtbid.Text = "" Then
        MsgBox "Please Select The Record You Want To Edit", vbInformation, "Please Select A Record To Edit"
        Exit Sub
    End If
    bookdata.Recordset.Edit
    EnableAll
End If
End If

End Sub

Private Sub cmdedit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblhelp.Caption = "Click This Button To Edit The Currently Selected Record"
End Sub

Private Sub cmdfirst_Click()

On Error Resume Next

If bookdata.Recordset.RecordCount = 0 Then
    MsgBox "There are no available records", vbInformation, "No Available Records"
ElseIf bookdata.Recordset.BOF = True Then
    MsgBox "You are on the first record", vbInformation, "First Record"
Else
    bookdata.Recordset.MoveFirst
End If
    With bookdata.Recordset
    
     txtbid.Text = ![B_ID]
     txtname.Text = ![B_Name]
     txtaut.Text = ![author]
     txtpub.Text = ![Publisher]
     txtISBN.Text = ![isbn]
     txtCat.Text = ![Category]
     txtshelf.Text = ![Shelf]
     txtstat.Text = ![Status]
     
    End With

End Sub

Private Sub cmdfirst_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblhelp.Caption = "Click This Button To Jump To The First Record Of The Database"
End Sub

Private Sub cmdlast_Click()
On Error Resume Next

If bookdata.Recordset.RecordCount = 0 Then
    MsgBox "There are no available records", vbInformation, "No Available Records"
ElseIf bookdata.Recordset.EOF = True Then
    MsgBox "You are on the last record", vbInformation, "Last Record"
Else
    bookdata.Recordset.MoveLast
End If
    With bookdata.Recordset
    
     txtbid.Text = ![B_ID]
     txtname.Text = ![B_Name]
     txtaut.Text = ![author]
     txtpub.Text = ![Publisher]
     txtISBN.Text = ![isbn]
     txtCat.Text = ![Category]
     txtshelf.Text = ![Shelf]
     txtstat.Text = ![Status]
     
    End With
    
End Sub

Private Sub cmdlast_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblhelp.Caption = "Click This Button To Jump To The Last Record Of The Database"
End Sub

Private Sub cmdnext_Click()
On Error Resume Next

If bookdata.Recordset.RecordCount = 0 Then
    MsgBox "There are no available records", vbInformation, "No Available Records"
ElseIf bookdata.Recordset.EOF = True Then
    MsgBox "You are on the last record", vbInformation, "Last Record"
Else
    bookdata.Recordset.MoveNext
End If
    With bookdata.Recordset
    
     txtbid.Text = ![B_ID]
     txtname.Text = ![B_Name]
     txtaut.Text = ![author]
     txtpub.Text = ![Publisher]
     txtISBN.Text = ![isbn]
     txtCat.Text = ![Category]
     txtshelf.Text = ![Shelf]
     txtstat.Text = ![Status]
     
    End With
End Sub

Private Sub cmdnext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblhelp.Caption = "Click This Button To Move To The Next Record Of The Database"
End Sub

Private Sub cmdpre_Click()

On Error Resume Next

If bookdata.Recordset.RecordCount = 0 Then
    MsgBox "There are no available records", vbInformation, "No Available Records"
ElseIf bookdata.Recordset.BOF = True Then
    MsgBox "You are on the first record", vbInformation, "First Record"
Else
    bookdata.Recordset.MovePrevious
End If
    With bookdata.Recordset
    
     txtbid.Text = ![B_ID]
     txtname.Text = ![B_Name]
     txtaut.Text = ![author]
     txtpub.Text = ![Publisher]
     txtISBN.Text = ![isbn]
     txtCat.Text = ![Category]
     txtshelf.Text = ![Shelf]
     txtstat.Text = ![Status]
     
    End With
End Sub

Private Sub cmdref_Click()
Call ClearAll
bookdata.Refresh
cmbbid.Text = ""
End Sub

Private Sub cmdpre_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblhelp.Caption = "Click This Button To Move To The Previous Record Of The Database"
End Sub

Private Sub cmdsave_Click()

If txtname.Text = "" Then
    MsgBox "Please Enter A Book Name In The Text Box.", vbInformation, "Enter Book Name"
    txtname.SetFocus
    Exit Sub
End If

If txtISBN.Text = "" Then
    MsgBox "Please Enter A Book ISBN In The Text Box.", vbInformation, "Enter Book ISBN"
    txtISBN.SetFocus
    Exit Sub
End If

If Not IsNumeric(txtISBN.Text) Then
    MsgBox "Please Enter A Valid Book ISBN", vbInformation, "ISBN Not Acceptable"
    txtISBN.SetFocus
    Exit Sub
End If

If txtCat.Text = "" Then
    
If cmbCat.Enabled = True Then
    MsgBox "Please Select A Book Category From The Drop-Down Menu.", vbInformation, "Select Book Category"
    cmbCat.SetFocus
Else
    txtCat.SetFocus
    MsgBox "Please Type In The Book's In The Text Box.", vbInformation, "Type In Book Category"
    Exit Sub
End If
End If
 
If txtshelf.Text = "" Then
    MsgBox "Please Type In The Shelf's Code Where The Book Will Be Kept.", vbInformation, "Type In Shelf Code"
    txtshelf.SetFocus
    Exit Sub
End If

  With bookdata.Recordset
   If bookdata.Recordset.RecordCount = 0 Then
    ![B_ID] = txtbid.Text
    ![B_Name] = txtname.Text
    ![author] = txtaut.Text
    ![Publisher] = txtpub.Text
    ![isbn] = txtISBN.Text
    ![Category] = txtCat.Text
    ![Shelf] = txtshelf.Text
    ![Status] = txtstat.Text
    .Update
   Else

    ![B_ID] = txtbid.Text
    ![B_Name] = txtname.Text
    ![author] = txtaut.Text
    ![Publisher] = txtpub.Text
    ![isbn] = txtISBN.Text
    ![Category] = txtCat.Text
    ![Shelf] = txtshelf.Text
    ![Status] = txtstat.Text
   .Update
   
   End If
End With
MsgBox "Update Successful.", vbInformation, "Update Successful"
Call ClearAll
Call DisableAll
bookdata.Refresh
cmbbid.Text = ""

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command6_Click()

End Sub

Private Sub cmdsave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblhelp.Caption = "Click This Button To Complete The Update Process"
End Sub

Private Sub Form_Load()


With cmbCat
    .AddItem "Physics"
    .AddItem "Chemistry"
    .AddItem "Biology"
    .AddItem "English Language"
    .AddItem "English Literature"
    .AddItem "Other"
End With
lblclock.Caption = Format(Time, "HH:MM:SS")
lbldate.Caption = Format(Date, "long date")


End Sub

Private Sub lblclock_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblclock.ToolTipText = Format(Date, "long date")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblhelp.Caption = ""
End Sub

Private Sub mnuadd_Click()
If bookdata.RecordsetType = 2 - snapshot Then
    MsgBox "You Do Not Have The Right To Make Any Changes To The Database.", vbCritical, "No Right To Perform Action"
Else
EnableAll
cmbbid.Text = ""
ClearAll

If bookdata.Recordset.RecordCount = 0 Then
    bookdata.Recordset.AddNew
    txtbid.Text = "B00001"
Else
If bookdata.Recordset.RecordCount < 9 Then
    bookdata.Recordset.MoveLast
    txtbid.Text = bookdata.Recordset.Fields("B_ID").Value
    bid = Right(txtbid.Text, 1)
    bid = (bid + 1)
    txtbid.Text = "B0000" + CStr(bid)
    bookdata.Recordset.AddNew
ElseIf bookdata.Recordset.RecordCount >= 9 And 99 Then
    bookdata.Recordset.MoveLast
    txtbid.Text = bookdata.Recordset.Fields("B_ID").Value
    bid = Right(txtbid.Text, 2)
    bid = (bid + 1)
    txtbid.Text = "B000" + CStr(bid)
    bookdata.Recordset.AddNew
ElseIf bookdata.Recordset.RecordCount >= 99 And 999 Then
    bookdata.Recordset.MoveLast
    txtbid.Text = bookdata.Recordset.Fields("B_ID").Value
    bid = Right(txtbid.Text, 2)
    bid = (bid + 1)
    txtbid.Text = "B00" + CStr(bid)
    bookdata.Recordset.AddNew
ElseIf bookdata.Recordset.RecordCount >= 999 And 9999 Then
    bookdata.Recordset.MoveLast
    txtbid.Text = bookdata.Recordset.Fields("B_ID").Value
    bid = Right(txtbid.Text, 2)
    bid = (bid + 1)
    txtbid.Text = "B0" + CStr(bid)
    bookdata.Recordset.AddNew
ElseIf bookdata.Recordset.RecordCount >= 999 And 9999 Then
    bookdata.Recordset.MoveLast
    txtbid.Text = bookdata.Recordset.Fields("B_ID").Value
    bid = Right(txtbid.Text, 2)
    bid = (bid + 1)
    txtbid.Text = "B" + CStr(bid)
    bookdata.Recordset.AddNew
End If
    
End If
txtname.SetFocus
txtstat.Text = "Available"
End If

End Sub

Private Sub mnucancel_Click()
ClearAll
DisableAll
cmbbid.Text = ""
bookdata.Refresh
End Sub

Private Sub mnudel_Click()
If bookdata.RecordsetType = 2 - snapshot Then
    MsgBox "You Do Not Have The Right To Make Any Changes To The Database.", vbCritical, "No Right To Perform Action"
Else
If bookdata.Recordset.RecordCount = 0 Then
    MsgBox "There are no records to be deleted!", vbCritical, "No Records!!!"
Else
    If txtbid.Text = "" Then
        MsgBox "Please Select The Record You Want To Delete", vbInformation, "Please Select A Record To Delete"
        Exit Sub
    End If
    choice = MsgBox("Are you sure want to delete the record?", vbYesNo, "Delete Confirmation")
    If choice = vbYes Then
        bookdata.Recordset.Delete
        bookdata.Refresh
        MsgBox "The Record Has Been Deleted From The Data Successfully", vbInformation, "Delete Successful"
    End If
End If

cmbbid.Text = ""
ClearAll
End If
End Sub

Private Sub mnuedit_Click()
If bookdata.RecordsetType = 2 - snapshot Then
    MsgBox "You Do Not Have The Right To Make Any Changes To The Database.", vbCritical, "No Right To Perform Action"
Else
If bookdata.Recordset.RecordCount = 0 Then
    MsgBox "There are no records to be edited!", vbCritical, "No Records!!!"
Else
    If txtbid.Text = "" Then
        MsgBox "Please Select The Record You Want To Edit", vbInformation, "Please Select A Record To Edit"
        Exit Sub
    End If
    bookdata.Recordset.Edit
    EnableAll
End If
End If
End Sub

Private Sub mnusave_Click()
If txtname.Text = "" Then
    MsgBox "Please Enter A Book Name In The Text Box.", vbInformation, "Enter Book Name"
    txtname.SetFocus
    Exit Sub
End If

If txtISBN.Text = "" Then
    MsgBox "Please Enter A Book ISBN In The Text Box.", vbInformation, "Enter Book ISBN"
    txtISBN.SetFocus
    Exit Sub
End If

If Not IsNumeric(txtISBN.Text) Then
    MsgBox "Please Enter A Valid Book ISBN", vbInformation, "ISBN Not Acceptable"
    txtISBN.SetFocus
    Exit Sub
End If

If txtCat.Text = "" Then
   
If cmbCat.Enabled = True Then
    MsgBox "Please Select A Book Category From The Drop-Down Menu.", vbInformation, "Select Book Category"
    cmbCat.SetFocus
Else
    txtCat.SetFocus
    MsgBox "Please Type In The Book's In The Text Box.", vbInformation, "Type In Book Category"
    Exit Sub
End If
End If
If IsNumeric(txtaut.Text) Then
    MsgBox "Author Name Cannot Be A Number", vbInformation, "Enter Valid Author Name"
    txtaut.SetFocus
    Exit Sub
End If
 
If txtshelf.Text = "" Then
    MsgBox "Please Type In The Shelf's Code Where The Book Will Be Kept.", vbInformation, "Type In Shelf Code"
    txtshelf.SetFocus
    Exit Sub
End If

  With bookdata.Recordset
If bookdata.Recordset.RecordCount = 0 Then
    ![B_ID] = txtbid.Text
    ![B_Name] = txtname.Text
    ![author] = txtaut.Text
    ![Publisher] = txtpub.Text
    ![isbn] = txtISBN.Text
    ![Category] = txtCat.Text
    ![Shelf] = txtshelf.Text
    ![Status] = txtstat.Text
    .Update
Else

    ![B_ID] = txtbid.Text
    ![B_Name] = txtname.Text
    ![author] = txtaut.Text
    ![Publisher] = txtpub.Text
    ![isbn] = txtISBN.Text
    ![Category] = txtCat.Text
    ![Shelf] = txtshelf.Text
    ![Status] = txtstat.Text
   .Update
   
End If
End With
MsgBox "Update Successful.", vbInformation, "Update Successful"
Call ClearAll
Call DisableAll
bookdata.Refresh
cmbbid.Text = ""

End Sub

Private Sub Timer1_Timer()
lblclock.Caption = Format(Time, "HH:MM:SS")
lbldate.Caption = Format(Date, "long date")
End Sub
