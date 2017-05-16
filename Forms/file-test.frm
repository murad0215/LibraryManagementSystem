VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   9060
   StartUpPosition =   3  'Windows Default
   Begin MSDBCtls.DBList DBList1 
      Bindings        =   "file-test.frx":0000
      Height          =   1620
      Left            =   6960
      TabIndex        =   8
      Top             =   3240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   2858
      _Version        =   393216
      ListField       =   "B_Name"
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "file-test.frx":0014
      Height          =   1215
      Left            =   6600
      OleObjectBlob   =   "file-test.frx":0028
      TabIndex        =   7
      Top             =   5160
      Width           =   2295
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\Library Management System\DB\LMS97DB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tblbook"
      Top             =   6600
      Width           =   4335
   End
   Begin VB.TextBox txt2 
      Height          =   495
      Left            =   7080
      TabIndex        =   6
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   5400
      TabIndex        =   5
      Top             =   1560
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7680
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txt 
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      Height          =   3990
      Left            =   2880
      Pattern         =   "*.jpg;*.bmp"
      TabIndex        =   3
      Top             =   2280
      Width           =   3615
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   2640
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
   End
   Begin VB.DirListBox Dir1 
      Height          =   4590
      Left            =   600
      TabIndex        =   1
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Photo"
      Height          =   735
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Source_file As Variant


Private Sub Command1_Click()
FileCopy txt2.Text, "d:\" & txt.Text & ".jpg"
MsgBox "Copy Complete"
End Sub

Private Sub Command2_Click()
On Error GoTo Error
    CommonDialog1.InitDir = App.Path
    CommonDialog1.Filter = "GIF, BMP, JPG, JPEG|*.gif; *.jpg; *. bmp; *.jpeg"
    CommonDialog1.ShowOpen
    Me.Caption = ""

    'AnimatedGIF1.FileName =

    
    txt2.Text = CommonDialog1.FileName
Error:

End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
Source_file = Dir1.Path & "\" & File1.FileName
MsgBox Source_file
End Sub

Private Sub txt2_Change()
Data1.RecordSource = "select * From tblbook where ucase(mid(B_Name,1, " & Len(txt2.Text) & "))='" & txt2.Text & "'"
Data1.Refresh

'DBList1.DataSource = Data1
'DBList1.ListField = B_Name

'With DBList1
 



End Sub
