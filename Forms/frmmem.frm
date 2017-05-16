VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmmem 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Member Information Entry & Retrivation Form"
   ClientHeight    =   10875
   ClientLeft      =   2205
   ClientTop       =   -915
   ClientWidth     =   10800
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmmem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10875
   ScaleWidth      =   10800
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8520
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7440
      Top             =   5520
   End
   Begin VB.CommandButton cmdsrch 
      Caption         =   "SEA&RCH"
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
      Left            =   240
      Picture         =   "frmmem.frx":08CA
      TabIndex        =   38
      Top             =   4440
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdprint 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&PRINT"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Picture         =   "frmmem.frx":866C
      TabIndex        =   39
      Top             =   4440
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
      Left            =   7800
      Picture         =   "frmmem.frx":105E6
      TabIndex        =   34
      Top             =   3840
      Width           =   1455
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
      Left            =   6540
      Picture         =   "frmmem.frx":18560
      TabIndex        =   29
      Top             =   7560
      Width           =   1935
   End
   Begin VB.FileListBox filphotos 
      Height          =   480
      Left            =   240
      Pattern         =   "*.jpg;*.jpeg;*.bmp;*.gif;*.png"
      TabIndex        =   25
      Top             =   3840
      Visible         =   0   'False
      Width           =   2415
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
      Height          =   405
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   3480
      Width           =   375
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
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2040
      Width           =   2655
   End
   Begin VB.TextBox txtadd 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   4320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   4560
      Width           =   2655
   End
   Begin VB.TextBox txtcnt 
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
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2520
      Width           =   2295
   End
   Begin VB.TextBox txtocc 
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
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   3960
      Width           =   2295
   End
   Begin VB.TextBox txtmid 
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
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtgen 
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
      Left            =   4320
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   8
      Top             =   3000
      Width           =   390
   End
   Begin VB.TextBox txtren 
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
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   6000
      Width           =   1575
   End
   Begin VB.TextBox txtreg 
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
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   5520
      Width           =   1575
   End
   Begin VB.TextBox txteml 
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
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   6480
      Width           =   2295
   End
   Begin VB.OptionButton optf 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Female"
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
      Height          =   270
      Left            =   5760
      MaskColor       =   &H000000FF&
      TabIndex        =   4
      Top             =   3120
      Width           =   255
   End
   Begin VB.OptionButton optm 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Male"
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
      Height          =   270
      Left            =   4800
      MaskColor       =   &H000000FF&
      TabIndex        =   3
      Top             =   3120
      Width           =   210
   End
   Begin VB.TextBox txtphoto 
      Height          =   495
      Left            =   12360
      TabIndex        =   1
      Top             =   3240
      Width           =   375
   End
   Begin VB.Data memdata 
      Caption         =   "memdata"
      Connect         =   "Access"
      DatabaseName    =   "e:\Library Management System\DB\LMS97DB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT * FROM tblmember ORDER BY tblmember.M_ID"
      Top             =   6600
      Visible         =   0   'False
      Width           =   1980
   End
   Begin MSDBGrid.DBGrid mgrid 
      Bindings        =   "frmmem.frx":21F96
      Height          =   2175
      Left            =   60
      OleObjectBlob   =   "frmmem.frx":21FAC
      TabIndex        =   0
      Top             =   8280
      Width           =   10695
   End
   Begin MSDBCtls.DBCombo cmbmid 
      Bindings        =   "frmmem.frx":2385D
      Height          =   390
      Left            =   5640
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   688
      _Version        =   393216
      ListField       =   "M_ID"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Left            =   9240
      Picture         =   "frmmem.frx":23873
      TabIndex        =   37
      Top             =   3840
      Width           =   1455
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
      Left            =   7800
      Picture         =   "frmmem.frx":2B7ED
      TabIndex        =   36
      Top             =   3240
      Width           =   1455
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
      Left            =   9240
      Picture         =   "frmmem.frx":33767
      TabIndex        =   33
      Top             =   2640
      Width           =   1455
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
      Left            =   7800
      Picture         =   "frmmem.frx":3B6E1
      TabIndex        =   35
      Top             =   2640
      Width           =   1455
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
      Left            =   4620
      Picture         =   "frmmem.frx":4365B
      TabIndex        =   32
      Top             =   7560
      Width           =   1935
   End
   Begin VB.CommandButton cmdpre 
      BackColor       =   &H00C0FFFF&
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
      Left            =   2700
      Picture         =   "frmmem.frx":4D091
      TabIndex        =   30
      Top             =   7560
      Width           =   1935
   End
   Begin VB.CommandButton cmdfirst 
      BackColor       =   &H00FFFFFF&
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
      Left            =   780
      Picture         =   "frmmem.frx":56AC7
      TabIndex        =   31
      Top             =   7560
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
      Left            =   8040
      TabIndex        =   40
      Top             =   10440
      Width           =   2655
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
      Left            =   6480
      TabIndex        =   41
      Top             =   10440
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
      Left            =   0
      TabIndex        =   42
      Top             =   10440
      Width           =   6495
   End
   Begin VB.Label Label10 
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
      Left            =   780
      TabIndex        =   28
      Top             =   480
      Width           =   8175
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Female"
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
      Left            =   6120
      TabIndex        =   27
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Male"
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
      Left            =   5160
      TabIndex        =   26
      Top             =   3120
      Width           =   495
   End
   Begin VB.Image imgStudent 
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   240
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label5 
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2760
      TabIndex        =   24
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label lblcat 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Contact"
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
      Left            =   2760
      TabIndex        =   23
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lbladd 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Left            =   2760
      TabIndex        =   22
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2760
      TabIndex        =   21
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblbid 
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
      Height          =   375
      Left            =   2760
      TabIndex        =   20
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
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
      Left            =   2760
      TabIndex        =   19
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Occupation"
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
      Left            =   2760
      TabIndex        =   18
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Renew Date"
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
      Left            =   2760
      TabIndex        =   17
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Registration Date"
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
      Left            =   2760
      TabIndex        =   16
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail"
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
      Left            =   2760
      TabIndex        =   15
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Menu mnumenu 
      Caption         =   "&Menu"
      Begin VB.Menu mnuadd 
         Caption         =   "&ADD NEW"
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
      Begin VB.Menu mnuprint 
         Caption         =   "&PRINT MEMBERSHIP CARD     "
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu mnurenew 
      Caption         =   "Rene&w Membership"
   End
End
Attribute VB_Name = "frmmem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Double
Dim DIALOGUE_FILE As Variant

Private Sub EnableAll()

txtname.Enabled = True
filphotos.Visible = True
cmbmid.Enabled = False
mnurenew.Enabled = False
mnuadd.Enabled = False
mnuedit.Enabled = False
mnudel.Enabled = False
MNUPRINT.Enabled = False
mnusave.Enabled = True
mnucancel.Enabled = True
cmdsrch.Visible = True
filphotos.Visible = True
filphotos.Refresh
cmdadd.Enabled = False
cmddel.Enabled = False
cmdedit.Enabled = False
cmdprint.Enabled = False
cmdsave.Enabled = True
cmdcancel.Enabled = True
cmdfirst.Enabled = False
cmdnext.Enabled = False
cmdpre.Enabled = False
cmdlast.Enabled = False
txtname.Locked = False
txtocc.Locked = False
txtcnt.Locked = False
txteml.Locked = False
txtadd.Locked = False
optm.Enabled = True
optf.Enabled = True

End Sub

Private Sub DisableAll()

txtname.Enabled = False
cmbmid.Enabled = True
filphotos.Visible = False
mnurenew.Enabled = True
mnusave.Enabled = False
mnucancel.Enabled = False
mnuadd.Enabled = True
mnudel.Enabled = True
mnuedit.Enabled = True
MNUPRINT.Enabled = True
cmdsrch.Visible = False
filphotos.Visible = False
filphotos.Refresh
cmdadd.Enabled = True
cmddel.Enabled = True
cmdedit.Enabled = True
cmdprint.Enabled = True
cmdsave.Enabled = False
cmdcancel.Enabled = False
cmdfirst.Enabled = True
cmdnext.Enabled = True
cmdpre.Enabled = True
cmdlast.Enabled = True
txtname.Locked = True
txtocc.Locked = True
txtcnt.Locked = True
txteml.Locked = True
txtadd.Locked = True
optm.Enabled = False
optf.Enabled = False

End Sub

Private Sub ClearAll()

txtmid.Text = ""
txtname.Text = ""
txtgen.Text = ""
txtadd.Text = ""
txtcnt.Text = ""
txtbis.Text = ""
txtocc.Text = ""
txteml.Text = ""
txtreg.Text = ""
txtren.Text = ""
optm.Value = False
optf.Value = False

End Sub

Private Sub cmbmid_Change()
If memdata.Recordset.RecordCount = 0 Then
    MsgBox "There are no records!!", vbCritical, "No Records!!!"
Else

    memdata.Recordset.Bookmark = cmbmid.SelectedItem
    
    With memdata.Recordset
    
     txtmid.Text = ![M_ID]
     txtname.Text = ![M_Name]
     txtgen.Text = ![Gender]
     txtocc.Text = ![occ]
     txtcnt.Text = ![contact]
     txteml.Text = ![email]
     txtreg.Text = ![Reg_Date]
     txtren.Text = ![Renew_Date]
     txtadd.Text = ![Address]
     txtbis.Text = ![Issued]
        
    End With

End If
    dispic

End Sub

Private Sub cmbmid_Click(Area As Integer)

If memdata.Recordset.RecordCount = 0 Then
    MsgBox "There are no records!!", vbCritical, "No Records!!!"
Else

    memdata.Recordset.Bookmark = cmbmid.SelectedItem
    
    With memdata.Recordset
    
     txtmid.Text = ![M_ID]
     txtname.Text = ![M_Name]
     txtgen.Text = ![Gender]
     txtocc.Text = ![occ]
     txtcnt.Text = ![contact]
     txteml.Text = ![email]
     txtreg.Text = ![Reg_Date]
     txtren.Text = ![Renew_Date]
     txtadd.Text = ![Address]
     txtbis.Text = ![Issued]
        
    End With

End If
    dispic

End Sub

Private Sub cmdadd_Click()
If memdata.RecordsetType = 2 - snapshot Then
    MsgBox "You Do Not Have Any Right To Make Changes In The Database", vbCritical, "Not Enough Right To Perform Action"
Else
EnableAll
cmbmid.Text = ""
ClearAll

If memdata.Recordset.RecordCount = 0 Then
    memdata.Recordset.AddNew
    txtmid.Text = "M00001"
Else
If memdata.Recordset.RecordCount < 9 Then
    memdata.Recordset.MoveLast
    txtmid.Text = memdata.Recordset.Fields("M_ID").Value
    M_ID = Right(txtmid.Text, 1)
    M_ID = (M_ID + 1)
    txtmid.Text = "M0000" + CStr(M_ID)
    memdata.Recordset.AddNew
ElseIf memdata.Recordset.RecordCount >= 9 Then
    memdata.Recordset.MoveLast
    txtmid.Text = memdata.Recordset.Fields("M_ID").Value
    M_ID = Right(txtmid.Text, 2)
    M_ID = (M_ID + 1)
    txtmid.Text = "M000" + CStr(M_ID)
    memdata.Recordset.AddNew
ElseIf memdata.Recordset.RecordCount >= 99 Then
    memdata.Recordset.MoveLast
    txtmid.Text = memdata.Recordset.Fields("M_ID").Value
    M_ID = Right(txtmid.Text, 2)
    M_ID = (M_ID + 1)
    txtmid.Text = "M00" + CStr(M_ID)
    memdata.Recordset.AddNew
ElseIf memdata.Recordset.RecordCount >= 999 Then
    memdata.Recordset.MoveLast
    txtmid.Text = memdata.Recordset.Fields("M_ID").Value
    M_ID = Right(txtmid.Text, 2)
    M_ID = (M_ID + 1)
    txtmid.Text = "M0" + CStr(M_ID)
    memdata.Recordset.AddNew
ElseIf memdata.Recordset.RecordCount >= 9999 Then
    memdata.Recordset.MoveLast
    txtmid.Text = memdata.Recordset.Fields("M_ID").Value
    M_ID = Right(txtmid.Text, 2)
    M_ID = (M_ID + 1)
    txtmid.Text = "M" + CStr(M_ID)
    memdata.Recordset.AddNew
End If

End If
txtname.SetFocus
txtbis.Text = "0"
txtreg.Text = Date
txtren.Text = Date + 365
End If

End Sub

Private Sub cmdadd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblhelp.Caption = "Click This Button To Add A New Record"
End Sub

Private Sub cmdcancel_Click()
DisableAll
ClearAll
cmbmid.Text = ""
memdata.Refresh
filphotos.Refresh
memdata.Recordset.MoveFirst
    With memdata.Recordset
    
     txtmid.Text = ![M_ID]
     txtname.Text = ![M_Name]
     txtgen.Text = ![Gender]
     txtocc.Text = ![occ]
     txtcnt.Text = ![contact]
     txteml.Text = ![email]
     txtreg.Text = ![Reg_Date]
     txtren.Text = ![Renew_Date]
     txtadd.Text = ![Address]
     txtbis.Text = ![Issued]
     
    End With
dispic

End Sub

Private Sub cmdcancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblhelp.Caption = "Click This Button To Cancel The Update Process"
End Sub

Private Sub cmddel_Click()
If memdata.RecordsetType = 2 - snapshot Then
    MsgBox "You Do Not Have Any Right To Make Changes In The Database", vbCritical, "Not Enough Right To Perform Action"
Else
If memdata.Recordset.RecordCount = 0 Then
    MsgBox "There are no records to be deleted!", vbCritical, "No Records!!!"
Else
    If txtmid.Text = "" Then
        MsgBox "Please Select The Record You Want To Delete", vbInformation, "Please Select A Record To Delete"
        Exit Sub
    End If
      choice = MsgBox("Are you sure want to delete the record?", vbYesNo, "Delete Confirmation")
    If choice = vbYes Then
        memdata.Recordset.Delete
        memdata.Refresh
        MsgBox "The Record Has Been Deleted From The Data Successfully", vbInformation, "Delete Successful"
    End If
End If

cmbmid.Text = ""
ClearAll
End If

End Sub

Private Sub cmddel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblhelp.Caption = "Click This Button To Delete The Currently Selected Record"
End Sub

Private Sub cmdedit_Click()
If memdata.RecordsetType = 2 - snapshot Then
    MsgBox "You Do Not Have Any Right To Make Changes In The Database", vbCritical, "Not Enough Right To Perform Action"
Else
If memdata.Recordset.RecordCount = 0 Then
    MsgBox "There are no records to be edited!", vbCritical, "No Records!!!"
Else
    If txtmid.Text = "" Then
        MsgBox "Please Select The Record You Want To Edit", vbInformation, "Please Select A Record To Edit"
        Exit Sub
    End If
    memdata.Recordset.Edit
    EnableAll
End If
End If

End Sub

Private Sub cmdedit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblhelp.Caption = "Click This Button To Edit The Currently Selected Record"
End Sub

Private Sub cmdfirst_Click()

On Error Resume Next

If memdata.Recordset.RecordCount = 0 Then
    MsgBox "There are no available records", vbInformation, "No Available Records"
ElseIf memdata.Recordset.BOF = True Then
    MsgBox "You are on the first record", vbInformation, "First Record"
Else
    memdata.Recordset.MoveFirst
End If

    With memdata.Recordset
    
     txtmid.Text = ![M_ID]
     txtname.Text = ![M_Name]
     txtgen.Text = ![Gender]
     txtocc.Text = ![occ]
     txtcnt.Text = ![contact]
     txteml.Text = ![email]
     txtreg.Text = ![Reg_Date]
     txtren.Text = ![Renew_Date]
     txtadd.Text = ![Address]
     txtbis.Text = ![Issued]
     
    End With
      dispic

End Sub

Private Sub cmdfirst_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblhelp.Caption = "Click This Button To Jump To The First Record Of The Database"
End Sub

Private Sub cmdlast_Click()

On Error Resume Next

If memdata.Recordset.RecordCount = 0 Then
    MsgBox "There are no available records", vbInformation, "No Available Records"
ElseIf memdata.Recordset.EOF = True Then
    MsgBox "You are on the last record", vbInformation, "First Record"
Else
    memdata.Recordset.MoveLast
End If

    With memdata.Recordset
    
     txtmid.Text = ![M_ID]
     txtname.Text = ![M_Name]
     txtgen.Text = ![Gender]
     txtocc.Text = ![occ]
     txtcnt.Text = ![contact]
     txteml.Text = ![email]
     txtreg.Text = ![Reg_Date]
     txtren.Text = ![Renew_Date]
     txtadd.Text = ![Address]
     txtbis.Text = ![Issued]
     
     End With
     dispic

End Sub

Private Sub cmdlast_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblhelp.Caption = "Click This Button To Jump To The Last Record Of The Database"
End Sub

Private Sub cmdnext_Click()

On Error Resume Next

If memdata.Recordset.RecordCount = 0 Then
    MsgBox "There are no available records", vbInformation, "No Available Records"
ElseIf memdata.Recordset.EOF = True Then
    MsgBox "You are on the last record", vbInformation, "First Record"
Else
    memdata.Recordset.MoveNext
End If

    With memdata.Recordset
    
     txtmid.Text = ![M_ID]
     txtname.Text = ![M_Name]
     txtgen.Text = ![Gender]
     txtocc.Text = ![occ]
     txtcnt.Text = ![contact]
     txteml.Text = ![email]
     txtreg.Text = ![Reg_Date]
     txtren.Text = ![Renew_Date]
     txtadd.Text = ![Address]
     txtbis.Text = ![Issued]
     
     End With
       dispic
    
End Sub

Private Sub cmdpnt_Click()

If memdata.RecordsetType = 2 - snapshot Then
    MsgBox "You Do Not Have Any Right To Perform The Action", vbCritical, "Not Enough Right To Perform Action"
Else
    memberid = txtmid.Text
    If denvLMS.rscomemnew.State = adStateOpen Then
        denvLMS.rscomemnew.Close
    End If
    denvLMS.comemnew Trim(memberid)
    rptmemnew.Show
End If
End Sub

Private Sub cmdnext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblhelp.Caption = "Click This Button To Move To The Next Record Of The Database"
End Sub

Private Sub cmdpre_Click()

On Error Resume Next
If memdata.Recordset.RecordCount = 0 Then
    MsgBox "There are no available records", vbInformation, "No Available Records"
ElseIf memdata.Recordset.BOF = True Then
    MsgBox "You are on the first record", vbInformation, "First Record"
Else
    memdata.Recordset.MovePrevious
End If
    With memdata.Recordset
    
     txtmid.Text = ![M_ID]
     txtname.Text = ![M_Name]
     txtgen.Text = ![Gender]
     txtocc.Text = ![occ]
     txtcnt.Text = ![contact]
     txteml.Text = ![email]
     txtreg.Text = ![Reg_Date]
     txtren.Text = ![Renew_Date]
     txtadd.Text = ![Address]
     txtbis.Text = ![Issued]
    End With
    dispic

End Sub

Private Sub cmdref_Click()

Call ClearAll
memdata.Refresh
cmbmid.Text = ""

End Sub

Private Sub cmdpre_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblhelp.Caption = "Click This Button To Move To The Previous Record Of The Database"
End Sub

Private Sub cmdprint_Click()
If memdata.RecordsetType = 2 - snapshot Then
    MsgBox "You Do Not Have Any Right To Perform The Action", vbCritical, "Not Enough Right To Perform Action"
       
Else
If txtmid.Text = "" Then
    MsgBox "Please Select A Member First Before Printing His/Her Membership Card", vbInformation + vbOKOnly, "Select Member ID"
    Exit Sub
End If
    memberid = txtmid.Text
    If denvLMS.rscomemnew.State = adStateOpen Then
        denvLMS.rscomemnew.Close
    End If
    denvLMS.comemnew Trim(memberid)
    rptmemnew.Show
End If
End Sub

Private Sub cmdprint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblhelp.Caption = "Click This Button To Print The Membership Card Of The Currently Selected Member"
End Sub

Private Sub cmdsave_Click()

If txtname.Text = "" Then
    MsgBox "Please Enter A Member's Name In The Text Box.", vbInformation, "Enter Member's Name"
    txtname.SetFocus
    Exit Sub
End If
If IsNumeric(txtname.Text) Then
    MsgBox "Please Enter A Valid Member's Name In The Text Box.", vbInformation, "Enter Member's Name"
    txtname.SetFocus
    Exit Sub
End If
If txtgen.Text = "" Then
    MsgBox "Please Select The Member's Gender From The Options Available.", vbInformation, "Select Member's Gender"
    optm.SetFocus
    Exit Sub
End If
If txtadd.Text = "" Then
    MsgBox "Please Enter The Address Of The Member In The Text Box.", vbInformation, "Enter Address"
    txtadd.SetFocus
    Exit Sub
End If
If txtcnt.Text = "" Then
    MsgBox "Please Enter A Contact Number", vbInformation, "Contact Number Required"
    txtcnt.SetFocus
    Exit Sub
End If

If Not IsNumeric(txtcnt.Text) Then
    MsgBox "Please Enter A Valid Contact Number", vbInformation, "Contact Number Not Acceptable"
    txtcnt.SetFocus
    Exit Sub
End If
  
With memdata.Recordset
    If memdata.Recordset.RecordCount = 0 Then
    
    ![M_ID] = txtmid.Text
    ![M_Name] = txtname.Text
    ![Gender] = txtgen.Text
    ![Address] = txtadd.Text
    ![contact] = txtcnt.Text
    ![Issued] = txtbis.Text
    ![occ] = txtocc.Text
    ![email] = txteml.Text
    ![Reg_Date] = txtreg.Text
    ![Renew_Date] = txtren.Text
    ![Pic-file] = txtphoto.Text
    
    .Update
    
    Else
    
    ![M_ID] = txtmid.Text
    ![M_Name] = txtname.Text
    ![Gender] = txtgen.Text
    ![Address] = txtadd.Text
    ![contact] = txtcnt.Text
    ![Issued] = txtbis.Text
    ![occ] = txtocc.Text
    ![email] = txteml.Text
    ![Reg_Date] = txtreg.Text
    ![Renew_Date] = txtren.Text
    ![Pic-file] = txtphoto.Text
    
    .Update
   
    End If

End With

MsgBox "Update Successful.", vbInformation, "Update Successful"
Call ClearAll
Call DisableAll
memdata.Refresh
cmbmid.Text = ""
cmbmid.Enabled = True


End Sub

Private Sub cmdsave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblhelp.Caption = "Click This Button To Complete The Update Process"
End Sub

Private Sub cmdsrch_Click()

On Error Resume Next

    CommonDialog1.InitDir = App.Path
    CommonDialog1.Filter = "GIF, BMP, JPG, JPEG|*.gif; *.jpg; *. bmp; *.jpeg"
    CommonDialog1.ShowOpen
    Me.Caption = ""

    DIALOGUE_FILE = CommonDialog1.FileName

FileCopy DIALOGUE_FILE, "D:\Library Management System\MEMBER PHOTOS\" & txtmid.Text & ".jpg"
filphotos.Refresh
'MsgBox "Copy Complete"

End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdsrch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblhelp.Caption = "Click This Button To Search For The Member's Photo"
End Sub

Private Sub filphotos_Click()
Photo = filphotos.FileName
txtphoto.Text = filphotos.FileName
imgStudent.Picture = LoadPicture(filphotos.Path + "\" + Photo)
End Sub

Private Sub filphotos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblhelp.Caption = "This Is The List Of Photos To Select One From"
End Sub

Private Sub Form_Load()
filphotos.Path = App.Path + "\MEMBER PHOTOS\"
lblclock.Caption = Format(Time, "HH:MM:SS")
lbldate.Caption = Format(Date, "long date")
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblhelp.Caption = ""
End Sub


Private Sub mgrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblhelp.Caption = ""
End Sub

Private Sub mnuadd_Click()
If memdata.RecordsetType = 2 - snapshot Then
    MsgBox "You Do Not Have Any Right To Make Changes In The Database", vbCritical, "Not Enough Right To Perform Action"
Else
EnableAll
cmbmid.Text = ""
ClearAll

If memdata.Recordset.RecordCount = 0 Then
    memdata.Recordset.AddNew
    txtmid.Text = "M00001"
Else
If memdata.Recordset.RecordCount < 9 Then
    memdata.Recordset.MoveLast
    txtmid.Text = memdata.Recordset.Fields("M_ID").Value
    M_ID = Right(txtmid.Text, 1)
    M_ID = (M_ID + 1)
    txtmid.Text = "M0000" + CStr(M_ID)
    memdata.Recordset.AddNew
ElseIf memdata.Recordset.RecordCount >= 9 Then
    memdata.Recordset.MoveLast
    txtmid.Text = memdata.Recordset.Fields("M_ID").Value
    M_ID = Right(txtmid.Text, 2)
    M_ID = (M_ID + 1)
    txtmid.Text = "M000" + CStr(M_ID)
    memdata.Recordset.AddNew
ElseIf memdata.Recordset.RecordCount >= 99 Then
    memdata.Recordset.MoveLast
    txtmid.Text = memdata.Recordset.Fields("M_ID").Value
    M_ID = Right(txtmid.Text, 2)
    M_ID = (M_ID + 1)
    txtmid.Text = "M00" + CStr(M_ID)
    memdata.Recordset.AddNew
ElseIf memdata.Recordset.RecordCount >= 999 Then
    memdata.Recordset.MoveLast
    txtmid.Text = memdata.Recordset.Fields("M_ID").Value
    M_ID = Right(txtmid.Text, 2)
    M_ID = (M_ID + 1)
    txtmid.Text = "M0" + CStr(M_ID)
    memdata.Recordset.AddNew
ElseIf memdata.Recordset.RecordCount >= 9999 Then
    memdata.Recordset.MoveLast
    txtmid.Text = memdata.Recordset.Fields("M_ID").Value
    M_ID = Right(txtmid.Text, 2)
    M_ID = (M_ID + 1)
    txtmid.Text = "M" + CStr(M_ID)
    memdata.Recordset.AddNew
End If

End If
txtname.SetFocus
txtbis.Text = "0"
txtreg.Text = Date
txtren.Text = Date + 365
End If
End Sub

Private Sub mnucancel_Click()
DisableAll
ClearAll
cmbmid.Text = ""
memdata.Refresh
filphotos.Refresh
memdata.Recordset.MoveFirst
    With memdata.Recordset
    
     txtmid.Text = ![M_ID]
     txtname.Text = ![M_Name]
     txtgen.Text = ![Gender]
     txtocc.Text = ![occ]
     txtcnt.Text = ![contact]
     txteml.Text = ![email]
     txtreg.Text = ![Reg_Date]
     txtren.Text = ![Renew_Date]
     txtadd.Text = ![Address]
     txtbis.Text = ![Issued]
     
    End With
dispic
End Sub

Private Sub mnudel_Click()
If memdata.RecordsetType = 2 - snapshot Then
    MsgBox "You Do Not Have Any Right To Make Changes In The Database", vbCritical, "Not Enough Right To Perform Action"
Else
If memdata.Recordset.RecordCount = 0 Then
    MsgBox "There are no records to be deleted!", vbCritical, "No Records!!!"
Else
    If txtmid.Text = "" Then
        MsgBox "Please Select The Record You Want To Delete", vbInformation, "Please Select A Record To Delete"
        Exit Sub
    End If
      choice = MsgBox("Are you sure want to delete the record?", vbYesNo, "Delete Confirmation")
    If choice = vbYes Then
        memdata.Recordset.Delete
        memdata.Refresh
        MsgBox "The Record Has Been Deleted From The Data Successfully", vbInformation, "Delete Successful"
    End If
End If

cmbmid.Text = ""
ClearAll
End If

End Sub

Private Sub mnuedit_Click()
If memdata.RecordsetType = 2 - snapshot Then
    MsgBox "You Do Not Have Any Right To Make Changes In The Database", vbCritical, "Not Enough Right To Perform Action"
Else
If memdata.Recordset.RecordCount = 0 Then
    MsgBox "There are no records to be edited!", vbCritical, "No Records!!!"
Else
    If txtmid.Text = "" Then
        MsgBox "Please Select The Record You Want To Edit", vbInformation, "Please Select A Record To Edit"
        Exit Sub
    End If
    memdata.Recordset.Edit
    EnableAll
End If
End If

End Sub

Private Sub MNUPRINT_Click()
If memdata.RecordsetType = 2 - snapshot Then
    MsgBox "You Do Not Have Any Right To Perform The Action", vbCritical, "Not Enough Right To Perform Action"
Else
    memberid = txtmid.Text
    If denvLMS.rscomemnew.State = adStateOpen Then
        denvLMS.rscomemnew.Close
    End If
    denvLMS.comemnew Trim(memberid)
    'rptmemnew.mempic.Picture = "d:\azam.jpg"
    rptmemnew.Show
End If
End Sub

Private Sub mnurenew_Click()
If memdata.RecordsetType = 2 - snapshot Then
    MsgBox "You Do Not Have Any Right To Make Changes In The Database", vbCritical, "Not Enough Right To Perform Action"
Else
memberid = InputBox("Please Enter The Member's ID In The Following TextBox", "Enter Member's ID")
memdata.RecordSource = "select * from tblmember where M_ID='" & memberid & "'"
memdata.Refresh
If Not memdata.Recordset.RecordCount = 0 Then
    If memberid = memdata.Recordset.Fields("M_ID").Value Then
        txtmid.Text = memberid
        regdate = memdata.Recordset.Fields("Reg_Date").Value
        rendate = memdata.Recordset.Fields("Renew_Date").Value
        regdate = Date
        rendate = Date + 365
        With memdata.Recordset
            .Edit
            ![M_ID] = txtmid.Text
            ![Reg_Date] = regdate
            ![Renew_Date] = rendate
            .Update
        End With
        MsgBox "Congratulations! The Membership Renew Process Is Now Complete", vbInformation, "Renew Process Successful"
    End If
    Else
        MsgBox "No Member Found With ID: " & memberid, vbExclamation, "Error"
        memdata.RecordSource = "Select * from tblmember"
        memdata.Refresh
End If
    memdata.RecordSource = "Select * from tblmember"
    memdata.Refresh
    ClearAll
    mgrid.Refresh
End If
End Sub

Private Sub mnusave_Click()

If txtname.Text = "" Then
    MsgBox "Please Enter A Member's Name In The Text Box.", vbInformation, "Enter Member's Name"
    txtname.SetFocus
    Exit Sub
End If
If IsNumeric(txtname.Text) Then
    MsgBox "Please Enter A Valid Member's Name In The Text Box.", vbInformation, "Enter Member's Name"
    txtname.SetFocus
    Exit Sub
End If
If txtgen.Text = "" Then
    MsgBox "Please Select The Member's Gender From The Options Available.", vbInformation, "Select Member's Gender"
    optm.SetFocus
    Exit Sub
End If
If txtadd.Text = "" Then
    MsgBox "Please Enter The Address Of The Member In The Text Box.", vbInformation, "Enter Address"
    txtadd.SetFocus
    Exit Sub
End If
If txtcnt.Text = "" Then
    MsgBox "Please Enter A Contact Number", vbInformation, "Contact Number Required"
    txtcnt.SetFocus
    Exit Sub
End If

If Not IsNumeric(txtcnt.Text) Then
    MsgBox "Please Enter A Valid Contact Number", vbInformation, "Contact Number Not Acceptable"
    txtcnt.SetFocus
    Exit Sub
End If
  
With memdata.Recordset
    If memdata.Recordset.RecordCount = 0 Then
    
    ![M_ID] = txtmid.Text
    ![M_Name] = txtname.Text
    ![Gender] = txtgen.Text
    ![Address] = txtadd.Text
    ![contact] = txtcnt.Text
    ![Issued] = txtbis.Text
    ![occ] = txtocc.Text
    ![email] = txteml.Text
    ![Reg_Date] = txtreg.Text
    ![Renew_Date] = txtren.Text
    ![Pic-file] = txtphoto.Text
    
    .Update
    
    Else
    
    ![M_ID] = txtmid.Text
    ![M_Name] = txtname.Text
    ![Gender] = txtgen.Text
    ![Address] = txtadd.Text
    ![contact] = txtcnt.Text
    ![Issued] = txtbis.Text
    ![occ] = txtocc.Text
    ![email] = txteml.Text
    ![Reg_Date] = txtreg.Text
    ![Renew_Date] = txtren.Text
    ![Pic-file] = txtphoto.Text
    
    .Update
   
    End If

End With

MsgBox "Update Successful.", vbInformation, "Update Successful"
Call ClearAll
Call DisableAll
memdata.Refresh
cmbmid.Text = ""
cmbmid.Enabled = True

End Sub

Private Sub optf_Click()

If optm.Value = True Then
    optm.Value = False
End If
optf.Value = True
txtgen.Text = "F"

End Sub

Private Sub optm_Click()

If optf.Value = True Then
    optf.Value = False
End If
optm.Value = True
txtgen.Text = "M"

End Sub

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

Private Sub RepCdg1_ObjectSelected(SelObj As String)

End Sub

Private Sub Timer1_Timer()
lblclock.Caption = Format(Time, "HH:MM:SS")
lbldate.Caption = Format(Date, "long date")
End Sub
