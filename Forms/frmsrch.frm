VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmsrch 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SEARCH AND REPORT"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9705
   FillColor       =   &H80000000&
   ForeColor       =   &H8000000A&
   Icon            =   "frmsrch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   9705
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optretdate 
      BackColor       =   &H00E0E0E0&
      Caption         =   "n"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   4440
      TabIndex        =   37
      Top             =   3600
      Width           =   255
   End
   Begin VB.OptionButton optidate 
      BackColor       =   &H00E0E0E0&
      Caption         =   "n"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   4440
      TabIndex        =   35
      Top             =   2640
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5280
      Top             =   1080
   End
   Begin VB.CommandButton cmdrpt 
      Height          =   615
      Left            =   5280
      Picture         =   "frmsrch.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CommandButton cmdsrch 
      Height          =   615
      Left            =   3240
      Picture         =   "frmsrch.frx":866C
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5040
      Width           =   1935
   End
   Begin VB.OptionButton optstat 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Status"
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   360
      TabIndex        =   6
      Top             =   4320
      Width           =   255
   End
   Begin MSDataGridLib.DataGrid srchgrid 
      Height          =   1335
      Left            =   120
      TabIndex        =   14
      Top             =   5880
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   2355
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtsrch 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   5040
      Width           =   3015
   End
   Begin VB.OptionButton optbid 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ID"
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   360
      TabIndex        =   2
      Top             =   2880
      Width           =   255
   End
   Begin VB.OptionButton optaut 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Author"
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   360
      TabIndex        =   3
      Top             =   3240
      Width           =   255
   End
   Begin VB.OptionButton optisbn 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ISBN"
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   360
      TabIndex        =   4
      Top             =   3600
      Width           =   255
   End
   Begin VB.OptionButton optcat 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Category"
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   360
      TabIndex        =   5
      Top             =   3960
      Width           =   255
   End
   Begin VB.OptionButton optbname 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Name"
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   360
      TabIndex        =   1
      Top             =   2520
      Width           =   255
   End
   Begin VB.OptionButton optmid 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   2160
      TabIndex        =   8
      Top             =   3000
      Width           =   255
   End
   Begin VB.OptionButton optgen 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   2160
      TabIndex        =   9
      Top             =   3360
      Width           =   255
   End
   Begin VB.OptionButton optocc 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Occupation"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   2160
      TabIndex        =   10
      Top             =   3720
      Width           =   255
   End
   Begin VB.OptionButton optreg 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Registration Date"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   2160
      TabIndex        =   11
      Top             =   4080
      Width           =   255
   End
   Begin VB.OptionButton optmname 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   2160
      TabIndex        =   7
      Top             =   2640
      Width           =   255
   End
   Begin VB.OptionButton optren 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Renew Date"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   2160
      TabIndex        =   12
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   3480
      TabIndex        =   39
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   4800
      TabIndex        =   38
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   4800
      TabIndex        =   36
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Issue"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   3480
      TabIndex        =   34
      Top             =   2040
      Width           =   2295
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
      TabIndex        =   33
      Top             =   7320
      Width           =   6615
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
      Left            =   6645
      TabIndex        =   32
      Top             =   7320
      Width           =   1455
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
      Left            =   8085
      TabIndex        =   31
      Top             =   7320
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Book"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   -120
      TabIndex        =   28
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Renew date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   2520
      TabIndex        =   27
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Registration Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   2520
      TabIndex        =   26
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Occupation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   2520
      TabIndex        =   25
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   2520
      TabIndex        =   24
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   2520
      TabIndex        =   23
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   2520
      TabIndex        =   22
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   765
      TabIndex        =   21
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   765
      TabIndex        =   20
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ISBN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   765
      TabIndex        =   19
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Author"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   765
      TabIndex        =   18
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   765
      TabIndex        =   17
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   765
      TabIndex        =   16
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LIBRARY MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   480
      TabIndex        =   15
      Top             =   360
      Width           =   8415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Member"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   1560
      TabIndex        =   13
      Top             =   2040
      Width           =   2295
   End
End
Attribute VB_Name = "frmsrch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdback_Click()

frmmain.Show
Unload Me

End Sub

Private Sub cmdrpt_Click()

If optidate.Value = True Then
    idate = txtsrch.Text
    If denvLMS.rscomday_Grouping.State = adStateOpen Then
        denvLMS.rscomday_Grouping.Close
    End If
    denvLMS.comday_Grouping Trim(idate)
    rptday.Show
End If

If optretdate.Value = True Then
    retdate = txtsrch.Text
    If denvLMS.rscomret_Grouping.State = adStateOpen Then
        denvLMS.rscomret_Grouping.Close
    End If
    denvLMS.comret_Grouping Trim(retdate)
    rptret.Show
End If

If optstat.Value = True Then
    bkstat = txtsrch.Text
    If denvLMS.rscombkstat_Grouping.State = adStateOpen Then
        denvLMS.rscombkstat_Grouping.Close
    End If
    denvLMS.combkstat_Grouping Trim(bkstat)
    rptbkstat.Show
End If

If optbid.Value = True Then
    bid = txtsrch.Text
    If denvLMS.rscombid.State = adStateOpen Then
        denvLMS.rscombid.Close
    End If
    denvLMS.combid Trim(bid)
    rptbid.Show
End If

If optmid.Value = True Then
    memid = txtsrch.Text
    If denvLMS.rscommid.State = adStateOpen Then
        denvLMS.rscommid.Close
    End If
    denvLMS.commid Trim(memid)
    rptmid.Show
End If

If optcat.Value = True Then
    cat = txtsrch.Text
    If denvLMS.rscomcat_Grouping.State = adStateOpen Then
        denvLMS.rscomcat_Grouping.Close
    End If
    denvLMS.comcat_Grouping Trim(cat)
    rptcat.Show
End If

If optgen.Value = True Then
    gen = txtsrch.Text
    If denvLMS.rscomgen_Grouping.State = adStateOpen Then
        denvLMS.rscomgen_Grouping.Close
    End If
    denvLMS.comgen_Grouping Trim(gen)
    rptgen.Show
End If

If optbname.Value = True Then
    bname = txtsrch.Text
    If denvLMS.rscombname_Grouping.State = adStateOpen Then
        denvLMS.rscombname_Grouping.Close
    End If
    denvLMS.combname_Grouping Trim(bname)
    rptbname.Show
End If

If optaut.Value = True Then
    author = txtsrch.Text
    If denvLMS.rscomaut_Grouping.State = adStateOpen Then
        denvLMS.rscomaut_Grouping.Close
    End If
    denvLMS.comaut_Grouping Trim(author)
    rptaut.Show
End If

If optisbn.Value = True Then
    isbn = txtsrch.Text
    If denvLMS.rscomISBN_Grouping.State = adStateOpen Then
        denvLMS.rscomISBN_Grouping.Close
    End If
    denvLMS.comISBN_Grouping Trim(isbn)
    rptISBN.Show
End If

If optmname.Value = True Then
    mname = txtsrch.Text
    If denvLMS.rscomname.State = adStateOpen Then
        denvLMS.rscomname.Close
    End If
    denvLMS.comname Trim(mname)
    rptmname.Show
End If

If optocc.Value = True Then
    occ = txtsrch.Text
    If denvLMS.rscomocc_Grouping.State = adStateOpen Then
        denvLMS.rscomocc_Grouping.Close
    End If
    denvLMS.comocc_Grouping Trim(occ)
    rptocc.Show
End If

If optreg.Value = True Then
    regd = txtsrch.Text
    If denvLMS.rscomreg_Grouping.State = adStateOpen Then
        denvLMS.rscomreg_Grouping.Close
    End If
    denvLMS.comreg_Grouping Trim(regd)
    rptreg.Show
End If

If optren.Value = True Then
    rend = txtsrch.Text
    If denvLMS.rscomren_Grouping.State = adStateOpen Then
        denvLMS.rscomren_Grouping.Close
    End If
    denvLMS.comren_Grouping Trim(rend)
    rptren.Show
End If
    
End Sub

Private Sub cmdrpt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblhelp.Caption = "Click This Button To Generate Report For The Record(s) Searched"
End Sub

Private Sub cmdsrch_Click()

If optstat.Value = True Then
    bkstat = txtsrch.Text
    If denvLMS.rscombkstat_Grouping.State = adStateOpen Then
        denvLMS.rscombkstat_Grouping.Close
    End If
    denvLMS.combkstat_Grouping Trim(bkstat)
    srchgrid.DataMember = "combkstat"
    Set srchgrid.DataSource = denvLMS
End If

If optidate.Value = True Then
    idate = txtsrch.Text
    If denvLMS.rscomday_Grouping.State = adStateOpen Then
        denvLMS.rscomday_Grouping.Close
    End If
    denvLMS.comday_Grouping Trim(idate)
    srchgrid.DataMember = "comday"
    Set srchgrid.DataSource = denvLMS
End If

If optretdate.Value = True Then
    retdate = txtsrch.Text
    If denvLMS.rscomret_Grouping.State = adStateOpen Then
        denvLMS.rscomret_Grouping.Close
    End If
    denvLMS.comret_Grouping Trim(retdate)
    srchgrid.DataMember = "comret"
    Set srchgrid.DataSource = denvLMS
End If

If optbid.Value = True Then
    bid = txtsrch.Text
    If denvLMS.rscombid.State = adStateOpen Then
        denvLMS.rscombid.Close
    End If
    denvLMS.combid Trim(bid)
    srchgrid.DataMember = "combid"
    Set srchgrid.DataSource = denvLMS
End If

If optmid.Value = True Then
    memid = txtsrch.Text
    If denvLMS.rscommid.State = adStateOpen Then
        denvLMS.rscommid.Close
    End If
    denvLMS.commid Trim(memid)
    srchgrid.DataMember = "commid"
    Set srchgrid.DataSource = denvLMS
End If

If optcat.Value = True Then
    cat = txtsrch.Text
    If denvLMS.rscomcat_Grouping.State = adStateOpen Then
        denvLMS.rscomcat_Grouping.Close
    End If
    denvLMS.comcat_Grouping Trim(cat)
    srchgrid.DataMember = "comcat"
    Set srchgrid.DataSource = denvLMS
End If

If optgen.Value = True Then
    gen = txtsrch.Text
    If denvLMS.rscomgen_Grouping.State = adStateOpen Then
        denvLMS.rscomgen_Grouping.Close
    End If
    denvLMS.comgen_Grouping Trim(gen)
    srchgrid.DataMember = "comgen"
    Set srchgrid.DataSource = denvLMS
End If

If optbname.Value = True Then
    bname = txtsrch.Text
    If denvLMS.rscombname_Grouping.State = adStateOpen Then
        denvLMS.rscombname_Grouping.Close
    End If
    denvLMS.combname_Grouping Trim(bname)
    srchgrid.DataMember = "combname"
    Set srchgrid.DataSource = denvLMS
End If

If optaut.Value = True Then
    author = txtsrch.Text
    If denvLMS.rscomaut_Grouping.State = adStateOpen Then
        denvLMS.rscomaut_Grouping.Close
    End If
    denvLMS.comaut_Grouping CStr(author)
    srchgrid.DataMember = "comaut"
    Set srchgrid.DataSource = denvLMS
End If

If optisbn.Value = True Then
    isbn = txtsrch.Text
    If denvLMS.rscomISBN_Grouping.State = adStateOpen Then
        denvLMS.rscomISBN_Grouping.Close
    End If
    denvLMS.comISBN_Grouping Trim(isbn)
    srchgrid.DataMember = "comisbn"
    Set srchgrid.DataSource = denvLMS
End If

If optmname.Value = True Then
    mname = txtsrch.Text
    If denvLMS.rscomname.State = adStateOpen Then
        denvLMS.rscomname.Close
    End If
    denvLMS.comname Trim(mname)
    srchgrid.DataMember = "comname"
    Set srchgrid.DataSource = denvLMS
End If

If optocc.Value = True Then
    occ = txtsrch.Text
    If denvLMS.rscomocc_Grouping.State = adStateOpen Then
        denvLMS.rscomocc_Grouping.Close
    End If
    denvLMS.comocc_Grouping Trim(occ)
    srchgrid.DataMember = "comocc"
    Set srchgrid.DataSource = denvLMS
End If

If optreg.Value = True Then
    regd = txtsrch.Text
    If denvLMS.rscomreg_Grouping.State = adStateOpen Then
        denvLMS.rscomreg_Grouping.Close
    End If
    denvLMS.comreg_Grouping Trim(regd)
    srchgrid.DataMember = "comreg"
    Set srchgrid.DataSource = denvLMS
End If

If optren.Value = True Then
    rend = txtsrch.Text
    If denvLMS.rscomren_Grouping.State = adStateOpen Then
        denvLMS.rscomren_Grouping.Close
    End If
    denvLMS.comren_Grouping Trim(rend)
    srchgrid.DataMember = "comren"
    Set srchgrid.DataSource = denvLMS
End If
    
End Sub

Private Sub Image2_Click()

End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdsrch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblhelp.Caption = "Click This Button To Search For The Record"
End Sub

Private Sub Form_Load()
lblclock.Caption = Format(Time, "HH:MM:SS")
lbldate.Caption = Format(Date, "long date")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblhelp.Caption = ""
End Sub

Private Sub optaut_Click()
txtsrch.MaxLength = 0
txtsrch.Text = ""
End Sub

Private Sub optbid_Click()
txtsrch.MaxLength = 6
txtsrch.Text = ""
End Sub

Private Sub optbname_Click()
txtsrch.MaxLength = 0
txtsrch.Text = ""
End Sub

Private Sub optcat_Click()
txtsrch.MaxLength = 0
txtsrch.Text = ""
End Sub

Private Sub optgen_Click()
txtsrch.MaxLength = 1
txtsrch.Text = ""
End Sub

Private Sub optisbn_Click()
txtsrch.MaxLength = 0
txtsrch.Text = ""
End Sub

Private Sub optmid_Click()
txtsrch.MaxLength = 6
txtsrch.Text = ""
End Sub

Private Sub optmname_Click()
txtsrch.MaxLength = 0
txtsrch.Text = ""
End Sub

Private Sub optocc_Click()
txtsrch.MaxLength = 0
txtsrch.Text = ""
End Sub

Private Sub optreg_Click()
txtsrch.MaxLength = 8
txtsrch.Text = ""
End Sub

Private Sub optren_Click()
txtsrch.MaxLength = 8
txtsrch.Text = ""
End Sub

Private Sub optstat_Click()
txtsrch.MaxLength = 0
txtsrch.Text = ""
End Sub

Private Sub srchgrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblhelp.Caption = ""
End Sub

Private Sub Timer1_Timer()
lblclock.Caption = Format(Time, "HH:MM:SS")
lbldate.Caption = Format(Date, "long date")
End Sub
