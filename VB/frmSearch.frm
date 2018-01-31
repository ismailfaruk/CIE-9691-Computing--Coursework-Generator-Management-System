VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSearchReport 
   Caption         =   "Search/Report Form"
   ClientHeight    =   5640
   ClientLeft      =   7080
   ClientTop       =   3765
   ClientWidth     =   5775
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   5775
   Begin VB.OptionButton optGType 
      Caption         =   "Generator Type"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2280
      TabIndex        =   13
      Top             =   1920
      Width           =   1455
   End
   Begin VB.OptionButton optCID 
      Caption         =   "Customer ID"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdfrmMain 
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      TabIndex        =   11
      Top             =   4920
      Width           =   2295
   End
   Begin VB.OptionButton optTDate 
      Caption         =   "Transaction Date"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   3720
      TabIndex        =   10
      Top             =   1920
      Width           =   1695
   End
   Begin VB.OptionButton optCName 
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   7
      Top             =   1440
      Width           =   1455
   End
   Begin VB.OptionButton optGID 
      Caption         =   "Generator ID"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   6
      Top             =   1920
      Width           =   1575
   End
   Begin VB.OptionButton optTID 
      Caption         =   "Transaction ID"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox txtSearchReport 
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Frame frameSearch 
      Caption         =   "Search By"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   5295
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Report"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   3360
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid grdSearch 
      Bindings        =   "frmSearch.frx":0000
      Height          =   855
      Left            =   240
      TabIndex        =   12
      Top             =   3960
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1508
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
   Begin VB.Label lbfrmSearchReport 
      Alignment       =   2  'Center
      Caption         =   "Search/Report Form"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   8
      Top             =   600
      Width           =   3855
   End
   Begin VB.Label lblDoricEng 
      Alignment       =   2  'Center
      Caption         =   "Doric Engineering Pvt. Ltd "
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "frmSearchReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdfrmMain_Click()
Unload Me
frmMain.Show
End Sub
Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdfrmCustomer_Click()
Unload Me
frmCustomer.Show
End Sub

Private Sub cmdfrmGenerator_Click()
Unload Me
frmGenerator.Show
End Sub

Private Sub cmdfrmSearch_Click()
Unload Me
frmSearch.Show
End Sub

Private Sub cmdfrmTransaction_Click()
Unload Me
frmTransaction.Show
End Sub

Private Sub cmdReport_Click()
If optTID.Value = True Then
txtSearchReport.SetFocus
If deGMS.rscomTID.State = adStateOpen Then
    deGMS.rscomTID.Close
End If
deGMS.comTID Trim(txtSearchReport.Text)
rptTID.Show

ElseIf optGID.Value = True Then
txtSearchReport.SetFocus
If deGMS.rscomGID.State = adStateOpen Then
    deGMS.rscomGID.Close
End If
deGMS.comGID Trim(txtSearchReport.Text)
rptGID.Show

ElseIf optCID.Value = True Then
txtSearchReport.SetFocus
If deGMS.rscomCID.State = adStateOpen Then
    deGMS.rscomCID.Close
End If
deGMS.comCID Trim(txtSearchReport.Text)
rptCID.Show

ElseIf optCName.Value = True Then
txtSearchReport.SetFocus
If deGMS.rscomCName.State = adStateOpen Then
    deGMS.rscomCName.Close
End If
deGMS.comCName Trim(txtSearchReport.Text)
rptCName.Show

ElseIf optGType.Value = True Then
txtSearchReport.SetFocus
If deGMS.rscomGType.State = adStateOpen Then
    deGMS.rscomGType.Close
End If
deGMS.comGType Trim(txtSearchReport.Text)
rptGType.Show

ElseIf optTDate.Value = True Then
txtSearchReport.SetFocus
If deGMS.rscomTDate.State = adStateOpen Then
    deGMS.rscomTDate.Close
End If
deGMS.comTDate Trim(txtSearchReport.Text)
rptTDate.Show

End If
End Sub

Private Sub cmdSearch_Click()

If optTID.Value = True Then
txtSearchReport.SetFocus
If deGMS.rscomTID.State = adStateOpen Then
    deGMS.rscomTID.Close
End If
deGMS.comTID Trim(txtSearchReport.Text)
grdSearch.DataMember = "comTID"
MsgBox "Search Result Shown", vbInformation, "Search"

ElseIf optGID.Value = True Then
If deGMS.rscomGID.State = adStateOpen Then
    deGMS.rscomGID.Close
End If
deGMS.comGID Trim(txtSearchReport.Text)
grdSearch.DataMember = "comGID"
MsgBox "Search Result Shown", vbInformation, "Search"

ElseIf optGType.Value = True Then
If deGMS.rscomGType.State = adStateOpen Then
    deGMS.rscomGType.Close
End If
deGMS.comGType Trim(txtSearchReport.Text)
grdSearch.DataMember = "comGType"
MsgBox "Search Result Shown", vbInformation, "Search"

ElseIf optCID.Value = True Then
If deGMS.rscomCID.State = adStateOpen Then
    deGMS.rscomCID.Close
End If
deGMS.comCID Trim(txtSearchReport.Text)
grdSearch.DataMember = "comCID"
MsgBox "Search Result Shown", vbInformation, "Search"

ElseIf optCName.Value = True Then
If deGMS.rscomCName.State = adStateOpen Then
    deGMS.rscomCName.Close
End If
deGMS.comCName Trim(txtSearchReport.Text)
grdSearch.DataMember = "comCName"
MsgBox "Search Result Shown", vbInformation, "Search"

ElseIf optTDate.Value = True Then
txtSearchReport.SetFocus
If deGMS.rscomTDate.State = adStateOpen Then
    deGMS.rscomTDate.Close
End If
deGMS.comTDate Trim(txtSearchReport.Text)
grdSearch.DataMember = "comTDate"
MsgBox "Search Result Shown", vbInformation, "Search"

End If
End Sub

Private Sub optCID_Click()
txtSearchReport.Text = ""
txtSearchReport.SetFocus
End Sub

Private Sub optCName_Click()
txtSearchReport.Text = ""
txtSearchReport.SetFocus
End Sub

Private Sub optGID_Click()
txtSearchReport.Text = ""
txtSearchReport.SetFocus
End Sub

Private Sub optGType_Click()
txtSearchReport.Text = ""
txtSearchReport.SetFocus
End Sub

Private Sub optTDate_Click()
txtSearchReport.Text = ""
txtSearchReport.SetFocus
End Sub

Private Sub optTID_Click()
txtSearchReport.Text = ""
txtSearchReport.SetFocus
End Sub
