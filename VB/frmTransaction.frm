VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmTransaction 
   Caption         =   "Transaction Form"
   ClientHeight    =   6330
   ClientLeft      =   6885
   ClientTop       =   2985
   ClientWidth     =   6435
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   6435
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
      Height          =   495
      Left            =   3600
      TabIndex        =   28
      Top             =   5760
      Width           =   2775
   End
   Begin VB.TextBox txtBAcNum 
      BackColor       =   &H00FFFFFF&
      DataField       =   "BankAcNumber"
      DataSource      =   "adcTransaction"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtBName 
      BackColor       =   &H00FFFFFF&
      DataField       =   "BankName"
      DataSource      =   "adcTransaction"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtVat 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Vat"
      DataSource      =   "adcTransaction"
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtTPrice 
      BackColor       =   &H00FFFFFF&
      DataField       =   "TotalPrice"
      DataSource      =   "adcTransaction"
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtPrice 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Price"
      DataSource      =   "adcTransaction"
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdNewTransaction 
      Caption         =   "New Transaction"
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
      Left            =   120
      TabIndex        =   9
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
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
      Left            =   4200
      TabIndex        =   8
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
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
      Left            =   2040
      TabIndex        =   7
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
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
      Left            =   5280
      TabIndex        =   6
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
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
      TabIndex        =   5
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox txtTID 
      BackColor       =   &H80000016&
      DataField       =   "TID"
      DataSource      =   "adcTransaction"
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtTDate 
      BackColor       =   &H00FFFFFF&
      DataField       =   "TDate"
      DataSource      =   "adcTransaction"
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtGID 
      BackColor       =   &H00FFFFFF&
      DataField       =   "GID"
      DataSource      =   "adcTransaction"
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtDDate 
      BackColor       =   &H00FFFFFF&
      DataField       =   "DDate"
      DataSource      =   "adcTransaction"
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtCID 
      BackColor       =   &H00FFFFFF&
      DataField       =   "CID"
      DataSource      =   "adcTransaction"
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid grdTransaction 
      Bindings        =   "frmTransaction.frx":0000
      Height          =   855
      Left            =   120
      TabIndex        =   27
      Top             =   4800
      Width           =   6255
      _ExtentX        =   11033
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
   Begin MSAdodcLib.Adodc adcGenerator 
      Height          =   330
      Left            =   120
      Top             =   6600
      Visible         =   0   'False
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=GMS.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=GMS.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tblGenerator"
      Caption         =   "Generator Info"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Constantia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adcCustomer 
      Height          =   330
      Left            =   3240
      Top             =   6600
      Visible         =   0   'False
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=GMS.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=GMS.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tblCustomer"
      Caption         =   "Customer Info"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Constantia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adcTransaction 
      Height          =   495
      Left            =   120
      Top             =   5760
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=GMS.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=GMS.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tblTransaction"
      Caption         =   "Transaction Info"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Constantia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo dcmbCID 
      Bindings        =   "frmTransaction.frx":001D
      Height          =   315
      Left            =   2040
      TabIndex        =   30
      Top             =   2400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Locked          =   -1  'True
      ListField       =   "CID"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo dcmbGID 
      Bindings        =   "frmTransaction.frx":0037
      Height          =   315
      Left            =   2040
      TabIndex        =   29
      Top             =   1800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Locked          =   -1  'True
      ListField       =   "GID"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lbBName 
      Caption         =   "&Bank Name"
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
      Left            =   3600
      TabIndex        =   26
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label lbBAcNum 
      Caption         =   "&Bank Ac Number"
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
      Left            =   3600
      TabIndex        =   25
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label lbfrmTransaction 
      Alignment       =   2  'Center
      Caption         =   "Transaction Form"
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
      Left            =   960
      TabIndex        =   22
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
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   5895
   End
   Begin VB.Label lbTPrice 
      Caption         =   "&Total Price"
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
      Left            =   3600
      TabIndex        =   19
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lbVat 
      Caption         =   "&Vat"
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
      Left            =   3600
      TabIndex        =   16
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lbDDate 
      Caption         =   "&DeliveryDate"
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
      Left            =   120
      TabIndex        =   15
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label lbPrice 
      Caption         =   "&Price"
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
      Left            =   3600
      TabIndex        =   14
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lbTDate 
      Caption         =   "&TransactionDate"
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
      Left            =   120
      TabIndex        =   13
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label lbCID 
      Caption         =   "&CustomerID"
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
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label lbGID 
      Caption         =   "&GeneratorID"
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
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label lbTID 
      Caption         =   "&TransactionID"
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
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   1575
   End
End
Attribute VB_Name = "frmTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim stockconf As Boolean

Private Sub cmdCancel_Click()
adcTransaction.Recordset.Cancel
adcTransaction.Refresh
stockconf = False
txtTID.Locked = True
txtGID.Locked = True
dcmbGID.Locked = True
txtCID.Locked = True
dcmbCID.Locked = True
txtTDate.Locked = True
txtDDate.Locked = True
txtPrice.Locked = True
txtVat.Locked = True
txtTPrice.Locked = True
txtBName.Locked = True
txtBAcNum.Locked = True
End Sub

Private Sub cmdDelete_Click()
d = MsgBox("Are you sure that you want to delete?", vbYesNo, "Delete.....")
If d = vbYes Then
    adcTransaction.Recordset.Delete
    MsgBox "Record Deleted Successfully", vbInformation, "Deleted....."
    'adcTransaction.Refresh
Else:
    Exit Sub
End If
End Sub

Private Sub cmdEdit_Click()
txtTID.Locked = False
txtGID.Locked = False
dcmbGID.Locked = False
txtCID.Locked = False
dcmbCID.Locked = False
txtTDate.Locked = False
txtDDate.Locked = False
txtPrice.Locked = False
txtVat.Locked = False
txtTPrice.Locked = False
txtBName.Locked = False
txtBAcNum.Locked = False
MsgBox "Data Edit Has Successfully Been Enabled", vbInformation, "Edit....."
End Sub

Private Sub cmdfrmMain_Click()
If txtGID = "" Then
    MsgBox "Cannot Switch To Main Menu", vbCritical, "Empty Field"
    txtGID.SetFocus
ElseIf txtCID = "" Then
    MsgBox "Cannot Switch To Main Menu", vbCritical, "Empty Field"
    txtCID.SetFocus
ElseIf txtDDate = "" Then
    MsgBox "Cannot Switch To Main Menu", vbCritical, "Empty Field"
    txtDDate.SetFocus
ElseIf txtPrice = "" Then
    MsgBox "Cannot Switch To Main Menu", vbCritical, "Empty Field"
    txtPrice.SetFocus
ElseIf txtVat = "" Then
    MsgBox "Cannot Switch To Main Menu", vbCritical, "Empty Field"
    txtVat.SetFocus
ElseIf txtTPrice = "" Then
    MsgBox "Cannot Switch To Main Menu", vbCritical, "Empty Field"
    txtTPrice.SetFocus
ElseIf txtBName = "" Then
    MsgBox "Cannot Switch To Main Menu", vbCritical, "Empty Field"
    txtBName.SetFocus
ElseIf txtBAcNum = "" Then
    MsgBox "Cannot Switch To Main Menu", vbCritical, "Empty Field"
    txtBAcNum.SetFocus
Else: Unload Me
    frmMain.Show
End If
End Sub

Private Sub cmdNewTransaction_Click()
If txtGID = "" Then
    MsgBox "Cannot Add New Record", vbCritical, "Empty Field"
    txtGID.SetFocus
ElseIf txtCID = "" Then
    MsgBox "Cannot Add New Record", vbCritical, "Empty Field"
    txtCID.SetFocus
ElseIf txtDDate = "" Then
    MsgBox "Cannot Add New Record", vbCritical, "Empty Field"
    txtDDate.SetFocus
ElseIf txtPrice = "" Then
    MsgBox "Cannot Add New Record", vbCritical, "Empty Field"
    txtPrice.SetFocus
ElseIf txtVat = "" Then
    MsgBox "Cannot Add New Record", vbCritical, "Empty Field"
    txtVat.SetFocus
ElseIf txtTPrice = "" Then
    MsgBox "Cannot Add New Record", vbCritical, "Empty Field"
    txtTPrice.SetFocus
ElseIf txtBName = "" Then
    MsgBox "Cannot Add New Record", vbCritical, "Empty Field"
    txtBName.SetFocus
ElseIf txtBAcNum = "" Then
    MsgBox "Cannot Add New Record", vbCritical, "Empty Field"
    txtBAcNum.SetFocus
Else
    txtTID.Locked = False
    txtGID.Locked = False
    dcmbGID.Locked = False
    txtCID.Locked = False
    dcmbCID.Locked = False
    txtTDate.Locked = False
    txtDDate.Locked = False
    txtPrice.Locked = False
    txtVat.Locked = False
    txtTPrice.Locked = False
    txtBName.Locked = False
    txtBAcNum.Locked = False

    adcTransaction.Recordset.MoveLast
    xTID = Right(txtTID, 5)
        adcTransaction.Recordset.AddNew
    txtTID = "T" + CStr(xTID + 1)
End If

txtTDate.Text = Date
txtDDate.Text = DateAdd("d", 15, txtTDate)
stockconf = True
MsgBox "Record Added Successfully", vbInformation, "New Transaction"

End Sub

Private Sub cmdSave_Click()
If txtGID = "" Then
    MsgBox "Cannot Save Record", vbCritical, "Empty Field"
    txtGID.SetFocus
ElseIf txtCID = "" Then
    MsgBox "Cannot Save Record", vbCritical, "Empty Field"
    txtCID.SetFocus
ElseIf txtDDate = "" Then
    MsgBox "Cannot Save Record", vbCritical, "Empty Field"
    txtDDate.SetFocus
ElseIf txtPrice = "" Then
    MsgBox "Cannot Save Record", vbCritical, "Empty Field"
    txtPrice.SetFocus
ElseIf txtVat = "" Then
    MsgBox "Cannot Save Record", vbCritical, "Empty Field"
    txtVat.SetFocus
ElseIf txtTPrice = "" Then
    MsgBox "Cannot Save Record", vbCritical, "Empty Field"
    txtTPrice.SetFocus
ElseIf txtBName = "" Then
    MsgBox "Cannot Save Record", vbCritical, "Empty Field"
    txtBName.SetFocus
ElseIf txtBAcNum = "" Then
    MsgBox "Cannot Save Record", vbCritical, "Empty Field"
    txtBAcNum.SetFocus
Else: adcTransaction.Recordset.Update

If stockconf = True Then
 s = adcGenerator.Recordset.Fields("Stock").Value
 s = s - 1
 With adcGenerator.Recordset
  .Fields("Stock") = s
  .Update
 End With
 stockconf = False
End If
    MsgBox "Record Saved Successfully", vbInformation, "Saved....."
End If
End Sub

Private Sub dcmbCID_Change()
adcCustomer.Recordset.Bookmark = dcmbCID.SelectedItem
txtCID.Text = dcmbCID
txtBName = adcCustomer.Recordset.Fields("BankName").Value
txtBAcNum = adcCustomer.Recordset.Fields("BankAcNum").Value
End Sub

Private Sub dcmbGID_Change()
adcGenerator.Recordset.Bookmark = dcmbGID.SelectedItem
txtGID.Text = dcmbGID
txtPrice = adcGenerator.Recordset.Fields("SalePrice").Value
End Sub

Private Sub Form_Load()
adcTransaction.Recordset.Sort = "TID"
End Sub

Private Sub txtPrice_Change()
txtVat = (Val(txtPrice) / 100) * 15
txtTPrice = Val(txtPrice) + Val(txtVat)
End Sub

Private Sub txtTDate_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  If txtTID Like "DD/MM/YYYY" Then
  txtGID.SetFocus
    Else
    MsgBox "Correct Format: DD/MM/YYYY", vbCritical, "Error Input"
  End If
 End If
End Sub

Private Sub txtTID_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  If txtTID Like "T#####" Then
  txtGID.SetFocus
    Else
    MsgBox "Correct Format: T#####", vbCritical, "Error Input"
  End If
 End If
End Sub


