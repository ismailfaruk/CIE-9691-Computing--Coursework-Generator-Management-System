VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCustomer 
   Caption         =   "Customer Form"
   ClientHeight    =   6915
   ClientLeft      =   -120
   ClientTop       =   135
   ClientWidth     =   6045
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   6045
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtAddress 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Address"
      DataSource      =   "adcCustomer"
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   4560
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   31
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtEmail 
      BackColor       =   &H00FFFFFF&
      DataField       =   "EmailID"
      DataSource      =   "adcCustomer"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtGender 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Gender"
      DataSource      =   "adcCustomer"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtBName 
      BackColor       =   &H00FFFFFF&
      DataField       =   "BankName"
      DataSource      =   "adcCustomer"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtAName 
      BackColor       =   &H00FFFFFF&
      DataField       =   "AssetName"
      DataSource      =   "adcCustomer"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtDesignation 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Designation"
      DataSource      =   "adcCustomer"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox txtPType 
      BackColor       =   &H00FFFFFF&
      DataField       =   "PropertyType"
      DataSource      =   "adcCustomer"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtBAcNum 
      BackColor       =   &H00FFFFFF&
      DataField       =   "BankAcNum"
      DataSource      =   "adcCustomer"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox txtCName 
      BackColor       =   &H00FFFFFF&
      DataField       =   "CustomerName"
      DataSource      =   "adcCustomer"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtPhNum 
      BackColor       =   &H00FFFFFF&
      DataField       =   "PhoneNumber"
      DataSource      =   "adcCustomer"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtCID 
      BackColor       =   &H80000016&
      DataField       =   "CID"
      DataSource      =   "adcCustomer"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
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
      Height          =   495
      Left            =   3120
      TabIndex        =   5
      Top             =   6360
      Width           =   2775
   End
   Begin VB.CommandButton cmdNewCustomer 
      Caption         =   "New Customer"
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
      TabIndex        =   4
      Top             =   4800
      Width           =   1695
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
      Left            =   3960
      TabIndex        =   3
      Top             =   4800
      Width           =   975
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
      Left            =   1800
      TabIndex        =   2
      Top             =   4800
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
      Left            =   4920
      TabIndex        =   1
      Top             =   4800
      Width           =   975
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
      Left            =   2880
      TabIndex        =   0
      Top             =   4800
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc adcCustomer 
      Height          =   495
      Left            =   120
      Top             =   6360
      Width           =   2895
      _ExtentX        =   5106
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
   Begin MSDataGridLib.DataGrid grdCustomer 
      Bindings        =   "frmCustomer.frx":0000
      Height          =   855
      Left            =   120
      TabIndex        =   29
      Top             =   5400
      Width           =   5775
      _ExtentX        =   10186
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
   Begin VB.ComboBox cmbGender 
      DataField       =   "Gender"
      DataSource      =   "adcCustomer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmCustomer.frx":001A
      Left            =   1560
      List            =   "frmCustomer.frx":0024
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label lbAddress 
      Caption         =   "&Address"
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
      Left            =   3240
      TabIndex        =   30
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label lbEmailID 
      Caption         =   "&Email ID"
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
      TabIndex        =   28
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label lbGender 
      Caption         =   "&Gender"
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
      TabIndex        =   24
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label lbfrmCustomer 
      Alignment       =   2  'Center
      Caption         =   "Customer Form"
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
      Left            =   1080
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
      TabIndex        =   23
      Top             =   0
      Width           =   6135
   End
   Begin VB.Label lbDesignation 
      Caption         =   "&Designation"
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
      TabIndex        =   21
      Top             =   4200
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
      Left            =   3240
      TabIndex        =   20
      Top             =   4080
      Width           =   1335
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
      Left            =   3240
      TabIndex        =   19
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label lbAName 
      Caption         =   "&Asset Name"
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
      Left            =   3240
      TabIndex        =   18
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lbPhNum 
      Caption         =   "&Phone Number"
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
      TabIndex        =   17
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label lbPType 
      Caption         =   "&Property Type"
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
      Left            =   3240
      TabIndex        =   16
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label lbCname 
      Caption         =   "&Customer Name"
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
      Top             =   1800
      Width           =   1335
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
      TabIndex        =   14
      Top             =   1200
      Width           =   1335
   End
End
Attribute VB_Name = "frmCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbGender_Change()
txtGender = cmbGender
End Sub

Private Sub cmdCancel_Click()
adcCustomer.Recordset.Cancel
adcCustomer.Refresh
adcCustomer.Recordset.Sort = "CID"
txtCID.Locked = True
txtCName.Locked = True
txtGender.Locked = True
cmbGender.Locked = True
txtPhNum.Locked = True
txtPType.Locked = True
txtAName.Locked = True
txtAddress.Locked = True
txtDesignation.Locked = True
txtBName.Locked = True
txtBAcNum.Locked = True
End Sub

Private Sub cmdDelete_Click()
d = MsgBox("Are you sure that you want to delete?", vbYesNo, "Delete.....")
If d = vbYes Then
    adcCustomer.Recordset.Delete
    MsgBox "Record Deleted Successfully", vbInformation, "Deleted....."
    adcCustomer.Refresh
    adcCustomer.Recordset.Sort = "CID"
Else:
    Exit Sub
End If
End Sub

Private Sub cmdEdit_Click()
txtCID.Locked = False
txtCName.Locked = False
txtGender.Locked = False
cmbGender.Locked = False
txtPhNum.Locked = False
txtPType.Locked = False
txtAName.Locked = False
txtAddress.Locked = False
txtDesignation.Locked = False
txtBName.Locked = False
txtBAcNum.Locked = False
MsgBox "Data Edit Has Successfully Been Enabled", vbInformation, "Edit....."
End Sub

Private Sub cmdfrmMain_Click()
If txtCName = "" Then
    MsgBox "Cannot Switch To Main Menu", vbCritical, "Empty Field"
    txtCName.SetFocus
ElseIf txtGender = "" Then
    MsgBox "Cannot Switch To Main Menu", vbCritical, "Empty Field"
    txtGender.SetFocus
ElseIf txtPhNum = "" Then
    MsgBox "Cannot Switch To Main Menu", vbCritical, "Empty Field"
    txtPhNum.SetFocus
ElseIf txtPType = "" Then
    MsgBox "Cannot Switch To Main Menu", vbCritical, "Empty Field"
    txtPType.SetFocus
ElseIf txtAName = "" Then
    MsgBox "Cannot Switch To Main Menu", vbCritical, "Empty Field"
    txtName.SetFocus
ElseIf txtAddress = "" Then
    MsgBox "Cannot Switch To Main Menu", vbCritical, "Empty Field"
    txtAddress.SetFocus
ElseIf txtDesignation = "" Then
    MsgBox "Cannot Switch To Main Menu", vbCritical, "Empty Field"
    txtDesignation.SetFocus
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

Private Sub cmdNewCustomer_Click()
If txtCName = "" Then
    MsgBox "Cannot Add New Record", vbCritical, "Empty Field"
    txtCName.SetFocus
ElseIf txtGender = "" Then
    MsgBox "Cannot Add New Record", vbCritical, "Empty Field"
    txtGender.SetFocus
ElseIf txtPhNum = "" Then
    MsgBox "Cannot Add New Record", vbCritical, "Empty Field"
    txtPhNum.SetFocus
ElseIf txtPType = "" Then
    MsgBox "Cannot Add New Record", vbCritical, "Empty Field"
    txtPType.SetFocus
ElseIf txtAName = "" Then
    MsgBox "Cannot Add New Record", vbCritical, "Empty Field"
    txtName.SetFocus
ElseIf txtAddress = "" Then
    MsgBox "Cannot Add New Record", vbCritical, "Empty Field"
    txtAddress.SetFocus
ElseIf txtDesignation = "" Then
    MsgBox "Cannot Add New Record", vbCritical, "Empty Field"
    txtDesignation.SetFocus
ElseIf txtBName = "" Then
    MsgBox "Cannot Add New Record", vbCritical, "Empty Field"
    txtBName.SetFocus
ElseIf txtBAcNum = "" Then
    MsgBox "Cannot Add New Record", vbCritical, "Empty Field"
    txtBAcNum.SetFocus
ElseIf txtEmail = "" Then
    MsgBox "Cannot Add New Record", vbCritical, "Empty Field"
    txtEmail.SetFocus
Else
    txtCID.Locked = False
    txtCName.Locked = False
    txtGender.Locked = False
    cmbGender.Locked = False
    txtPhNum.Locked = False
    txtPType.Locked = False
    txtAName.Locked = False
    txtAddress.Locked = False
    txtDesignation.Locked = False
    txtBName.Locked = False
    txtBAcNum.Locked = False
    txtEmail.Locked = False
    
    adcCustomer.Recordset.MoveLast
    xCID = Right(txtCID, 4)
        adcCustomer.Recordset.AddNew
    txtCID = "C" + CStr(xCID + 1)
End If

txtCID.SetFocus
MsgBox "Record Added Successfully", vbInformation, "New Customer"
End Sub

Private Sub cmdSave_Click()
If txtCID = "" Then
    MsgBox "Cannot Save Record", vbCritical, "Empty Field"
    txtCID.SetFocus
ElseIf txtCName = "" Then
    MsgBox "Cannot Save Record", vbCritical, "Empty Field"
    txtCName.SetFocus
ElseIf txtGender = "" Then
    MsgBox "Cannot Save Record", vbCritical, "Empty Field"
    txtGender.SetFocus
ElseIf txtPhNum = "" Then
    MsgBox "Cannot Save Record", vbCritical, "Empty Field"
    txtPhNum.SetFocus
ElseIf txtPType = "" Then
    MsgBox "Cannot Save Record", vbCritical, "Empty Field"
    txtPType.SetFocus
ElseIf txtAName = "" Then
    MsgBox "Cannot Save Record", vbCritical, "Empty Field"
    txtName.SetFocus
ElseIf txtAddress = "" Then
    MsgBox "Cannot Add New Record", vbCritical, "Empty Field"
    txtAddress.SetFocus
ElseIf txtDesignation = "" Then
    MsgBox "Cannot Save Record", vbCritical, "Empty Field"
    txtDesignation.SetFocus
ElseIf txtBName = "" Then
    MsgBox "Cannot Save Record", vbCritical, "Empty Field"
    txtBName.SetFocus
ElseIf txtBAcNum = "" Then
    MsgBox "Cannot Save Record", vbCritical, "Empty Field"
    txtBAcNum.SetFocus
ElseIf txtEmail = "" Then
    MsgBox "Cannot Save Record", vbCritical, "Empty Field"
    txtEmail.SetFocus
Else: adcCustomer.Recordset.Update
    MsgBox "Record Saved Successfully", vbInformation, "Saved....."
End If
End Sub

Private Sub Form_Load()
adcCustomer.Recordset.Sort = "CID"
End Sub

Private Sub txtBAcNum_KeyPress(KeyAscii As Integer)
If IsNumeric(txtBAcNum.Text) = True Then
Else: MsgBox "Bank Account Number can be numeric only", vbCritical, "Error"
    txtBAcNum.Text = ""
    txtBAcNum.SetFocus
End If
End Sub

Private Sub txtBAcNum_LostFocus()
'If IsNumeric(txtBAcNum.Text) = True Then
'Else: MsgBox "Bank Account Number can be numeric only", vbCritical, "Error"
'    txtBAcNum.Text = ""
'    txtBAcNum.SetFocus
'End If
End Sub

Private Sub txtBName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtBName = "" Then
    MsgBox "Empty Field", vbCritical
    txtBName.SetFocus
ElseIf KeyAscii = 13 Then
    txtBAcNum.SetFocus
End If
End Sub

Private Sub txtBName_LostFocus()
'If txtBName = "" Then
'    MsgBox "Empty Field", vbCritical
'    txtBName.SetFocus
'End If
End Sub

Private Sub txtCID_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  If txtCID Like "C####" Then
  txtCName.SetFocus
    Else
    MsgBox "Correct Format: C####", vbCritical, "Error Input"
  End If
 End If
End Sub

Private Sub txtCName_Change()
If Len(txtCName.Text) > 25 Then
MsgBox "Customer Name cannot be more than 25 characters", vbCritical, "Error"
End If
End Sub

Private Sub txtEmail_Validate(Cancel As Boolean)
If InStr(1, txtEmail, "@") = 0 Then
        MsgBox "Email Address Not Valid", vbCritical, "Error Input"
        txtEmail.SetFocus
        Exit Sub
End If
End Sub

