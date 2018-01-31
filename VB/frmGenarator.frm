VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmGenerator 
   Caption         =   "Generator Form"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6300
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   6300
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtSPrice 
      BackColor       =   &H00FFFFFF&
      DataField       =   "SalePrice"
      DataSource      =   "adcGenerator"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox txtPrice 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Price"
      DataSource      =   "adcGenerator"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtStock 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Stock"
      DataSource      =   "adcGenerator"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtPCase 
      BackColor       =   &H00FFFFFF&
      DataField       =   "PotectionClass"
      DataSource      =   "adcGenerator"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtPower 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Power_kW"
      DataSource      =   "adcGenerator"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox txtVMR 
      BackColor       =   &H00FFFFFF&
      DataField       =   "VoltageModulationRate"
      DataSource      =   "adcGenerator"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtCurrent 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Current_A"
      DataSource      =   "adcGenerator"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtVoltage 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Voltage_V"
      DataSource      =   "adcGenerator"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox txtManCountry 
      BackColor       =   &H00FFFFFF&
      DataField       =   "ManufacturingCountry"
      DataSource      =   "adcGenerator"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtMan 
      BackColor       =   &H00FFFFFF&
      DataField       =   "ManufacturerCo"
      DataSource      =   "adcGenerator"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtModel 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Model"
      DataSource      =   "adcGenerator"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtGType 
      BackColor       =   &H00FFFFFF&
      DataField       =   "GType"
      DataSource      =   "adcGenerator"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1800
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
      Left            =   3360
      TabIndex        =   10
      Top             =   6960
      Width           =   2535
   End
   Begin VB.TextBox txtGID 
      BackColor       =   &H80000016&
      DataField       =   "GID"
      DataSource      =   "adcGenerator"
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
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
      Left            =   3000
      TabIndex        =   4
      Top             =   5400
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
      Left            =   5160
      TabIndex        =   3
      Top             =   5400
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
      Left            =   1920
      TabIndex        =   2
      Top             =   5400
      Width           =   1095
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
      Left            =   4080
      TabIndex        =   1
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmdNewGenerator 
      Caption         =   "New Generator"
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
      TabIndex        =   0
      Top             =   5400
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc adcGenerator 
      Height          =   495
      Left            =   120
      Top             =   6960
      Width           =   3135
      _ExtentX        =   5530
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
   Begin MSDataGridLib.DataGrid grdGenerator 
      Bindings        =   "frmGenarator.frx":0000
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   6000
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
   Begin MSDataListLib.DataCombo dcmbGType 
      Bindings        =   "frmGenarator.frx":001B
      Height          =   315
      Left            =   1800
      TabIndex        =   33
      Top             =   1800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "GType"
      Text            =   ""
   End
   Begin VB.Label lbSPrice 
      Caption         =   "&Sale Price"
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
      Left            =   3360
      TabIndex        =   35
      Top             =   4200
      Width           =   1455
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
      Left            =   3360
      TabIndex        =   32
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label lbStock 
      Caption         =   "&Stock"
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
      Left            =   3360
      TabIndex        =   30
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label lbPCase 
      Caption         =   "&Protection Case"
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
      Left            =   3360
      TabIndex        =   28
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label lbPower 
      Caption         =   "&Power_kW"
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
      TabIndex        =   26
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label lbVMR 
      Caption         =   "&V Modulation Rate"
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
      Left            =   3360
      TabIndex        =   24
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label lbCurrent 
      Caption         =   "&Current_A"
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
      Left            =   3360
      TabIndex        =   22
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lbVoltage 
      Caption         =   "&Voltage_V"
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
      TabIndex        =   20
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label LbManCountry 
      Caption         =   "&Manufacturing Country"
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
      TabIndex        =   18
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label lbMan 
      Caption         =   "&Manufacturer"
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
      TabIndex        =   16
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label lbModel 
      Caption         =   "&Model"
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
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label lbGtype 
      Caption         =   "&GType"
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
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label lbfrmGenerator 
      Alignment       =   2  'Center
      Caption         =   "Generator Form"
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
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   6135
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
      TabIndex        =   7
      Top             =   1200
      Width           =   1575
   End
End
Attribute VB_Name = "frmGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbGType_Change()
txtGType.Text = cmbGType
End Sub

Private Sub cmdCancel_Click()
adcGenerator.Recordset.Cancel
adcGenerator.Refresh
txtGID.Locked = True
txtGType.Locked = True
txtModel.Locked = True
txtMan.Locked = True
txtManCountry.Locked = True
txtPower.Locked = True
txtVoltage.Locked = True
txtCurrent.Locked = True
txtVMR.Locked = True
txtPCase.Locked = True
txtStock.Locked = True
txtPrice.Locked = True
txtSPrice.Locked = True
End Sub

Private Sub cmdDelete_Click()
d = MsgBox("Are you sure that you want to delete?", vbYesNo, "Delete.....")
If d = vbYes Then
    adcGenerator.Recordset.Delete
    MsgBox "Record Deleted Successfully", vbInformation, "Deleted....."
    adcGenerator.Refresh
Else: Exit Sub
End If
End Sub

Private Sub cmdEdit_Click()
txtGID.Locked = False
txtGType.Locked = False
txtModel.Locked = False
txtMan.Locked = False
txtManCountry.Locked = False
txtPower.Locked = False
txtVoltage.Locked = False
txtCurrent.Locked = False
txtVMR.Locked = False
txtPCase.Locked = False
txtStock.Locked = False
txtPrice.Locked = False
txtSPrice.Locked = False
MsgBox "Data Edit Has Successfully Been Enabled", vbInformation, "Edit....."
End Sub

Private Sub cmdClose_Click()
End
End Sub

Private Sub cmdfrmMain_Click()
Unload Me
frmMain.Show
End Sub

Private Sub cmdNewGenerator_Click()
If txtGType = "" Then
    MsgBox "Cannot Add New Record", vbCritical, "Empty Field"
    txtGType.SetFocus
ElseIf txtModel = "" Then
    MsgBox "Cannot Add New Record", vbCritical, "Empty Field"
    txtModel.SetFocus
ElseIf txtMan = "" Then
    MsgBox "Cannot Add New Record", vbCritical, "Empty Field"
    txtMan.SetFocus
ElseIf txtManCountry = "" Then
    MsgBox "Cannot Add New Record", vbCritical, "Empty Field"
    txtManCountry.SetFocus
ElseIf txtPower = "" Then
    MsgBox "Cannot Add New Record", vbCritical, "Empty Field"
    txtPower.SetFocus
ElseIf txtVoltage = "" Then
    MsgBox "Cannot Add New Record", vbCritical, "Empty Field"
    txtVoltage.SetFocus
ElseIf txtCurrent = "" Then
    MsgBox "Cannot Add New Record", vbCritical, "Empty Field"
    txtCurrent.SetFocus
ElseIf txtVMR = "" Then
    MsgBox "Cannot Add New Record", vbCritical, "Empty Field"
    txtVMR.SetFocus
ElseIf txtPCase = "" Then
    MsgBox "Cannot Add New Record", vbCritical, "Empty Field"
    txtPCase.SetFocus
ElseIf txtStock = "" Then
    MsgBox "Cannot Add New Record", vbCritical, "Empty Field"
    txtStock.SetFocus
ElseIf txtPrice = "" Then
    MsgBox "Cannot Add New Record", vbCritical, "Empty Field"
    txtPrice.SetFocus
ElseIf txtSPrice = "" Then
    MsgBox "Cannot Add New Record", vbCritical, "Empty Field"
    txtSPrice.SetFocus
Else
    txtGID.Locked = False
    txtGType.Locked = False
    txtModel.Locked = False
    txtMan.Locked = False
    txtManCountry.Locked = False
    txtPower.Locked = False
    txtVoltage.Locked = False
    txtCurrent.Locked = False
    txtVMR.Locked = False
    txtPCase.Locked = False
    txtStock.Locked = False
    txtPrice.Locked = False
    txtSPrice.Locked = False
    
    adcGenerator.Recordset.MoveLast
    xGID = Right(txtGID, 3)
    adcGenerator.Recordset.AddNew
    txtGID = "G" + CStr(xGID + 1)
End If
MsgBox "Record Added Successfully", vbInformation, "New Generator"
End Sub

Private Sub cmdSave_Click()
If txtGType = "" Then
    MsgBox "Cannot Save Record", vbCritical, "Empty Field"
    txtGType.SetFocus
ElseIf txtModel = "" Then
    MsgBox "Cannot Save Record", vbCritical, "Empty Field"
    txtModel.SetFocus
ElseIf txtMan = "" Then
    MsgBox "Cannot Save Record", vbCritical, "Empty Field"
    txtMan.SetFocus
ElseIf txtManCountry = "" Then
    MsgBox "Cannot Save Record", vbCritical, "Empty Field"
    txtManCountry.SetFocus
ElseIf txtPower = "" Then
    MsgBox "Cannot Save Record", vbCritical, "Empty Field"
    txtPower.SetFocus
ElseIf txtVoltage = "" Then
    MsgBox "Cannot Save Record", vbCritical, "Empty Field"
    txtVoltage.SetFocus
ElseIf txtCurrent = "" Then
    MsgBox "Cannot Save Record", vbCritical, "Empty Field"
    txtCurrent.SetFocus
ElseIf txtVMR = "" Then
    MsgBox "Cannot Save Record", vbCritical, "Empty Field"
    txtVMR.SetFocus
ElseIf txtPCase = "" Then
    MsgBox "Cannot Save Record", vbCritical, "Empty Field"
    txtPCase.SetFocus
ElseIf txtStock = "" Then
    MsgBox "Cannot Save Record", vbCritical, "Empty Field"
    txtStock.SetFocus
ElseIf txtPrice = "" Then
    MsgBox "Cannot Save Record", vbCritical, "Empty Field"
    txtPrice.SetFocus
Else: adcGenerator.Recordset.Update
    MsgBox "Record Saved Successfully", vbInformation, "Saved....."
End If
End Sub

Private Sub dcmbGType_Change()
txtGType = dcmbGType
End Sub

Private Sub Form_Load()
adcGenerator.Recordset.Sort = "GID"
End Sub

Private Sub txtGID_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  If txtGID Like "G###" Then
    Else
    MsgBox "Correct Format: G###", vbCritical, "Error Input"
  End If
 End If
End Sub

Private Sub txtGtype_GotFocus()
g = MsgBox("Is generator type already in database?", vbYesNo, "Generator Type")
If g = vbYes Then
    dcmbGType.Visible = True
Else:
    dcmbGType.Visible = False
    txtGType.SetFocus
End If
End Sub

Private Sub txtModel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtModel = "" Then
    MsgBox "Empty Field", vbCritical
    txtModel.SetFocus
ElseIf KeyAscii = 13 Then
    txtMan.SetFocus
End If
End Sub

Private Sub txtModel_LostFocus()
'If txtModel = "" Then
'    MsgBox "Empty Field", vbCritical
'    txtModel.SetFocus
'End If
End Sub

Private Sub txtPrice_Change()
'txtSPrice.Text = Val(txtPrice.Text) + (Val(txtPrice.Text) * 20 / 100)
End Sub

Private Sub txtStock_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Val(txtStock.Text) > 0 And Val(txtStock.Text) <= 50 Then
Else: MsgBox "Stock is out of range", vbCritical, "Error"
    txtStock.Text = ""
    txtStock.SetFocus
    Exit Sub
End If
If KeyAscii = 13 Then
    txtPrice.SetFocus
End If
End If
End Sub

