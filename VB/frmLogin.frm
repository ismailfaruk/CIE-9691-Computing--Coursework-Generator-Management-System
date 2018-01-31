VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLogin 
   Caption         =   "Login Form"
   ClientHeight    =   4545
   ClientLeft      =   6750
   ClientTop       =   3435
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   7200
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "adcLogin"
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   1845
      Width           =   2325
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5040
      TabIndex        =   7
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox txtUserName 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "adcLogin"
      Height          =   345
      Left            =   1920
      TabIndex        =   6
      Top             =   1335
      Width           =   2325
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5040
      TabIndex        =   5
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox txtNewPassword 
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   2400
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtCNewPassword 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   3000
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdChPassword 
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5040
      TabIndex        =   2
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      TabIndex        =   1
      Top             =   3720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5040
      TabIndex        =   0
      Top             =   3720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc adcLogin 
      Height          =   330
      Left            =   840
      Top             =   3720
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      RecordSource    =   "tblLogin"
      Caption         =   "Login"
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
   Begin VB.Label lbfrmLogin 
      Alignment       =   2  'Center
      Caption         =   "Login Form"
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
      TabIndex        =   14
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
      Left            =   -120
      TabIndex        =   15
      Top             =   0
      Width           =   6135
   End
   Begin VB.Label lbPassword 
      Caption         =   "&Password"
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
      Left            =   240
      TabIndex        =   13
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label lbOldPassword 
      Caption         =   "&Old Password"
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
      Left            =   240
      TabIndex        =   12
      Top             =   1800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lbNewPassword 
      Caption         =   "&New Password"
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
      Left            =   240
      TabIndex        =   11
      Top             =   2400
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lbUserName 
      Caption         =   "&User Name"
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
      Left            =   240
      TabIndex        =   10
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lbCNewPassword 
      Caption         =   "&Confirm New Password"
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
      Left            =   240
      TabIndex        =   9
      Top             =   3000
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
lbPassword.Visible = True
cmdLogin.Visible = True
cmdChPassword.Visible = True
cmdClose.Visible = True
lbOldPassword.Visible = False
lbNewPassword.Visible = False
txtNewPassword.Visible = False
lbCNewPassword.Visible = False
txtCNewPassword.Visible = False
cmdDone.Visible = False
cmdCancel.Visible = False
End Sub

Private Sub cmdChPassword_Click()
lbPassword.Visible = False
cmdLogin.Visible = False
cmdChPassword.Visible = False
cmdClose.Visible = False
lbOldPassword.Visible = True
lbNewPassword.Visible = True
txtNewPassword.Visible = True
lbCNewPassword.Visible = True
txtCNewPassword.Visible = True
cmdDone.Visible = True
cmdCancel.Visible = True
'txtUserName = ""
'txtPassword = ""
End Sub

Private Sub cmdClose_Click()
End
End Sub

Private Sub cmdDone_Click()
If deGMS.rscomUserName.State = adStateOpen Then
    deGMS.rscomUserName.Close
End If

deGMS.comUserName Trim(txtUserName.Text)
If deGMS.rscomUserName.RecordCount = 0 Then
    MsgBox "Invalied User Name", vbCritical
    Exit Sub
'Else: adcLogin.Recordset.Bookmark = txtUserName.Text
End If

If txtPassword.Text <> deGMS.rscomUserName.Fields("Password").Value Then
    MsgBox "Old Password Is Incorrect", vbCritical
    Exit Sub
Else
 If txtNewPassword.Text = "" Then
MsgBox "Empty Field", vbCritical
txtNewPassword.SetFocus
Exit Sub
ElseIf txtCNewPassword.Text = "" Then
MsgBox "Empty Field", vbCritical
txtCNewPassword.SetFocus
Exit Sub
ElseIf txtNewPassword.Text <> txtCNewPassword Then
MsgBox "Passwords Do Not Match", vbCritical
txtNewPassword.SetFocus
Exit Sub
End If

 With adcLogin.Recordset
  .Fields("Password") = txtNewPassword.Text
  .Update
 End With

MsgBox "Password Changed Successfully", vbInformation

lbPassword.Visible = True
cmdLogin.Visible = True
cmdChPassword.Visible = True
cmdClose.Visible = True
lbOldPassword.Visible = False
lbNewPassword.Visible = False
txtNewPassword.Visible = False
lbCNewPassword.Visible = False
txtCNewPassword.Visible = False
cmdDone.Visible = False
cmdCancel.Visible = False

End If
End Sub

Private Sub cmdLogin_Click()
'check for correct user name and password

If deGMS.rscomUserName.State = adStateOpen Then
    deGMS.rscomUserName.Close
End If

deGMS.comUserName Trim(txtUserName.Text)
If deGMS.rscomUserName.RecordCount = 0 Then
    MsgBox "Invalied User Name", vbCritical
    Exit Sub
End If

If txtPassword.Text = deGMS.rscomUserName.Fields("Password").Value Then
    MsgBox "Welcome", vbInformation
    Unload Me
    frmMain.Show
Else: MsgBox "Invalid Password", vbCritical
    txtPassword.SetFocus
End If

End Sub


