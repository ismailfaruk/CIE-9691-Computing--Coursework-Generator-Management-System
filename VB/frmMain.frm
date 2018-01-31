VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Main Menu"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7815
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   7815
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdfrmGenerator 
      Caption         =   "Generator"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5400
      TabIndex        =   5
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5400
      TabIndex        =   4
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton cmdfrmSearchReport 
      Caption         =   "Search/Report"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   3
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton cmdfrmCustomer 
      Caption         =   "Customer"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      TabIndex        =   2
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton cmdfrmTransaction 
      Caption         =   "Transaction"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton cmdfrmLogin 
      Caption         =   "Logout"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      TabIndex        =   0
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label lbfrmMainMenu 
      Alignment       =   2  'Center
      Caption         =   "Main Menu"
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
      Left            =   1920
      TabIndex        =   6
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
      Left            =   840
      TabIndex        =   7
      Top             =   0
      Width           =   6135
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
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

Private Sub cmdfrmLogin_Click()
Unload Me
frmLogin.Show
End Sub

Private Sub cmdfrmSearchReport_Click()
Unload Me
frmSearchReport.Show
End Sub

Private Sub cmdfrmTransaction_Click()
Unload Me
frmTransaction.Show
End Sub
