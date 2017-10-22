VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDevelopment 
   Caption         =   " "
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   7140
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frsSearching 
      Caption         =   "Frame3"
      Height          =   2295
      Left            =   360
      TabIndex        =   36
      Top             =   5400
      Width           =   6375
      Begin MSFlexGridLib.MSFlexGrid flxResult 
         Height          =   1215
         Left            =   1080
         TabIndex        =   21
         Top             =   840
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   2143
         _Version        =   393216
         Cols            =   4
      End
      Begin VB.TextBox txtFields 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   3480
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   20
         Top             =   360
         Width           =   2535
      End
      Begin VB.ComboBox cmbFields 
         Height          =   315
         Left            =   1080
         TabIndex        =   19
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label12 
         Caption         =   "Select Field"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CheckBox chkSelect 
      Caption         =   "Select"
      Height          =   375
      Left            =   360
      TabIndex        =   18
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   5760
      TabIndex        =   17
      Top             =   4320
      Width           =   975
   End
   Begin VB.Frame frsConfirmation 
      Caption         =   "Confirmation"
      Height          =   1575
      Left            =   5520
      TabIndex        =   35
      Top             =   2520
      Width           =   1455
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   495
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   495
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame frsManipulation 
      Caption         =   "Manipulation"
      Height          =   2175
      Left            =   5520
      TabIndex        =   34
      Top             =   240
      Width           =   1455
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox txtMail 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   2640
      Width           =   3975
   End
   Begin VB.TextBox txtContact 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2040
      Width           =   3975
   End
   Begin VB.TextBox txtState 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   11
      Top             =   4440
      Width           =   3975
   End
   Begin VB.TextBox txtPin 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3840
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   10
      Top             =   3840
      Width           =   1575
   End
   Begin VB.TextBox txtCity 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   3840
      Width           =   1455
   End
   Begin VB.TextBox txtArea 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3000
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   3240
      Width           =   2415
   End
   Begin VB.TextBox txtHno 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox txtDOJ 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3960
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtDOB 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtLname 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3960
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox txtFname 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox txtOfficerID 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label lblMail 
      Caption         =   "Email Address"
      Height          =   375
      Left            =   360
      TabIndex        =   33
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblContact 
      Caption         =   "Contact"
      Height          =   375
      Left            =   360
      TabIndex        =   32
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lblState 
      Caption         =   "State "
      Height          =   375
      Left            =   360
      TabIndex        =   31
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label lblPin 
      Caption         =   "Pincode"
      Height          =   375
      Left            =   3000
      TabIndex        =   30
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label lblCity 
      Caption         =   "City"
      Height          =   375
      Left            =   360
      TabIndex        =   29
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label lblArea 
      Caption         =   "Area"
      Height          =   375
      Left            =   2520
      TabIndex        =   28
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label lblHno 
      Caption         =   "House No."
      Height          =   375
      Left            =   360
      TabIndex        =   27
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label lblDOJ 
      AutoSize        =   -1  'True
      Caption         =   "Date of Join :"
      Height          =   435
      Left            =   3000
      TabIndex        =   26
      Top             =   1440
      Width           =   1065
   End
   Begin VB.Label lblLname 
      Caption         =   "Last Name :"
      Height          =   375
      Left            =   3000
      TabIndex        =   25
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblDOB 
      Caption         =   "Date of Birth"
      Height          =   375
      Left            =   360
      TabIndex        =   24
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblFname 
      Caption         =   "First Name :"
      Height          =   375
      Left            =   360
      TabIndex        =   23
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblOfficerID 
      Caption         =   "Officer ID :"
      Height          =   375
      Left            =   360
      TabIndex        =   22
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmDevelopment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkSelect_Click()

End Sub
