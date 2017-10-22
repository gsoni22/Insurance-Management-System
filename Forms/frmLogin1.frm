VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   1980
   ClientLeft      =   5565
   ClientTop       =   3555
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5196.85
   ScaleMode       =   0  'User
   ScaleWidth      =   4859.045
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Login"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Clear"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   1440
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Left            =   480
      Top             =   120
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   1605
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.TextBox txtUsername 
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox txtPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   840
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   120
      Picture         =   "frmLogin1.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "Connecting....."
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   1035
   End
   Begin VB.Label lblUsername 
      AutoSize        =   -1  'True
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1440
      TabIndex        =   6
      Top             =   240
      Width           =   945
   End
   Begin VB.Label lblPassword 
      Caption         =   "Password"
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
      Left            =   1440
      TabIndex        =   5
      Top             =   840
      Width           =   855
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rstlogin As New ADODB.Recordset
Private i As Byte

Private Sub reset()
    txtPassword.Text = Empty
    txtUsername.Text = Empty
    cmdCancel.Enabled = False
    txtUsername.SetFocus
End Sub

Private Sub cmdcancel_Click()
    reset
End Sub

Private Sub cmdexit_Click()
    If (MsgBox("Do you want to exit ?", vbYesNo + vbDefaultButton2 + vbExclamation, "Confirm") = vbYes) Then
        End
    End If
End Sub

Private Sub cmdLogin_Click()
    If rstlogin.State = adStateOpen Then rstlogin.Close
        If (Len(txtUsername.Text) > 4) Or (Len(txtPassword.Text) > 4) Then
            rstlogin.Open "select password from users where user_name = '" & txtUsername.Text & "'", cnn, adOpenKeyset, adLockOptimistic
               If (rstlogin.Fields(0).Value = txtPassword.Text) Then
                    frmLogin.BorderStyle = 0
                    frmLogin.Height = frmLogin.Height + 700
                    lblStatus.Visible = True
                    ProgressBar1.Visible = True
                    Timer1.Enabled = True
                    Timer1.Interval = 200
                    cmdCancel.Enabled = False
                    cmdExit.Enabled = False
                    cmdLogin.Enabled = False
                    Exit Sub
                End If
        End If
    MsgBox "Invalid User or Password"
    reset
End Sub

Private Sub Form_Activate()
    lblStatus.Visible = False
    ProgressBar1.Visible = False
    txtPassword.Enabled = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    rstlogin.Close
End Sub

Private Sub Timer1_Timer()
    ProgressBar1.Value = i
    i = i + 10
    lblStatus.Caption = i & "%"
    If i = 100 Then
        Timer1.Enabled = False
        Dim rstTemp As New ADODB.Recordset
       ' rstTemp.Open "select user_name from users where user_name = '" & txtUsername.Text) & "'", cnn, adOpenKeyset, adLockOptimistic
        If LCase(txtUsername.Text) = "admin" Then
                frmAdmin.Show
            Else
                frmUser.Show
            End If
            Unload Me
    End If
    
End Sub

Private Sub txtPassword_Change()
    If txtPassword.Text = Empty Then
        cmdCancel.Enabled = False
        cmdLogin.Enabled = False
    Else
        cmdCancel.Enabled = True
        cmdLogin.Enabled = True
    End If
End Sub

Private Sub txtUsername_Change()
    If txtUsername.Text = Empty Then
        cmdCancel.Enabled = False
        txtPassword.Enabled = False
    Else
        cmdCancel.Enabled = True
        txtPassword.Enabled = True
    End If
End Sub

Private Sub txtusername_KeyPress(KeyAscii As Integer)
    
    If (Not (KeyAscii = 8 Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 47 And KeyAscii <= 58) Or KeyAscii = 95)) Then
        KeyAscii = 0
    End If
End Sub

