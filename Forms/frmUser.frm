VERSION 5.00
Begin VB.Form frmUser 
   Caption         =   "Hello "
   ClientHeight    =   2895
   ClientLeft      =   5175
   ClientTop       =   3765
   ClientWidth     =   3930
   LinkTopic       =   "Form1"
   ScaleHeight     =   2895
   ScaleWidth      =   3930
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtNew 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1680
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox txtConfirm 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1680
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtOld 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1680
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label lblConfirm 
      Caption         =   "Confirm Password"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblNew 
      Caption         =   "New Password"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblOld 
      Caption         =   "Old Password"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim user As String
'Dim name As String
Dim rstUser As New ADODB.Recordset
Private Sub reset()
    txtNew.Text = Empty
    txtConfirm.Text = Empty
    txtNew.SetFocus
End Sub
Private Sub cmdexit_Close()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    txtOld.Text = Empty
    txtNew.Text = Empty
    txtConfirm.Text = Empty
    txtOld.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdsave_Click()
    If rstUser.Fields(0).Value = txtOld.Text Then
            If Len(txtNew.Text) >= 6 Then
                If txtNew.Text = txtConfirm.Text Then
                    rstUser.Fields(0).Value = txtNew.Text
                    rstUser.Update
                    MsgBox "Password Sucessfully Changed"
                    Exit Sub
                Else
                    MsgBox "Incorrect Confirm Password"
                    reset
                End If
            Else
                MsgBox "Password must be atleast 6 Character"
                reset
            End If
    Else
        MsgBox "Incorrect Old Password"
        cmdClear_Click
    End If
End Sub

Private Sub Form_Load()
    user = frmLogin.txtUsername.Text
    user = StrConv(user, vbProperCase)
    
    frmUser.Caption = frmUser.Caption + user
    rstUser.Open "select password from users where user_name='" & frmLogin.txtUsername & "'", cnn, adOpenKeyset, adLockOptimistic
End Sub

