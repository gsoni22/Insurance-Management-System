VERSION 5.00
Begin VB.Form frmAdmin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administrator"
   ClientHeight    =   4155
   ClientLeft      =   3585
   ClientTop       =   2970
   ClientWidth     =   6585
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   6585
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
      Height          =   555
      Left            =   4320
      TabIndex        =   13
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Frame fraModify 
      Caption         =   "Manupulation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   11
      Top             =   2880
      Width           =   3975
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
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
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
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
         Left            =   1440
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
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
         Left            =   2640
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame fraRecord 
      Caption         =   "Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   2160
      TabIndex        =   9
      Top             =   240
      Width           =   4335
      Begin VB.TextBox txtPassword 
         Enabled         =   0   'False
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1680
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox txtUsername 
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Ca&ncel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2400
         TabIndex        =   5
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1080
         TabIndex        =   4
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox txtConfirm 
         Enabled         =   0   'False
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1680
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label lblPassword 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   825
      End
      Begin VB.Label lblConfirm 
         AutoSize        =   -1  'True
         Caption         =   "Confirm Password"
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
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   1635
      End
      Begin VB.Label lblUserName 
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
         Left            =   120
         TabIndex        =   10
         Top             =   540
         Width           =   945
      End
   End
   Begin VB.ListBox lstUsers 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2400
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label lblCaption 
      Caption         =   "User List"
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
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstadmin As New ADODB.Recordset
Dim insert As Boolean
Dim st As Boolean
Private isEditNewMode As Boolean

Private Sub EnableDisable(ByVal status As Boolean)
   fraModify.Enabled = status
   fraRecord.Enabled = Not status
   lstUsers.Enabled = status
   cmdExit.Enabled = status
End Sub
Private Sub showtxt()
    txtUsername.Text = Empty
    txtPassword.Text = Empty
    txtConfirm.Text = Empty
    'txtusername.SetFocus
End Sub
Private Sub cmdadd_Click()
    
    EnableDisable False
    showtxt
    txtUsername.SetFocus
    insert = True
    isEditNewMode = True
End Sub

Private Sub cmdcancel_Click()
    showtxt
    EnableDisable True
    lstUsers.ListIndex = 0
    isEditNewMode = False
End Sub

Private Sub cmddelete_Click()
    If Len(txtUsername.Text) = 0 Then
        MsgBox "please Select Record which you want to delete"
    Else
        If rstadmin.State = adStateOpen Then rstadmin.Close
        rstadmin.Open "select * from users where user_name='" & txtUsername.Text & "'", cnn, adOpenKeyset, adLockOptimistic
            If rstadmin.Fields(0).Value = "admin" Then
                MsgBox "You can't Delete Administrator"
                cmdAdd.SetFocus
                Exit Sub
            End If
            
      If MsgBox("Have you sure to delete current Record", vbYesNo + vbDefaultButton2 + vbCritical, "Delete") = vbYes Then
            rstadmin.Delete
            rstadmin.Update
            rstadmin.Close
            showtxt
            MsgBox "Record Succussfull Deleted", vbExclamation
        End If
    End If
    lstUsers.Clear
    If rstadmin.State = adStateOpen Then rstadmin.Close
    rstadmin.Open "select * from users", cnn, adOpenKeyset, adLockOptimistic

    rstadmin.MoveFirst
    For i = 0 To rstadmin.RecordCount - 1
        lstUsers.AddItem rstadmin.Fields(0).Value
        rstadmin.MoveNext
    Next
    lstUsers.ListIndex = 0
    cmdAdd.SetFocus
End Sub

Private Sub cmdedit_Click()
    txtPassword.Text = ""
    txtConfirm.Text = ""
    EnableDisable False
    txtPassword.SetFocus
    isEditNewMode = True
End Sub


Private Sub cmdexit_Click()
    If MsgBox("Have you sure to Exit", vbYesNo + vbDefaultButton2 + vbCritical, "Exit") = vbYes Then
        End
    End If
End Sub

Private Sub cmdsave_Click()
    Dim rsTemp As New ADODB.Recordset
    rsTemp.Open "select user_name from users where user_name = '" & txtUsername.Text & "'", cnn, adOpenKeyset, adLockOptimistic
    
    If rsTemp.RecordCount > 0 Then
        MsgBox "Already Exists..."
        Exit Sub
    End If
    

    If (Len(txtUsername.Text) > 4) Then
    If Not ((Len(txtPassword.Text) < 4) Or (Len(txtConfirm.Text) < 4)) Then
        If (txtPassword.Text = txtConfirm.Text) Then
            
            If insert = True Then
                
                rstadmin.MoveFirst
                For i = 0 To lstUsers.ListCount - 1
                    If (txtUsername.Text = lstUsers.List(i)) Then
                        MsgBox "can't add duplicate record Error", vbInformation
                        showtxt
                        Exit Sub
                    End If
                    rstadmin.MoveNext
                Next
                rstadmin.AddNew
                rstadmin.Fields(0).Value = txtUsername.Text
                rstadmin.Fields(1).Value = txtPassword.Text
                rstadmin.Update
                  
                MsgBox "Record Successfully Added", vbInformation
                insert = False
                showtxt
                lstUsers.Clear
                rstadmin.MoveFirst
                For i = 0 To rstadmin.RecordCount - 1
                    lstUsers.AddItem rstadmin.Fields(0).Value
                    rstadmin.MoveNext
                Next
                showtxt
                lstUsers.ListIndex = 0
            Else
                txtUsername.Enabled = False
                If rstadmin.State = adStateOpen Then rstadmin.Close
                rstadmin.Open "select * from users", cnn, adOpenKeyset, adLockOptimistic
                rstadmin.Fields(1).Value = txtPassword.Text
                rstadmin.Update
                MsgBox "Record Successfully Updated", vbInformation
                
                lstUsers.ListIndex = 0
            End If
            EnableDisable True
            showtxt
            cmdAdd.SetFocus
            
        Else
            MsgBox "password must be same", vbCritical
            txtPassword.SetFocus
        End If
    Else
        MsgBox "Password Must be at least 4 Character"
        txtPassword.Text = ""
        txtConfirm.Text = ""
        txtPassword.SetFocus
    End If
Else
        MsgBox "Username must at least 4 character"
        txtUsername.SetFocus
    End If
    isEditNewMode = False
End Sub

Private Sub Form_Activate()
    For i = 0 To rstadmin.RecordCount - 1
        lstUsers.AddItem rstadmin.Fields(0).Value
        rstadmin.MoveNext
    Next
    lstUsers.ListIndex = 0
    EnableDisable True
End Sub

Private Sub Form_Load()
    If rstadmin.State = adStateOpen Then rstadmin.Close
    rstadmin.Open "select * from users", cnn, adOpenKeyset, adLockOptimistic
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If isEditNewMode = True Then
        Cancel = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rstadmin.Close
End Sub

Private Sub txtpassword_KeyPress(KeyAscii As Integer)
    If (Right(txtPassword.Text, 1) = " " And KeyAscii = 32) Then KeyAscii = 0
End Sub

Private Sub txtusername_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 32) Then
        KeyAscii = 0
    End If
End Sub

Private Sub lstUsers_click()
    txtPassword.Text = "***************"
    txtConfirm.Text = "***************"
    For i = 0 To lstUsers.ListCount - 1
        If lstUsers.Selected(i) = True Then
            txtUsername.Text = lstUsers.List(i)
        End If
    Next
End Sub

