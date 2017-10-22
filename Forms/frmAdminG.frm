VERSION 5.00
Begin VB.Form frmAdmin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administrator:User Management"
   ClientHeight    =   3540
   ClientLeft      =   3585
   ClientTop       =   2970
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   8205
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   435
      Left            =   6360
      TabIndex        =   14
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Frame frmModify 
      Caption         =   "Manipulation"
      Height          =   855
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   5415
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   3960
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   2280
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   480
         TabIndex        =   0
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame fraRecord 
      Caption         =   "Information"
      Height          =   2295
      Left            =   2640
      TabIndex        =   7
      Top             =   120
      Width           =   5415
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Ca&ncel"
         Height          =   315
         Left            =   3240
         TabIndex        =   5
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   315
         Left            =   1800
         TabIndex        =   4
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtretype 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1800
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox txtusername 
         Height          =   285
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   1
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox txtpassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1800
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label lblConfirm 
         AutoSize        =   -1  'True
         Caption         =   "Confirm Password"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Width           =   1260
      End
      Begin VB.Label lblUName 
         AutoSize        =   -1  'True
         Caption         =   "User Name"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   795
      End
      Begin VB.Label lblPassword 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   1290
      End
   End
   Begin VB.ListBox lstUsers 
      Height          =   2205
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   2415
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
   frmModify.Enabled = status
   fraRecord.Enabled = Not status
   lstUsers.Enabled = status
   cmdExit.Enabled = status
End Sub
Private Sub showtxt()
    txtusername.Text = Empty
    txtpassword.Text = Empty
    txtretype.Text = Empty
    'txtusername.SetFocus
End Sub
Private Sub cmdadd_Click()
    
    EnableDisable False
    showtxt
    txtusername.SetFocus
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
    If Len(txtusername.Text) = 0 Then
        MsgBox "please Select Record which you want to delete"
    Else
        If rstadmin.State = adStateOpen Then rstadmin.Close
        rstadmin.Open "select * from users where user_name='" & txtusername.Text & "'", cnn, adOpenKeyset, adLockOptimistic
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
    txtpassword.Text = ""
    txtretype.Text = ""
    EnableDisable False
    txtpassword.SetFocus
    isEditNewMode = True
End Sub


Private Sub cmdexit_Click()
    If MsgBox("Have you sure to Exit", vbYesNo + vbDefaultButton2 + vbCritical, "Exit") = vbYes Then
        End
    End If
End Sub

Private Sub cmdsave_Click()
    Dim rsTemp As New ADODB.Recordset
    rsTemp.Open "select user_name from users where user_name = '" & txtusername.Text & "'", cnn, adOpenKeyset, adLockOptimistic
    
    If rsTemp.RecordCount > 0 Then
        MsgBox "Already Exists..."
        Exit Sub
    End If
    

    If (Len(txtusername.Text) > 4) Then
    If Not ((Len(txtpassword.Text) < 4) Or (Len(txtretype.Text) < 4)) Then
        If (txtpassword.Text = txtretype.Text) Then
            
            If insert = True Then
                
                rstadmin.MoveFirst
                For i = 0 To lstUsers.ListCount - 1
                    If (txtusername.Text = lstUsers.List(i)) Then
                        MsgBox "can't add duplicate record Error", vbInformation
                        showtxt
                        Exit Sub
                    End If
                    rstadmin.MoveNext
                Next
                rstadmin.AddNew
                rstadmin.Fields(0).Value = txtusername.Text
                rstadmin.Fields(1).Value = txtpassword.Text
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
                txtusername.Enabled = False
                If rstadmin.State = adStateOpen Then rstadmin.Close
                rstadmin.Open "select * from users", cnn, adOpenKeyset, adLockOptimistic
                rstadmin.Fields(1).Value = txtpassword.Text
                rstadmin.Update
                MsgBox "Record Successfully Updated", vbInformation
                
                lstUsers.ListIndex = 0
            End If
            EnableDisable True
            showtxt
            cmdAdd.SetFocus
            
        Else
            MsgBox "password must be same", vbCritical
            txtpassword.SetFocus
        End If
    Else
        MsgBox "Password Must be at least 4 Character"
        txtpassword.Text = ""
        txtretype.Text = ""
        txtpassword.SetFocus
    End If
Else
        MsgBox "Username must at least 4 character"
        txtusername.SetFocus
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
    If (Right(txtpassword.Text, 1) = " " And KeyAscii = 32) Then KeyAscii = 0
End Sub

Private Sub txtusername_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 32) Then
        KeyAscii = 0
    End If
End Sub

Private Sub lstUsers_click()
    txtpassword.Text = "***************"
    txtretype.Text = "***************"
    For i = 0 To lstUsers.ListCount - 1
        If lstUsers.Selected(i) = True Then
            txtusername.Text = lstUsers.List(i)
        End If
    Next
End Sub
