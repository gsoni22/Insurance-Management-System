VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   5595
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim PASSWORD As String
PASSWORD = InputBox("Enter")
Dim spcl, alp, num, cap, sm As Byte
    'Left(" ",len(password))
    sm = InStr(0, " ", PASSWORD)
End Sub
