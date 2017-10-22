Attribute VB_Name = "Module1"
Public cnn As New ADODB.Connection
'Public rstlogin As New ADODB.Recordset
Public Sub main()
    cnn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Nitesh\26\DataBase\user.mdb"
    cnn.Open
    frmLogin.Show
End Sub
