VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Database"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Uses ADO 2.5
Dim db As New ADODB.Connection
Sub Open_Db_Connection()
    With db
    ' If You Use Access 97 then change then
    ' Change the Provider 30 3.51
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\Security_Test.mdb"
    '
        .CursorLocation = adUseClient
        .Open
    End With
End Sub


Sub Login(rs As ADODB.Recordset, ByVal pLogin As String, ByVal pPassword As String, LoginStatus As Boolean)
    
    Set rs = db.Execute("SELECT USER_NAME, PASSWORD FROM t_Logins " _
    & "WHERE USER_NAME = " & Chr(34) & pLogin & Chr(34))
    ' There Should always only be 1 record,
    ' This is enforced in the Database
    'i.e Username in unique
    If rs.EOF = False Then
        If (rs.Fields("USER_NAME").Value = pLogin And rs.Fields("PASSWORD").Value = pPassword) Then
            LoginStatus = True
        Else
            LoginStatus = False
        End If
    Else
        LoginStatus = False
    End If
    
End Sub

Sub Close_Db_Connection()
    db.Close
End Sub

Function BadLogin() As String
    BadLogin = "<HTML>" & Chr(vbKeyReturn)
    BadLogin = BadLogin & "Bad Login, Not Logged Into System"
    BadLogin = BadLogin & "</HTML>"
End Function
