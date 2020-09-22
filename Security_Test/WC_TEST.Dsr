VERSION 5.00
Begin {17016CEE-E118-11D0-94B8-00A0C91110ED} WC_TEST 
   ClientHeight    =   4875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4215
   _ExtentX        =   7435
   _ExtentY        =   8599
   MajorVersion    =   0
   MinorVersion    =   8
   StateManagementType=   1
   ASPFileName     =   ""
   DIID_WebClass   =   "{AABFD6B5-070F-11D5-B301-000102A90980}"
   DIID_WebClassEvents=   "{AABFD6B4-070F-11D5-B301-000102A90980}"
   TypeInfoCookie  =   0
   BeginProperty WebItems {193556CD-4486-11D1-9C70-00C04FB987DF} 
      WebItemCount    =   0
   EndProperty
   NameInURL       =   "Test"
End
Attribute VB_Name = "WC_TEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim cls_db As New cls_Database
Private Sub WebClass_BeginRequest()
    If Session("Login") <> "LoggedIn" Then
        Response.Redirect ("Login.asp")
    Else
        ' Do Nothing
    End If
End Sub

Private Sub WebClass_Start()
    If Session("Login") <> "LoggedIn" Then
        Response.Write (cls_db.BadLogin)
    Else
        Response.Write ("Started, Test.asp Login OK")
    End If
End Sub
