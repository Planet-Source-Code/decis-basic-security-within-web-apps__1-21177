VERSION 5.00
Begin {17016CEE-E118-11D0-94B8-00A0C91110ED} WC_Login 
   ClientHeight    =   7920
   ClientLeft      =   750
   ClientTop       =   1425
   ClientWidth     =   6975
   _ExtentX        =   12303
   _ExtentY        =   13970
   MajorVersion    =   0
   MinorVersion    =   8
   StateManagementType=   1
   ASPFileName     =   ""
   DIID_WebClass   =   "{12CBA1F6-9056-11D1-8544-00A024A55AB0}"
   DIID_WebClassEvents=   "{12CBA1F5-9056-11D1-8544-00A024A55AB0}"
   TypeInfoCookie  =   27
   BeginProperty WebItems {193556CD-4486-11D1-9C70-00C04FB987DF} 
      WebItemCount    =   2
      BeginProperty WebItem1 {FA6A55FE-458A-11D1-9C71-00C04FB987DF} 
         MajorVersion    =   0
         MinorVersion    =   8
         Name            =   "tmp_Login"
         DISPID          =   1280
         Template        =   "Logon.htm"
         Token           =   "WC@"
         DIID_WebItemEvents=   "{CAA2F50B-065D-11D5-B300-000102A90980}"
         ParseReplacements=   0   'False
         AppendedParams  =   ""
         HasTempTemplate =   0   'False
         UsesRelativePath=   -1  'True
         OriginalTemplate=   "C:\Security_Test\HTML Templates\Logon.htm"
         TagPrefixInfo   =   2
         BeginProperty Events {193556D1-4486-11D1-9C70-00C04FB987DF} 
            EventCount      =   0
         EndProperty
         BeginProperty BoundTags {FA6A55FA-458A-11D1-9C71-00C04FB987DF} 
            AttribCount     =   0
         EndProperty
      EndProperty
      BeginProperty WebItem2 {FA6A55FE-458A-11D1-9C71-00C04FB987DF} 
         MajorVersion    =   0
         MinorVersion    =   8
         Name            =   "WI_LOGIN"
         DISPID          =   1281
         Template        =   ""
         Token           =   "WC@"
         DIID_WebItemEvents=   "{CAA2F53B-065D-11D5-B300-000102A90980}"
         ParseReplacements=   0   'False
         AppendedParams  =   ""
         HasTempTemplate =   0   'False
         UsesRelativePath=   0   'False
         OriginalTemplate=   ""
         TagPrefixInfo   =   2
         BeginProperty Events {193556D1-4486-11D1-9C70-00C04FB987DF} 
            EventCount      =   0
         EndProperty
         BeginProperty BoundTags {FA6A55FA-458A-11D1-9C71-00C04FB987DF} 
            AttribCount     =   1
            BeginProperty Attrib0 {FA6A55FC-458A-11D1-9C71-00C04FB987DF} 
               TagType         =   1
               Attribute       =   "ACTION"
               State           =   2
               TagName         =   "frmSearch"
               OriginalURL     =   ""
               Parent          =   ""
               Template        =   "tmp_Login"
               BoundEvent      =   ""
               BoundItem       =   "WI_LOGIN"
               Suffix          =   ""
               UsesAnonymousName=   0
               TagNumber       =   0
            EndProperty
         EndProperty
      EndProperty
   EndProperty
   NameInURL       =   "Login"
End
Attribute VB_Name = "WC_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text
Dim cls_Db As New cls_Database
Dim rs As New ADODB.Recordset
Private Sub WebClass_Start()
    tmp_Login.WriteTemplate
End Sub

Private Sub WI_LOGIN_Respond()
Dim LoginStatus As Boolean
    Call cls_Db.Open_Db_Connection
    
    Call cls_Db.Login(rs, Trim$(Request.Item("txtLogin")), Trim$(Request.Item("txtPassword")), LoginStatus)
    If LoginStatus = True Then
        Session("Login") = "LoggedIn"
        Response.Redirect ("Test.asp") ' Start WC_TEST
    Else
        Session("Login") = ""
        Response.Write cls_Db.BadLogin
    End If
    Call cls_Db.Close_Db_Connection
End Sub
