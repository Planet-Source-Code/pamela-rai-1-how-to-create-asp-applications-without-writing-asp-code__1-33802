VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Change"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mstrConnect As String

Private mobjContext As ASPTypeLibrary.ScriptingContext

Public Sub OnStartPage(objContext As ASPTypeLibrary.ScriptingContext)
   Set mobjContext = objContext
End Sub

Public Sub OnEndPage()
   Set mobjContext = Nothing
End Sub

Private Sub Class_Initialize()
   mstrConnect = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & App.Path & "\Sample.mdb"
End Sub

Private Sub Class_Terminate()
   Set mobjContext = Nothing
End Sub

Public Function ProcessRequest() As String
    Dim objTemplate As clsTemplateManager
    Dim strReturn As String
    Dim ID As Long
    Dim Name As String
    
    With mobjContext.Request
        ID = .Form("id")
        Name = .Form("txtName")
    End With


    '**normally I have this in another object or module**
    Dim oConn As New ADODB.Connection
    Dim oRs As New ADODB.Recordset
    Dim sList As String
    
    oConn.Open mstrConnect
    oConn.Execute "update tblNames set Name='" & Name & "' where id=" & ID
    '********************************************************
    
    'the data has changed - now redirect to the list
    mobjContext.Response.Redirect "Exec.asp?prj=MyProject&cls=View"

       
exitPoint:
    Set objTemplate = Nothing
End Function













