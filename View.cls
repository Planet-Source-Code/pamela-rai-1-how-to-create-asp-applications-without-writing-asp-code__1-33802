VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "View"
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


    '**normally I have this in another object or module**
    Dim oConn As New ADODB.Connection
    Dim oRs As New ADODB.Recordset
    Dim sList As String
    
    oConn.Open mstrConnect
    oRs.Open "select * from tblNames", oConn
    
    Do While Not oRs.EOF
        sList = sList & "<option value=" & oRs!ID & ">" & _
            oRs!Name & "</option>"
    
        oRs.MoveNext
    Loop
    Set oRs = Nothing
    Set oConn = Nothing
    '********************************************************
    
    'Template Manager
    Set objTemplate = New clsTemplateManager
    With objTemplate
       ' Add values and Tokens
       .TokenAdd sList, "NAMELIST"
       
       ' Set Template Name and Tokens
       .TemplatePath = App.Path & "\Templates" & "\SeeName.htm"
       .TokenPrefix = gconHTMLTokenPrefix
       .TokenSuffix = gconHTMLTokenSuffix
       
       ' Combine tokens and template and return HTML page
       strReturn = .WriteResponse
    End With
    
    ' Write out the page
    Call mobjContext.Response.Write(strReturn)
       
exitPoint:
    Set objTemplate = Nothing
End Function













