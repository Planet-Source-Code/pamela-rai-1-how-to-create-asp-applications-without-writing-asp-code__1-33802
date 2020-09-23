VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTemplateManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mintPrefixLen As Integer
Private mintSuffixLen As Integer

Private mstrTemplatePath As String
Private mstrTokenPrefix As String
Private mstrTokenSuffix As String
Private mcolTokenValues As Collection

Public Sub TokenAdd(strValue As String, strToken As String)
   mcolTokenValues.Add strValue, strToken
End Sub

Public Property Get TokenSuffix() As String
   TokenSuffix = mstrTokenSuffix
End Property

Public Property Let TokenSuffix(strTokenSuffix As String)
   mstrTokenSuffix = strTokenSuffix
End Property

Public Property Get TokenPrefix() As String
   TokenPrefix = mstrTokenPrefix
End Property

Public Property Let TokenPrefix(strTokenPrefix As String)
   mstrTokenPrefix = strTokenPrefix
End Property

Public Property Get TemplatePath() As String
   TemplatePath = mstrTemplatePath
End Property

Public Property Let TemplatePath(strTemplatePath As String)
   mstrTemplatePath = strTemplatePath
End Property

Public Function WriteResponse(Optional ByVal strTemplatePath As String, Optional ByVal strTokenPrefix As String, Optional ByVal strTokenSuffix As String, Optional ByVal colTokenValues As Collection) As String
   'Dim objFSO As Scripting.FileSystemObject
'   Dim objTemplate As Scripting.TextStream
   Dim strHTML As String
   Dim strIn As String
   Dim iHandle As Integer
   
   If Len(strTemplatePath) > 0 Then
      mstrTemplatePath = strTemplatePath
   End If
   If Len(strTokenPrefix) > 0 Then
      mstrTokenPrefix = strTokenPrefix
   End If
   If Len(strTokenSuffix) > 0 Then
      mstrTokenSuffix = strTokenSuffix
   End If
   If Not colTokenValues Is Nothing Then
      Set mcolTokenValues = colTokenValues
   End If
      
      
   iHandle = FreeFile
   Open mstrTemplatePath For Input As iHandle
   
'   Set objFSO = New Scripting.FileSystemObject
'   Set objTemplate = objFSO.OpenTextFile(mstrTemplatePath, ForReading)
   mintPrefixLen = Len(mstrTokenPrefix)
   mintSuffixLen = Len(mstrTokenSuffix)

    Do While Not EOF(iHandle)
        Line Input #iHandle, strIn
        Call ProcessTokens(strIn)
   
        strHTML = strHTML & strIn & vbCrLf
    Loop
   
'   strIn = objTemplate.ReadLine
'   Do Until objTemplate.AtEndOfStream
'      Call ProcessTokens(strIn)
'
'      strHTML = strHTML & strIn & vbCrLf
'      strIn = objTemplate.ReadLine
'   Loop

    Close #iHandle
'   objTemplate.Close
'   Set objTemplate = Nothing
'   Set objFSO = Nothing
   
   WriteResponse = strHTML
End Function

Private Sub ProcessTokens(strIn As String)
   Dim strToken As String
   Dim boolPerform As Boolean
   Dim intPos As Integer
   Dim strValue As String
   Dim strTokenName As String
   
   boolPerform = True
   Do While boolPerform
      intPos = InStr(strIn, mstrTokenPrefix)
      If intPos > 0 Then
         boolPerform = True
         strToken = Mid$(strIn, intPos)
         strToken = Mid$(strToken, 1, InStr(strToken, mstrTokenSuffix) + 1)
         strTokenName = Mid$(strToken, mintPrefixLen + 1, Len(strToken) - mintPrefixLen - mintSuffixLen)
      Else
         boolPerform = False
      End If

      If boolPerform Then
         strValue = GetTokenValue(strTokenName)
         strIn = Replace(strIn, strToken, strValue)
      End If
   Loop
End Sub

Private Function GetTokenValue(strTokenName As String) As String
   On Error GoTo ProcError
   
   GetTokenValue = mcolTokenValues(strTokenName)
   
ProcExit:
   Exit Function
ProcError:
   GetTokenValue = ""
   Resume ProcExit
End Function

Private Sub Class_Initialize()
   Set mcolTokenValues = New Collection
End Sub
