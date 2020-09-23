<%@ Language=VBScript %>
<%
	dim oInstance
	dim sClass, sProject
	dim sClassString
	
	sClass = Request.Item("cls")
	sProject = Request.Item("prj")
	sClassString = sProject & "." & sClass 
	set oInstance = server.CreateObject(sClassString)
	oInstance.ProcessRequest
	set oInstance = nothing
%>