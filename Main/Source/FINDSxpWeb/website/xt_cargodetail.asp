<%@Language="VBScript" %>
<%	Option Explicit 
	On Error Resume Next
%>
<!--#include file="include/checkuserlogin.asp"-->
<!--#include file="include/webfunctions.asp"-->
<%	
	function CheckQuotes(str_in_Value, str_in_Heading)
		CheckQuotes = ""
		if instr(1, str_in_Value, "'") Or instr(1, str_in_Value, """") Then CheckQuotes = "Invalid " & str_in_Heading & ": " & "Quotes are not allowed.<br>"
	End Function
	
	Dim objConn
	Dim i
	' Use this value to store any errors that came up
	Dim strErrorMessage
	' Use this value to send data back to the previous page if there's an error
	Dim strValueStorage
	
	strErrorMessage = ""
	strValueStorage = ""
	' Container num
	strErrorMessage = strErrorMessage & CheckQuotes(request("containernum"), "Container Number")
	if request("containernum") = "" Then strErrorMessage = strErrorMessage & "Invalid Container Number: " & "A value must be entered.<br>"
	strValueStorage = strValueStorage & "&containernum=" & request("containernum") & "&bb=" & request("bb")
	
	' Seal nums
	for i = 1 to 5
		strErrorMessage = strErrorMessage & CheckQuotes(request("seal" & i), "Seal " & i)
		strValueStorage = strValueStorage & "&seal" & i & "=" & request("seal" & i)
	Next
	
	' Return error if there were any errors
	if strErrorMessage <> "" Then 
		strErrorMessage = replace(strErrorMessage, " ", "%20") 
		Response.redirect("cargodetail.asp?tab=0" & "&err=" & strErrorMessage & strValueStorage & "&cargoindex=" & request("cargoindex"))
	Else
		Set objConn = Server.CreateObject("ADODB.Connection")
	
		objConn.Open Application("DB_CONNECT_STRING")
	
		objConn.Execute("fwpContainerUpdate '" & ucase(request("cargoindex")) & "', '" & ucase(request("containernum")) & "', '" & ucase(request("seal1")) & "', '" & ucase(request("seal2")) & "', '" & ucase(request("seal3")) & "', '" & ucase(request("seal4")) & "', '" & ucase(request("seal5")) & "', '" & ucase(session("UserIndex")) & "'")
		objConn.Close
		Set objConn = Nothing
		Response.redirect("cargodetail.asp?confirm=" & replace("This container was successfully updated.<br>", " ", "%20") & "&bb=" & request("bb") & "&cargoindex=" & request("cargoindex") & "&tab=0")
	End If
	
	CheckError ("Error found at end of page.")
%>