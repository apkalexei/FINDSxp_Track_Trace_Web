<%@Language="VBScript" %>
<%	Option Explicit  
	On Error Resume Next
%>
<!--#include file="include/checkuserlogin.asp"-->
<!--#include file="include/webfunctions.asp"-->
	<html>
	<title>
		<%= request("addressdisplay") %> Address&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	</title>
	<body bgColor=#ffffff>
<%
	Dim objConn, rs
	Dim strQuery
	' Ensure that the user is allowed to access this address type
	Set objConn = Server.CreateObject("ADODB.Connection")
	Set rs = Server.CreateObject("ADODB.RecordSet")
	objConn.Open Application("DB_CONNECT_STRING")
	strQuery = "fwpOceanDetailAddressSelect '" & request("tracknum") & "', '" & session("UserIndex") & "','" & request("addresstype") & "'"
	rs.Open strQuery, objConn
	CheckError ("Calling " & strQuery)
%>
	<FONT face="Tahoma, Arial, Helvetica, sans-serif" color=#003768 size=2><b>
		<%' output the address to the screen %>
		<%= rs("LINE1")%><BR>
		<%= rs("LINE2")%><BR>
		<%= rs("LINE3")%><BR>
		<%= rs("LINE4")%><BR>
		<%= rs("LINE5")%><BR>
		<%= rs("LINE6")%><BR>
		<% CheckError ("Outputting address records.") %>
		<center><input type="button" value="OK" onClick="self.close();" id=button1 name=button1></center>
	</b></FONT>
	</body>
	</html>
<%
	rs.Close
	objConn.Close
	Set rs = Nothing
	Set objConn = Nothing
		
	
	CheckError ("Error found at end of page.") %>