<%@ language=vbscript %>
<%	Option Explicit 
	On Error Resume Next
	Response.Buffer = False

	response.write(Application("DB_CONNECT_STRING"))
%>
<!--#include file="include/jakewebfunctions.asp"-->
<%
	Dim strUserName
	Dim strPassword
	
	Dim rs, objConn

	Const TOO_MANY_FAILED_LOGINS = 30

	'strUserName = getQueryItem("username", "A")
	'strPassword = getQueryItem("password", "A")

	strUserName = "mk"
	strPassword = "km"

 	set objConn = server.createobject("ADODB.Connection")
	set rs = server.createobject("ADODB.RecordSet")

	CheckError ("Creating objects")
	
	' If user failed on too many login attempts, don't allow them to try to log in anymore
	If (session("NumFailedLogins") >= TOO_MANY_FAILED_LOGINS) Then Response.Redirect("login.asp?username=" & strUserName & "&err=" & "Invalid%20UserName/Password%20combination.")
	
	If getQueryItem("logoff", "A") <> "yesplease" Then
		'initiate connection to database
		objConn.Open Application("DB_CONNECT_STRING")

		response.write(err.number)
		response.write(err.description)

		' Is user in the database?
		rs.Open "fwpUserLogin '" & strUserName & "', '" & strPassword & "'", objconn
		CheckError ("Opening recordset")
		If rs("UserIndex") = 0 Then 
			' First, keep track of the failed login id, and if enough logins fail in one session, report it to the event log and don't let this user login anymore
			if session("NumFailedLogins") = "" Then 
				session("NumFailedLogins") = 0
			else
				session("NumFailedLogins") = session("NumFailedLogins") + 1
			end if
			' too many failed logins
			if session("NumFailedLogins") >= TOO_MANY_FAILED_LOGINS then 
				' Don't allow this user to even attempt to login anymore
				LogEvent "This user has failed on " & session("NumFailedLogins") & " login attempts.  The last username/password attempted was, '" & request("username") & "'/'" & request("password") & "'"
				Response.Redirect("login.asp?username=" & strUserName & "&err=" & "Invalid%20UserName/Password%20combination.")
			end if
			
			' Set the user's session variable to say that she's not logged in, jump back to login screen
			session("UserIndex") = rs("UserIndex")
			Response.Redirect("login.asp?username=" & strUserName & "&err=" & "Invalid%20UserName/Password%20combination.")
		Else 
			' Set the user's session variable to say that he is logged in, and go to the start screen
			session("UserIndex") = rs("UserIndex")
			session("UserName") = rs("UserName")
			session("UserRole") = rs("UserRole")
			
			'Response.Redirect("start.asp")	
		End If
	Else
		' Set the user's session variable to say that she's not logged in, jump back to login screen
		'session.Abandon
		'Response.Redirect("login.asp")
	End If
	rs.close
	objConn.Close
	set rs = nothing
	set objConn = nothing
	
	CheckError ("Error found at end of page.")
%>

