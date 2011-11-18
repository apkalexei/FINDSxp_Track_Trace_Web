<%
	if session("UserIndex") = 0 Then Response.Redirect("login.asp")
	'Response.Write("UserIndex: " & session("UserIndex"))
%>