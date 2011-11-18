<% 
	CheckError ("Error found at beginning of include file: 'webfunctions.asp'.")

	Function getQueryItem(str_in_ParameterName, str_in_Type)
		'D For Date, N for numeric, A for alphanumeric
		If Request(str_in_ParameterName) <> "" Then 
			getQueryItem = Replace(Request(str_in_ParameterName), "'", "''")
			If str_in_Type = "D" Then
				if Not isDate(getQueryItem) then getQueryItem = ""
			End If
		Else
			If str_in_Type = "N" Then
				getQueryItem = 0
			Else
				getQueryItem = ""
			End If
		End If
	End Function
	
	Function GetUserInfo()
		Dim strTemp
	
		strTemp = "Error Number: " & err.number & vbcrlf
		strTemp = strTemp & "Windows Error Description: " & err.description & vbcrlf
		strTemp = strTemp & "Page: " & Request.ServerVariables("URL") & vbcrlf
		strTemp = strTemp & "Querystring: " & Request.QueryString & vbcrlf
		strTemp = strTemp & "User Index: " & session("userindex") & vbcrlf
		strTemp = strTemp & "Browser Type: " & Request.ServerVariables("HTTP_USER_AGENT") & vbcrlf
		strTemp = strTemp & "Remote IP Address: " & Request.ServerVariables("REMOTE_ADDR") & vbcrlf
		strTemp = strTemp & "Remote Host: " & Request.ServerVariables("REMOTE_HOST") & vbcrlf
		strTemp = strTemp & "Website IP Address: " & Request.ServerVariables("LOCAL_ADDR")	
		GetUserInfo = strTemp
	End Function
	
	Sub CheckError(str_in_Description)
		if err.number <> 0 Then
			Dim strTemp


			strTemp = strTemp & vbcrlf & "Page Error Description: " & str_in_Description & vbcrlf & GetUserInfo
		
			' Write the error to the event log
			Dim x
			Set x = Server.CreateObject("prjEventLogWrite.clsEventLog")
			x.EventLogWrite(strTemp)
			Set x = Nothing
			'Response.write strTemp 
			Response.Redirect("error.asp")
		End If
	End Sub
	
	Sub LogEvent(str_in_Description)
		' This function is used when information is logged without redirecting the user to the error page
		Dim strTemp

		strTemp = strTemp & vbcrlf & "Page Error Description: " & str_in_Description & vbcrlf & GetUserInfo
		' Write the error to the event log
		Dim x
		Set x = Server.CreateObject("prjEventLogWrite.clsEventLog")
		x.EventLogWrite(strTemp)
		Set x = Nothing
	End Sub
	
	CheckError ("Error found at end of include file: 'webfunctions.asp'.")
%>