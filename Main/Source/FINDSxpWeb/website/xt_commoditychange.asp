<%@Language="VBScript" %>
<%	Option Explicit 
	On Error Resume Next
%>
<!--#include file="include/checkuserlogin.asp"-->
<!--#include file="include/webfunctions.asp"-->
<%	
	function IsValidNumber(str_in_Number, int_in_Min, int_in_Max, int_in_MaxNumDecimals)
		Dim strReturn
		
		strReturn=""
		
		if str_in_Number = "" Then
			' This is OK
		elseif Not isNumeric(str_in_Number) Then
			strReturn = "Only numeric values are allowed.<br>"
		elseif CDbl(str_in_Number) > int_in_Max OR CDbl(str_in_Number) < int_in_Min Then 
			strReturn = "The value must fall between " & int_in_Min & " and " & int_in_Max & ".<br>"
		Elseif int_in_MaxNumDecimals = 0 Then
			if instr(1, str_in_Number, ".") Then 
				strReturn = "" & "Decimal values are not allowed.<br>"
			End If
		ElseIf instr(1, str_in_Number, ".") Then
			' Has no "."
			If CDbl(len(str_in_Number) - instr(1, str_in_Number, "." )) > CDbl(int_in_MaxNumDecimals) Then 
				strReturn = "The value cannot exceed " & int_in_MaxNumDecimals & " decimal places.<br>"
			End If
		End If
		
		IsValidNumber = strReturn
	end function
	
	function CheckQty(str_in_Number, str_in_Heading)
		Dim strTemp
		
		strTemp = IsValidNumber(str_in_Number, 0, 999999, 0)
		if strTemp <> "" Then strTemp = "Invalid " & str_in_Heading & ": " & strTemp
		CheckQty = strTemp
	End function
	
	function CheckMeasurement(str_in_Number, str_in_Heading)	
		Dim strTemp
		
		strTemp = IsValidNumber(str_in_Number, 0, 999999999.999, 3)
		if strTemp <> "" Then strTemp = "Invalid " & str_in_Heading & ": " & strTemp
		CheckMeasurement = strTemp
	End Function
	
	Dim objConn
	Dim i
	Dim strErrorMessage
	Dim strValueStorage

	strErrorMessage = ""
	strValueStorage = ""
	 
	' Check for invalid values
	strValueStorage = strValueStorage & "&CmdyIndex=" & request("CmdyIndex") & "&bb=" & request("bb")
	
	strErrorMessage = strErrorMessage & CheckMeasurement(request("NetLBS"), "LBS(Net) For " & request("commodityid"))
	strValueStorage = strValueStorage & "&netlbs=" & request("NetLBS")
	
	strErrorMessage = strErrorMessage & CheckMeasurement(request("NetKGS"), "KGS(Net) For " & request("commodityid"))
	strValueStorage = strValueStorage & "&NetKGS=" & request("NetKGS")
	
	strErrorMessage = strErrorMessage & CheckMeasurement(request("GrossLBS"), "LBS(Gross) For " & request("commodityid"))
	strValueStorage = strValueStorage & "&GrossLBS=" & request("GrossLBS")
	
	strErrorMessage = strErrorMessage & CheckMeasurement(request("GrossKGS"), "LBS(Gross) For " & request("commodityid"))
	strValueStorage = strValueStorage & "&GrossKGS=" & request("GrossKGS")
	
	strErrorMessage = strErrorMessage & CheckMeasurement(request("CFT"), "CFT For " & request("commodityid"))
	strValueStorage = strValueStorage & "&CFT=" & request("CFT")
	
	strErrorMessage = strErrorMessage & CheckMeasurement(request("CBM"), "CBM For " & request("commodityid"))
	strValueStorage = strValueStorage & "&CBM=" & request("CBM")
	
	strErrorMessage = strErrorMessage & CheckQty(request("PieceCount"), "Piece Count For " & request("commodityid"))
	strValueStorage = strValueStorage & "&PieceCount=" & request("PieceCount")
	
	strValueStorage = strValueStorage & "&PieceTypeDesc=" & request("PieceTypeDesc")
	
	
	if strErrorMessage <> "" Then
		strErrorMessage = replace(strErrorMessage, " ", "%20") 
		Response.redirect("cargodetail.asp?tab=1" & strValueStorage & "&cargoindex=" & request("cargoindex") & "&err=" & strErrorMessage)
	Else
		Dim strInsertQuery
		
		Set objConn = Server.CreateObject("ADODB.Connection")
	
		objConn.Open Application("DB_CONNECT_STRING")
		strInsertQuery = "fwpCmdyUpdate '" & ucase(request("CmdyIndex")) & "', '" & request("GrossLBS") & "', '" & request("GrossKGS") & "', '" & request("NetLBS") & "', '" & request("NetKGS") & "', '" & request("CFT") & "', '" & request("CBM") & "', '" & request("PieceCount") & "', '" & ucase(request("PieceTypeDesc")) & "', '" & ucase(session("UserIndex")) & "'"
		objConn.Execute(strInsertQuery)
		
		objConn.Close
		Set objConn = Nothing
		Response.Write(strInsertQuery)
		CheckError("Error while calling " & strInsertQuery)
		Response.Redirect("cargodetail.asp?tab=1&cargoindex=" & request("cargoindex") & "&bb=" & request("bb") & "&confirm=" & replace("The commodity, '" & request("commodityid") & "' was successfully updated.<br>", " ", "%20"))
	End If
	
	CheckError ("Error found at end of page.")
%>