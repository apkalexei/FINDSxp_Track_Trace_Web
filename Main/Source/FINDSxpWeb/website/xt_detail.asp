<%@Language="VBScript" %>
<%	Option Explicit 
	On Error Resume Next
%>
<!--#include file="include/checkuserlogin.asp"-->
<!--#include file="include/webfunctions.asp"-->
<%	
	Dim strTemp
	Dim strReportIndex
	
	strReportIndex = request("doctype")
	
	Dim strReportFormat
	Dim rs, objConn
	Dim strOceanIndex, strCargoIndex, strInlandIndex
	
	Dim blnGetCargo
	Dim blnGetInland
	Dim intTemp
	Dim strTempIndex
	
	strReportFormat = request("docformat")
	' EVERY REPORT REQUIRES an OCEAN INDEX...  If this is the only one, go ahead and generate the report
	strOceanIndex = request("tracknum")
	
	Set objConn = Server.CreateObject("ADODB.Connection")
	Set rs = Server.CreateObject("ADODB.RecordSet")
	
	objConn.Open Application("DB_CONNECT_STRING")
	rs.Open("fwpReportParamSelect '" & strReportIndex & "'"), objConn
	
	' Determine which extra parameters (besides OceanIndex) are required
	do while not rs.EOF
		if rs("ParamType") = "OceanCargoIndex" Then blnGetCargo = True
		if rs("ParamType") = "InlandIndex" Then blnGetInland = True	
		rs.MoveNext
	loop
	
	rs.Close
	
	CheckError ("Preparing data.")
	
	if Not blnGetCargo AND Not blnGetInland Then 
		objConn.Close
		Set rs = Nothing
		Set objConn = Nothing
		CheckError ("Before redirection.")
		' Ship it off!  Get the report immediately.
		Response.Redirect("getreport.asp?tracknum=" & strOceanIndex & "&doctype=" & strReportIndex & "&docformat=" & strReportFormat)
	Else
		' User entered info required...  Get parameters first
		%>
		<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
		<HTML>
		<HEAD>
			<TITLE>RF International - Transaction Details Screen</TITLE>
		</HEAD>
		<BODY text=#ffffff vLink=#ffcc66 aLink=#0099cc link=#ffcc66 bgColor=#006699 leftMargin=8 topMargin=8>

		<TABLE height="100%" cellSpacing=0 cellPadding=0 width="100%" border=1>
			<TR vAlign=center align=middle>
				<TD>
					<TABLE cellSpacing=0 cellPadding=0 border=0>
						<tr>
							<TD vAlign=top align=center colSpan=2 rowSpan=2><img alt='' src='pictures/shipataglancetitle.jpg'></td>
						</TR>
						<TR>
						</TR>
						<TR>
							<TD valign=top rowSpan=3 align=center>
								<form action=getreport.asp>
									<input type=hidden name=tracknum value='<%= strOceanIndex %>'>
									<input type=hidden name=doctype value='<%= strReportIndex %>'>
									<input type=hidden name=docformat value='<%= strReportFormat %>'>
						<%			if blnGetCargo Then 
										Response.Write("<FONT face='Tahoma, Arial, Helvetica, sans-serif' size=2><b>Please select the Container to use for this report.</b></font><br><br>")
										Response.Write("<select name=cargoindex>")
										rs.Open("fwpOceanCargoIndexSelect '" & strOceanIndex & "', '" & session("UserIndex") & "'"), objConn 
										' Count the number of containers found
										intTemp = 0
										do while Not rs.Eof
											strTempIndex = rs("OceanCargoIndex")
											response.write("<option value='" & strTempIndex & "'>" & rs("ContainerDescription") & "</option>")
											rs.MoveNext
											intTemp = intTemp + 1
										Loop
										
										Response.write("</select>")
										rs.Close
										if intTemp = 0 Then 
											Err.Raise -1239992, "xt_detail.asp", "Calling '" & "fwpOceanCargoIndexSelect '" & strOceanIndex & "', '" & session("UserIndex") & "'" & "'.No parameters were returned for this transaction."
										elseif intTemp = 1 Then 
											Response.redirect("getreport.asp?tracknum=" & strOceanIndex & "&doctype=" & strReportIndex & "&docformat=" & strReportFormat & "&cargoindex=" & strTempIndex)
										end if
									End If  %>
									<input type=submit value='Get Report'>
								</form>												
							</TD>
						</TR>
					</TABLE>
				</TD>
			</TR>
		</TABLE>

		</BODY></HTML>
<%
		CheckError ("Displaying parameter input form")
		
		objConn.Close
		Set rs = Nothing
		Set objConn = Nothing
	
	End If
	
	CheckError ("Error found at end of page.")
%>