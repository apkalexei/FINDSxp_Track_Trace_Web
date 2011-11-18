<%@Language="VBScript" %>
<%	Option Explicit 
	On Error Resume Next
%>
<!--#include file="include/webfunctions.asp"-->
<!--#include file="include/checkuserlogin.asp"-->

<%
	CheckError ("Error found at beginning of page.")

	Const REDIRECT_REPORTS_DIRECTORY = "tempdocuments/"
	
	Dim objExporter
	Dim strDocFileName, strTrackNum, strCargoIndex
	Dim strFileExtention
	Dim strTempFileName
	Dim objConn, rs
	Dim strReportFileName
	dim strQuery
	Dim fso
	Dim folder
	Dim files
	Dim file
	Dim intRandom
	Dim strTempLeftOfFileName
	
	'Check that the user is allowed to see the document we are about to create
	set objConn = server.createobject("ADODB.Connection")
	set rs = server.createobject("ADODB.RecordSet")
	
	objConn.Open Application("DB_CONNECT_STRING")
	strQuery = "fwpUserReportAuthenticate '" & request("tracknum") & "', '" & session("UserIndex") & "', '" & request("doctype") & "'"
	rs.Open strQuery, objconn
	CheckError ("Error found calling " & strQuery & ".")

	strReportFileName = rs("RptFile")
		
	rs.close
	objConn.Close
	set rs = nothing
	set objConn = nothing
		
	if request("docformat") = "0" Then
		strFileExtention = ".pdf"
	elseif request("docformat") = "1" Then
		strFileExtention = ".doc"
	elseif request("docformat") = "2" Then
		strFileExtention = ".html"
	End If
	
	strCargoIndex = Trim(request("cargoindex"))
	if len(strCargoIndex) < 1 Then strCargoIndex = ""
	
	' Use a random number to generate the report file name (so that an unauthorized user cannot guess the report name)
	randomize
	intRandom = int(rnd * 10000000)
	strTempFileName = strReportFileName & request("tracknum") & "_" & intRandom & strFileExtention
	strTempLeftOfFileName = strReportFileName & request("tracknum") & "_" & intRandom
	strDocFileName = Application("EXPORT_REPORTS_DIRECTORY") & strTempFileName
	
	strTrackNum = Trim(request("tracknum"))
	
	CheckError ("Error during Preparation.")
	
	Set objExporter = Server.CreateObject("prjReportExporter.clsReportExporter")
	
	CheckError ("Error creating the report object.")
	' Add the parameters - they must be added in order
	objExporter.addParameter(CLng(strTrackNum))
	if strCargoIndex <> "" Then objExporter.addParameter(CLng(strCargoIndex))
	
	CheckError ("Error Adding parameters to the report.")
		
	objExporter.ExportReport request("docformat"), Application("CRYSTAL_REPORTS_DIRECTORY") & strReportFileName & ".rpt", strDocFileName
	
	' Delete the last file that was generated
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	
	' Delete associated all files associated with the last report generated (HTML exports sometimes include images)
	if session("LeftTextOfExportFilesToDelete") <> "" Then
		Set folder = fso.GetFolder(Application("EXPORT_REPORTS_DIRECTORY"))
		Set files = folder.Files

		For each file in files
			if session("LeftTextOfExportFilesToDelete") <> "" and left(ucase(file.name), len(session("LeftTextOfExportFilesToDelete"))) = ucase(session("LeftTextOfExportFilesToDelete")) then 
				file.delete
			end if
		Next
		session("LeftTextOfExportFilesToDelete") = ""
		Set folder = nothing
		set files = nothing
		set file = nothing
	End If
	
	Set fso = Nothing
	
	' Save this report in the session object for deleting
	session("LeftTextOfExportFilesToDelete") = strTempLeftOfFileName
	
	CheckError ("Exporting the report.")
	
	Set objExporter = Nothing
		
	'Response.Redirect(REDIRECT_REPORTS_DIRECTORY & strTempFileName)
	CheckError ("Error found at end of page.")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
	<TITLE>RF International - Report Generation</TITLE>
     <link rel="stylesheet" type="text/css" href="rfi.css" />
</HEAD>
<BODY text=#ffffff vLink=#ffcc66 aLink=#0099cc link=#ffcc66 bgColor=#ffffff leftMargin=8 topMargin=8>
<!--#include file="include/header.html"-->
<TABLE height="100%" cellSpacing=0 cellPadding=0 width="100%" border=0>
	<TR vAlign=center align=middle>
		<TD align=center>
			<TABLE cellSpacing=0 cellPadding=0 border=0>
				<tr>
					<TD width=vAlign=top align=center colSpan=2 rowSpan=2>
						<b><font size=6>D</font><font size=5>OCUMENT</font><font size=6>&nbsp;G</font><font size=5>ENERATION</font></b>
					</td>
				</tr>
				<tr>
				</tr>
				<TR>
					<TD valign=top align=center height=160 colspan=2>
						<FORM name=login action='<%= REDIRECT_REPORTS_DIRECTORY & strTempFileName %>' method=get>
							
							<TABLE cellSpacing=2 cellPadding=2 border=0 name="form">
								<TR vAlign=center>
									<TH align=right><FONT face="Tahoma, Arial, Helvetica, sans-serif" size=2><% Response.Write("Your document has been prepared.  Please click <a href='" & REDIRECT_REPORTS_DIRECTORY & strTempFileName & "'><font color=blue>here</font color></a> to view it.") %></FONT></TH>
								</tr>
								<tr>
									<td align=center>
										<br><br><br><a target="_top" href='xt_login.asp?logoff=yesplease'><FONT face="Tahoma, Arial, Helvetica, sans-serif" size=2 color=#003768><b>Logoff</b></font></a>&nbsp;<font color=#003768><b>|</b></font>&nbsp;<a target="_top" href='start.asp'><FONT face="Tahoma, Arial, Helvetica, sans-serif" size=2 color=#003768><b>Main Page</b></font></a>&nbsp;<font color=#003768><b>|</b></font>&nbsp;<a href='<%= Session("searchpage") %>'><FONT face="Tahoma, Arial, Helvetica, sans-serif" size=2 color=#003768><b>Search</b></a></font>&nbsp;<font color=#003768><b>|</b></font>&nbsp;<a href='<%= Session("detailpage") %>'><FONT face="Tahoma, Arial, Helvetica, sans-serif" size=2 color=#003768><b>Transaction Detail</b></font></a>
									</td>
								</tr>
							</TABLE>
						
						</FORM>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
</TABLE>
  <!--#include file="include/footer.html"-->
</BODY>
</HTML>

<% CheckError ("Error found at end of page.")
%>