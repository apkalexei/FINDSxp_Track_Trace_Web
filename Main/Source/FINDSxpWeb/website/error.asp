<%@ language=vbscript %>
<% On Error Resume Next %>
<!--#include file="include/webfunctions.asp"-->
<%
	Dim strUserName
	Dim strPassword
	Dim strError
	
	strUserName = getQueryItem("username", "A")
	strPassword = getQueryItem("password", "A")
	strError = getQueryItem("err", "A")
	
	Dim i
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
	<TITLE>RF International - Error Screen</TITLE>
     <link rel="stylesheet" type="text/css" href="rfi.css" />
</HEAD>
<BODY text=#ffffff vLink=#ffcc66 aLink=#0099cc link=#ffcc66 bgColor=#ffffff 
leftMargin=8 topMargin=8>
<!--#include file="include/header.html"-->
<TABLE height="100%" cellSpacing=0 cellPadding=0 width="100%" border=0>
	<TR vAlign=center align=middle>
		<TD>
			<TABLE cellSpacing=0 cellPadding=0 border=0>
				<tr>
					<TD vAlign=top align=center colSpan=2 rowSpan=2>&nbsp;<br>

					<b><font size=6>E</font><font size=5>RROR!</font></b></td>
				</tr>
				<TR>
					<td></td>
				</TR>
				<TR>
					<TD valign=top rowSpan=3 align=center height=160>
						<TABLE cellSpacing=2 cellPadding=2 width=319 border=0>
								<TR>
									<td colspan=2 align=center><FONT face="Tahoma, Arial, Helvetica, sans-serif" size=2><b>The page you have requested is unavailable.  Please press the back button in your browser to return to Ship-at-a-Glance.</b></font></td>
								</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
</TABLE>
 <!--#include file="include/footer.html"-->
</BODY></HTML>

