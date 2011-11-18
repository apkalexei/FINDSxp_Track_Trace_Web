<%@Language="VBScript" %>
<%
	Option Explicit
	On Error Resume Next

	' Constants used in case the page names change
	Const THIS_PAGE_NAME = "detail.asp"
	Const ADDRESS_PAGE_NAME = "address.asp"

	Session("detailpage") = "detail.asp?" & Request.QueryString

	' VARIABLE DECLARATIONS
	Dim objWebReviewScreen
	Dim objConn, rs
	Dim strTemp
	Dim blnFirstRecord
%>
<!--#include file="include/checkuserlogin.asp"-->
<!--#include file="include/webfunctions.asp"-->
<%
	' Ensure that the user is allowed to access this transaction type
	if instr(1, session("availtxtypes"), "~" & request("MnTransType") & "~") = 0 Then
		LogEvent "The user tried to access the the transtype '" & request("MnTransType") & "'.  The available transaction types are in this string: '" & session("availtxtypes") & "'."
		Response.Redirect("error.asp")
	End If

	CheckError ("Error while checking user's access to the requested info.")
%>
<html>
<HEAD>
	<title>RF International - Transaction Details Screen</title>
	<link rel="stylesheet" type="text/css" href="rfi.css" />

<script language=javascript>
	//-- This function opens the window named floater.
	function openWindow(strTrackNum, strInAddressType, strInDisplayAddressType)
	{
		// Pass in the tracknum (1), the address type ('AC'), and the Address Type to display ('Consignee'-> puts 'Consignee Address' in title bar of address window)
		<%'	The WebReviewScreen object creates the construct for addresses to call this function %>
		winStats='toolbar=no,location=no,directories=no,menubar=no,scrollbars=no,width=375,height=150';
		floater=window.open((document.URL.substr(0, document.URL.indexOf('<%= THIS_PAGE_NAME %>')) + '<%= ADDRESS_PAGE_NAME %>?tracknum=' + strTrackNum + '&addresstype=' + strInAddressType + '&addressdisplay=' + strInDisplayAddressType).replace(" ", "+"), "", winStats);
	}
</script>
</HEAD>
<%
	' Browser dependant code to fix netscape 4.7's nasty radio button problem, where it uses the bgcolor of the main screen for the radio bgcolor (ugly)
	if instr(1, request.servervariables("HTTP_USER_AGENT"), "Mozilla/4.77") then
		response.write "<BODY background='pictures/bgcolorpixel.gif' text=#000000 vLink=#ffcc66 aLink=#0099cc link=#ffcc66 bgColor=#c0c0c0 leftMargin=8 topMargin=8>"
	else
		Response.Write "<BODY text=#000000 vLink=#ffcc66 aLink=#0099cc link=#ffcc66 bgColor=#006699 leftMargin=8 topMargin=8>"
	end if
%>
<BODY text=#000000 vLink=#ffcc66 aLink=#0099cc link=#ffcc66 bgColor=#006699 leftMargin=8 topMargin=8>
<!--#include file="include/header.html"-->
<a name="screentop">
<TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
	<TR vAlign=center align=middle>
		<TD>
			<TABLE cellSpacing=0 cellPadding=0 border=0>
				<tr>
					<TD vAlign=top align=center colSpan=2 rowSpan=2><!--<img alt='' src='pictures/shipataglancetitle.jpg'><br>-->
						<b><font size=6>T</font><font size=5>RANSACTION</font><font size=6>&nbsp;D</font><font size=5>ETAILS</font></b><br>
						<b>File #: <%= request("tracknum") %></b></font>
						<br>
						<a target="_top" href='xt_login.asp?logoff=yesplease'><FONT face="Tahoma, Arial, Helvetica, sans-serif" size=2 color=#003768><b>Logoff</b></font></a>&nbsp;<font color=#003768><b>|</b></font>&nbsp;<a target="_top" href='start.asp'><FONT face="Tahoma, Arial, Helvetica, sans-serif" size=2 color=#003768><b>Main Page</b></font></a>&nbsp;<font color=#003768><b>|</b></font>&nbsp;<a href='<%= Session("searchpage") %>'><FONT face="Tahoma, Arial, Helvetica, sans-serif" size=2 color=#003768><b>Search</b></font></a>
					</td>
				</tr>
			</table>
		</td>
	</TR>
</TABLE>

<% CheckError ("Error while displaying doc/format tables.") %>

<%
	' This object generates the entire set of tables for the data screen for us
	Set objWebReviewScreen = Server.CreateObject("prjWebReviewScreen.clsWebReviewScreen")
	CheckError ("Error while creating objWebReviewScreen.")
	objWebReviewScreen.prepareReviewScreenData Application("DB_CONNECT_STRING"), request("tracknum"), Session("UserRole"), Session("UserIndex")
	CheckError ("Error while preparing Review Screen Data.")
%>

<br>
<center>
<table border=0>
	<tr>
		<td>
			 <table cellpadding=0 cellspacing=0 border=0>
				<tr>
					<td colspan=3>
						<table bgcolor=#FFFFFF border=0 cellspacing=0 cellpadding=0>
							<tr bgcolor=#FFFFFF>
							</tr>
							<tr>
								<td colspan=2>
									<table>
										<tr>
											<td width=50>&nbsp;
											</td>
											<td bgcolor=#b0b0b0 width=140><center><font face="arial, helvetica, Tahoma"><b>Contents</b></font></center>
											</td>
											<td width=50>
											</td>
											<td bgcolor=#b0b0b0 width=200><center><font face="arial, helvetica, Tahoma"><b>Reports</b></font></center>
											</td>
											<td width=50>&nbsp;
											</td>
										</tr>
										<tr>
											<td width=50>&nbsp;
											</td>
											<td valign=top>
												<table border=0>
													<tr>
														<td width=10>&nbsp;</td>
														<td align=center>
															<table cellpadding=0 cellspacing=0>
																<%= objWebReviewScreen.getJavaScriptTableDataCode %>
															</table>
														</td>
														<td>

														</td>
													</tr>

												</table>
											</td>
											<td width=50>&nbsp;
											</td>
											<td align=center>
												<TABLE cellSpacing=0 cellPadding=0 border=0>

															<%
															' For each fwpUserReportSelect, output the radio button
															Set objConn = Server.CreateObject("ADODB.Connection")
															Set rs = Server.CreateObject("ADODB.RecordSet")
															objConn.Open Application("DB_CONNECT_STRING")
															' Get the reports available to the user
															rs.Open "fwpUserReportSelect '" & session("UserRole") & "', " & request("MnTransType"), objConn
															' Only display report information if the user has reports available
															if not rs.EOF Then
															%>
																<%' DOCUMENT SELECTION RADIO BUTTONS (PDF, WORD, HTML) %>
																<TR>
																	<td valign=top>
																		<form id=docselection name=docselection action="xt_detail.asp">
																			<b><u><font color=000000><b>FORMAT</b></font></u></b><br>
																			<input type=hidden name=tracknum value=<%= request("tracknum") %>>
																			<input type=radio name=docformat value='0' checked><FONT face="Tahoma, Arial, Helvetica, sans-serif" color=#000000 size=2><b>PDF (Printable)</font></b><br>
																			&nbsp;&nbsp;&nbsp;&nbsp;<A HREF="http://www.adobe.com/products/acrobat/readstep.html"><FONT face="Tahoma, Arial, Helvetica, sans-serif" color=blue size=1>Get PDF Viewer here</font></A><br>
																			<input type=radio name=docformat value='1' ><FONT face="Tahoma, Arial, Helvetica, sans-serif" color=#000000 size=2><b>Word</font></b><br>
																			<input type=radio name=docformat value='2' ><FONT face="Tahoma, Arial, Helvetica, sans-serif" color=#000000 size=2><b>HTML</b></font><br>
																	</td>
																</tr>
																<tr>
																	<td valign=top>
																			<b><u><font color=000000><b>DOCUMENT</b></u></b><br>
															<%
																		' DOCUMENT TYPE RADIO BUTTONS (For each crystal report available)
																		rs.MoveFirst
																		' The first button gets 'checked' when this page is loaded
																		blnFirstRecord = True

																		do while Not rs.EOF
																			Response.write("<input type=radio name=doctype value='" & rs("ReportIndex") & "'")
																			if blnFirstRecord Then
																				' Select the first button (so that the user is unable to leave the report unselected)
																				Response.Write(" checked")
																				blnFirstRecord = False
																			End If
																			Response.Write("><FONT face='Tahoma, Arial, Helvetica, sans-serif' color=#000000 size=2><b><nobr>" & rs("RptName") & "</nobr></b></font><br>")
																			rs.MoveNext
																		Loop
											%>
																	</td>
																</tr>
																<tr>
																	<td height=15>&nbsp;
																	</td>
																</tr>
																<tr>
																	<td  valign=Center>
																			<% 'Submit the form -- Javascript is required to submit it via a hyperlink %>
																			<a href='JavaScript: document.docselection.submit()'><FONT face="Tahoma, Arial, Helvetica, sans-serif" size=2 color=blue><center><b>GENERATE DOCUMENT</b></center></font></a>
																		</form>
																	</TD>
																</TR>
											<%
															Else
																Response.Write("(NONE AVAILABLE)")
															End If

															rs.Close
															set rs = Nothing
															set objConn = Nothing

															%>
												</TABLE>
											</td>
											<td width=50>&nbsp;
											</td>
										</tr>
									</table>
								</td>
							</tr>
							<tr>
								<td colspan=2>
									<%
										' Get the table, and output it to the screen
										Response.Write(objWebReviewScreen.getReviewScreenTable())
									%>
								<br>
								</td>
							</tr>
						</table>
					</td>
					<td width=12 bgcolor=#FFFFFF valign=top></td>
				</tr>
				<tr>
					<td bgcolor=#FFFFFF colspan=6></td>
				</tr>
			</table>
		</td>
	</tr>
</table>
</center>
<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>
<!--#include file="include/footer.html"-->
</body>
</html>
<% CheckError ("Error found at end of page.") %>
