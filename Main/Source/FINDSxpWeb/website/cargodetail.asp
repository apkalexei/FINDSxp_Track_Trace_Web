<%@Language="VBScript" %>
<% Option Explicit 
   On Error Resume Next	

%>
<!--#include file="include/checkuserlogin.asp"-->
<!--#include file="include/webfunctions.asp"-->

<%
	' ************************* VARIABLE DECLARATIONS *************************
	' Array for tab names
	Dim aryTabName()
	
	' Current tab that the user is on
	Dim strTab
	Dim strTemp 
	
	Dim i
	Dim intTemp
		
	Dim strThisPageName

	' ************************* PREPARE DATA FOR HTML OUTPUT ************************* 
	strThisPageName = "cargodetail.asp"

	strTab = getQueryItem("tab", "N")
	
	CheckError ("Error preparing data")
    
    Redim aryTabName(3)
    
    ' Get the tab names
	aryTabName(0) = "CONTAINER DETAILS"
	aryTabName(1) = "COMMODITY DETAILS"
	aryTabName(2) = "INLAND DETAILS"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/loose.dtd">
<HTML>
<HEAD>
	<title>RF International - Cargo Details Screen</title>
        <link rel="stylesheet" type="text/css" href="rfi.css" />
</HEAD>

<BODY bgcolor=#FFFFFF>

<%'Outermost table %>
 <center>

	<!--<img src='pictures/shipataglancetitle.jpg'><br>-->
	<b><font size=6>C</font><font size=5>ARGO</font><font size=6>&nbsp;R</font><font size=5>EVIEW</font></b>
	<%	' get the tracknum from the detail page hyperlink
		strTemp = session("detailpage")
		' Cut off the unneeded left portion of the string
		strTemp = right(strTemp, len(strTemp) - instr(1, strTemp, "tracknum=") - 8)
		' Cut off the unneeded right portion of the string
		strTemp = left(strTemp, instr(1, strTemp, "&") - 1)
	%>
	
	<br><font size = 3><b>File #: <%= strTemp %></b></font><br>

	<a target="_top" href='xt_login.asp?logoff=yesplease'><FONT face="Tahoma, Arial, Helvetica, sans-serif" size=2 color=#003768><b>Logoff</b></font></a>&nbsp;<font color=#003768><b>|</b></font>&nbsp;<a target="_top" href='start.asp'><FONT face="Tahoma, Arial, Helvetica, sans-serif" size=2 color=#003768><b>Main Page</b></font></a>&nbsp;<font color=#003768><b>|</b></font>&nbsp;<a href='<%= Session("searchpage") %>'><FONT face="Tahoma, Arial, Helvetica, sans-serif" size=2 color=#003768><b>Search</b></a></font>&nbsp;<font color=#003768><b>|</b></font>&nbsp;<a href='<%= Session("detailpage") %>'><FONT face="Tahoma, Arial, Helvetica, sans-serif" size=2 color=#003768><b>Transaction Detail</b></font></a>
 </center>
<% if request("err") <> "" Then 
		response.write("<br><b>ERROR!<br>" & request("err") & "</b>") 
   elseif request("confirm") <> "" Then
		response.write("<br><b>" & request("confirm") & "</b>") 
   end if
%>

 <br>
 <table cellpadding=0 cellspacing=0 border=0 bgcolor=#FFFFFF>
				<tr>
					<td valign=top>
						<table cellpadding=0 cellspacing=0 border=0>
							<tr>
								<td bgcolor=#FFFFFF height=31></td>
							</tr>
						</table>
					</td>
					
					<td>

						<table bgcolor=#FFFFFF border=0 cellpadding=0 cellspacing=0 bordercolor=red>
						 	<tr>
						  		<td>
						  			<% 'This 2nd table is used for the chunks of the table ie. tabs=Row1, spacer=row2 ,tabledata=Row3, SearchBox/etc=Row4 %>
									<table bgcolor=#FFFFFF cellpadding="0" cellspacing="0" border=0>
										<tr>
											<td align=bottom bgcolor=#FFFFFF>
												<%'Display the tabs %>
												<table cellpadding="0" cellspacing="0" border=0>

													<tr>
													
													<td>
													<table cellspacing=0 cellpadding=0 border=0 bgcolor=#FFFFFF>
													<tr>
														<% for i = 0 to Ubound(aryTabName) - 1 
																' Hide the container tab if it's not available (for a break bulk shipment only)
																if i & "" = 0 AND request("bb") Then 
																' If the user tries to access the container tab, the container tab will try to display and the associated SP will raise an error -> event log
																elseif strTab & "" = i & "" Then  %>
																	<td align=left width=118><table cellspacing=0 cellpadding=0 border=0>
																			<tr>
																				<td width=9></td>
																				<td><table cellspacing=0 cellpadding=0 border=1>
																						<tr>
																							<td height=1 bgcolor=#FFFFFF></td>
																						</tr>
																						<tr>
																							<td bgcolor=#FFFFF height=29 width=100 align=center><b><font size=2 face="Tahoma, Arial, Helvetica, sans-serif color=blue"><%= aryTabName(i) %></font></b></td>
																						</tr>
																					</table></td>
																				<td></td>
																			</tr>
																			<tr>
																				<td  colspan=3 bgcolor=#FFFFFF></td>
																			</tr>
																			</tr>
																		</table></td>
																		<% if i =  Ubound(aryTabName) - 1 Then Response.Write("<td valign=bottom></td>")
																%>
															<% Else %>
																<td align=left valign=bottom width=100><table cellspacing=0 cellpadding=0 border=0>
																		<tr>
																			<td width=7></td>
																			<td><table cellspacing=0 cellpadding=0 border=1>
																					<tr>
																						<td height=1 bgcolor="#FFFFFF"></td>
																					</tr>
																					<tr>
																						<td bgcolor=#FFFFFF height=22 width=100 align=center><a href='<%= strThisPageName & "?tab=" & i & "&cargoindex=" & request("cargoindex") & "&bb=" & request("bb") %>'><b><font size=1 color=blue face="Tahoma, Arial, Helvetica, sans-serif"><b><%= aryTabName(i) %></font></b></a></td>
																					</tr>
																				</table>
																			</td>
																		</tr>
																		<tr>
																			<td colspan=3 bgcolor="#FFFFFF"></td>
																		</tr>
																	</table></td>
																	<% if i =  Ubound(aryTabName) - 1 Then Response.Write("<td valign=bottom></td>")
																	%>
															<% End If %>
														<% Next %>
													</tr>
													</table>
													</td>
													</tr>
												</table>
											</td>
										</tr>
										<tr>
											<td bgcolor=#FFFFFF height=10></td>
										</tr>
										<tr>
											<td height=10></td>
										</tr>
										
										<tr>
											<td>
												<% 	
													Dim objConn, rs
													Dim blnReadOnly
	
													Dim strSeal(5)
													Dim strContainerNum %>
												
												<% If strTab = "0" Then  %>
													<!--#include file="include/cargotabcontainer.asp"-->
												<% elseif strTab = "1" Then %>
													<!--#include file="include/cargotabcommodity.asp"-->
												<% elseif strTab = "2" Then %>
													<!--#include file="include/cargotabinland.asp"-->
												<% end if %>
											</td>
										</tr>
									</table>
								</td>
							</tr>

						</table>
					</td>
					<td valign=top>
						<table cellpadding=0 cellspacing=0 border=0>
							<tr>
								<td bgcolor=#FFFFFF height=31></td>
							</tr>
						</table>
					</td>
					<td width=12 bgcolor=#FFFFFF valign=top></td>
				</tr>
				<%'BOTTOM BAR - Page number and Prev/Next links %>

				<tr>
					<td bgcolor=#FFFFFF colspan=4></td>
				</tr>
</table>
</BODY>
</HTML>
<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>
<% CheckError ("Error found at end of page.") %>
