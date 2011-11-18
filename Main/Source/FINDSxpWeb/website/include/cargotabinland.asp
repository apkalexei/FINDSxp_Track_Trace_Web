<%
	CheckError ("Error found at beginning of include file cargotabinland.asp.")
	
	Dim rs2
	
	Set objConn = Server.CreateObject("ADODB.Connection")
	Set rs2 = Server.CreateObject("ADODB.RecordSet")
	objConn.Open Application("DB_CONNECT_STRING")
%>
	<table bgcolor=#c0c0c0 border=0 cellspacing=0 cellpadding=0>		
							<tr>
								<td colspan=2>
									<center>
									<table bgcolor=#FFFFFF border=0>
										<tr>
											<td valign=top height=15 colspan=5>

												<table>
													<TR>
														<TD align=middle>
															<TABLE bgColor=#d0d0d0 border=1>
															<TBODY>
																<TR>
																	<TD vAlign=top noWrap><FONT size=2><B>&nbsp;Tripleg #&nbsp;</B></FONT></TD>
																	<TD vAlign=top noWrap><FONT size=2><B>&nbsp;Pickup&nbsp;</B></FONT></TD>
																	<TD vAlign=top noWrap><FONT size=2><B>&nbsp;Dropoff&nbsp;</B></FONT></TD>
																	<TD vAlign=top noWrap><FONT size=2><B>&nbsp;Provider&nbsp;</B></FONT></TD>
																	<TD vAlign=top noWrap><FONT size=2><B>&nbsp;Load Time&nbsp;</B></FONT></TD>
																	<TD vAlign=top noWrap><FONT size=2><B>&nbsp;Drop Time&nbsp;</B></FONT></TD>
																</TR>
																<%' Inland Info %>
																<% 
																	Dim strQuery
																	strQuery = "fwpInlandReview '" & request("cargoindex") & "', '" & session("UserIndex") & "'"
																	rs2.Open strQuery, objConn
																	'Response.Write(strQuery)
																	CheckError (strQuery)
																        Do While Not rs2.EOF
																			Response.Write "<TR><TD vAlign=top nowrap><FONT face='tahoma, arial, helvetica' size=2>&nbsp;" & rs2("TripLegNo") & "&nbsp;</FONT></TD>"
																			Response.Write "<TD vAlign=top nowrap><FONT face='tahoma, arial, helvetica' size=2>&nbsp;" & rs2("PickupLocation") & "&nbsp;</FONT></TD>"
																			Response.Write "<TD vAlign=top nowrap><FONT face='tahoma, arial, helvetica' size=2>&nbsp;" & rs2("DropoffLocation") & "&nbsp;</FONT></TD>"
																			Response.Write "<TD vAlign=top nowrap><FONT face='tahoma, arial, helvetica' size=2>&nbsp;" & rs2("ServiceProvider") & "&nbsp;</FONT></TD>"
																			Response.Write "<TD vAlign=top nowrap><FONT face='tahoma, arial, helvetica' size=2>&nbsp;" & rs2("EstLoadDateTime") & "&nbsp;</FONT></TD>"
																			Response.Write "<TD vAlign=top nowrap><FONT face='tahoma, arial, helvetica' size=2>&nbsp;" & rs2("EstDropoffDateTime") & "&nbsp;</FONT></TD></TR>"
																           rs2.MoveNext
																        Loop
																        
																        rs2.Close
																%>
															</TBODY>
															</TABLE>
														</TD>
													</TR>
													<TR>
														<TD height=15></TD>
													</TR>
												</table>

											</td>
										</tr>
										

									</table>
									</center>
								</td>
							</tr>
						</table>
<% 
	objConn.Close
	set objconn = nothing
	CheckError ("Error found at end of include file cargotabinland.asp.") %>