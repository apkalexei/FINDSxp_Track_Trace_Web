<%
	CheckError ("Error found at beginning of include file cargotabcommodity.asp.")
	
	Set objConn = Server.CreateObject("ADODB.Connection")
	Set rs = Server.CreateObject("ADODB.RecordSet")
	objConn.Open Application("DB_CONNECT_STRING")
	rs.Open "fwpOceanCommodityDetail '" & request("cargoindex") & "', '" & session("UserIndex") & "'", objConn
	CheckError ("Calling fwpOceanCommodityDetail '" & request("cargoindex") & "'")
	
	Dim strNetLBS, strNetKGS, strGrossLBS, strGrossKGS, strCFT, strCBM, strPieceCount, strPieceTypeDesc
	
	' Fill the piece type rs with the piece types from the system
	Dim rs3 
	Dim aryPieceTypes
	Dim strTempPT
	Dim j
	
	Dim strCmdyIndex
	Dim strUserCmdyID
	
	blnReadOnly = rs("fReadOnly")
	
	set rs3 = Server.CreateObject("ADODB.Recordset") 
	rs3.Open "fwpPieceTypeLookup", objConn
checkerror("2")
%>

<table bgcolor=#c0c0c0 border=0 cellspacing=0 cellpadding=0>	
	<tr>
		<td colspan=2>
			<center>
			<table bgcolor=#FFFFFF border=0 cellpadding=0 cellspacing=0>
			<%	
				j = 0
				Do While not rs.Eof 
					if request("err") <> "" AND trim(rs("CmdyIndex")) = request("CmdyIndex") Then
						' error AND this was the commodity that was modified - get data from URL (data that the user entered)
						strNetLBS = request("NetLBS")
						strNetKGS = request("NetKGS")
						strGrossLBS = request("GrossLBS")
						strGrossKGS = request("GrossKGS")
						strCFT = request("CFT")
						strCBM = request("CBM")
						strPieceCount = request("PieceCount")
						strPieceTypeDesc = request("PieceTypeDesc")
						strUserCmdyID = request("UserCmdyID")
						strCmdyIndex = request("CmdyIndex")
					else
						'no errors - get data from database
						strNetLBS = rs("NetLBS")
						strNetKGS = rs("NetKGS")
						strGrossLBS = rs("GrossLBS")
						strGrossKGS = rs("GrossKGS")
						strCFT = rs("CFT")
						strCBM = rs("CBM")
						strPieceCount = rs("PieceCount")
						strPieceTypeDesc = rs("PieceTypeDesc")
						strUserCmdyID = rs("UserCmdyID")
						strCmdyIndex = rs("CmdyIndex")
					end if 
											
					if NOT blnReadOnly Then Response.Write("<form action='xt_commoditychange.asp' id=form" & j & " name=form" & j & " method=get>")
				%>
						<input type=hidden name=bb value=<%= request("bb") %>>
						<input type=hidden name=commodityid value='<%= strUserCmdyID %>'>
						<input type=hidden name=CmdyIndex value='<%= strCmdyIndex %>'>
						<input type=hidden name=cargoindex value='<%= request("cargoindex") %>'>
						<tr>
							<td valign=center bgcolor=#b0b0b0 colspan=3 align=left>
								<table cellspacing=0 cellpadding=0 border=0>
									<tr>
										<td width=300 height=25>
											<font size=4><%= "&nbsp;&nbsp;Commodity ID:&nbsp;" & rs("UserCmdyID") %></font>
										</td>
										<td colspan=2 bgcolor=#b0b0b0>
											<% if Not blnReadOnly Then Response.Write("<center><a href=""javascript: document.form" & j & ".submit()""><image src='pictures/save.gif' border=0></center></a>") %>
										</td>
									</tr>
								</table>
							</td>
						</tr>
						<tr>
							<td valign=top height=15 colspan=5>
							</td>
						</tr>
						<tr>
							<td width=15>
							</td>
							<td>
								<b><font size=2>Commodity Description</font></b>
								<table border=1 bgcolor=#d0d0d0>
									<tr>
										<td><font face='tahoma, arial, helvetica' size=2>&nbsp;<%= rs("CmdyDesc") %>&nbsp;</font></td>
									</tr>
								</table>
								<b><font size=2>Weights and Measures</font></b>
								<table border=1 bgcolor=#d0d0d0>
									<tr>
										<td><b><font size=2>LBS (Net)</font></b></td>
										<td><b><font size=2>KGS (Net)</font></b></td>
										<td><b><font size=2>LBS (Gross)</font></b></td>
										<td><b><font size=2>KGS (Gross)</font></b></td>
										<td><b><font size=2>CFT</font></b></td>
										<td><b><font size=2>CBM</font></b></td>
										<td><b><font size=2>Piece Count</font></b></td>
										<td><b><font size=2>Piece Type</font></b></td>
									</tr>		
									<tr>
										<% checkerror("3.5")
											if blnReadOnly Then  %>
											<td><font face='tahoma, arial, helvetica' size=2>&nbsp;<%= strNetLBS %>&nbsp;</font></td>
											<td><font face='tahoma, arial, helvetica' size=2>&nbsp;<%= strNetKGS %>&nbsp;</font></td>
											<td><font face='tahoma, arial, helvetica' size=2>&nbsp;<%= strGrossLBS %>&nbsp;</font></td>
											<td><font face='tahoma, arial, helvetica' size=2>&nbsp;<%= strGrossKGS %>&nbsp;</font></td>
											<td><font face='tahoma, arial, helvetica' size=2>&nbsp;<%= strCFT %>&nbsp;</font></td>
											<td><font face='tahoma, arial, helvetica' size=2>&nbsp;<%= strCBM %>&nbsp;</font></td>
											<td><font face='tahoma, arial, helvetica' size=2>&nbsp;<%= strPieceCount %>&nbsp;</font></td>
											<td><font face='tahoma, arial, helvetica' size=2>&nbsp;<%= strPieceTypeDesc %>&nbsp;</font></td>
										<% else %>
											<td><input name='NetLBS' size=9 maxlength=13 value='<%= strNetLBS %>'></input></td>
											<td><input name='NetKGS' size=9 maxlength=13 value='<%= strNetKGS %>'></input></td>
											<td><input name='GrossLBS' size=9 maxlength=13 value='<%= strGrossLBS %>'></input></td>
											<td><input name='GrossKGS' size=9 maxlength=13 value='<%= strGrossKGS %>'></input></td>
											<td><input name='CFT' size=9 maxlength=13 value='<%= strCFT %>'></input></td>
											<td><input name='CBM' size=9 maxlength=13 value='<%= strCBM %>'></input></td>
											<td><input name='PieceCount' size=5 maxlength=6 value='<%= strPieceCount %>'></input></td>
											<td>
												<select name=PieceTypeDesc>
													<% 
													   ' Fill the piece type dropdown
													   rs3.MoveFirst
													   Do While Not rs3.Eof
															if trim(strPieceTypeDesc) = trim(rs3("ID")) Then 
																strTempPT = " SELECTED"
															else 
																strTempPT = ""
															end if
															Response.Write("<option value='" & rs3("ID") & "'" & strTempPT & ">" & rs3("ID") & "</option>")
															rs3.MoveNext
													   Loop
													%>
												</select>
											</td>
										<% end if %>
									</tr>										
								</table>
								<b><font size=2>Hazardous Info</font></b>
								<table border=1 bgcolor=#d0d0d0>
									<tr>
										<td><b><font size=2>UNN #</font></b></td>
										<td><b><font size=2>IMDG Page</font></b></td>
										<td><b><font size=2>Class</font></b></td>
										<td><b><font size=2>Package Group</font></b></td>
										<td><b><font size=2>Hazardous Label</font></b></td>
										<td><b><font size=2>Flashpoint</font></b></td>
										<td><b><font size=2>DOT Code</font></b></td>
									</tr>
									<tr>
										<td><font face='tahoma, arial, helvetica' size=2>&nbsp;<%= rs("UNNUmber") %>&nbsp;</font></td><td><font face='arial, helvetica' size=1>&nbsp;<%= rs("IMDGPage") %>&nbsp;</font></td><td><font face='arial, helvetica' size=1>&nbsp;<%= rs("Class") %>&nbsp;</font></td><td><font face='arial, helvetica' size=1>&nbsp;<%= rs("PkgGroup") %>&nbsp;</font></td><td><font face='arial, helvetica' size=1>&nbsp;<%= rs("HazLabel") %>&nbsp;</font></td><td><font face='arial, helvetica' size=1>&nbsp;<% If isNull(rs("Flashpoint")) Then Response.Write(" ") ELSE Response.write(replace(rs("Flashpoint"), "degrees", "º")) End If %>&nbsp;</font></td><td><font face='arial, helvetica' size=1>&nbsp;<%= rs("DOTCode") %>&nbsp;</font></td>
									</tr>									
								</table>
								
							</td>
							<td width=15>
							</td>
						</tr>
						<tr>
							<td valign=top height=15 colspan=5>
							</td>
						</tr>
					<%	if NOT blnReadOnly Then Response.Write("</form>")
					 
						rs.MoveNext
						checkerror("4")
						j = j + 1
				Loop %>
			</table>
			</center>
		</td>
	</tr>
</table>
<% 
	rs3.Close
	set rs3 = Nothing
	rs.close
	objConn.Close
	set rs = nothing
	set objconn = nothing
	CheckError ("Error found at end of include file cargotabcommodity.asp.") %>