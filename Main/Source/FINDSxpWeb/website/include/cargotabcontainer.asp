<%
	CheckError ("Error found at beginning of include file cargotabcontainer.asp.")
	
	Set objConn = Server.CreateObject("ADODB.Connection")
	Set rs = Server.CreateObject("ADODB.RecordSet")
	objConn.Open Application("DB_CONNECT_STRING")
	rs.Open "fwpOceanContainerDetail '" & request("cargoindex") & "', '" & session("UserIndex") & "'", objConn
	CheckError ("Calling fwpOceanContainerDetail '" & request("cargoindex") & "'")

	CheckError ("0")
	blnReadOnly = rs("fReadOnly")
	CheckError ("0.1")
	' If everything is allright, then get the data from the database
	function getItem(str_in_String)
		getItem = ""
		
		if Not isNull(str_in_String) Then getItem = str_in_String
	End Function
	
	if request("err") = "" Then
		strSeal(0) = getItem(rs("SealNo1"))
		strSeal(1) = getItem(rs("SealNo2"))
		strSeal(2) = getItem(rs("SealNo3"))
		strSeal(3) = getItem(rs("SealNo4"))
		strSeal(4) = getItem(rs("SealNo5"))
		strContainerNum = rs("ContainerNum")
	' Otherwise, use the user-entered data
	Else
		strSeal(0) = request("seal1")
		strSeal(1) = request("seal2")
		strSeal(2) = request("seal3")
		strSeal(3) = request("seal4")
		strSeal(4) = request("seal5")
		strContainerNum = request("containernum")
	End If
	CheckError ("0.1")
	
	if NOT blnReadOnly Then Response.Write("<form action='xt_cargodetail.asp' id=form1 name=form1 method=get>") %>
		<input type=hidden name=cargoindex value=<%= request("cargoindex") %>>
		<input type=hidden name=bb value=<%= request("bb") %>>
		<table bgcolor=#FFFFFF border=1 cellspacing=0 cellpadding=0>
                	<tr vAlign=center align=left>
				<td valign=center bgcolor=#b0b0b0 colspan=3 align=left>
					<font face="Tahoma, Arial, Helvetica, sans-serif" size=3><b>&nbsp;&nbsp;&nbsp;Container Number: </b></font>	
						<% if blnReadOnly Then %>
							<font face="Tahoma, Arial, Helvetica, sans-serif" size=4><%= strContainerNum %></font>
						<% else %>
							<input name='containernum' size=15 maxlength=11 value="<%= replace(strContainerNum, """", "&quot;") %>"></input>
						<% end if %>
				</td>
			</tr>
                	<tr vAlign=center align=left>
				<td valign=center bgcolor=#b0b0b0 colspan=3 align=left>
					<font face="Tahoma, Arial, Helvetica, sans-serif" size=3><b>&nbsp;&nbsp;&nbsp;Container Type: </b><font size=4><%= rs("ContTypeDesc") %></font>
				</td>
			</tr>
			<tr>
				<td bgcolor=#FFFFF>
					<br>
				</td>
			</tr>
			<tr>
				<td colspan=3>
					<table bgcolor=lightgrey cellspacing=2 cellpadding=2 border=0>
						<tr>
							<td width=11>
							</td>
							<td valign=top align=Left>
								<font face="Tahoma, Arial, Helvetica, sans-serif" size=2><b><u>Seal Number</u></b></font>
								<%' Seal Number table %>
								<table bgcolor=#b0b0b0 border=0 cellspacing=0 cellpadding=0>
									<%	CheckError ("1")
										Dim TempSeal
										For i = 1 to 5 
											' If the seal number is not null, or NOT read only
											if isNull(rs("SealNo" & i)) Then 
												TempSeal = ""
											Else
												TempSeal = rs("SealNo" & i)
											End If
											If TempSeal <> "" OR NOT blnReadOnly Then
									%>
												<tr>
													<td><b><font face="Tahoma, Arial, Helvetica, sans-serif" size=2><%= i %>: </font></b></td><td>
														<% if blnReadOnly Then %>
															<font face="Tahoma, Arial, Helvetica, sans-serif" size=2>&nbsp;<%= TempSeal %>&nbsp;</font>
														<% else %>
															<input size=15 maxlength=15 id=text2 name='seal<%= i %>' value="<%= replace(strSeal(i - 1), """", "&quot;") %>"></input>
														<% end if %>
													</td>
												</tr>
										<%	End If
										Next 
										 CheckError ("2")
										%>
								</table>
							</td>
							<td width=10>
							</td>
							<td valign=top align=Left>
								<font face="Tahoma, Arial, Helvetica, sans-serif" size=2><b><u>Reefer Info<% if rs("fNonOperating") Then Response.Write("<font size=2>&nbsp;(Non-Operating)</font>") %></u></b></font>
								<%' Reefer info table %>
								<table border=0>
									<tr align=left>
										<td><b><font face="Tahoma, Arial, Helvetica, sans-serif" size=2>Min Temp:</font></b></td><td><font color=black face="Tahoma, Arial, Helvetica, sans-serif" size=2>&nbsp;<% if rs("fNonOperating") Or isNull(rs("MinTemp")) Then Response.Write("N/A") Else Response.Write(replace(rs("MinTemp"), "degrees", "º")) %>&nbsp;</font></td>
									</tr>
									<tr align=left>
										<td><b><font face="Tahoma, Arial, Helvetica, sans-serif" size=2>Max Temp:</font></b></td><td><font color=black face="Tahoma, Arial, Helvetica, sans-serif" size=2>&nbsp;<% if rs("fNonOperating") Or isNull(rs("MaxTemp")) Then Response.Write("N/A") Else Response.Write(replace(rs("MaxTemp"), "degrees", "º")) %>&nbsp;</font></td>
									</tr>
								</table>
							</td>
							<td width=10>
							</td>
							<td valign=top align=Left>
								<font face="Tahoma, Arial, Helvetica, sans-serif" size=2><b><u>Out of Gauge Info</u></b></font>
								<%' Out of gauge info table %>
								<table border=0>
									<tr>
										<td><b><font face="Tahoma, Arial, Helvetica, sans-serif" size=2>Length:</font></b></td><td><font color=black face="Tahoma, Arial, Helvetica, sans-serif" size=2>&nbsp;<%= rs("Length") & " " & rs("MeasureUnitID") %>&nbsp;</font></td>
									</tr>
									<tr>
										<td><b><font face="Tahoma, Arial, Helvetica, sans-serif" size=2>Width:</font></b></td><td><font color=black face="Tahoma, Arial, Helvetica, sans-serif" size=2>&nbsp;<%= rs("Width") & " " & rs("MeasureUnitID") %>&nbsp;</font></td>
									</tr>
									<tr>
										<td><b><font face="Tahoma, Arial, Helvetica, sans-serif" size=2>Height:</font></b></td><td><font color=black face="Tahoma, Arial, Helvetica, sans-serif" size=2>&nbsp;<%= rs("Height") & " " & rs("MeasureUnitID") %>&nbsp;</font></td>
									</tr>
									<% if NOT blnReadOnly Then %>
										<tr height=78>
										<td>
										</td>
										
										<td valign=bottom>
											<a href="javascript: document.form1.submit()"><img src='pictures/save.gif' border=0></center></a>
										</td>
										</tr>
									<% else %>
										<td height=10>&nbsp;</td>
									<% end if %>
								</table>
							</td>
							<td width=100>
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
<%	if NOT blnReadOnly Then Response.Write("</form>")
 
	rs.close
	objConn.Close
	set rs = nothing
	set objconn = nothing
	CheckError ("Error found at end of include file cargotabcontainer.asp.") %>