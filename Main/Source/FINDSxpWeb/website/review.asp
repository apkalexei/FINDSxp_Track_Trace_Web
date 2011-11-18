<%@Language="VBScript" %>
<% Option Explicit
   On Error Resume Next

	' Used because this page name will change - update any links referencing 'me' with this const
	Const strThisPageName = "review.asp"

	Session("searchpage") = "review.asp?" & Request.QueryString
%>
<!--#include file="include/checkuserlogin.asp"-->
<!--#include file="include/webfunctions.asp"-->
<%
	CheckError ("Error at the beginning of the page.")
%>

<!--#include file="include/shipatglancefunctions.asp"-->

<%
	' ************************* VARIABLE DECLARATIONS *************************

	' Number of rows to fetch from the database for each screen
	Const intRecordsToShow = 10

	' For each page, this constant defines how many search items to display
	Const BOOKING_NUM_PARAMS = 2
	Const BOL_NUM_PARAMS = 2
	Const RATES_NUM_PARAMS = 8

	' Enumeration for the column array indexes
	Const COLUMN_ARRAY_HEADING = 0
	Const COLUMN_ARRAY_RS_HEADING = 1
	Const COLUMN_ARRAY_TYPE = 2
	'Const COLUMN_ARRAY_IS_HIDDEN = 3
	'Const COLUMN_ARRAY_IS_MULTILINE = 4

	' Enumeration for the column data types (for the COLUMN_ARRAY_TYPE array index)
	Const CT_STANDARD = 2
	Const CT_DATE = 4
	Const CT_NUMBER = 8
	Const CT_FLAG = 16
	Const CT_INCLAUSE = 32
	Const CT_MULTILINE = 64
	Const CT_PRIMARYKEY = 128
	Const CT_HIDDEN = 256

	' Constants for building the SQL query
	Const WHERE_EQUAL = " = "
	Const WHERE_LIKE = " LIKE "
	Const WILDCARD = "%"
	Const SQT = "'"
	Const BOL_FLAG = 1
	Const BOK_FLAG = 2

	' Array storing all column info (heading, rs heading, data type, ishidden, and ismultiline - see constants)
	Dim aryColumnHeadings()
	' Array for storing each search column dropdown
	Dim clbLUField()
	' Array for storing each search column text value
	Dim txtField()
	' Array for storing the 'and' / 'or' operator fields
	Dim clbOperatorField()
	' Array for tab names
	Dim aryTabName()

	' Booking or BOL - passed to this page from depending on what user selected on previous page
	Dim MnTransType
	' Current tab that the user is on
	Dim strTab

	' Date search variables (from user query)
	Dim strFromDate, strToDate, strDateType, strDateRangeOn

	' Temporary storage for all user entered data - used for generating links that won't lose what the user entered
	Dim strUserQueryInfo
	' Which sp to be used for the current tab
	Dim strStoredProcedureToUse

	Dim i, j
	Dim intTemp

	Dim rs, objConn

	' The page number -- Start @ page 1, NEXT-> page 2
	Dim intCurrentPage

	' Temp storage for the SQL Query based on user entered info
	Dim strSQLQuery

	Dim sSqlConnect
	Dim nNumParams

	Dim intNumTabs

	Dim strRecordSetPrimaryKey

	' If there are no dates, don't display the date search area
	Dim blnHasDate
	Dim strScreenTitle

	Dim intNumRecords
	Dim intRecordsRemaining
 %>

 <%
	' ************************* PREPARE DATA FOR HTML OUTPUT *************************
 	set objConn = server.createobject("ADODB.Connection")
	set rs = server.createobject("ADODB.RecordSet")

	' Get info from the request object
	MnTransType = getQueryItem("MnTransType", "T")

	' Ensure that the user is allowed to access this transaction type
	if instr(1, session("availtxtypes"), "~" & MnTransType & "~") = 0 Then
		LogEvent "The user tried to access the the transtype '" & MnTransType & "'.  The available transaction types are in this string: '" & session("availtxtypes") & "'."
		Response.Redirect("error.asp")
	End If

	intCurrentPage = getQueryItem("Page", "N")

	strFromDate = getQueryItem("fromdate", "D")
	strToDate = getQueryItem("todate", "D")
	strDateType = getQueryItem("datetype", "A")
	strDateRangeOn = getQueryItem("daterangeon", "A")

	strTab = getQueryItem("tab", "N")

	CheckError ("Error preparing data")

	' Set the num params/screen title based on the transaction type
	if MnTransType="0" then
		nNumParams = BOOKING_NUM_PARAMS
		strScreenTitle = "<b><font size=6>B</font><font size=5>OOKING</FONT></b>"
	elseif MnTransType="1" then
		nNumParams = BOL_NUM_PARAMS
		strScreenTitle = "<b><font size=6>B</font><font size=5>ILL&nbsp;</FONT><font size=5>OF&nbsp;</FONT><font size=6>L</font><font size=5>ADING</FONT></b>"
	elseif MnTransType="2" then
		nNumParams = RATES_NUM_PARAMS
		strScreenTitle = "<b><font size=6>R</font><font size=5>ATE</FONT></b>"
	end if

	' Size the user-entered info arrays
	Redim clbLUField(nNumParams)
	Redim txtField(nNumParams)
	Redim clbOperatorField(nNumParams - 1)

	'Fill the lookup field arrays with user entered data (from request object)
	For i = 0 To nNumParams - 1
		clbLUField(i) = getQueryItem("clbLUField" & i, "A")
		txtField(i) = getQueryItem("txtField" & i, "A")
		if i <> nNumParams - 1 Then clbOperatorField(i) = getQueryItem("clbOperatorField" & i, "A")
	Next

	CheckError ("Error filling arrays")

	' Default the Transaction type (Booking/BOL/Rates)
	If MnTransType = "" then MnTransType = "0"

	'create connection string
	sSqlConnect = Application("DB_CONNECT_STRING")

	'initiate connection to database
    objConn.Open sSqlConnect

    Redim aryTabName(0)

    ' Get the tab names and associated view name
    rs.Open "fwpUserSAAGViewSelect '" & Session("UserRole") & "', " & MnTransType, objConn
    CheckError ("Error opening fwpUserSAAGViewSelect '" & Session("UserRole") & "', ")
    if Not isNull(rs) Then
		if not rs.eof Then
			rs.MoveFirst
			do while not rs.EOF
				if strTab & "" = rs("oTabNo") & "" Then strStoredProcedureToUse = rs("ViewName")
				Redim Preserve aryTabName (UBound(aryTabName) + 1)
				aryTabName(UBound(aryTabName) - 1) = rs("TabName")
				rs.MoveNext
			loop
		end if
    End If
    CheckError ("Error looping through fwpUserSAAGViewSelect rs")
    rs.Close

    ' Select NO records in order to get the column names
    rs.Open "select * from " & strStoredProcedureToUse & " WHERE 0=1", objConn
	CheckError("Calling select * from " & strStoredProcedureToUse & " WHERE 0=1...  It's likely that there was no valid view defined for this user.  This may be because the user tried to access a transaction or a tab that he isn't allowed to see.")

	'create the column defintions - heading, rs heading, data type, ishidden, and ismultiline
	Dim nFirstPos
	Dim numFields

	For i = 0 To rs.Fields.Count - 1
	ReDim Preserve aryColumnHeadings(2, i + 1)
		nFirstPos = InStr(1, rs.Fields(i).Name, "_")
		If nFirstPos > 0 Then
			numFields = numFields + 1
			While nFirstPos > 0
				nFirstPos = InStr(nFirstPos + 1, rs.Fields(i).Name, "_")
				If nFirstPos > 0 Then numFields = numFields + 1
			Wend
		Else
			numFields = 0
		End If

		If numFields > 0 Then
			For j = 1 To numFields
				Select Case Mid(rs.Fields(i).Name, j * 2, 1)
				  Case "d"
				    aryColumnHeadings(COLUMN_ARRAY_TYPE, i) = aryColumnHeadings(COLUMN_ARRAY_TYPE, i) Or CT_DATE
				    blnHasDate = True
				  Case "f"
				    aryColumnHeadings(COLUMN_ARRAY_TYPE, i) = aryColumnHeadings(COLUMN_ARRAY_TYPE, i) Or CT_FLAG
				  Case "n"
				    aryColumnHeadings(COLUMN_ARRAY_TYPE, i) = aryColumnHeadings(COLUMN_ARRAY_TYPE, i) Or CT_NUMBER
				  Case "p"
				    aryColumnHeadings(COLUMN_ARRAY_TYPE, i) = aryColumnHeadings(COLUMN_ARRAY_TYPE, i) Or CT_PRIMARYKEY
				    strRecordSetPrimaryKey = rs.Fields(i).Name
				  Case "h"
					aryColumnHeadings(COLUMN_ARRAY_TYPE, i) = aryColumnHeadings(COLUMN_ARRAY_TYPE, i) Or CT_HIDDEN
				    'aryColumnHeadings(COLUMN_ARRAY_IS_HIDDEN, i) = True
				  Case "i"
				    aryColumnHeadings(COLUMN_ARRAY_TYPE, i) = aryColumnHeadings(COLUMN_ARRAY_TYPE, i) Or CT_INCLAUSE
				  Case "m"
				    aryColumnHeadings(COLUMN_ARRAY_TYPE, i) = aryColumnHeadings(COLUMN_ARRAY_TYPE, i) Or CT_MULTILINE
				    'aryColumnHeadings(COLUMN_ARRAY_IS_MULTILINE, i) = True
				End Select
			Next

		  	aryColumnHeadings(COLUMN_ARRAY_HEADING, i) = Mid(rs.Fields(i).Name, (numFields * 2) + 1)
			numFields = 0
		Else
			' No parameters found, set to standard
			aryColumnHeadings(COLUMN_ARRAY_TYPE, i) = CT_STANDARD
			'aryColumnHeadings(COLUMN_ARRAY_IS_HIDDEN, i) = False
			aryColumnHeadings(COLUMN_ARRAY_HEADING, i) = rs.Fields(i).Name
		End If

		'Fetch the name used to get the recordset data
		aryColumnHeadings(COLUMN_ARRAY_RS_HEADING, i) = rs.Fields(i).Name
	Next

    CheckError ("Error Creating column definitions.")
	rs.Close

	' Don't fetch data if the user didn't search!
	if request("submit1.x") <> "" OR request("page") <> "" Then
		'Create the sql query string using functions robbed from Finds2
		strSQLQuery = "from " & strStoredProcedureToUse & " (NOLOCK) WHERE "
		strSQLQuery = strSQLQuery & fcnSHIPGLANCE_ConstructSQLStatement(nNumParams, MnTransType, strFromDate, strToDate)
		' Open the recordset, for use in the rest of the page
		' Select as few records as we need - page 1->x records, page2->2x records, etc.
		rs.Open "select count(" & strRecordSetPrimaryKey & ") cnt " & strSQLQuery, objConn
		intNumRecords = rs("cnt")
		intRecordsRemaining = intNumRecords - (intRecordsToShow * (intCurrentPage + 1))
		rs.Close
		rs.MaxRecords = intRecordsToShow * (intCurrentPage + 1)

		rs.Open "select * " & strSQLQuery, objConn
	End If

	' This stores info besides the page number for use with hyperlinks, so that we don't lose the user's search criteria if we don't want to
	strUserQueryInfo = ""
	strUserQueryInfo = strUserQueryInfo & "daterangeon=" & strDateRangeOn & "&datetype=" & strDateType & "&fromdate=" & strFromDate & "&todate=" & strToDate & "&tab=" & strTab & "&MnTransType=" & MnTransType
	For i = 0 To nNumParams - 1
		strUserQueryInfo = strUserQueryInfo & "&clbLUField" & i & "=" & clbLUField(i)
		strUserQueryInfo = strUserQueryInfo & "&txtField" & i & "=" & replace(replace(txtField(i), "''", "%27"), """", "&quot;")
		if i <> nNumParams - 1 Then strUserQueryInfo = strUserQueryInfo & "&clbOperatorField" & i & "=" & clbOperatorField(i)
	Next

	CheckError ("Error looping through building query")
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/loose.dtd">
<HTML>
<HEAD>
	<title>RF International - Transaction Search Screen</title>
	<link rel="stylesheet" type="text/css" href="rfi.css" />
</HEAD>

<BODY bgcolor=#FFFFFF>
<!--#include file="include/header.html"-->

<%'Outermost table %>
 <center>
	<!--<img src='pictures/shipataglancetitle.jpg'><br>-->
	<%= strScreenTitle%>&nbsp;</font><b><font size=6>T</font><font size=5>RANSACTION</font><font size=6>&nbsp;S</font><font size=5>EARCH</font></b><br><br>
 </center>
 <table cellpadding=0 cellspacing=0 border=0 bgcolor=#FFFFFF>
				<tr>
					<td valign=top>
						<table cellpadding=0 cellspacing=0 border=0>
							<tr>
								<td bgcolor=#FFFFFF height=31></td>
							</tr>
							<tr>
								<td></td>
							</tr>
						</table>
					</td>

					<td>
						<table bgcolor=#FFFFFF border=0 cellpadding=0 cellspacing=0 bordercolor=red>
						 	<tr>
						  		<td>
						  			<% 'This 2nd table is used for the chunks of the table ie. tabs=Row1, spacer=row2 ,tabledata=Row3, SearchBox/etc=Row4 %>
									<table bgcolor=#FFFFFF cellpadding="0" cellspacing="0" border=0>
										<tr align=left>
											<td align=bottom bgcolor=#FFFFFF>
												<%'Display the tabs %>
												<table cellpadding="0" cellspacing="0" border=0>

													<tr align=left>

													<td>
													<table cellspacing=0 cellpadding=0 border=0 bgcolor=#FFFFFF>
													<tr align=left>
														<% for i = 0 to Ubound(aryTabName) - 1 %>
															<% if strTab & "" = i & "" Then %>
																<td align=left width=118><table cellspacing=0 cellpadding=0 border=0>
																		<tr>
																			<td width=9></td>
																			<td><table cellspacing=0 cellpadding=0 border=1>
																					<tr align=left>
																						<td height=1 bgcolor=#FFFFFF></td>
																					</tr>
																					<tr>
																						<td bgcolor=#FFFFFF height=29 width=100 align=center><b><font size=2 face="Tahoma, Arial, Helvetica, sans-serif"><%= aryTabName(i) %></b></font></td>
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
																						<td bgcolor=#FFFFFF height=22 width=100 align=center><a href='<%= strThisPageName & "?tab=" & i & "&MnTransType=" & MnTransType %>'><b><font size=1 color=blue face="Tahoma, Arial, Helvetica, sans-serif"><%= aryTabName(i) %></font></b></a></td>
																					</tr>
																				</table>
																			</td>
																			<td></td>
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
										<% ' SEARCH FORM %>
										<tr bgcolor=#FFFFFF align=left>
											<td valign=top>
												<table cellpadding=0 cellspacing=0 border=0>
													<form id='searchform' name='searchform' method=get action='<%= strThisPageName %>'>
													<input type=hidden name=anchor value='sagtable'>
													<input type=hidden name=tab value='<%= strTab %>'>
													<input type=hidden name=MnTransType value='<%= MnTransType %>'>
														<tr>
															<td valign=top colspan=2 nowrap>
																<b>Search for values:</b><br>
																<% For i = 0 to nNumParams - 1 %>
																	&nbsp;&nbsp;
																	<select name='clbLUField<%= i %>' onChange="javascript:txtField<%= i %>.value=''">
																		<% For j = 0 to UBound(aryColumnHeadings, 2) - 1
																				' If it's not a date and it's not hidden, then it must be a visible non-date heading
																				'If (aryColumnHeadings(COLUMN_ARRAY_TYPE, j) And CT_DATE <> 0) And (aryColumnHeadings(COLUMN_ARRAY_TYPE, j) And CT_HIDDEN <> 0) Then
																				If ((aryColumnHeadings(COLUMN_ARRAY_TYPE, j) And CT_DATE) = 0) And ((aryColumnHeadings(COLUMN_ARRAY_TYPE, j) And CT_HIDDEN) = 0) Then %>
																					<option value='<%= aryColumnHeadings(COLUMN_ARRAY_RS_HEADING, j) %>'<% if clbLUField(i) = aryColumnHeadings(COLUMN_ARRAY_RS_HEADING, j) Then Response.Write(" SELECTED") %>><%= aryColumnHeadings(0, j)%></option>
																		<%		End IF
																		   Next	%>
																	</select>
																	&nbsp;=&nbsp;&nbsp;<input name='txtField<%= i %>' size=15 value="<%= replace(replace(txtField(i), "''", "'"), """", "&quot;") %>">
																	<% If i <> nNumParams - 1 Then %>
																		<select name='clbOperatorField<%= i %>'>
																			<option value='AND'<% if clbOperatorField(i) = "AND" Then Response.Write(" SELECTED") %>>AND</option>
																			<option value='OR'<% if clbOperatorField(i) = "OR" Then Response.Write(" SELECTED") %>>OR</option>
																		</select>
																	<% End If %>
																	<% If i Mod 2 = 1 Then Response.Write("<br>") %>
																<% Next %>
															</td>
														</tr>
														<tr>
															<td bgcolor=#FFFFFF height=8></td>
														</tr>
														<tr>
														<% 'Date Range %>
															<td width=270 valign=top align=left>
																<% If blnHasDate Then %>
																	<b>Date Range:</b><br>
																	<input type='CheckBox' name=daterangeon value="on" onClick="javascript:if (daterangeon.checked && todate.value == '' && fromdate.value == '') {todate.value='<%= Date() %>'; fromdate.value='<%= Date() - 30 %>'}"<% if strDateRangeOn = "on" Then Response.Write (" Checked")%>>&nbsp;Date Range On:&nbsp;
																	<select name=datetype>
																		<% For i = 0 to UBound(aryColumnHeadings, 2) - 1
																				If ((aryColumnHeadings(COLUMN_ARRAY_TYPE, i) And CT_DATE) <> 0) Then %>
																					<option value='<%= aryColumnHeadings(COLUMN_ARRAY_RS_HEADING, i) %>'<% if strDateType = aryColumnHeadings(COLUMN_ARRAY_RS_HEADING, i) Then Response.Write(" SELECTED") %>><%= aryColumnHeadings(0, i) %></option>
																			<%  End If
																		   Next %>
																	</select>
																	<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;From:&nbsp;<input size=11 name=fromdate value='<%= strFromDate %>'>
																	<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;To:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input size=11 name=todate value='<%= strToDate %>'>
																<% End If %>
															</td>
															<td width=400 valign=top>
																<table cellpadding=0 cellspacing=0 align=left>
																	<tr>
																		<td height=15>
																		</td>
																	</tr>
																	<tr align=left>
																		<td height=30 valign=top>
																			<input border=0 type=image id=submit1 name=submit1 src='pictures/search.gif'>&nbsp;&nbsp;<a href='searchtips.asp'>Helpful Tips</a>
																		</td>
																	</tr>
																<% if request("submit1.x") <> "" OR request("page") <> "" Then %>
																	<tr align=left>
																		<td bgcolor=#d0d0d0 align=center valign=center height=45>
																			<%  if intNumRecords = 0 Then
																					Response.Write("&nbsp;&nbsp;&nbsp;<b>No records were returned.</b>&nbsp;&nbsp;&nbsp;")
																				else %>
																					<b>&nbsp;Displaying Records <%= (intRecordsToShow * (intCurrentPage)) + 1 %>-<% if (intRecordsToShow * (intCurrentPage + 1)) >= intNumRecords Then Response.Write(intNumRecords) Else Response.Write intRecordsToShow * (intCurrentPage + 1) End If %> of <%= intNumRecords %>.&nbsp;</b><br>
																					<b><% 	' Show prev/next links
																							if intCurrentPage <> 0 Then  %>
																								<a href='<%= strThisPageName %>?page=<%= intCurrentPage - 1 %>&<%= strUserQueryInfo %>'>Previous</a>
																						 <% else %>
																								Previous
																						 <% end if
																							Response.Write("&nbsp;|&nbsp;")
																						   ' Don't Show Link if there are too many records
																						   if (intRecordsToShow * (intCurrentPage + 1)) < intNumRecords Then %>
																								<a href='<%= strThisPageName %>?page=<%= intCurrentPage + 1 %>&<%= strUserQueryInfo %>'>Next</a>
																						<% else %>
																								Next
																						<% end if %></td></b>
																			<%	end if %>
																		</tr>
																	<% end if %>
																</table>
															</td>
														</tr>
														<% 'End Date Range %>
													<% 'End Search form %>
													</form>
												</table>
											</td>
										</tr>
										<tr>
											<td height=10></td>
										</tr>
										<%'DISPLAY THE TABLE - COLUMN HEADINGS AND DATA %>
										<tr>
											<td>
												<table bgcolor=#d0d0d0 border=1 cellpadding=0 cellspacing=0 >
													<tr>
														<%	CheckError ("Error before displaying the data table")
															' Show the column headings
															if MnTransType = 0 Then
																Response.Write("<td><font size=2><b>&nbsp;Booking&nbsp;</b></font></td>")
															Elseif MnTransType = 1 Then
																Response.Write("<td><font size=2><b>&nbsp;B/L&nbsp;</b></font></td>")
															Elseif  MnTransType = 2 Then
																Response.Write("<td><font size=2><b>&nbsp;Rates&nbsp;</b></font></td>")
															End If

															For i = 0 to UBound(aryColumnHeadings, 2) - 1
																'If it's not hidden
																if ((aryColumnHeadings(COLUMN_ARRAY_TYPE, i) And CT_HIDDEN) = 0) Then Response.Write("<td nowrap><font size=2><b>&nbsp;" & aryColumnHeadings(0, i) & "&nbsp;</b></font></td>")
															Next
															CheckError ("Error showing column headings")
														%>
													</tr>
													<%  CheckError ("Error Displaying column headings")

														'Display the table
														' Count  if the number of records for use in displaying prev/next link
														intTemp = 0
														if request("submit1.x") <> "" OR request("page") <> "" Then
															if Not rs.EOF Then
																rs.Move((intCurrentPage) * intRecordsToShow)
																While Not rs.eof
																	' Create a row for each record
																	Response.Write("<tr>")
																	Response.Write("<td><FONT FACE='tahoma, arial, helvetica' SIZE=1>&nbsp;<a class='linkStyle' href='detail.asp?tracknum=" & rs(strRecordSetPrimaryKey) & "&MnTransType=" & MnTransType & "'>VIEW</a>&nbsp;</font></td>")
																	For j = 0 to UBound(aryColumnHeadings, 2) - 1
																		' If it's NOT hidden
																		if ((aryColumnHeadings(COLUMN_ARRAY_TYPE, j) And CT_HIDDEN) = 0) Then
																			Response.Write("<td nowrap>")
																			Dim strTemp
																			strTemp = rs(aryColumnHeadings(COLUMN_ARRAY_RS_HEADING, j))
																			if isnull(strTemp) Then strTemp = ""
																			If (aryColumnHeadings(COLUMN_ARRAY_TYPE, j) AND CT_MULTILINE) <> 0 Then strTemp = Replace (strTemp, vbCrLf, "&nbsp;<BR>&nbsp;")
																			Response.Write("<FONT FACE='tahoma, arial, helvetica' SIZE=1>&nbsp;" & strTemp & "&nbsp;</font>")
																			Response.Write("</td>")
																		End If
																	Next
																	Response.Write("</tr>")
																	intTemp = intTemp + 1
																	rs.MoveNext
																Wend
															End If
														End If
														CheckError ("Error Displaying the data table")%>
												</table>
											</td>
										</tr>
										<%'SPACER %>
										<tr>
											<td height=10 bgcolor=#FFFFFF></td>
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
				<tr>
					<td bgcolor=#FFFFFF colspan=4></td>
				</tr>
</table>
<center>
	<a target="_top" href='xt_login.asp?logoff=yesplease'><FONT face="Tahoma, Arial, Helvetica, sans-serif" size=2 color=#003768><b>Logoff</b></font></a>&nbsp;<font color=#003768><b>|</b></font>&nbsp;<a target="_top" href='start.asp'><FONT face="Tahoma, Arial, Helvetica, sans-serif" size=2 color=#003768><b>Main Page</b></font>
</center>
<!--#include file="include/footer.html"-->
</BODY>
</HTML>
<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>
<% CheckError ("Error found at end of page.") %>
