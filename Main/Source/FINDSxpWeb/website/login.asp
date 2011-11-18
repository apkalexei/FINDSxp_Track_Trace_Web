
<%@ language=vbscript %>
<!--#include file="include/webfunctions.asp"-->
<%
	On Error Resume Next
	Response.Buffer = False

	Dim strUserName
	Dim strError

	strUserName = getQueryItem("username", "A")
	strError = getQueryItem("err", "A")

	Dim i

	Set bc = Server.CreateObject("MSWC.BrowserType")

	if bc.Version >= 4 Then
		' Do nothing
	elseif bc.Browser="Default" Then
		' Do nothing
	else
		Response.Redirect("oldbrowser.asp")
	end if

%>
<script type="text/javascript">

    var gaJsHost = (("https:" == document.location.protocol) ? "https://ssl." : "http://www.");

    document.write(unescape("%3Cscript src='" + gaJsHost + "google-analytics.com/ga.js' type='text/javascript'%3E%3C/script%3E"));

</script>

<script type="text/javascript">

    try {

        var pageTracker = _gat._getTracker("UA-6622798-2");

        pageTracker._trackPageview();

    } catch (err) { }</script>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<TITLE>RF International - Login Screen</TITLE>
	 <link rel="stylesheet" type="text/css" href="rfi.css" />
</HEAD>
<BODY text=#ffffff  bgColor=#ffffff >

<!--#include file="include/header.html"-->
<TABLE height="72%" cellSpacing=0 cellPadding=0 width="100%" border=0 align="center">
	<TR vAlign=center align=middle>
		<TD align=center>
			<TABLE cellSpacing=0 cellPadding=0 border=0>
				<tr>
					<TD width=vAlign=top align=center colSpan=1 rowSpan=1>

                                                <br>
						<b><font size=6>FINDS</font><font size=5>xp<font size=6>&nbsp;U</font><font size=5>SER</font><font size=6>&nbsp;L</font><font size=5>OGIN</font></b>
					</td>
				</tr>
				<tr>
				</tr>
				<TR>
			<% if strError <> "" Then %>
				<TR>
					<td colspan=2 align=center>
                                        <FONT face="Tahoma, Arial, Helvetica, sans-serif" size=2><b><%= strError %></b></font></td>
				</TR>
			<% end if %>
					<TD valign=top align=center height=160 colspan=2>
						<FORM name=login action='xt_login.asp' method=get>
							<TABLE cellSpacing=1 cellPadding=1 width=50 border=0 name="form">
								<TR vAlign=center>
									<TH align=right><FONT face="Tahoma, Arial, Helvetica, sans-serif" size=2>User ID:</FONT></TH>
									<TD><INPUT id=autho title="Enter your User ID" size=15 name=username value=<%= strUserName %>> </TD>
								</TR>
								<TR vAlign=center>
									<TH align=right><FONT face="Tahoma, Arial, Helvetica, sans-serif" size=2>Password:</FONT></TH>
									<TD><INPUT id=password title="Enter your password" type=password size=15 name=password></TD>
								</TR>
							</TABLE>
							<TABLE cellSpacing=1 cellPadding=1 width=50 border=0 name="forma">
								<TR vAlign=center>
									<TD align=right><INPUT title="Start your Pacer session" accessKey=s type=submit value=Login name=submit></TD>
								</TR>
							</TABLE>
						</FORM>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
</TABLE>
  <!--#include file="include/footer.html"-->
</BODY></HTML>

<% CheckError ("Error found at end of page.")
%>
