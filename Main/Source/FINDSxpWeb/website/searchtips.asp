<%@ language=vbscript %>
<% On Error Resume Next %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
	<TITLE>RF International - Login Screen</TITLE>
<link rel="stylesheet" type="text/css" href="rfi.css" />
</HEAD>
  <BODY text=#ffffff bgColor=#ffffff >

<!--#include file="include/checkuserlogin.asp"-->
<!--#include file="include/webfunctions.asp"-->

<!--#include file="include/header.html"-->
<TABLE height="72%" cellSpacing=0 cellPadding=0 width="100%" border=0 align="center">



	<TR vAlign=center align=middle>
		<TD>
			<TABLE cellSpacing=0 cellPadding=0 border=0>
				<tr>
					<TD vAlign=top align=center colSpan=2 rowSpan=2>
                                        </td>
				</TR>
				<TR>
				</TR>
				<TR>
					<td width=185>
					</td>
					<TD valign=top rowSpan=3 align=center height=160>
						<FORM name=login action='start.asp' method=post>
							<TABLE cellSpacing=2 cellPadding=2 width=319 border=0 name="form">
								<TR vAlign=center>
									<td>
<h3>The Basics of Ship-at-a-Glance Search</h3>
        <p> <a name="basic"><b>Basic Search</b> </a>
		<p><b>(UNDER CONTRUCTION)</b>
									</td>
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
<% CheckError ("Error found at end of page.") %>