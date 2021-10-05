<%@ LANGUAGE="VBSCRIPT" %>

<% 
'--- Module: FORMRESP.ASP
'---
'--- Simple file upload form processing. 
'---
'--- Copyright (c) 1997, 1998, 1999 Software Artisans, Inc.
'--- Mail: info@softartisans.com   http://www.softartisans.com
%>

<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=iso-950">
<TITLE>Simple File Upload Results</TITLE>
</HEAD>
<BODY>
<%
'---
'--- Instanciate SA-FileUp
'---The following line creates an instance of the SA-FileUp object:
Set upl = Server.CreateObject("SoftArtisans.FileUp")
'---
'--- Set the default path to store uploaded files. 
'---
upl.Path ="d:\64.224.187.199\johnny\images\product"
'upl.Path = "c:\abc\abc2"
'upl.Save 
'upl.Path = " http://www.e-r3.com/johnny/images/product"
'upl.Path = "\\172.30.1.11\wwwroot\new1226\index\image\news"
%>
<% if upl.IsEmpty Then %>
The file that you uploaded was empty. Most likely, you did not specify a valid
filename to your browser or you left the filename field blank. Please try again.
<% ElseIf upl.ContentDisposition <> "form-data" Then %>
Your upload did not succeed, most likely because your browser
does not support Upload via this mechanism.
<br>
For Internet Explorer Users:
<UL>
<LI>For Windows 95 or Windows NT 4.0:
	<UL>
	<LI><A HREF="http://www.microsoft.com/ie/">Download</A> V3.02 or later of Internet Explorer
	<LI><A HREF="http://www.microsoft.com/ie/download">Download</A> the File Upload Add-on
	<LI>For further information, See Knowledge Base Article <A HREF="http://www.microsoft.com/kb/articles/Q165/2/87.htm">Q165287</A>
	</UL>
<LI>For Windows 3.1, WFW 3.11 (Windows 16-bit), or Windows NT 3.51:
	<UL><A HREF="http://www.microsoft.com/ie/">Download</A> V3.02A or later of Internet Explorer for 16-bit Windows
	</UL>
</UL>
For Netscape Users:
<UL>
<LI><A HREF="http://home.netscape.com">Download</A> a version of Netscape Navigator or Communicator of 2.x or later
</UL>
For users of other browsers:
<UL>
<LI>Your browser must support a standard called RFC 1867. Please check with your browser vendor for
support of this standard.
</UL>
<%Else %>
<P>The file was successfully transmitted by the user.</P>
<% 
	on error resume next
	'---
	'--- Save the file now. If you want to preserve the original user's filename, use
	'---
	upl.Save%>
<% End If %>
<!--#include file="database.asp"-->
<!--#include file="data.inc"-->
<%
set rs=server.CreateObject("adodb.recordset")
 RS.Open "product",conn,1,3
  rs.MoveFirst
 
 do while not rs.EOF 
    if rs("id")=int(upl.Form("id")) then
    img1=upl.UserFilename
       leg=len(img1)
       for i=leg to 1 step -1
         if mid(img1,i,1)="\" or mid(img1,i,1)="/" then
            exit for
         else
            y=y+1
         end if
       next
       filename=right(img1,y)
       rs("img1")=filename
    end if
    rs.MoveNext
 loop  
  rs.Close 
    set rs=nothing
    conn.Close 
   set conn=nothing
%>
	<%if Err <> 0 Then %>
<H1><FONT COLOR="#ff0000">An error occurred when saving the file on the server.</FONT></H1>
Possible causes include:
<UL>
  <LI>An incorrect filename was specified
  <LI>File permissions do not allow writing to the specified area
</UL>
Please check the SA-FileUp documentation for more troubleshooting information,
or send e-mail to <A HREF="mailto:info@softartisans.com">info@softartisans.com</A>

<%	Else 
		Response.Write("Upload saved successfully to " & upl.ServerName)
	End If %>
<P>&nbsp;</P>
<FONT SIZE="-1"><CENTER>
<TABLE WIDTH="80%" BORDER="1" CELLSPACING="2" CELLPADDING="0" HEIGHT="206">
<TR>
<TD COLSPAN="2"><P><CENTER>Information About The Uploaded File</CENTER></TD></TR>
<TR>
<TD WIDTH="30%" HEIGHT="27" ALIGN="RIGHT" VALIGN="TOP">&nbsp;User's filename</TD>
<TD WIDTH="70%"><%=upl.UserFilename%>&nbsp;</TD></TR>
<TR>
<TD WIDTH="30%" HEIGHT="27" ALIGN="RIGHT" VALIGN="TOP">Size in bytes&nbsp;</TD>
<TD WIDTH="70%"><%=upl.TotalBytes%>&nbsp;</TD></TR>
<TR>
<TD WIDTH="30%" HEIGHT="27" ALIGN="RIGHT" VALIGN="TOP">Content Type</TD>
<TD WIDTH="70%"><%=upl.ContentType%>&nbsp;</TD></TR>
<TR>
<TD WIDTH="30%" HEIGHT="27" ALIGN="RIGHT" VALIGN="TOP">Content Disposition</TD>
<TD WIDTH="70%"><%=upl.ContentDisposition%>&nbsp;</TD></TR>
<TR>
<TD WIDTH="30%" HEIGHT="27" ALIGN="RIGHT" VALIGN="TOP">MIME Version</TD>
<TD WIDTH="70%"><%=upl.MimeVersion%>&nbsp;</TD></TR>
<TR>
<TD WIDTH="30%" HEIGHT="27" ALIGN="RIGHT" VALIGN="TOP">Content Transfer Encoding</TD>
<TD WIDTH="70%"><%=upl.ContentTransferEncoding%>&nbsp;</TD></TR>
</TABLE>
</FONT></CENTER>

</BODY>
</HTML>
