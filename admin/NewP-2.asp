<%@ LANGUAGE="VBSCRIPT" %>

<% 
id=Request.QueryString("id") 
Response.Write id & "ok"
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Developer Studio">
<META HTTP-EQUIV="Content-Type" content="text/html; charset=iso-950">
<TITLE>Simple Upload Example</TITLE>
</HEAD>
<BODY BGCOLOR="#ffffff">
<H2><CENTER>   上載產品圖片</CENTER></H2>

<P>&nbsp;</P>
<%
'---
'--- Note the special form definition tag: ENCTYPE="multipart/form-data"
'---
%>
<FORM ACTION="NewP-3.asp" ENCTYPE="multipart/form-data" METHOD="post" id=form1 name=form1>
<TABLE WIDTH="100%">
<TR>
	<TD ALIGN="right" VALIGN="top">圖檔Enter Filename:</TD>
<%
'---
'--- Note the use of the TYPE="FILE" specification
'---
%>
	<TD ALIGN="left"><INPUT TYPE="file" NAME="FILE1"><BR>
	<B><I><SMALL>Note: if a button labeled "Browse..." does not appear, then your
	browser does not support File Upload. For Internet Explorer 3.02 users, a
	free add-on is available from Microsoft. Please see the SA-FileUp documentation
	for more information.</SMALL></I></B>
	</TD>
</TR>
<TR>
	<TD ALIGN="right"><INPUT TYPE="hidden" NAME="id" value=<%=id%>></TD>
	<TD ALIGN="left"><INPUT TYPE="submit" NAME="SUB1" VALUE="Upload File"></TD>
</TR>
</TABLE>
</FORM>

</BODY>
</HTML>
