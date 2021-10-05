<%@ LANGUAGE="VBSCRIPT" %>

<% 
'id=Request.QueryString("id") 
'Response.Write id & "ok"
%>
<HTML>
<HEAD>

<TITLE>Simple Upload Example</TITLE>
</HEAD>
<BODY BGCOLOR="#ffffff">
<H2><CENTER>   上載Logo圖片</CENTER></H2>

<P>&nbsp;</P>
<%
'---
'--- Note the special form definition tag: ENCTYPE="multipart/form-data"
'---
%>
<FORM ACTION="logo-p1.asp"  METHOD="post" ENCTYPE="multipart/form-data">
<TABLE WIDTH="100%">
<TR>
	<TD ALIGN="right" VALIGN="top">圖檔 </TD>

	<TD ALIGN="left"><INPUT TYPE="file" NAME="FILE1"><BR>
	
	</TD>
</TR>
<TR>
	<TD ALIGN="right">注意事項:</TD>
	<TD ALIGN="left">建議圖片尺寸大小為????????</TD>
</TR>
<TR>
	<TD ALIGN="center"colspan=2><INPUT TYPE="submit" NAME="SUB1" VALUE="Upload File"></TD>
	
</TR>
</TABLE>
</FORM>

</BODY>
</HTML>
