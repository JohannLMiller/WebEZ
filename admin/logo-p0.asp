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
<H2><CENTER>   �W��Logo�Ϥ�</CENTER></H2>

<P>&nbsp;</P>
<%
'---
'--- Note the special form definition tag: ENCTYPE="multipart/form-data"
'---
%>
<FORM ACTION="logo-p1.asp"  METHOD="post" ENCTYPE="multipart/form-data">
<TABLE WIDTH="100%">
<TR>
	<TD ALIGN="right" VALIGN="top">���� </TD>

	<TD ALIGN="left"><INPUT TYPE="file" NAME="FILE1"><BR>
	
	</TD>
</TR>
<TR>
	<TD ALIGN="right">�`�N�ƶ�:</TD>
	<TD ALIGN="left">��ĳ�Ϥ��ؤo�j�p��????????</TD>
</TR>
<TR>
	<TD ALIGN="center"colspan=2><INPUT TYPE="submit" NAME="SUB1" VALUE="Upload File"></TD>
	
</TR>
</TABLE>
</FORM>

</BODY>
</HTML>
