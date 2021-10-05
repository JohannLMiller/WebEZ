<%@ Language=VBScript %>
<!--#include file="data.inc"-->
<%
set rs=server.CreateObject("adodb.recordset")
	SQLStr="select * from template"
	set rs=conn.execute(SQLStr)
	rs.MoveFirst 
Response.Redirect rs("template")&"/admin-main.asp"



%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>

<P>&nbsp;</P>

</BODY>
</HTML>
<%

rs.Close 
set rs=nothing
conn.close
set conn=nothing	



%>