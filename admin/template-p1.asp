<%@ Language=VBScript %>
<!--#include file="database.asp"-->
<!--#include file="data.inc"-->
<%
template=Request.Form("template")

set rs=server.CreateObject("adodb.recordset")
RS.Open "template",conn,1,3
rs.MoveFirst 
rs("template")=template
rs.Update 
Response.Write "網站樣式選擇(或修改)完成<br>"     
'Response.Write "<A HREF='headline-p0.asp'>headline</A>"
rs.Close 
set rs=nothing
conn.close
set conn=nothing   
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>



</BODY>
</HTML>
