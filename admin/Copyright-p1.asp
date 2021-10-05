<%@ Language=VBScript %>
<!--#include file="database.asp"-->
<!--#include file="data.inc"-->
<%
copyright=Request.Form("copyright")
set rs=server.CreateObject("adodb.recordset")
RS.Open "template",conn,1,3
rs.MoveFirst 
rs("copyright")=copyright
rs.Update 
Response.Write "copyright(或修改)完成<br>"     

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
