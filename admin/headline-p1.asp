<%@ Language=VBScript %>
<!--#include file="database.asp"-->
<!--#include file="data.inc"-->
<%
headline=Request.Form("headline")
set rs=server.CreateObject("adodb.recordset")
RS.Open "template",conn,1,3
rs.MoveFirst 
rs("headline")=headline
rs.Update 
Response.Write "banner Headline (�έק�)����<br>"     

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
