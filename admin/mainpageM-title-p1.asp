<%@ Language=VBScript %>
<!--#include file="database.asp"-->
<!--#include file="data.inc"-->
<%
title=Request.Form("title")
set rs=server.CreateObject("adodb.recordset")
RS.Open "mainpagetitle",conn,1,3
rs.MoveFirst 
rs("title")=title
rs.Update 
Response.Write "title (或修改)完成<br>"     

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
