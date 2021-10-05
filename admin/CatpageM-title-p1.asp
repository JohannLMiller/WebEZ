<%@ Language=VBScript %>
<!--#include file="database.asp"-->
<!--#include file="data.inc"-->
<%
catNo=Cint(Request.Form("catNo"))
title=Request.Form("title")
set rs=server.CreateObject("adodb.recordset")
RS.Open "cat",conn,1,3
  rs.MoveFirst
do while not rs.EOF  
	if rs("catNo")=catNo then
	rs("title")=title
	rs.Update 
	end if
	rs.MoveNext 
loop
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
