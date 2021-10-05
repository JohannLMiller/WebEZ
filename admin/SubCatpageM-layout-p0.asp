<%@ Language=VBScript %>
<!--#include file="database.asp"-->
<!--#include file="data.inc"-->
<% 

catNo=Request.Form("catNo")
subcatNo=Request.Form("subcatNo")
layout=Request.Form("layout")
Response.Write catNo &"catNo<br>"
Response.Write subcatNo & "subcatNo<br>"
Response.Write layout & "layout<br>"

set rs=server.CreateObject("adodb.recordset")
 RS.Open "product",conn,1,3
  rs.MoveFirst
 
 do while not rs.EOF 
	if rs("cat")=catNo and rs("subcatNo")=subcatNo then
       rs("layout")=layout
      rs.Update 
	end if
    rs.MoveNext
 loop  
     
Response.Write "排版修改完成"
%> 
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Content-Type" content="text/html; charset=big5">

<BODY>


  </BODY>
</HTML>
