<%@ Language=VBScript %>
<!--#include file="database.asp"-->
<!--#include file="data.inc"-->
<% 

catNo=Request.Form("catNo")
layout=Request.Form("layout")
Response.Write catNo
Response.Write layout

set rs=server.CreateObject("adodb.recordset")
 RS.Open "catcontent",conn,1,3
  rs.MoveFirst
 
 do while not rs.EOF 
	if rs("catNo")=catNo then
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
