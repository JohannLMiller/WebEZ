<%@ Language=VBScript %>
<!--#include file="database.asp"-->
<!--#include file="data.inc"-->
<% 

id=Request.Form("id")

layout=Request.Form("layout")

Response.Write id & "<br>"
Response.Write layout
set rs=server.CreateObject("adodb.recordset")
 RS.Open "mainpage",conn,1,3
  rs.MoveFirst
 
 do while not rs.EOF 
		if rs("id")=int(id) then
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
