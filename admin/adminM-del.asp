<%@ Language=VBScript %>

<!--#include file="database.asp"-->
<!--#include file="data.inc"-->
<%
autoNo=Request.QueryString("autoNo")
set rs=server.CreateObject("adodb.recordset")
	
	autoNo=int(autoNo)
	RS.Open "admin",conn,1,3
    RS.MoveFirst
	  do while not rs.EOF 
	     if rs("autoNo")=autoNo then
	     Application.Lock 
	        rs.delete
	     Application.UnLock 
	     end if
	     rs.movenext
	  loop
	
rs.Close 
set rs=nothing

conn.close
set conn=nothing	

Response.Redirect "adminM.asp"

%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Content-Type" content="text/html; charset=big5">

</HEAD>
<BODY>

<P>&nbsp;</P>

</BODY>
</HTML>
