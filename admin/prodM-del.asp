<%@ Language=VBScript %>
<!--#include file="database.asp"-->
<!--#include file="data.inc"-->
<%
'subcatNo=int(Request.QueryString("subcatNo"))
catNo=Request.QueryString("catNo")
id=int(Request.QueryString("id"))

set rs=server.CreateObject("adodb.recordset")
	
	RS.Open "product",conn,1,3
    RS.MoveFirst
	  do while not rs.EOF 
	     if rs("id")=id then
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
Response.Write "刪除檔案"
Response.Redirect "prodM.asp?catNo="& catNo & "&flag=send"

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
