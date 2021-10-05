<%@ Language=VBScript %>

<!--#include file="database.asp"-->
<!--#include file="data.inc"-->
<%
ID=Request.Form("ID")
PWD=Request.Form("PWD")
PS=Request.Form("PS")



set rs=server.CreateObject("adodb.recordset")
	rs.open "admin",conn,1,3	
	
	''在 Table 中加入新資料
	
	Application.Lock 
	rs.addnew
	'會員資料'
	rs("ID").value=ID '---1
	rs("PWD").value=PWD '----2
	rs("PS").value=PS '----3
	

	rs.update
	
	Application.UnLock
	
    rs.Close 
    set rs=nothing
    conn.Close 
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
