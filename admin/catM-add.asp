<%@ Language=VBScript %>

<!--#include file="database.asp"-->
<!--#include file="data.inc"-->
<%
catNo=Request.Form("catNo")
cat=Request.Form("cat")



set rs=server.CreateObject("adodb.recordset")
	rs.open "cat",conn,1,3	
	
	''在 Table 中加入新資料
	
	Application.Lock 
	rs.addnew
	'資料'
	rs("cat").value=cat '----1
	rs.update
	Application.UnLock
	
    rs.Close 
    set rs=nothing
    conn.Close 
   set conn=nothing
  


Response.Redirect "catM.asp"
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
