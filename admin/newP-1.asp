<%@ Language=VBScript %>
<!--#include file="database.asp"-->
<!--#include file="data.inc"-->
<%
CatNo=Request.Form("CatNo")
subcatNo=Request.Form("subcatNo")
NAME=Request.Form("NAME")
itemNo=Request.Form("itemNo")
ModelNo=Request.Form("ModelNo")
content1=Request.Form("content1")
'Response.Write "CatNo:" & CatNo &"<br>"
'Response.Write "subcatNo:" & subcatNo &"<br>"
'Response.Write "NAME:" & NAME &"<br>"
'Response.Write "itemNo:" & itemNo &"<br>"
'Response.Write "ModelNo:" & ModelNo &"<br>"

set rs=server.CreateObject("adodb.recordset")
	rs.open "product",conn,1,3	
	
	''在 Table 中加入新資料
	
	Application.Lock 
	rs.addnew
	'資料'
	rs("cat").value=CatNo '----1
	rs("subcatNo").value=subcatNo '----2
	rs("NAME").value=NAME '----3
	rs("itemNo").value=itemNo '----4
	rs("ModelNo").value=ModelNo '----5
	rs("content1").value=content1 '----6
	
	rs.update
	Application.UnLock
	id=rs("id")
    rs.Close 
    set rs=nothing
    conn.Close 
   set conn=nothing

Response.Write "文字部份上載成功"
Response.Write "<A HREF='newP-2.asp?id=" & id & "'>繼續上載圖片</A>"


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
