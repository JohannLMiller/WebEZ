<%@ Language=VBScript %>
<HTML>
<HEAD>
<TITLE>修改圖文</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=big5">

</HEAD>
<BODY>
<% Set Upload = Server.CreateObject("Persits.Upload.1") 
'Count = Upload.Save("C:\abc")
Count = Upload.SaveVirtual("../images/product")
%> 
<!--#include file="database.asp"-->
<!--#include file="data.inc"-->
<%
set rs=server.CreateObject("adodb.recordset")
 RS.Open "catcontent",conn,1,3
  rs.MoveFirst
 do while not rs.EOF 
    if rs("id")=int(upload.Form("id")) then
    Response.Write "filename:" & filename & "<br>"
    
			For Each File in Upload.Files
			  filename=File.ExtractFileName
			  
			  if filename<>"" then
			  rs("img1")=filename
			  end if
			next 
			  rs("title")=upload.Form("title")
			  rs("content1")=upload.Form("content1")
			  rs.Update
			
			
    end if
    rs.MoveNext
 loop  

  rs.Close 
    set rs=nothing
    conn.Close 
   set conn=nothing
%>

<%if Err <> 0 Then %>
 "錯誤發生"

<%Else 
	Response.Write "成功"
 End If %>

</BODY>
</HTML>







