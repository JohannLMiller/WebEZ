<%@ Language=VBScript %>
<HTML>
<HEAD>
<TITLE>Upload Logo</TITLE>

<meta http-equiv="Content-Type" content="text/html; charset=big5">

</HEAD>
<BODY>
<% Set Upload = Server.CreateObject("Persits.Upload.1") 
'Count = Upload.Save("C:\abc")

Count = Upload.SaveVirtual("../images/product")%>



<!--#include file="database.asp"-->
<!--#include file="data.inc"-->
<%  
set rs=server.CreateObject("adodb.recordset")
 RS.Open "product",conn,1,3
 
'For Each File in Upload.Files  
'Response.Write File.Name & "= " & File.Path & " (" & File.Size &")<BR>"

 Application.Lock 
	rs.addnew
			For Each File in Upload.Files
			  filename=File.ExtractFileName
			  if filename<>"" then
			  rs("img1")=filename
			  end if
			next 
			  rs("Cat")=upload.Form("catNo")
			  rs("SubCatNo")=upload.Form("subcatNo")
			  rs("title")=upload.Form("title")
              rs("content1")=upload.Form("content1")
              rs("layout")=upload.Form("layout")
              rs.update
	
	
	
	
      
Application.UnLock 
'Next
%>  
<P>  
 

<%
  rs.Close 
    set rs=nothing
    conn.Close 
   set conn=nothing
%>
<%if Err <> 0 Then %>
 "錯誤發生"

<%	
Else 
Response.Write  "檔案成功上傳"
End If
%>

