<%@ LANGUAGE="VBSCRIPT" %>
<HTML>
<HEAD>
<TITLE>Upload Logo</TITLE>
</HEAD>
<BODY>
<% Set Upload = Server.CreateObject("Persits.Upload.1") 
'Count = Upload.Save("C:\abc")

Count = Upload.SaveVirtual("../images/logo")
%> 
<!--#include file="database.asp"-->
<!--#include file="data.inc"-->

<%  
For Each File in Upload.Files  %>

圖檔寬度為  <%=File.ImageWidth%>pix<br>
圖檔高度為  <%=File.ImageHeight%>pix<br>
圖檔名稱為  <%=File.ExtractFileName%> 
<%
set rs=server.CreateObject("adodb.recordset")
 RS.Open "template",conn,1,3
  rs.MoveFirst
 '''''''''''''''
 do while not rs.EOF 
   
    logo=File.ExtractFileName
       leg=len(logo)
       for i=leg to 1 step -1
         if mid(logo,i,1)="\" or mid(logo,i,1)="/" then
            exit for
         else
            y=y+1
         end if
       next
       filename=right(logo,y)
     ''寫入資料庫
       rs("logo")=filename
   '' end if
    rs.MoveNext
 loop  
 '''''''''''''''
  rs.Close 
    set rs=nothing
    conn.Close 
   set conn=nothing
%>
<%Next%> 
<br>
<% = Count %> 個檔案上傳成功!!<br> 
</BODY>
</HTML>








