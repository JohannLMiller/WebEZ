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

���ɼe�׬�  <%=File.ImageWidth%>pix<br>
���ɰ��׬�  <%=File.ImageHeight%>pix<br>
���ɦW�٬�  <%=File.ExtractFileName%> 
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
     ''�g�J��Ʈw
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
<% = Count %> ���ɮפW�Ǧ��\!!<br> 
</BODY>
</HTML>








