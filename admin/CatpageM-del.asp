<%@ Language=VBScript %>
<!--#include file="database.asp"-->
<!--#include file="data.inc"-->
<%id=int(Request.QueryString("id"))
set rs=server.CreateObject("adodb.recordset")
	
	RS.Open "catcontent",conn,1,3
    RS.MoveFirst
	  do while not rs.EOF 
	     if rs("id")=id then
	     img1=rs("img1")
	     fp=server.MapPath("../images/product/")  & "/" & img1
	     'fp="../images/product/" & img1
	     Response.Write fp
	     set fso=server.CreateObject("Scripting.FileSystemObject")
	     if fso.FileExists(fp) then
	      fso.DeleteFile fp,true
	      Response.Write img1 & "此檔案已刪除"
	     else 
	     Response.Write "無圖片"
	     end if
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

Response.Write "此段文字已刪除"

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





