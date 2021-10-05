<%@ Language=VBScript %>
<!--#include file="data.inc"-->
<%
set rs=server.CreateObject("adodb.recordset")
'SQLStr="select * from mainpage "
rs.open "mainpage",conn,1,3	
'set rs=conn.execute(SQLStr)
Response.Write rs.CursorType 

if not rs.EOF then
	rs.MoveLast 
	layout=rs("layout")
else
	layout=1
end if

%>




<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Content-Type" content="text/html; charset=iso-950">

</HEAD>
<BODY>
請輸入文字及圖片
<%
''以下的form必須配合SAfileup元件使用語法傳送圖片及文字並在下一步驟中寫入資料庫中
''其中有3個元素 title  & file1 & content1
%>
<FORM action="mainPageM-add-p0.asp" ENCTYPE="multipart/form-data" METHOD="post" id=form1 name=form1>
  <TABLE WIDTH=100% BORDER=0 CELLSPACING=1 CELLPADDING=1>
    <TR> 
      <TD align="left" valign="top">Title</TD>
      <TD valign="top" align="left"> 
        <input type="text" name="title" >
      </TD>
    </TR>
    <TR> 
      <TD align="left" valign="top">文字</TD>
      <TD valign="top" align="left"> 
      
        <textarea id="text1" name="content1" cols="100" rows="10" >
		
		</textarea>
      </TD>
    </TR>
    <TR> 
      <TD align="left" valign="top">編修圖片</TD>
      <TD valign="top" align="left"> 
        <INPUT type="file" id=file1 name=file1>
      </TD>
    </TR>
    <TR> 
      <TD colspan="2" align="left" valign="top"> 
        <div align="center"> 
        <INPUT type="hidden" id=text2 name=layout value="<%=layout%>">
          <input type="submit" name="Submit" value="Submit">
        </div>
      </TD>
    </TR>
  </TABLE>


</FORM>
</BODY>
</HTML>
<%
rs.Close 
set rs=nothing
conn.close
set conn=nothing	
%>