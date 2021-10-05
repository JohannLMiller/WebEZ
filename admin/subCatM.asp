<%@ Language=VBScript %>
<!--#include file="database.asp"-->
<!--#include file="data.inc"-->
<%

sub showsubcat()
set rs=server.CreateObject("adodb.recordset")
SQLstr="select * from subcat where catNo='" & catNo & "'"
set rs=conn.Execute (SQLstr)
if not rs.eof then

set rs1=server.CreateObject("adodb.recordset")
SQLstr="select * from cat where catNo=" & catNo 
set rs1=conn.Execute (SQLstr)
Response.Write "主類別:" & rs1("cat")
rs1.Close 
set rs1=nothing


Response.Write "<table WIDTH='75%' BORDER='1' CELLSPACING='1' CELLPADDING='1'>"
Response.Write " <tr> "
Response.Write "   <td>系統識別編號</td>"
Response.Write "   <td>次類別名稱</td>"
Response.Write "   <td>修改</td>"
Response.Write "  <td>刪除</td>"
Response.Write " </tr>"
  do while not rs.EOF 
Response.Write "  <tr> "
Response.Write "<td>" & rs("subcatNo") & "</td>"
Response.Write "<td>" & rs("subcat") & "</td>"
Response.Write "<td><a HREF=subcatM-edit.asp?subcatNo=" & rs("subcatNo") & "&amp;subcat=" & rs("subcat")& "&amp;catNo=" & rs("catNo") & ">修改</a></td>"
Response.Write "<td><a HREF=subcatM-del.asp?subcatNo=" & rs("subcatNo")& "&amp;catNo=" & rs("catNo") & ">刪除</a></td>"
Response.Write "</tr>"
  
	rs.MoveNext 
	loop
Response.Write "</table>"
Response.Write "<p>"
Response.Write "<p>"
	
	

Response.Write "<form name='form1' method='post' action='subcatM-add.asp'>"
Response.Write "  <table width='80%' border='1'>"
Response.Write "    <tr> "
Response.Write "      <td>次類別名稱</td>"
Response.Write "      <td>"
Response.Write "<input type='text' name='subcat'><input type='hidden'name='catNo' value=" & catNo & ">"
Response.Write "      </td>"
Response.Write "      <td> "
Response.Write "        <input type='submit' name='Submit' value='新增'>"
Response.Write "      </td>"
Response.Write "    </tr>"
Response.Write "  </table>"
Response.Write "</form>"



else
Response.Write "此類別下無次類別"
Response.Write "<p>"
Response.Write "<p>"

set rs1=server.CreateObject("adodb.recordset")
SQLstr="select * from cat where catNo=" & catNo 
set rs1=conn.Execute (SQLstr)
Response.Write "主類別:" & rs1("cat")
rs1.Close 
set rs1=nothing
	
	

Response.Write "<form name='form1' method='post' action='subcatM-add.asp'>"
Response.Write "  <table width='80%' border='1'>"
Response.Write "    <tr> "
Response.Write "      <td>次類別名稱</td>"
Response.Write "      <td>"
Response.Write "<input type='text' name='subcat'><input type='hidden'name='catNo' value=" & catNo & ">"
Response.Write "      </td>"
Response.Write "      <td> "
Response.Write "        <input type='submit' name='Submit' value='新增'>"
Response.Write "      </td>"
Response.Write "    </tr>"
Response.Write "  </table>"
Response.Write "</form>"
end if
end sub


%>
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=JavaScript>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Content-Type" content="text/html; charset=big5">

</HEAD>
<BODY>
次類別管理<br>

<%
set rs=server.CreateObject("adodb.recordset")
SQLstr="select * from cat"
set rs=conn.Execute (SQLstr)
do while not rs.EOF
VP=rs("catNo")
op= op+"<option" & " " & "value=" & VP & ">" & rs("cat") & "</option>"
rs.MoveNext 
loop
rs.Close
set rs=nothing 

%>
<FORM action="subcatM.asp" method=POST id=form1 name=form1>



<INPUT type="hidden" id=text2 name=cat value="<%=op%>">
<INPUT type="hidden" id=text1 name=send value=send>
</FORM>
<p>&nbsp;</p>
<%
if Request.QueryString("catNo")<>"" then
catNo=Request.QueryString("catNo")
call showsubcat()
else
	flag=Request.Form("send")
	catNo=Request.Form("cat")
	if flag<>"" then
	call showsubcat()
	end if
end if
%>
</BODY>
</HTML>
