<%@ Language=VBScript %>
<!--#include file="database.asp"-->
<!--#include file="data.inc"-->
<%
sub showsubcat()
set rs=server.CreateObject("adodb.recordset")
SQLstr="select * from subcat where catNo='" & catNo & "'"
set rs=conn.Execute (SQLstr)
if not rs.eof then
			do while not rs.EOF
			VP=rs("subcatNo")
			op2= op2+"<option" & " " & "value=" & VP & ">" & rs("subcat") & "</option>"
			rs.MoveNext 
			loop
			rs.Close
			set rs=nothing 
		set rs1=server.CreateObject("adodb.recordset")
		SQLstr="select * from cat where catNo=" & catNo 
		set rs1=conn.Execute (SQLstr)
		catName=rs1("cat")
		CatNo=rs1("catNo")
		rs1.Close 
		set rs1=nothing

		Response.Write "<FORM action='newP-1.asp' method=POST id=form2 name=form2>"		
		Response.Write "您將於主類別"	
		Response.Write "<font color=red>" & catName & "</font>"
		Response.Write "<INPUT type='hidden' name=CatNo value=" & CatNo & ">" 		Response.Write ""
		Response.Write "之下的"	
		Response.Write "<SELECT id=select2 name=subcatNo>"
		Response.Write "<OPTION>請選擇次類別</OPTION>" & op2		Response.Write "</SELECT>"
		Response.Write "內新增產品<br>"
		
		Response.Write "產品名稱<INPUT type='text' id=text2 name=NAME><BR>"
		Response.Write "Item No.<INPUT type='text' id=text2 name=itemNo><BR>"
		Response.Write "Model No.<INPUT type='text' id=text2 name=ModelNo><BR>"
		'Response.Write "<INPUT type='file' id=file1 name=file1><BR>"		Response.Write "產品內容描述<TEXTAREA rows=10 cols=100 id=textarea1 name=content1>"
		Response.Write "</TEXTAREA>"
		
		
		
		
		Response.Write "<INPUT type='submit' value='下一步' id=submit1 name=submit1>"
		Response.Write "<INPUT type='reset' value='Reset' id=reset1 name=reset1>"
		Response.Write "</FORM>"	
		
else
set rs1=server.CreateObject("adodb.recordset")
		SQLstr="select * from cat where catNo=" & catNo 
		set rs1=conn.Execute (SQLstr)
		catName=rs1("cat")
		CatNo=rs1("catNo")
		rs1.Close 
		set rs1=nothing

Response.Write "您所選擇之<font color=red>" & catName  & "</font>主類別之下無次類別,請先新增次類別"
end if
end sub




%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Content-Type" content="text/html; charset=big5">

</HEAD>
<BODY>
新增產品<br>
請選擇主類別
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
<FORM action="newP.asp" method=POST id=form1 name=form1>

<SELECT id=select1 name=cat onchange="submit()">
<OPTION>請選擇主類別</OPTION>
<%=op%> </SELECT>
<INPUT type="hidden" id=text1 name=send value=send>
</FORM>
<p>&nbsp;</p>
<%
flag=Request.Form("send")
	catNo=Request.Form("cat")
	if flag<>"" then
	call showsubcat()
	end if
%>
</BODY>
</HTML>
