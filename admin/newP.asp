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
		Response.Write "�z�N��D���O"	
		Response.Write "<font color=red>" & catName & "</font>"
		Response.Write "<INPUT type='hidden' name=CatNo value=" & CatNo & ">" 		Response.Write ""
		Response.Write "���U��"	
		Response.Write "<SELECT id=select2 name=subcatNo>"
		Response.Write "<OPTION>�п�ܦ����O</OPTION>" & op2		Response.Write "</SELECT>"
		Response.Write "���s�W���~<br>"
		
		Response.Write "���~�W��<INPUT type='text' id=text2 name=NAME><BR>"
		Response.Write "Item No.<INPUT type='text' id=text2 name=itemNo><BR>"
		Response.Write "Model No.<INPUT type='text' id=text2 name=ModelNo><BR>"
		'Response.Write "<INPUT type='file' id=file1 name=file1><BR>"		Response.Write "���~���e�y�z<TEXTAREA rows=10 cols=100 id=textarea1 name=content1>"
		Response.Write "</TEXTAREA>"
		
		
		
		
		Response.Write "<INPUT type='submit' value='�U�@�B' id=submit1 name=submit1>"
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

Response.Write "�z�ҿ�ܤ�<font color=red>" & catName  & "</font>�D���O���U�L�����O,�Х��s�W�����O"
end if
end sub




%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Content-Type" content="text/html; charset=big5">

</HEAD>
<BODY>
�s�W���~<br>
�п�ܥD���O
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
<OPTION>�п�ܥD���O</OPTION>
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
