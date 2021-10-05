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

		Response.Write "<FORM action='prodM.asp' method=POST id=form2 name=form2>"		
		Response.Write "您於主類別"	
		Response.Write "<font color=red>" & catName & "</font>"
		Response.Write "<INPUT type='hidden' name=CatNo value=" & CatNo & ">" 		Response.Write ""
		Response.Write "之下的"	
		Response.Write "<br>"	
		Response.Write "<SELECT id=select2 name=subcatNo onchange='submit()'>"
		Response.Write "<OPTION>請選擇次類別</OPTION>" & op2		Response.Write "</SELECT>"
		Response.Write "之內<br>"
		Response.Write "<INPUT type='hidden' id=send2 name=send2 value=send2>"

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
修改/刪除產品<br>
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
<FORM action="prodM.asp" method=POST id=form1 name=form1>
<SELECT id=select1 name=cat onchange="submit()">
<OPTION>請選擇主類別</OPTION>
<%=op%> </SELECT>
<INPUT type="hidden" id=text1 name=send value=send>
</FORM>
<%
flag=Request.Form("send")
	catNo=Request.Form("cat")
	if flag<>"" then
	call showsubcat()
	end if
	
flag2=Request.Form("send2")
			catNo=Request.Form("catNo")
			subcatNo=Request.Form("subcatNo")
			if flag2<>"" then
			'call ShowProd()
			'Response.Write "subcatNo:" & subcatNo
			set rs=server.CreateObject("adodb.recordset")
				SQLStr="select * from product where cat='" & catNo  & "'"& "and subcatNo='" & subcatNo & "'"
				set rs=conn.execute(SQLStr)%>
				
			<table WIDTH="75%" BORDER="1" CELLSPACING="1" CELLPADDING="1">
			  <tr> 
			    <td>系統識別編號</td>
			    <td>產品名稱</td>
			    <td>修改</td>
			    <td>刪除</td>
			  </tr>
			  <%do while not rs.EOF %>
			  <tr> 
			    <td><%=rs("id")%></td>
			    <td><%=rs("name")%></td>
			    <td><a HREF="prodM-edit.asp?catNo=<%=rs("id")%>&amp;cat=<%=rs("id")%>">修改</a></td>
			    <td><a HREF="prodM-del.asp?id=<%=rs("id")%>&amp;catNo=<%=rs("cat")%>">刪除</a></td>
			  </tr>
			  <%
				rs.MoveNext 
				loop
				rs.Close 
				set rs=nothing
					
					%>
				</table>
			<%else
			'Response.Write "K." & "O"
			end if
			%>
</BODY>
</HTML>
