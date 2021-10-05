<%@ Language=VBScript %>
<!--#include file="database.asp"-->
<!--#include file="data.inc"-->
<%
catNo=Cint(Request.QueryString("catNo"))
set rs=server.CreateObject("adodb.recordset")
SQLstr="select * from cat where catNo=" & catNo
set rs=conn.execute(SQLStr)


%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<form name="form1" method="post" action="CatpageM-title-p1.asp">
   <INPUT type="hidden" id=text1 name=catNo value="<%=catNo%>">
  <input type="text" name="title" size="50" value="<%=rs("title")%>">
  <br>
  <input type="submit" name="Submit" value="Submit">
  <input type="reset" name="Submit2" value="Reset">
</form>
<P>&nbsp;</P>

</BODY>
</HTML>
