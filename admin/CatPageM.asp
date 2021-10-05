<%@ Language=VBScript %>

<!--#include file="database.asp"-->
<!--#include file="data.inc"-->
<%
id=cint(Request.QueryString("id"))
Response.Write id
set rs=server.CreateObject("adodb.recordset")
	SQLStr="select * from Catcontent where id=" & id
	set rs=conn.Execute (SQLstr)
	

%>





<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
請先修改文字,再修改圖片
<FORM action="mainPageM-edit.asp" method=POST id=form1 name=form1>
  <TABLE WIDTH=100% BORDER=0 CELLSPACING=1 CELLPADDING=1>
    <TR> 
      <TD>文字</TD>
      <TD> 
      <INPUT type="hidden" id=text2 name=id value=<%=id%>>
        <textarea id="text1" name="content1" cols="70" rows="6" >
		<%=rs("content1")%>
		</textarea>
      </TD>
    </TR>
    <TR> 
      <TD>編修圖片</TD>
      <TD> 
        <INPUT type="checkbox" id=checkbox1 name=checkbox1>
      </TD>
    </TR>
    <TR> 
      <TD colspan="2"> 
        <div align="center">
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
