<%@ Language=VBScript %>

<!--#include file="database.asp"-->
<!--#include file="data.inc"-->
<%
id=cint(Request.QueryString("id"))
Response.Write "Mainpage ��Ʈw�ѧO�N�X" & id & "<br>"
Response.Write "<A HREF=mainpageM-del.asp?id=" & id & ">�R�����q�Ϥ�</A><br>"
set rs=server.CreateObject("adodb.recordset")
	SQLStr="select * from mainpage where id=" & id
	set rs=conn.Execute (SQLstr)
%>



<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Content-Type" content="text/html; charset=iso-950">

</HEAD>
<BODY>
�έק��r�ιϤ� 
<%
''�H�U��form�����t�XASPupload����ϥλy�k�ǰe�Ϥ��Τ�r�æb�U�@�B�J���g�J��Ʈw��
''�䤤��4�Ӥ��� title & id & file1 & content1
%>
<br>
<a href="mainpageM-area-layout.asp?id=<%=id%>">�ק惡�q�Ϥ�ƪ�</a><br>
<FORM action="mainPageM-p1.asp" ENCTYPE="multipart/form-data" METHOD="post" id=form1 name=form1>
  <TABLE WIDTH=100% BORDER=0 CELLSPACING=1 CELLPADDING=1>
    <TR> 
      <TD align="left" valign="top">Title</TD>
      <TD valign="top" align="left"> 
        <input type="text" name="title" value="<%=rs("title")%>">
      </TD>
    </TR>
    <TR> 
      <TD align="left" valign="top">��r</TD>
      <TD valign="top" align="left"> 
        <INPUT type="hidden" id=text2 name=id value=<%=id%>>
        <textarea id="text1" name="content1" cols="100" rows="10" ><% Response.Write "'" & rs("content1") & "'" %>
		</textarea>
      </TD>
    </TR>
    <TR> 
      <TD align="left" valign="top">�s�׹Ϥ�</TD>
      <TD valign="top" align="left"> 
        <INPUT type="file" id=file1 name=file1>
      </TD>
    </TR>
    <TR> 
      <TD colspan="2" align="left" valign="top"> 
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










