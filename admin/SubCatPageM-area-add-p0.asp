<%@ Language=VBScript %>
<!--#include file="data.inc"-->
<%
''�ޤJlayout

catNo=Request.QueryString("catNo")
subcat=Request.QueryString("subcat")
subcatNo=Request.QueryString("subcatNo")

set rs=server.CreateObject("adodb.recordset")
rs.open "product",conn,1,3	

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
�Эק��r�ιϤ�
<%
catNo=Request.QueryString("catNo")
''�H�U��form�����t�XSAfileup����ϥλy�k�ǰe�Ϥ��Τ�r�æb�U�@�B�J���g�J��Ʈw��
''�䤤��5�Ӥ��� catNo & title  & file1 & content1 & layout
%>
<FORM action="SubCatPageM-area-add-p1.asp" ENCTYPE="multipart/form-data" METHOD="post" id=form1 name=form1>
  <TABLE WIDTH=100% BORDER=0 CELLSPACING=1 CELLPADDING=1>
    <TR> 
      <TD align="left" valign="top">Title</TD>
      <TD valign="top" align="left"> 
        <input type="text" name="title">
      </TD>
    </TR>
    <TR> 
      <TD align="left" valign="top">��r</TD>
      <TD valign="top" align="left"> 
      
        <textarea id="text1" name="content1" cols="100" rows="10" >
		
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
        <INPUT type="hidden" id=text2 name=catNo value="<%=catNo%>">
        <INPUT type="hidden" id=text3 name=layout value="<%=layout%>">
         <INPUT type="hidden" id=text4 name=subcatNo value=<%=subcatNo%>>
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
