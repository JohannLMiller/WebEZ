<!--#include file="data.inc"-->
<%
set rs=server.CreateObject("adodb.recordset")
	SQLStr="select * from template"
	set rs=conn.execute(SQLStr)
if err.number <> 0 then
		response.write "¼Æ¾Ú®w¾Þ§@¿ù»~¡J" + err.description
		err.clear
else
	if not rs.EOF then 
	rs.MoveFirst 
	%>

<table width=100% border="0" cellspacing="0"  bordercolor=green  cellpadding="0" align="center">
  <tr class="copyright">
    <td align="center" valign="middle"> 
      <div align="center"><%=rs("copyright")%><br>
      <A HREF="../admin/copyright-p0.asp" target=_blank><IMG SRC="../images/manage_copyright.gif" BORDER=0></A>
      
      </div>
    </td>
  </tr>
</table>
<%
	end if
end if
rs.Close 
set rs=nothing
conn.close
set conn=nothing	



%>