<!--#include file="data.inc"-->
<%
cat=Request.QueryString("cat")
catNo=Request.QueryString("catNo")
'subcat=Request.QueryString("subcat")
'subcatNo=Request.QueryString("subcatNo")
  

%>
  
 <table width="140" border="0" cellspacing="0" cellpadding="0" height="30">
        <tr> 
          <td background="images/button-bg.jpg" width="140" height="30" class="linktable"> 
          <a HREF="main.asp" target="_top">回首頁</a>
          </td>
        </tr>
      </table>
      
<%
set rs=server.CreateObject("adodb.recordset")
	SQLStr="select * from cat order by catNo "
	set rs=conn.execute(SQLStr)



if err.number <> 0 then
		response.write "數據庫操作錯誤︰" + err.description
		err.clear
else
if not rs.EOF then 
	do while not rs.EOF 
	%>
 <table width="140" border="0" cellspacing="0" cellpadding="0" height="30">
        <tr> 
          <td background="images/button-bg.jpg" width="140" height="30" class="linktable"> 
			<a HREF="main1.asp?catNo=<%=rs("catNo")%>&amp;cat=<%=rs("cat")%>"><%=rs("cat")%></a>
		  </td>
		</tr>
 </table>
<%
rs.MoveNext 
	loop
	
	end if
	
	
rs.Close 
set rs=nothing
	
	
end if

conn.close
set conn=nothing	
	

%>

