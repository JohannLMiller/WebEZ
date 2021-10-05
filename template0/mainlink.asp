<!--#include file="data.inc"-->
<%
cat=Request.QueryString("cat")
catNo=Request.QueryString("catNo")
'subcat=Request.QueryString("subcat")
'subcatNo=Request.QueryString("subcatNo")
  

%>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<a HREF="main.asp" target="_top">回首頁</a>  
<table width="100%" border="1" >
  <tr> 
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

    <td valign="top"><a HREF="main1.asp?catNo=<%=rs("catNo")%>&amp;bamp;cat=<%=rs("cat")%>"><%=rs("cat")%></a></td>
  </tr>

<%
rs.MoveNext 
	loop
	
	end if
	
	
rs.Close 
set rs=nothing
'conn.close
'set conn=nothing	
	
	
end if

%>
</table>

		
<%


conn.close
set conn=nothing	
	

%>

