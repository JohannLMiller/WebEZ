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
	<%if  cint(catNo)=rs("catNo")then %>
				
				<table width="140" border="0" cellspacing="0" cellpadding="0" height="30">
				        <tr> 
				          <td background="images/button-bg.jpg" width="140" height="30" class="linktable"> 
							<a HREF="admin-main1.asp?catNo=<%=rs("catNo")%>&amp;cat=<%=rs("cat")%>"><font color= red><%=rs("cat")%></font></a>
						  </td>
						</tr>
				 </table>
				 
				 
				 
				 
							 <%
								  set rs1=server.CreateObject("adodb.recordset")
												SQLStr="select * from subcat where catNo='" & catNo & "'"
												set rs1=conn.execute(SQLStr)
												
												if err.number <> 0 then
														response.write "數據庫操作錯誤︰" + err.description
														err.clear
												else
														if not rs1.EOF then
														do while not rs1.EOF %>
														
														<table width="140" border="1" bordercolor=pink  cellspacing="0" cellpadding="0" height="30">
														<tr>
														<td class="sublinktable" align=center>
														
								              		<A HREF="main2.asp?cat=<%=cat%>&amp;catNo=<%=rs1("catNo")%>&amp;subcat=<%=rs1("subcat")%>&amp;subcatNo=<%=rs1("subcatNo")%>" class="sub" ><%=rs1("subcat")%></A>
														
														</td>
														</tr>
														<table>
														<%
														rs1.MoveNext 
														loop
														end if
														
												rs1.Close 
												set rs1=nothing
												end if
									%>
									
				<%else%>	
				<table width="140" border="0" cellspacing="0" cellpadding="0" height="30">
				        <tr> 
				          <td background="images/button-bg.jpg" width="140" height="30" class="linktable"> 
							<a HREF="main1.asp?catNo=<%=rs("catNo")%>&amp;cat=<%=rs("cat")%>"><%=rs("cat")%></a>
						  </td>
						</tr>
				 </table>
				 <%end if%>
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

