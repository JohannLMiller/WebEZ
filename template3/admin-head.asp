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

		
			<a HREF="../admin/logo-p0.asp" target="_blank"><img SRC="../images/manage_logo.gif" border="0" WIDTH="96" HEIGHT="25"></a>
			<table width="780" border="0" cellspacing="0" cellpadding="0" height="100">				 
				  <tr>
				    <td background="images/banner-bg.jpg" width="780" class="banner" valign="top">
					<div align="center">
					<img src="../images/logo/<%=rs("logo")%>" align="left" border="0" HEIGHT="100">
					
					<br>
					<a HREF="../admin/headline-p0.asp" target="_blank">
					<img SRC="../images/manage_title.gif" border="0" WIDTH="75" HEIGHT="25"></a>
					<br><%=rs("headline")%>
					
				    
				    </div>
				    </td>
				    
				  </tr>
				</table>

				<%	
				end if	
				rs.Close 
				set rs=nothing
				conn.close
				set conn=nothing	
					
					
				end if

				%>
		