<!--#include file="data.inc"-->
	
				<%
				set rs=server.CreateObject("adodb.recordset")
					SQLStr="select * from template"
					set rs=conn.execute(SQLStr)
				if err.number <> 0 then
						response.write "�ƾڮw�ާ@���~�J" + err.description
						err.clear
				else
				if not rs.EOF then 
					rs.MoveFirst 
					%>



				<table width="100%" border="1" cellspacing="0" cellpadding="0">
				  <tr>
				    <td><img src="../images/logo/<%=rs("logo")%>" align=left>
   <center><b><strong><br><%=rs("headline")%></strong> <A HREF="../admin/headline-p0.asp" target=_blank>�޲z���D</A></b></center>
				    <A HREF="../admin/logo-p0.asp" target=_blank>�޲zLogo</A>
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
		