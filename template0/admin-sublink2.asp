<!--#include file="data.inc"-->
<%
cat=Request.QueryString("cat")
catNo=Request.QueryString("catNo")
subcat=Request.QueryString("subcat")
subcatNo=Request.QueryString("subcatNo")
%>

<%set rs1=server.CreateObject("adodb.recordset")
				SQLStr="select * from cat where catNo=" & cint(catNo) 
				set rs1=conn.execute(SQLStr)
				pagetitle=rs1("title")
%>		
	
	<TABLE WIDTH=100% ALIGN=center BORDER=1 bordercolor=pink  CELLSPACING=1 CELLPADDING=1>
	<TR>
		<TD>
			<center><STRONG><%=pagetitle%></STRONG>
			
			</center>
		</TD>
	</TR>
	</TABLE>





<table border="0" topmargin="0" marginwidth="0" marginheight="0" leftmargin="0" align=center>
  
  <tr> 
   

   <%
    set rs=server.CreateObject("adodb.recordset")
				SQLStr="select * from subcat where catNo='" & catNo & "'"
				set rs=conn.execute(SQLStr)
				
				if err.number <> 0 then
						response.write "�ƾڮw�ާ@���~�J" + err.description
						err.clear
				else
						if not rs.EOF then
						do while not rs.EOF %>
						<td>
						
                		<A HREF="admin-main2.asp?cat=<%=cat%>&amp;catNo=<%=rs("catNo")%>&amp;subcat=<%=rs("subcat")%>&amp;subcatNo=<%=rs("subcatNo")%>" ><font color="black"><%=rs("subcat")%></font></A>
						
						</td>
						<%
						rs.MoveNext 
						loop
						end if
						
				rs.Close 
				set rs=nothing
				end if
	%>
  </tr>
  </table>
<A HREF="../admin/subCatpageM-layout.asp?cat=<%=cat%>&amp;catNo=<%=catNo%>&amp;subcat=<%=subcat%>&amp;subcatNo=<%=subcatNo%>" target=_blank>�ק�Ҧ������Ϥ�ƪ��覡</A><br> 
  <A HREF="../admin/subcatPageM-area-add-p0.asp?cat=<%=cat%>&amp;catNo=<%=catNo%>&amp;subcat=<%=subcat%>&amp;subcatNo=<%=subcatNo%>" target=_blank>�s�W�Ϥ�q��</A><br>

<%

    set rs=server.CreateObject("adodb.recordset")
				SQLStr="select * from product where cat='" & catNo & "'" & " and subcatNo='" & subcatNo & "'" & "order by id"
				set rs=conn.execute(SQLStr)
				
				if err.number <> 0 then
						response.write "�ƾڮw�ާ@���~�J" + err.description
						err.clear
				else
						if not rs.EOF then
						do while not rs.EOF %>
						<%Dim objFS, strName
							On Error Resume Next
			 
					   Set objFS = Server.CreateObject("Scripting.FileSystemObject")
					   strName = "../images/product/" & rs("img1")
					   '�HServer����MapPath()��k���o���ɪ�������|�A�A�ǤJ
					   'FileSystemObject����FileExists()��k���A�P�_�ɮ׬O�_�s�b
					   Table1="<table  border='1'  bordercolor=yellow cellspacing='0' cellpadding='10'>"
						table1_1="<tr><td align=left><b>" 
						table1_2="</b></td></tr>"
						Table2= "<tr><td class='9pt-black'>" 
						TABLE2_1=""
						Table3= "</td>"
						Table4= "</tr>"
						Table5= "</table><br>"
					   
					   If objFS.FileExists(Server.MapPath(strName)) Then
							''�P�_layout�O���@��
							''1�O���ϥk��
							''2�O����k��
							''3�O�W��(�m��)�U��
							''4�O�W��(�a��)�U��
							''5�O�W��(�a�k)�U��
							''6�O�W��U��(�m��)
							''7�O�W��U��(�a��)
							''8�O�W��U��(�a�k)
							''9�O����r
						
							if rs("layout")=1 then
								imgsrc= "<img name='imgshown' src='" & strName & "' border=0 align=left>"
								contentText=rs("content1")
								showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2  & imgsrc & contentText & table3 & table4 & table5
							elseif rs("layout")=2 then
								imgsrc= "<img name='imgshown' src='" & strName & "' border=0 align=right>"
								contentText=rs("content1")
								showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2  & imgsrc & contentText & table3 & table4 & table5
							''�`�N�ƶ�
							''3���᪺�˦��ݳW�洣�X��ץ�,�ثe���ץ�
							elseif rs("layout")=3 then
								imgsrc="<td class='9pt-black' align='center'>" & "<img name='imgshown' src='" & strName & "' border=0 >" 
								contentText="<br>" & rs("content1") 
								showLayout=Table1 & Table2 & imgsrc & contentText & table3 & table4 & table5
							elseif rs("layout")=4 then
								imgsrc="<td class='9pt-black' align=left>" & "<img name='imgshown' src='" & strName & "' border=0 >" 
								contentText="<br>" & rs("content1") 
								showLayout=Table1 & Table2 & imgsrc & contentText & table3 & table4 & table5	
							elseif rs("layout")=5 then
								imgsrc="<td class='9pt-black' align=right>" & "<img name='imgshown' src='" & strName & "' border=0 >" 
								contentText="<br>" & rs("content1") 
								showLayout=Table1 & Table2 & imgsrc & contentText & table3 & table4 & table5	
							
							elseif rs("layout")=6 then
								contentText="<td class='9pt-black' >" & rs("content1") 
								imgsrc="<center>" & "<img name=imgshown src=" & strName & " border=0 >" & "</center>"
								showLayout=Table1 & Table2 & contentText & imgsrc & table3 & table4 & table5
							elseif rs("layout")=7 then
								contentText="<td class='9pt-black' >" & rs("content1") 
								imgsrc="<p align='left'>" & "<img name=imgshown src=" & strName & " border=0 >" & "</p>"
								showLayout=Table1 & Table2 & contentText & imgsrc & table3 & table4 & table5
							elseif rs("layout")=8 then
								contentText="<td class='9pt-black' >" & rs("content1") 
								imgsrc="<p align='right'>" & "<img name=imgshown src=" & strName & " border=0 >" & "</p>"
								showLayout=Table1 & Table2 & contentText & imgsrc & table3 & table4 & table5
							elseif rs("layout")=9 then
								contentText="<td class='9pt-black' >" & rs("content1") 
								imgsrc="<p align='right'>" & "<img name=imgshown src=" & strName & " border=0 >" & "</p>"
								showLayout=Table1 & Table2 & contentText & table3 & table4 & table5
							elseif rs("layout")=10 then
								contentText="<td class='9pt-black'>" & rs("content1") 
								imgsrc="<img name=imgshown src=" & strName & " border=0 >"
								showLayout=imgsrc 
							else ''���w�q�̥��H�˦��@���w�]�˦�
								imgsrc= "<img name='imgshown' src='" & strName & "' border=0 align=left>"
								contentText=rs("content1")
								showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2  & imgsrc & contentText & table3 & table4 & table5
							end if
							
					   Else
					        imgsrc=""
					       contentText=rs("content1")
						   showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2  & imgsrc & contentText & table3 & table4 & table5
					   
					   End If
				
				Response.Write "<br><A HREF=../admin/subcatpageM-area.asp?id=" & rs("id") & " target=_blank>�ק�H�U���Ϥ�</A>"
				Response.Write  showLayout 
				
				
				rs.MoveNext 
				loop
							
				end if
				rs.Close 
				set rs=nothing
				rs1.Close 
				set rs1=nothing
				conn.close
				set conn=nothing	
				end if
%>
