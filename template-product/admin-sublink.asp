<!--#include file="data.inc"-->
<%
cat=Request.QueryString("cat")
catNo=Request.QueryString("catNo")
%>
<%set rs1=server.CreateObject("adodb.recordset")
				SQLStr="select * from cat where catNo=" & cint(catNo) 
				set rs1=conn.execute(SQLStr)
				pagetitle=rs1("title")
%>		
<!--title-table-->	
	<table width="100%" ALIGN="center" BORDER="0" bordercolor="pink" CELLSPACING="0" CELLPADDING="0">
	<tr>
		<td class="mainpagetitle">
			<center>
			
			<a HREF="../admin/CatpageM-title-p0.asp?catNo=<%=catNo%>" target="_blank">
			<img SRC="../images/fix_title.gif" BORDER="0" WIDTH="75" HEIGHT="25">
			</a>
			<br>
			<%=pagetitle%>
			</center>
		</td>
	</tr>
	</table>
<!--sublink-table-->	
 <center>
<table border="0" topmargin="0" marginwidth="0" marginheight="0" leftmargin="0" align="center">


						<a HREF="../admin/GUIsubcatM.asp?catNo=<%=catNo%>" target="_blank">
						<img SRC="../images/manage_sublink.gif" BORDER="0" WIDTH="102" HEIGHT="25">
						</a>
						</center>  

   <%
    set rs=server.CreateObject("adodb.recordset")
				SQLStr="select * from subcat where catNo='" & catNo & "'"
				set rs=conn.execute(SQLStr)
				
				if err.number <> 0 then
						response.write "�ƾڮw�ާ@���~�J" + err.description
						err.clear
				else
						if not rs.EOF then%>
						
  <table border="0" topmargin="0" marginwidth="0" marginheight="0" leftmargin="0" align="center">
  <tr> 
						<%do while not rs.EOF %>
						<td class="sublinktable" align="center">
                		<a HREF="admin-main2.asp?cat=<%=cat%>&amp;catNo=<%=rs("catNo")%>&amp;subcat=<%=rs("subcat")%>&amp;subcatNo=<%=rs("subcatNo")%>" class="sub"><%=rs("subcat")%></a>
						</td>
						<%
						rs.MoveNext 
						loop
						end if%>
						
				<%rs.Close 
				set rs=nothing
				end if
	%>
  </tr>
  </table>

  <a HREF="../admin/CatpageM-layout.asp?catNo=<%=catNo%>" target="_blank"><img SRC="../images/manage_layout.gif" BORDER="0" WIDTH="108" HEIGHT="25"></a><br> 
  <center>
  <a HREF="../admin/catPageM-area-add-p0.asp?catNo=<%=catNo%>" target="_blank"><img SRC="../images/add_par_img.gif" BORDER="0" WIDTH="108" HEIGHT="25"></a><br>
  </center>
<%
    set rs=server.CreateObject("adodb.recordset")
				SQLStr="select * from catcontent where catNo='" & catNo & "'" & "order by id"
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
					 'sublinkcontent-table'
					  Table1="<table  border='0' width=100%  bordercolor=yellow cellspacing='0' cellpadding='0'>"
						table1_1="<tr><td align=left  class='title'><b>" 
						table1_2="</b></td></tr>"
						Table2= "<tr><td  class='text'>" 
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
										contentText = replace(contentText,chr(13),"<br>")
										contentText = replace(contentText,chr(10),"<br>")
										showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2  & imgsrc & contentText & table3 & table4 & table5
									''2�O����k��
									elseif rs("layout")=2 then
										imgsrc= "<img name='imgshown' src='" & strName & "' border=0 align=right>"
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										contentText = replace(contentText,chr(10),"<br>")
										showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2  & imgsrc & contentText & table3 & table4 & table5
									''�`�N�ƶ�
									''3���᪺�˦��ݳW�洣�X��ץ�,�ثe���ץ�
									''3�O�W��(�m��)�U��
									elseif rs("layout")=3 then
										imgsrc= "<img name='imgshown' src='" & strName & "' border=0 >"
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										contentText = replace(contentText,chr(10),"<br>")
										showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2 & "<CENTER>"  & imgsrc  & "</CENTER>" & table3 & table4 &  Table2  & contentText  & table3 & table4 & table5
									''4�O�W��U��(�m��)
									elseif rs("layout")=4 then
										imgsrc= "<img name='imgshown' src='" & strName & "' border=0 >"
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										contentText = replace(contentText,chr(10),"<br>")
										showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2 & contentText  & table3 & table4 &  Table2  & "<CENTER>"  & imgsrc  & "</CENTER>" & table3 & table4 & table5
									''5�O����r
									elseif rs("layout")=5 then
										imgsrc= "<img name='imgshown' src='" & strName & "' border=0 align=left>"
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										contentText = replace(contentText,chr(10),"<br>")
										showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2  & contentText & table3 & table4 & table5
									''6�O�ȹϤ�(�m��)
									''6�O�ȹϤ�(�m��)
									elseif rs("layout")=6 then
										imgsrc= "<img name='imgshown' src='" & strName & "' border=0 align=left>"
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										contentText = replace(contentText,chr(10),"<br>")
										'showLayout=Table1 & table1_1 & rs("title") & table1_2 &  table5 & "<CENTER>" & Table1 & table1_1 & imgsrc & table1_2 & table5 & "</CENTER>" 
										showLayout=Table1 & table1_1 & "<center>" & rs("title")& "</center>" & table1_2 &  table5 & "<CENTER>" & "<table><tr><td>" & imgsrc & "</td></tr></table>" & "</CENTER>"  
									end if
					   Else
									''1�O���ϥk��
									if rs("layout")=1 then
										imgsrc= ""
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										contentText = replace(contentText,chr(10),"<br>")
										showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2  & imgsrc & contentText & table3 & table4 & table5
									''2�O����k��
									elseif rs("layout")=2 then
										imgsrc= ""
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										contentText = replace(contentText,chr(10),"<br>")
										showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2  & imgsrc & contentText & table3 & table4 & table5
									''�`�N�ƶ�
									''3���᪺�˦��ݳW�洣�X��ץ�,�ثe���ץ�
									''3�O�W��(�m��)�U��
									elseif rs("layout")=3 then
										imgsrc= ""
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										contentText = replace(contentText,chr(10),"<br>")
										showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2 & "<CENTER>"  & imgsrc  & "</CENTER>" & table3 & table4 &  Table2  & contentText  & table3 & table4 & table5
									''4�O�W��U��(�m��)
									elseif rs("layout")=4 then
										imgsrc= ""
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										contentText = replace(contentText,chr(10),"<br>")
										showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2 & contentText  & table3 & table4 &  Table2  & "<CENTER>"  & imgsrc  & "</CENTER>" & table3 & table4 & table5
									''5�O����r
									elseif rs("layout")=5 then
										imgsrc= ""
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										contentText = replace(contentText,chr(10),"<br>")
										showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2  & contentText & table3 & table4 & table5
									''6�O�ȹϤ�(�m��)
									''6�O�ȹϤ�(�m��)
									elseif rs("layout")=6 then
										imgsrc= ""
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										contentText = replace(contentText,chr(10),"<br>")
										showLayout=Table1 & table1_1 & "<center>" & rs("title")& "</center>" & table1_2 &  table5 & "<CENTER>" & "<table><tr><td>" & imgsrc & "</td></tr></table>" & "</CENTER>"  
										'showLayout=Table1 & table1_1 & rs("title") & table1_2 &  "<tr><td>"  & "<CENTER>"  & imgsrc  & "</CENTER>" & table3 & table4 & table5
									
									end if
					   End If
				
				Response.Write "<A HREF=../admin/catpageM-area.asp?id=" & rs("id") & " target=_blank><IMG SRC='../images/fix_text_img.gif' BORDER=0></A>"
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
