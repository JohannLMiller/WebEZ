<!--#include file="data.inc"-->
				<%set rs1=server.CreateObject("adodb.recordset")
				SQLStr="select * from mainpagetitle"
				set rs1=conn.execute(SQLStr)
				if not rs1.EOF then
				rs1.MoveFirst 
				pagetitle=rs1("title")
				end if
				%>		
	<!--title-table-->
	<table width="100%" ALIGN="center" BORDER="0" bordercolor="pink" CELLSPACING="0" CELLPADDING="0">
	<tr>
		<td class="mainpagetitle">
		
			<center>
			<a HREF="../admin/mainpageM-title-p0.asp" target="_blank"><img SRC="../images/fix_title.gif" border="0" WIDTH="75" HEIGHT="25"></a><br>
			<%=pagetitle%>
			</center>
		</td>
	</tr>
	</table>
	<br>
	<a HREF="../admin/mainpageM-layout.asp" target="_blank"><img SRC="../images/manage_layout.gif" BORDER="0" WIDTH="108" HEIGHT="25"></a><br> 
	<center>
	<a HREF="../admin/mainpageM-add.asp" target="_blank"><img SRC="../images/add_par_img.gif" BORDER="0" WIDTH="108" HEIGHT="25"></a><br>
	</center>		
				<%
				rs1.Close 
				set rs1=nothing
				set rs=server.CreateObject("adodb.recordset")
				SQLStr="select * from mainpage order by id"
				set rs=conn.execute(SQLStr)
				if err.number <> 0 then
						response.write "�ƾڮw�ާ@���~�J" + err.description
						err.clear
				else
				
				if not rs.EOF then
				rs.MoveFirst 
				do while not rs.EOF 
				
					On Error Resume Next
			 
					   Set objFS = Server.CreateObject("Scripting.FileSystemObject")
					   strName = "../images/product/" & rs("img1")
					   '�HServer����MapPath()��k���o���ɪ�������|�A�A�ǤJ
					   'FileSystemObject����FileExists()��k���A�P�_�ɮ׬O�_�s�b
					''  mainpagecontent-table
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
							''4�O�W��U��(�m��)
							''5�O����r
							''6�O�ȹϤ�(�m��)
							
						
									''1�O���ϥk��
									if rs("layout")=1 then
										imgsrc= "<img name='imgshown' src='" & strName & "' border=0 align=left>"
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										'contentText = replace(contentText,chr(10),"<br>")
										contentText = replace(contentText,chr(32)," ")
										showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2  & imgsrc & contentText & table3 & table4 & table5
									''2�O����k��
									elseif rs("layout")=2 then
										imgsrc= "<img name='imgshown' src='" & strName & "' border=0 align=right>"
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										'contentText = replace(contentText,chr(10),"<br>")
										contentText = replace(contentText,chr(32)," ")
										showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2  & imgsrc & contentText & table3 & table4 & table5
									''�`�N�ƶ�
									
									''3�O�W��(�m��)�U��
									elseif rs("layout")=3 then
										imgsrc= "<img name='imgshown' src='" & strName & "' border=0 >"
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										'contentText = replace(contentText,chr(10),"<br>")
										contentText = replace(contentText,chr(32)," ")
										showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2 & "<CENTER>"  & imgsrc  & "</CENTER>" & table3 & table4 &  Table2  & contentText  & table3 & table4 & table5
									''4�O�W��U��(�m��)
									elseif rs("layout")=4 then
										imgsrc= "<img name='imgshown' src='" & strName & "' border=0 >"
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										'contentText = replace(contentText,chr(10),"<br>")
										contentText = replace(contentText,chr(32)," ")
										showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2 & contentText  & table3 & table4 &  Table2  & "<CENTER>"  & imgsrc  & "</CENTER>" & table3 & table4 & table5
									''5�O����r
									elseif rs("layout")=5 then
										imgsrc= "<img name='imgshown' src='" & strName & "' border=0 align=left>"
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										'contentText = replace(contentText,chr(10),"<br>")
										contentText = replace(contentText,chr(32)," ")
										showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2  & contentText & table3 & table4 & table5
									''6�O�ȹϤ�(�m��)
									''6�O�ȹϤ�(�m��)
									elseif rs("layout")=6 then
										imgsrc= "<img name='imgshown' src='" & strName & "' border=0 align=left>"
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										'contentText = replace(contentText,chr(10),"<br>")
										contentText = replace(contentText,chr(32)," ")
										'showLayout=Table1 & table1_1 & rs("title") & table1_2 &  table5 & "<CENTER>" & Table1 & table1_1 & imgsrc & table1_2 & table5 & "</CENTER>" 
										showLayout=Table1 & table1_1 & "<center>" & rs("title")& "</center>" & table1_2 &  table5 & "<CENTER>" & "<table><tr><td>" & imgsrc & "</td></tr></table>" & "</CENTER>"  
									end if
					   Else
									''1�O���ϥk��
									if rs("layout")=1 then
										imgsrc= ""
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										'contentText = replace(contentText,chr(10),"<br>")
										contentText = replace(contentText,chr(32)," ")
										showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2  & imgsrc & contentText & table3 & table4 & table5
									''2�O����k��
									elseif rs("layout")=2 then
										imgsrc= ""
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										'contentText = replace(contentText,chr(10),"<br>")
										contentText = replace(contentText,chr(32)," ")
										showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2  & imgsrc & contentText & table3 & table4 & table5
									''�`�N�ƶ�
									''3���᪺�˦��ݳW�洣�X��ץ�,�ثe���ץ�
									''3�O�W��(�m��)�U��
									elseif rs("layout")=3 then
										imgsrc= ""
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										'contentText = replace(contentText,chr(10),"<br>")
										contentText = replace(contentText,chr(32)," ")
										showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2 & "<CENTER>"  & imgsrc  & "</CENTER>" & table3 & table4 &  Table2  & contentText  & table3 & table4 & table5
									''4�O�W��U��(�m��)
									elseif rs("layout")=4 then
										imgsrc= ""
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										'contentText = replace(contentText,chr(10),"<br>")
										contentText = replace(contentText,chr(32)," ")
										showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2 & contentText  & table3 & table4 &  Table2  & "<CENTER>"  & imgsrc  & "</CENTER>" & table3 & table4 & table5
									''5�O����r
									elseif rs("layout")=5 then
										imgsrc= ""
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										'contentText = replace(contentText,chr(10),"<br>")
										contentText = replace(contentText,chr(32)," ")
										showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2  & contentText & table3 & table4 & table5
									''6�O�ȹϤ�(�m��)
									''6�O�ȹϤ�(�m��)
									elseif rs("layout")=6 then
										imgsrc= ""
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										'contentText = replace(contentText,chr(10),"<br>")
										contentText = replace(contentText,chr(32)," ")
										showLayout=Table1 & table1_1 & "<center>" & rs("title")& "</center>" & table1_2 &  table5 & "<CENTER>" & "<table><tr><td>" & imgsrc & "</td></tr></table>" & "</CENTER>"  
										'showLayout=Table1 & table1_1 & rs("title") & table1_2 &  "<tr><td>"  & "<CENTER>"  & imgsrc  & "</CENTER>" & table3 & table4 & table5
									
									end if
					   End If
				
				
						Response.Write "<br><A HREF=../admin/mainpageM.asp?id=" & rs("id") & " target=_blank><IMG SRC='../images/fix_text_img.gif' BORDER=0></A><br>"
						Response.Write  showLayout 
				
				
						rs.MoveNext 
						loop
					end if				
					end if
				rs.Close 
				set rs=nothing
				'conn.close
				'set conn=nothing	
				%>
				
				