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
<!--sublink2-title-table-->	
	<table width="100%" ALIGN="center" BORDER="0" bordercolor="pink" CELLSPACING="0" CELLPADDING="0">
	<tr>
		<td class="mainpagetitle">
			<center>
			<%=pagetitle%>
			</center>
		</td>
	</tr>
	</table>
	<br>



<!--sublink2-sublink-table-->	
 
   

   <%
    set rs=server.CreateObject("adodb.recordset")
				SQLStr="select * from subcat where catNo='" & catNo & "'"
				set rs=conn.execute(SQLStr)
				
				if err.number <> 0 then
						response.write "數據庫操作錯誤︰" + err.description
						err.clear
				else
						if not rs.EOF then%>
 <table border="0" topmargin="0" marginwidth="0" marginheight="0" leftmargin="0" align="center">
  <tr>
						
						
						<%
						do while not rs.EOF %>
						<td class="sublinktable" align="center">
							<%if  cint(subcatNo)=rs("subcatNo")then %>
							<table border="0" bordercolor=pink topmargin="0" marginwidth="0" marginheight="0" leftmargin="0" align="center">
							<tr>
							<td>
                			<a HREF="#" class="sub"><font color=red><%=rs("subcat")%></font></a>
							</td>
							</tr>
							</table>
							<%else%>
							<table border="0" bordercolor=black  topmargin="0" marginwidth="0" marginheight="0" leftmargin="0" align="center">
							<tr>
							<td>
                			<a HREF="admin-main2.asp?cat=<%=cat%>&amp;catNo=<%=rs("catNo")%>&amp;subcat=<%=rs("subcat")%>&amp;subcatNo=<%=rs("subcatNo")%>" class="sub"><%=rs("subcat")%></a>
							
							</td>
							</tr>
							</table>
							<%end if%>
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
 <br> 
<a HREF="../admin/subCatpageM-layout.asp?cat=<%=cat%>&amp;catNo=<%=catNo%>&amp;subcat=<%=subcat%>&amp;subcatNo=<%=subcatNo%>" target="_blank"><img SRC="../images/manage_layout.gif" BORDER="0" WIDTH="108" HEIGHT="25"></a><br> 
<center>
<a HREF="../admin/subcatPageM-area-add-p0.asp?cat=<%=cat%>&amp;catNo=<%=catNo%>&amp;subcat=<%=subcat%>&amp;subcatNo=<%=subcatNo%>" target="_blank"><img SRC="../images/add_par_img.gif" BORDER="0" WIDTH="108" HEIGHT="25"></a><br>
  </center>
<%

    set rs=server.CreateObject("adodb.recordset")
				SQLStr="select * from product where cat='" & catNo & "'" & " and subcatNo='" & subcatNo & "'" & "order by id"
				set rs=conn.execute(SQLStr)
				
				if err.number <> 0 then
						response.write "數據庫操作錯誤︰" + err.description
						err.clear
				else
						if not rs.EOF then
						do while not rs.EOF %>
						<%Dim objFS, strName
							On Error Resume Next
			 
					   Set objFS = Server.CreateObject("Scripting.FileSystemObject")
					   strName = "../images/product/" & rs("img1")
					   '以Server物件的MapPath()方法取得該檔的實體路徑，再傳入
					   'FileSystemObject物件的FileExists()方法中，判斷檔案是否存在
					 'sublink-content-table'
					 Table1="<table  border='0' width=100%  bordercolor=yellow cellspacing='0' cellpadding='0'>"
						table1_1="<tr><td align=left  class='title'><b>" 
						table1_2="</b></td></tr>"
						Table2= "<tr><td  class='text'>" 
						TABLE2_1=""
						Table3= "</td>"
						Table4= "</tr>"
						Table5= "</table><br>"
					   
					   If objFS.FileExists(Server.MapPath(strName)) Then
							''判斷layout是哪一種
							''1是左圖右文
							''2是左文右圖
							''3是上圖(置中)下文
							''4是上圖(靠左)下文
							''5是上圖(靠右)下文
							''6是上文下圖(置中)
							''7是上文下圖(靠左)
							''8是上文下圖(靠右)
							''9是全文字
						
							''1是左圖右文
									if rs("layout")=1 then
										imgsrc= "<img name='imgshown' src='" & strName & "' border=0 align=left>"
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										'contentText = replace(contentText,chr(10),"<br>")
										contentText = replace(contentText,chr(32)," ")
										showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2  & imgsrc & contentText & table3 & table4 & table5
									''2是左文右圖
									elseif rs("layout")=2 then
										imgsrc= "<img name='imgshown' src='" & strName & "' border=0 align=right>"
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										'contentText = replace(contentText,chr(10),"<br>")
										contentText = replace(contentText,chr(32)," ")
										showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2  & imgsrc & contentText & table3 & table4 & table5
									''注意事項
									
									''3是上圖(置中)下文
									elseif rs("layout")=3 then
										imgsrc= "<img name='imgshown' src='" & strName & "' border=0 >"
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										'contentText = replace(contentText,chr(10),"<br>")
										contentText = replace(contentText,chr(32)," ")
										showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2 & "<CENTER>"  & imgsrc  & "</CENTER>" & table3 & table4 &  Table2  & contentText  & table3 & table4 & table5
									''4是上文下圖(置中)
									elseif rs("layout")=4 then
										imgsrc= "<img name='imgshown' src='" & strName & "' border=0 >"
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										'contentText = replace(contentText,chr(10),"<br>")
										contentText = replace(contentText,chr(32)," ")
										showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2 & contentText  & table3 & table4 &  Table2  & "<CENTER>"  & imgsrc  & "</CENTER>" & table3 & table4 & table5
									''5是全文字
									elseif rs("layout")=5 then
										imgsrc= "<img name='imgshown' src='" & strName & "' border=0 align=left>"
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										'contentText = replace(contentText,chr(10),"<br>")
										contentText = replace(contentText,chr(32)," ")
										showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2  & contentText & table3 & table4 & table5
									''6是僅圖片(置中)
									''6是僅圖片(置中)
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
									''1是左圖右文
									if rs("layout")=1 then
										imgsrc= ""
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										'contentText = replace(contentText,chr(10),"<br>")
										contentText = replace(contentText,chr(32)," ")
										showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2  & imgsrc & contentText & table3 & table4 & table5
									''2是左文右圖
									elseif rs("layout")=2 then
										imgsrc= ""
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										'contentText = replace(contentText,chr(10),"<br>")
										contentText = replace(contentText,chr(32)," ")
										showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2  & imgsrc & contentText & table3 & table4 & table5
									''注意事項
									''3之後的樣式待規格提出後修正,目前未修正
									''3是上圖(置中)下文
									elseif rs("layout")=3 then
										imgsrc= ""
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										'contentText = replace(contentText,chr(10),"<br>")
										contentText = replace(contentText,chr(32)," ")
										showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2 & "<CENTER>"  & imgsrc  & "</CENTER>" & table3 & table4 &  Table2  & contentText  & table3 & table4 & table5
									''4是上文下圖(置中)
									elseif rs("layout")=4 then
										imgsrc= ""
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										'contentText = replace(contentText,chr(10),"<br>")
										contentText = replace(contentText,chr(32)," ")
										showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2 & contentText  & table3 & table4 &  Table2  & "<CENTER>"  & imgsrc  & "</CENTER>" & table3 & table4 & table5
									''5是全文字
									elseif rs("layout")=5 then
										imgsrc= ""
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										'contentText = replace(contentText,chr(10),"<br>")
										contentText = replace(contentText,chr(32)," ")
										showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2  & contentText & table3 & table4 & table5
									''6是僅圖片(置中)
									''6是僅圖片(置中)
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
				
				Response.Write "<br><A HREF=../admin/subcatpageM-area.asp?id=" & rs("id") & " target=_blank><IMG SRC='../images/fix_text_img.gif' BORDER=0></A>"
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
