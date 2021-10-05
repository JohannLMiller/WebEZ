<!--#include file="data.inc"-->
		<%set rs1=server.CreateObject("adodb.recordset")
				SQLStr="select * from mainpagetitle"
				set rs1=conn.execute(SQLStr)
				pagetitle=rs1("title")
				%>		
		<!--title-table-->
		<TABLE width=100% ALIGN=center BORDER=0 CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD  class="mainpagetitle" ><center><STRONG><%=pagetitle%></STRONG></center></TD>
	</TR>
</TABLE>

			
				<%
				rs1.Close 
				set rs1=nothing
				set rs=server.CreateObject("adodb.recordset")
				SQLStr="select * from mainpage order by id"
				set rs=conn.execute(SQLStr)
				if err.number <> 0 then
						response.write "數據庫操作錯誤︰" + err.description
						err.clear
				else
				rs.MoveFirst 
				do while not rs.EOF 
				
					On Error Resume Next
			 
					   Set objFS = Server.CreateObject("Scripting.FileSystemObject")
					   strName = "../images/product/" & rs("img1")
					   '以Server物件的MapPath()方法取得該檔的實體路徑，再傳入
					   'FileSystemObject物件的FileExists()方法中，判斷檔案是否存在
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
							''判斷layout是哪一種
							''1是左圖右文
							''2是左文右圖
							''3是上圖(置中)下文
							''4是上文下圖(置中)
							''5是全文字
							''6是僅圖片(置中)
							
			''1是左圖右文
									if rs("layout")=1 then
										imgsrc= "<img name='imgshown' src='" & strName & "' border=0 align=left>"
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										contentText = replace(contentText,chr(10),"<br>")
										showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2  & imgsrc & contentText & table3 & table4 & table5
									''2是左文右圖
									elseif rs("layout")=2 then
										imgsrc= "<img name='imgshown' src='" & strName & "' border=0 align=right>"
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										contentText = replace(contentText,chr(10),"<br>")
										showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2  & imgsrc & contentText & table3 & table4 & table5
									''注意事項
									''3之後的樣式待規格提出後修正,目前未修正
									''3是上圖(置中)下文
									elseif rs("layout")=3 then
										imgsrc= "<img name='imgshown' src='" & strName & "' border=0 >"
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										contentText = replace(contentText,chr(10),"<br>")
										showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2 & "<CENTER>"  & imgsrc  & "</CENTER>" & table3 & table4 &  Table2  & contentText  & table3 & table4 & table5
									''4是上文下圖(置中)
									elseif rs("layout")=4 then
										imgsrc= "<img name='imgshown' src='" & strName & "' border=0 >"
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										contentText = replace(contentText,chr(10),"<br>")
										showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2 & contentText  & table3 & table4 &  Table2  & "<CENTER>"  & imgsrc  & "</CENTER>" & table3 & table4 & table5
									''5是全文字
									elseif rs("layout")=5 then
										imgsrc= "<img name='imgshown' src='" & strName & "' border=0 align=left>"
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										contentText = replace(contentText,chr(10),"<br>")
										showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2  & contentText & table3 & table4 & table5
									''6是僅圖片(置中)
									''6是僅圖片(置中)
									elseif rs("layout")=6 then
										imgsrc= "<img name='imgshown' src='" & strName & "' border=0 align=left>"
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										contentText = replace(contentText,chr(10),"<br>")
										'showLayout=Table1 & table1_1 & rs("title") & table1_2 &  table5 & "<CENTER>" & Table1 & table1_1 & imgsrc & table1_2 & table5 & "</CENTER>" 
										showLayout=Table1 & table1_1 & "<center>" & rs("title")& "</center>" & table1_2 &  table5 & "<CENTER>" & "<table><tr><td>" & imgsrc & "</td></tr></table>" & "</CENTER>"  
									end if
					   Else
									''1是左圖右文
									if rs("layout")=1 then
										imgsrc= ""
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										contentText = replace(contentText,chr(10),"<br>")
										showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2  & imgsrc & contentText & table3 & table4 & table5
									''2是左文右圖
									elseif rs("layout")=2 then
										imgsrc= ""
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										contentText = replace(contentText,chr(10),"<br>")
										showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2  & imgsrc & contentText & table3 & table4 & table5
									''注意事項
									''3之後的樣式待規格提出後修正,目前未修正
									''3是上圖(置中)下文
									elseif rs("layout")=3 then
										imgsrc= ""
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										contentText = replace(contentText,chr(10),"<br>")
										showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2 & "<CENTER>"  & imgsrc  & "</CENTER>" & table3 & table4 &  Table2  & contentText  & table3 & table4 & table5
									''4是上文下圖(置中)
									elseif rs("layout")=4 then
										imgsrc= ""
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										contentText = replace(contentText,chr(10),"<br>")
										showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2 & contentText  & table3 & table4 &  Table2  & "<CENTER>"  & imgsrc  & "</CENTER>" & table3 & table4 & table5
									''5是全文字
									elseif rs("layout")=5 then
										imgsrc= ""
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										contentText = replace(contentText,chr(10),"<br>")
										showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2  & contentText & table3 & table4 & table5
									''6是僅圖片(置中)
									''6是僅圖片(置中)
									elseif rs("layout")=6 then
										imgsrc= ""
										contentText=rs("content1")
										contentText = replace(contentText,chr(13),"<br>")
										contentText = replace(contentText,chr(10),"<br>")
										showLayout=Table1 & table1_1 & "<center>" & rs("title")& "</center>" & table1_2 &  table5 & "<CENTER>" & "<table><tr><td>" & imgsrc & "</td></tr></table>" & "</CENTER>"  
										'showLayout=Table1 & table1_1 & rs("title") & table1_2 &  "<tr><td>"  & "<CENTER>"  & imgsrc  & "</CENTER>" & table3 & table4 & table5
									
									end if
					   End If
				
				
				
				Response.Write  showLayout 
				
				
				rs.MoveNext 
				loop
							
				end if
				rs.Close 
				set rs=nothing
				'conn.close
				'set conn=nothing	
				%>
				