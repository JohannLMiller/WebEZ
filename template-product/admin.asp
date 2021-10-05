<%@ Language=VBScript %>

<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="content.css" type="text/css">

</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<A HREF="../admin-index.asp">檢視修改結果</A><br>
<A HREF="../admin/template-p0.asp"target=_blank>管理樣式</A>
<TABLE WIDTH=100% ALIGN=center BORDER=1 CELLSPACING=0 CELLPADDING=0>
  <TR>
		<TD>
		<!--#include file="data.inc"-->
	
				<%
				set rs=server.CreateObject("adodb.recordset")
					SQLStr="select * from template"
					set rs=conn.execute(SQLStr)
				if err.number <> 0 then
						response.write "數據庫操作錯誤︰" + err.description
						err.clear
				else
				if not rs.EOF then 
					rs.MoveFirst 
					%>



				<table width="100%" border="1" cellspacing="0" cellpadding="0">
				  <tr>
				    <td><img src="../images/logo/<%=rs("logo")%>" align=left>
   <center><b><strong><br><%=rs("headline")%></strong> <A HREF="../admin/headline-p0.asp" target=_blank>管理標題</A></b></center>
				    <A HREF="../admin/logo-p0.asp" target=_blank>管理Logo</A>
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
		
		</TD>
		<TD></TD>
	</TR>
</TABLE>
<table width="100%" border="1" cellspacing="0" cellpadding="0">
  <tr>
	<td width="13%"  valign="top"> 
				  <!--#include file="data.inc"-->
					<%
					cat=Request.QueryString("cat")
					catNo=Request.QueryString("catNo")
					'subcat=Request.QueryString("subcat")
					'subcatNo=Request.QueryString("subcatNo")
					  

					%>
					
					<a HREF="admin.asp" target="_top">回首頁</a>  
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

					    <td valign="top"><a HREF="admin-main1.asp?catNo=<%=rs("catNo")%>&amp;bamp;cat=<%=rs("cat")%>"><%=rs("cat")%></a></td>
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

					<A HREF="../admin/CatM.asp" target=_blank>管理主連結</A>
	</td>
				
    <td valign="top">
    

     <!--#include file="data.inc"-->
				<%set rs1=server.CreateObject("adodb.recordset")
				SQLStr="select * from mainpagetitle"
				set rs1=conn.execute(SQLStr)
				pagetitle=rs1("title")
				%>		
		<td valign="top"> 
		<TABLE WIDTH=100% ALIGN=center BORDER=1 CELLSPACING=1 CELLPADDING=1>
	<TR>
		<TD><center><STRONG><%=pagetitle%></STRONG>
		<A HREF="../admin/mainpageM-title-p0.asp" target=_blank>修改標題</A>
		
		</center></TD>
	</TR>
</TABLE>
<A HREF="../admin/mainpageM-layout.asp" target=_blank>修改所有首頁圖文排版方式</A><br> 
<A HREF="../admin/mainpageM-add.asp" target=_blank>新增段落圖文</A><p> 
			
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
						Table1="<table  border='1' cellspacing='0' cellpadding='10'>"
						table1_1="<tr><td align=left><b>" 
						table1_2="</b></td></tr>"
						Table2= "<tr><td class='9pt-black'>" 
						TABLE2_1=""
						Table3= "</td>"
						Table4= "</tr>"
						Table5= "</table><br>"
							if rs("layout")=1 then
								imgsrc= "<img name='imgshown' src='" & strName & "' border=0 align=left>"
								contentText=rs("content1")
								showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2  & imgsrc & contentText & table3 & table4 & table5
							elseif rs("layout")=2 then
								imgsrc= "<img name='imgshown' src='" & strName & "' border=0 align=right>"
								contentText=rs("content1")
								showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2  & imgsrc & contentText & table3 & table4 & table5
							''注意事項
							''3之後的樣式待規格提出後修正,目前未修正
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
							else ''未定義者先以樣式一為預設樣式
								imgsrc= "<img name='imgshown' src='" & strName & "' border=0 align=left>"
								contentText=rs("content1")
								showLayout=Table1 & table1_1 & rs("title") & table1_2 &  Table2  & imgsrc & contentText & table3 & table4 & table5
							end if
					   Else
					       imgsrc=""
					     
					   End If
				
				
				Response.Write "<A HREF=../admin/mainpageM.asp?id=" & rs("id") & " target=_blank>修改以下之圖文</A>"
				Response.Write  showLayout 
				
				
				rs.MoveNext 
				loop
							
				end if
				rs.Close 
				set rs=nothing
				'conn.close
				'set conn=nothing	
				%>
				
				
								
		</td>
	</tr>
</table>
<!--#include file="data.inc"-->
<%
set rs=server.CreateObject("adodb.recordset")
	SQLStr="select * from template"
	set rs=conn.execute(SQLStr)
if err.number <> 0 then
		response.write "數據庫操作錯誤︰" + err.description
		err.clear
else
if not rs.EOF then 
	rs.MoveFirst 
	%>

<table width="100%" border="1" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td align="center" valign="middle"> 
      <div align="center"><%=rs("copyright")%><br>
      <A HREF="../admin/copyright-p0.asp" target=_blank>管理版權所有宣告</A>
      
      </div>
    </td>
  </tr>
</table>
<%
end if
end if
rs.Close 
set rs=nothing
conn.close
set conn=nothing	



%>
</body>
</html>

