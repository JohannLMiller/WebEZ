<%@ Language=VBScript %>
<% option explicit %>
<!--#include file="database.asp"-->
<%
	dim sql
	dim rs
	
	sql="select pwd from siteman"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,conn,1,1
    if err.number<>0 then 
		response.write "數據庫操作失敗︰"&err.description
    else
		if rs.bof and rs.eof then
			response.write "<center>對不起，數據庫無法操作。</center>"
   			rs.close
	    else
			if request("password")<>rs("pwd") then
		        response.write "<center>對不起，請輸入正確的站長口令。</center>"
			    rs.close		    
			else
				rs.Close
				session("adminOK")="true"
				set rs=nothing
				call endConnection()
				Response.Redirect "main.html"
			end if
		end if
	end if
	set rs=nothing
	
	call endConnection()	
%>