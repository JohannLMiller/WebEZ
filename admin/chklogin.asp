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
		response.write "�ƾڮw�ާ@���ѡJ"&err.description
    else
		if rs.bof and rs.eof then
			response.write "<center>�藍�_�A�ƾڮw�L�k�ާ@�C</center>"
   			rs.close
	    else
			if request("password")<>rs("pwd") then
		        response.write "<center>�藍�_�A�п�J���T�������f�O�C</center>"
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