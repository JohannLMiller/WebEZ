<%@ LANGUAGE="VBSCRIPT" %>
<% option explicit%>
<!--#include file="database.asp"-->
<% 
if request("action") = "�T �w" then

dim rs, sql
dim OldPassword  
dim newPassword  
dim newPassword1
dim ErrMsg
dim FoundError

ErrMsg = ""
FoundError=false

oldPassword=request("oldpasswd")
newPassword=request("newpasswd")
newPassword1=request("newpasswd1")
  
if newPassword="" or len(newPassword)>10 then
	FoundError=True
	ErrMsg=ErrMsg & "<li>�s���f�O���ର�ŨåB���פ���j�_10!</li>"
else
	if newPassword<>newPassword1 then
		FoundError=true
		ErrMsg=ErrMsg & "<li>�⦸��J���f�O����!</li>"
	else
set rs=server.createobject("adodb.recordset")
sql="select * from siteman where pwd ='"&OldPassword&"'"
rs.open sql,conn,1,1
if err.number<>0 then 
	response.write "�ƾڮw�ާ@���ѡJ"&err.description
	err.clear
else
	if rs.bof and rs.eof then
		FoundError=true
		ErrMsg=ErrMsg & "<li>�K�X���~!</li>"
	else
		sql = "UPDATE siteman SET pwd = '" + newPassword + "'"
		conn.execute sql
		if not err.number<>0 then 
			response.write "<p align=center>�K�X�ק令�\</p>"+chr(13)+chr(10)
		else 
		FoundError=true
		ErrMsg=ErrMsg & "<li>�ƾڮw�ާ@���ѡA�ХH�Z�A��!<br>"
		ErrMsg = ErrMsg + err.Description + "</li>"
		err.clear
		end if
	end if		
end if
	end if	'if newPassword<>newPassword1 then
end if	'if newPassword="" or len(newPassword)>10 then
if FoundError then
	response.write "<ul>"
	response.write ErrMsg
	response.write "</ul>"
end if
end if	'if request("action") = "�T �w" then
 %>
<html>
<head>
	<title>�޲z�K�X</title>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Content-Type" content="text/html; charset=big5">

</head>

<body>
<form method="POST" name="frmChangePass" action="adminpwd.asp">
  <p align="left"><strong>���A���K�X</strong></p>
  �ª��K�X: <input  name="oldpasswd" size="10" maxlength="10" type="password" value><br>
  �s���K�X: <input  name="newpasswd" size="10" maxlength="10" type="password" value><br>
  �K�X����: <input  name="newpasswd1" size="10" maxlength="10" type="password" value><br>
  <input type="submit" value="�T �w" name="action">
  <input type="reset" value="�M ��" name="action"></p>
</form>

</body>
</html>
