<%@ LANGUAGE="VBSCRIPT" %>
<% option explicit%>
<!--#include file="database.asp"-->
<% 
if request("action") = "確 定" then

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
	ErrMsg=ErrMsg & "<li>新的口令不能為空並且長度不能大于10!</li>"
else
	if newPassword<>newPassword1 then
		FoundError=true
		ErrMsg=ErrMsg & "<li>兩次輸入的口令不符!</li>"
	else
set rs=server.createobject("adodb.recordset")
sql="select * from siteman where pwd ='"&OldPassword&"'"
rs.open sql,conn,1,1
if err.number<>0 then 
	response.write "數據庫操作失敗︰"&err.description
	err.clear
else
	if rs.bof and rs.eof then
		FoundError=true
		ErrMsg=ErrMsg & "<li>密碼錯誤!</li>"
	else
		sql = "UPDATE siteman SET pwd = '" + newPassword + "'"
		conn.execute sql
		if not err.number<>0 then 
			response.write "<p align=center>密碼修改成功</p>"+chr(13)+chr(10)
		else 
		FoundError=true
		ErrMsg=ErrMsg & "<li>數據庫操作失敗，請以后再試!<br>"
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
end if	'if request("action") = "確 定" then
 %>
<html>
<head>
	<title>管理密碼</title>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Content-Type" content="text/html; charset=big5">

</head>

<body>
<form method="POST" name="frmChangePass" action="adminpwd.asp">
  <p align="left"><strong>更改你的密碼</strong></p>
  舊的密碼: <input  name="oldpasswd" size="10" maxlength="10" type="password" value><br>
  新的密碼: <input  name="newpasswd" size="10" maxlength="10" type="password" value><br>
  密碼校驗: <input  name="newpasswd1" size="10" maxlength="10" type="password" value><br>
  <input type="submit" value="確 定" name="action">
  <input type="reset" value="清 除" name="action"></p>
</form>

</body>
</html>
