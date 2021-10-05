<%@ Language=VBScript %>

<!--#include file="database.asp"-->
<!--#include file="data.inc"-->

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Content-Type" content="text/html; charset=big5">

<script language="JavaScript">
<!--
function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}

function MM_validateForm() { //v4.0
  var i,p,q,nm,test,num,min,max,errors='',args=MM_validateForm.arguments;
  for (i=0; i<(args.length-2); i+=3) { test=args[i+2]; val=MM_findObj(args[i]);
    if (val) { nm=val.name; if ((val=val.value)!="") {
      if (test.indexOf('isEmail')!=-1) { p=val.indexOf('@');
        if (p<1 || p==(val.length-1)) errors+='- '+nm+' must contain an e-mail address.\n';
      } else if (test!='R') {
        if (isNaN(val)) errors+='- '+nm+' must contain a number.\n';
        if (test.indexOf('inRange') != -1) { p=test.indexOf(':');
          min=test.substring(8,p); max=test.substring(p+1);
          if (val<min || max<val) errors+='- '+nm+' must contain a number between '+min+' and '+max+'.\n';
    } } } else if (test.charAt(0) == 'R') errors += '- '+nm+' is required.\n'; }
  } if (errors) alert('The following error(s) occurred:\n'+errors);
  document.MM_returnValue = (errors == '');
}
//-->
</script>
</head>
<body>
<%
''列出所有管理者ID
set rs=server.CreateObject("adodb.recordset")
	SQLStr="select * from admin  "
	set rs=conn.execute(SQLStr)%>
	<table WIDTH="75%" BORDER="1" CELLSPACING="1" CELLPADDING="1">
<tr>
		<td>系統識別編號</td>
		<td>使用者名稱</td>
		<td>備註</td>
		<td>修改</td>
		<td>刪除</td>
	</tr>
	<%do while not rs.EOF %>
	<tr>
		<td><%=rs("autoNo")%></td>
		<td><%=rs("id")%></td>
		<td><%=rs("ps")%></td>
		<td><a HREF="adminM-edit.asp?autoNo=<%=rs("autoNo")%>&amp;ID=<%=rs("ID")%>&amp;PWD=<%=rs("PWD")%>&amp;PS=<%=rs("PS")%>">修改</a></td>
		<td><a HREF="adminM-del.asp?autoNo=<%=rs("autoNo")%>">刪除</a></td>
	</tr>
	
	<%
	rs.MoveNext 
	loop
	%>
</table>


<p>&nbsp;</p>
<p>新增管理者</p>
<form name="form1" method="post" action="adminM-add.asp">
  <table width="80%" border="1">
    <tr>
      <td>ID</td>
      <td>Password</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>
        <input type="text" name="ID">
      </td>
      <td>
        <input type="text" name="PWD">
      </td>
      <td>
        <select name="menu1">
          <option selected>unnamed1</option>
        </select>
      </td>
      <td>
        <input type="text" name="PS">
      </td>
      <td>&nbsp;</td>
      <td>
        <input type="submit" name="Submit" value="Submit" onClick="MM_validateForm('ID','','R','PWD','','R');return document.MM_returnValue">
      </td>
    </tr>
  </table>
</form>
<p>&nbsp;</p>
</body>
</html>
<%
rs.Close 
set rs=nothing
conn.close
set conn=nothing	
	



%>