<%@ Language=VBScript %>
<!--#include file="database.asp"-->
<!--#include file="data.inc"-->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
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
</HEAD>
<BODY>
<%
''列出所有管理者ID
set rs=server.CreateObject("adodb.recordset")
	SQLStr="select * from cat order by catNo "
	set rs=conn.execute(SQLStr)%>
	
<table WIDTH="75%" BORDER="1" CELLSPACING="1" CELLPADDING="1">
  <tr> 
    <td>系統識別編號</td>
    <td>主類別名稱</td>
    <td>修改</td>
    <td>刪除</td>
  </tr>
  <%do while not rs.EOF %>
  <tr> 
    <td><%=rs("catNo")%></td>
    <td><%=rs("cat")%></td>
    <td><a HREF="catM-edit.asp?catNo=<%=rs("catNo")%>&amp;cat=<%=rs("cat")%>">修改</a></td>
    <td><a HREF="catM-del.asp?catNo=<%=rs("catNo")%>">刪除</a></td>
  </tr>
  <%
	rs.MoveNext 
	loop
	%>
</table>
<p>新增主類別</p>
<form name="form1" method="post" action="catM-add.asp">
  <table width="80%" border="1">
    <tr> 
      <td>主類別名稱</td>
      <td>
<input type="text" name="cat">
      </td>
      <td> 
        <input type="submit" name="Submit" value="Submit" onClick="MM_validateForm('cat','','R');return document.MM_returnValue">
      </td>
    </tr>
  </table>
</form>
<p>&nbsp;</p>

</BODY>
</HTML>
<%
rs.Close 
set rs=nothing
conn.close
set conn=nothing	
	



%>