<%@ Language=VBScript %>
<!--#include file="database.asp"-->
<!--#include file="data.inc"-->
<% 
catNo=Request.QueryString("catNo")
cat=Request.QueryString("cat")

if Request.Form("send")<>"" then

catNo=int(Request.Form("catNo"))
cat=Request.Form("cat")

set rs=server.CreateObject("adodb.recordset")
 RS.Open "cat",conn,1,3
  rs.MoveFirst
 
 do while not rs.EOF 
    if rs("catNo")=catNo then
       rs("cat")=cat
    end if
    rs.MoveNext
 loop  
     
Response.Redirect "catM.asp"
end if%> 
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
<FORM action="catM-edit.asp" id=FORM1 method=post name=edit>
<input type=hidden id=no name=catNo value="<%=catNo%>">
  <table width="580" border="0" cellpadding="0" cellspacing="1">
    <tr> 
      <td>&nbsp;</td>
      <td> 
        <table width="100%" border=0 cellpadding=4 cellspacing=1>
          <tr> 
            <td>主類別名稱 ：</td>
            <td > 
              <input type="text" name=cat value="<%=cat%>" size="60">
            </td>
            <td> 
              <input type="submit" value="確定修改" id=send name=send onClick="MM_validateForm('cat','','R');return document.MM_returnValue">
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
  <P> 　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　 
</FORM>

  </BODY>
</HTML>
