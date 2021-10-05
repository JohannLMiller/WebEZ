<%@ Language=VBScript %>
<!--#include file="database.asp"-->
<!--#include file="data.inc"-->
<% 
subcatNo=Request.QueryString("subcatNo")
subcat=Request.QueryString("subcat")
catNo=Request.QueryString("catNo")

if Request.Form("send")<>"" then

subcatNo=int(Request.Form("subcatNo"))
subcat=Request.Form("subcat")
catNo=Request.Form("catNo")

set rs=server.CreateObject("adodb.recordset")
 RS.Open "subcat",conn,1,3
  rs.MoveFirst
 
 do while not rs.EOF 
    if rs("subcatNo")=subcatNo then
       rs("subcat")=subcat
    end if
    rs.MoveNext
 loop  
     
Response.Redirect "subcatM.asp?catNo="& catNo
end if%> 
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Content-Type" content="text/html; charset=big5">


</HEAD>
<BODY>
<FORM action="subcatM-edit.asp" id=FORM1 method=post name=edit>
<input type=hidden id=no name=subcatNo value="<%=subcatNo%>">
<input type=hidden id=no name=catNo value="<%=catNo%>">

  <table width="580" border="0" cellpadding="0" cellspacing="1">
    <tr> 
      <td>&nbsp;</td>
      <td> 
        <table width="580" border=0 cellpadding=4 cellspacing=1>
          <tr> 
            <td width="108">次類別名稱 ：</td>
            <td width="292"> 
              <input type="text" name=subcat value="<%=subcat%>">
            </td>
            <td width="152"> 
              <input type="submit" value="確定修改" id=send name=send >
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
