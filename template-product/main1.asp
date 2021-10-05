<%@ Language=VBScript %>
<%title=Request.QueryString("cat")%>
<html>
<head>
<title><%=title%></title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="webez.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
		<!--#include file="head.asp"-->
<!--前端超級第一大table-->			
<table width="780" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="140" background="images/link-bg.jpg" valign="top"class="link-bg"><img src="images/link-top.jpg" width="140"><br>
      
      <!--#include file="mainlink.asp"-->
      
      <img src="images/link-down.jpg" width="140" >
    </td>
   <td valign="top" width=640  class="content"  background="images/content-bg.jpg">
		<!--#include file="sublink.asp"-->
		<!--#include file="copyright.asp"-->
	</td>
  </tr>
</table>



</body>
</html>









