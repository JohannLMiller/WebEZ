<%@ Language=VBScript %>
<%
catNo=Request.QueryString("catNo")
subcat=Request.QueryString("subcat")
subcatNo=Request.QueryString("subcatNo")

Response.Write catNo
Response.Write subcatNo

%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<form name="form1" method="post" action="subCatpageM-layout-p0.asp">
  <p>
   <input type="radio" name="layout" value="1">
    �˦��@<br>
    <br>
    <input type="radio" name="layout" value="2">
    �˦��G<br>
    <br>
    <input type="radio" name="layout" value="3">
    �˦��T<br>
    <br>
    <input type="radio" name="layout" value="4">
    �˦��|<br>
    <br>
    <input type="radio" name="layout" value="5">
    �˦���<br>
    <br>
    <input type="radio" name="layout" value="6">
    �˦���<br>
    <br>
  <p>
    <INPUT type="hidden" id=text1 name=catNo value=<%=catNo%>>
    <INPUT type="hidden" id=text2 name=subcatNo value=<%=subcatNo%>>
    <input type="submit" name="Submit" value="Submit">
  </p>
</form>
<P>&nbsp;</P>

</BODY>
</HTML>
