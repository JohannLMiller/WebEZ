<%@ Language=VBScript %>
<%
catNo=Request.QueryString("catNo")
Response.Write catNo
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<form name="form1" method="post" action="CatpageM-layout-p0.asp">
  <p>
   <input type="radio" name="layout" value="1">
    樣式一<br>
    <br>
    <input type="radio" name="layout" value="2">
    樣式二<br>
    <br>
    <input type="radio" name="layout" value="3">
    樣式三<br>
    <br>
    <input type="radio" name="layout" value="4">
    樣式四<br>
    <br>
    <input type="radio" name="layout" value="5">
    樣式五<br>
    <br>
    <input type="radio" name="layout" value="6">
    樣式六<br>
    <br>
  <p>
    <INPUT type="hidden" id=text1 name=catNo value=<%=catNo%>>
    <input type="submit" name="Submit" value="Submit">
  </p>
</form>
<P>&nbsp;</P>

</BODY>
</HTML>
