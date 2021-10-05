<%
 ' dim Connstr,conn
  Connstr = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & Server.MapPath("../admin/database/DMS.MDB")
  set conn=server.CreateObject("adodb.connection")         
  conn.Open Connstr
%>