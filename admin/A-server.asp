<% LANGUAGE="VBSCRIPT" %>
<HTML> 
<BODY> 
<% Set Upload = Server.CreateObject("Persits.Upload.1") 
'Count = Upload.Save("C:\abc")
Count = Upload.SaveVirtual("../images/logo/")

 %> 
<% = Count %> files uploaded. 
</BODY> 
</HTML>



