<!-- This example shows how to use the Directory Listing feature to delete
arbitrary files from a directory on the server.-->

<HTML>
<HEAD>
<TITLE>AspUpload - Delete Files</TITLE>
</HEAD>
<BODY>
<BASEFONT FACE="Arial" SIZE="2">
<H3>File Deletion</H3>
<P>
For security reasons the command <B>Upload.DeleteFile</B> (line 28)
is commented out. Un-comment it to make this sample functional.
Notice that this will make your root directory vulnerable
to an outside attack as anyone will be able to delete files from it.
<P>
<%
	Directory = "../images/logo" ' initial directory
	Set Upload = Server.CreateObject("Persits.Upload.1")
	Set Dir = Upload.Directory( Directory & "*.*", SORTBY_NAME)

	' perform deletions if this is a form submission
	If Request("Delete") <> "" Then
		For Each Item in Request("FileName")
			Response.Write "<B>Deleting File " & Item & "</B><BR>"
			
			' uncomment next line to enable deletions.
			'Upload.DeleteFile Directory & Item 
		Next
	End If
%>

<h3><% = Dir.Path %></h3>

<FORM ACTION="DeleteFiles.asp" METHOD="POST">
<TABLE BORDER="1" CELLSPACING="0" CELLPADDING="0">
<TR><TH>&nbsp;</TH><TH>Name</TH><TH>Size</TH></TR>
<% For Each File in Dir %>
<% If Not File.IsSubdirectory Then %>
<TR>
	<TD><INPUT TYPE="CHECKBOX" VALUE="<% = Server.HTMLEncode(File.FileName)%>" NAME="FileName"></TD>
	<TD><% = File.FileName %></TD>
	<TD><% = File.Size %></TD>
</TR>
<% End If
Next %>
<TR>
<TD COLSPAN="3"><INPUT TYPE="SUBMIT" NAME="Delete" VALUE="Delete selected files"></TD>
</TR>
</TABLE>
</FORM>

</BASEFONT>
</BODY>
</HTML>
