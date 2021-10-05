<% 
if session("adminOK") <> "true" then
	response.redirect "default.asp"
end if
 %>