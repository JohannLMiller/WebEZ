   <%
   dim conn   
   dim connstr
   
   'on error resume next
   call conn_init()

   sub conn_init()
       on error resume next
       connstr= "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & Server.MapPath("database/DMS.MDB")
       'connstr="DBQ="+server.mappath("DMS.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"       
       set conn=server.createobject("ADODB.CONNECTION")
       if err.number<>0 then 
           err.clear
           set conn=nothing
		   response.write "�ƾڮw�s���X���T"
		   Response.End
       else
           conn.open connstr
           if err then 
              err.clear
              set conn=nothing
			   response.write "�ƾڮw�s���X���T"
              Response.End 
           end if
       end if   
  end sub
	
  sub endConnection()
      conn.close
      set conn=nothing
  end sub
  
%>