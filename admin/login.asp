<%@ Language=VBScript %>
<!--#include file="database.asp"-->
<!--#include file="data.inc"-->
<%
ID=Request("ID")
PWD=Request("PWD")
strMsg=checklogin(ID,PWD)

if strMsg="success" then
Response.Redirect "main.asp"
elseif strMsg="�L���ϥΪ�" then
		showstr="�п�J���T��ƶi�J�޲z����"
elseif strMsg="�K�X���~" then
		showstr="�K�X���~"
end if
		

		function checklogin(id,pwd)
		Set rs = Server.CreateObject("ADODB.Recordset")	
			
			SQLstr = "SELECT * FROM admin " & _
					 "WHERE ID = '" & ID & "'" 
			'������d�ߪ�SQL�ԭz
			
			rs.Open SQLstr, conn 
			'�Q��Recordset�������SQL�ԭz, �ëإ߰O����
			if rs.EOF then '�ˬd�O�����ЬO�_����̫�@���O��
				checklogin = "�L���ϥΪ�"
				'�Y�O�����Цb�O�����@�}�ҮɫK����̫�@���O���h��ܨS���O��
			elseif rs("PWD") <> PWD then '�P�_�K�X�O�_���T
				checklogin = "�K�X���~"
			else '�q�L�ˬd��ܱb���P�K�X�����T���\�n�J
				checklogin = "success"
			end if





		end function

%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Content-Type" content="text/html; charset=big5">

</HEAD>
<BODY bgcolor="#FFFFFF">
<form name="form1" method="post" action="login.asp">
  <table width="100%" border="1" cellspacing="0" cellpadding="0" align="center" bordercolor="#52616b">
              <tr bgcolor="#90b5cf" bordercolor="#90b5cf"> 
                <td colspan="2" class="w10-title-small-time" valign="center" height="30" bgcolor="#90b5cf"> 
                  <div align="center"><b>Login</b></div>
                </td>
              </tr>
              <tr valign="center" bgcolor="#f2f2f2" bordercolor="#f2f2f2"> 
                <td class="w8-left-right30" colspan="2" height="40">
      <P align=center><FONT color=#52616b><%=showstr%></FONT></P></td>
              </tr>
              <tr valign="center" bgcolor="#f2f2f2" bordercolor="#f2f2f2"> 
                <td class="w9-right30" width="40%" height="27" bgcolor="#f2f2f2"> 
                  <div align="right"><font color="#52616b" 
     > ID: </font></div>
                </td>
                <td class="w9-left-right30" height="27"> 
                  <div align="left"> 
                    
        <input type="text" name="ID" size="25" style="COLOR: #444444; FONT-FAMILY: Verdana; FONT-SIZE: 8pt; HEIGHT: 17px" 
     >
                  </div>
                </td>
              </tr>
              <tr valign="center" bgcolor="#f2f2f2" bordercolor="#f2f2f2"> 
                <td class="w9-right30" width="40%" height="27" bgcolor="#f2f2f2"> 
                  <div align="right"><font color="#52616b" 
     >Password : </font></div>
                </td>
                <td class="w9-left-right30" height="27"> 
                  
      <input type="password" name="PWD" size="25" style="COLOR: #444444; FONT-FAMILY: Verdana; FONT-SIZE: 8pt; HEIGHT: 17px" 
     >
                </td>
              </tr>
              <tr bordercolor="#f2f2f2"> 
                <td colspan="2" class="w1" valign="center" bgcolor="#f2f2f2"> 
                  <table width="20%" border="1" cellspacing="10" cellpadding="0" bordercolor="#f2f2f2" align="center">
                    <tr bordercolor="#666666" bgcolor="#dddddd" valign="center"> 
                      <td class="w8" bgcolor="#90b5cf" bordercolor="#52616b"> 
                        
              <div align="center"><b><A href="#"><b>
                <input type="submit" name="Submit" value="�T�w">
                </b></a></b></div>
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
          </form>

</BODY>
</HTML>
<%
conn.close
set conn=nothing	



%>