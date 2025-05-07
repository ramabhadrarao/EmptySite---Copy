<HTML>
<HEAD><TITLE>SoftRobot LOGIN / SECURITY SERVER (sales@ERPWEB.com)</TITLE></HEAD>
<BODY BGCOLOR=#6699CC topmargin=0 leftmargin=1>
<%
USR=REQUEST("USR"): IF USR="" OR ISNULL(USR) THEN USR="GUEST"
PWD=REQUEST("PWD"): IF PWD="" OR ISNULL(PWD) THEN PWD="GUEST"
%>
<%
Set Conn=Server.CreateObject("ADODB.Connection")
Conn.Open Session("erp_ConnectionString")
%>
<%
Sub LIST1 ( Conn )
Set lst=Server.CreateObject("ADODB.Recordset")
lst.Open "Select CALENDERID, THISYEAR FROM CALENDER", Conn
%>
	<SELECT   name=CID>
	<%WHILE not lst.eof%>
	<OPTION VALUE=<%=lst("CALENDERID")%>><%=lst("THISYEAR")%>
	<%lst.MoveNext%>
	<%wend%>
	</SELECT>
	<%lst.Close%>
<%
End Sub
%>
<marquee bgcolor=black behavior=alternate loop=1><font size=2 face=arial color=white><b>Login / Security Server</b></font></marquee>
<table bgcolor=#6699CC width="153">
  <tr>
    <td bgcolor=#6699CC width="155">
    <a href=home.html target=main>
     <img border="0" src=softrobot.gif width="163" height="29">
    </a>
    </td>
  </tr>
  <tr>
    <td bgcolor="#336699" width="155">
    <b><font color="#FFFF00" face="arial" size="1">ERPWEB ITM / SoftRobot Login&nbsp;</font></b>
    </td>
  </tr>
  <tr>
    <td bgcolor=#6699CC width="155">
<font face=arial size=1 COLOR=black>
<form method=post action=menu.asp>
<b>Username:</b><br>
<INPUT TYPE=TEXT NAME="TEXT1" value=<%=USR%> size=12><br>
<b>Password:</b><br>
<INPUT TYPE=PASSWORD NAME="TEXT2" value=<%=PWD%> size=12><br>
<b>SoftRobot Development Tools:</b><br>
<select name=mode>
<option value=1>Testing Manager
<option value=2>Application Manager
<option value=6>MenuDesign Manager
<option value=3>Workflow Manager
<option value=4>Database Manager
<option value=5>Reports Manager
</select>
<b>Login Year:</b><br>
<%List1 Conn%>
<INPUT TYPE=SUBMIT VALUE="Login">
</form>
</FONT>
    </td>
  </tr>
  <tr>
    <td bgcolor="#336699" width="155">
    <b><font color="#FFFF00" face="arial" size="1">For Un-registered User</font></b>
    </td>
  </tr>
  <tr>
    <td bgcolor=#6699CC width="155">
<font face=arial size=1 COLOR=black>
  <a href=FORMADD.ASP?UID=85&DID=1448 target="main">New User Registration</a><br>
  (Sign as GUEST)
  </td>
  </tr>
</table>
<%
Conn.Close
Set Conn=nothing
%>
</BODY>
</HTML>
