<HTML>
<HEAD><TITLE>IRP SUPPLIER LOGIN (sales@erpweb)</TITLE></HEAD>
<BODY BGCOLOR=#6699CC topmargin=0 leftmargin=1>
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
<marquee bgcolor=black behavior=alternate loop=1><font size=2 face=arial color=white><b>Welcome to IRP</b></font></marquee>
<table bgcolor=#6699CC width="153">
  <tr>
    <td bgcolor=#6699CC width="155">
    <a href=home.html target=main>
    <img border="0" src="erppack.jpg" width="153" height="199">
    </a>
    </td>
  </tr>
  <tr>
    <td bgcolor="#336699" width="155">
    <b><font color="#FFFF00" face="arial" size="1">For Registered Supplier</font></b>
    </td>
  </tr>
  <tr>
    <td bgcolor=#6699CC width="155">
<font face=arial size=1 COLOR=black>
<form method=post action=suppmenu.asp>
<b>Username:</b><br>
<INPUT TYPE=TEXT NAME="TEXT1" value="SUPPLIER" size=12><br>
<b>Password:</b><br>
<INPUT TYPE=PASSWORD NAME="TEXT2" value="SUPPLIER" size=12><br>
<b>Supplier Number:</b><br>
<INPUT TYPE=PASSWORD NAME="SNO" value="1" size=12><br>
<b>Login Year:</b><br>
<%List1 Conn%>
<INPUT TYPE=SUBMIT VALUE="Login">
</form>
</FONT>
</td>
</tr>
<tr>
    <td bgcolor="#336699" width="155">
    <b><font color="#FFFF00" face="arial" size="1">For Un-Register Supplier</font></b>
    </td>
  </tr>
  <tr>
    <td width="155">
    <b><font face="arial" size="1"></font></b>
    </td>
  </tr>
</table>
<font face=arial size=1>
<img src="MYDOC.GIF" width=15>&nbsp;<a href="newuser.htm" target="main">Supplier Registration</a><br>
<img src="MYDOC.GIF" width=15>&nbsp;<a href="profile.asp?DID=1022&UID=57&FADD=False&FDEL=False&FVIEW=False&FEDIT=False&FILTER=False&FOFFHOLD=False&FREJECT=False&CID=1" target=main>Company Profile</a><br>
<img src="MYDOC.GIF" width=15>&nbsp;<a href=homepage.htm target=main>Company News</a><br>
<hr>
<img src="MYDOC.GIF" width=15>&nbsp;<a href=mailto:sales@erpweb>Contact us</a><br>
</font>
<%
Conn.Close
Set Conn=nothing
%>
</BODY>
</HTML>
