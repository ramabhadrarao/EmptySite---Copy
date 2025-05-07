<HTML>
<HEAD><TITLE>IRP CUSTOMER LOGIN (sales@erpweb)</TITLE></HEAD>
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
    <b><font color="#FFFF00" face="arial" size="1">For Registered Customer</font></b>
    </td>
</tr>
<tr>
<td bgcolor=#6699CC width="155">
<font face=arial size=1 COLOR=black>
<form method=post action=custmenu.asp>
<b>Username:</b><br>
<INPUT TYPE=TEXT NAME="TEXT1" value="CUSTOMER" size=12><br>
<b>Password:</b><br>
<INPUT TYPE=PASSWORD NAME="TEXT2" value="CUSTOMER" size=12><br>
<b>Customer Number:</b><br>
<INPUT TYPE=PASSWORD NAME=CNO value=1 size=12><br>
<b>Login Year:</b><br>
<%List1 Conn%>
<INPUT TYPE=SUBMIT VALUE="Login">
</form>
</FONT>
</td>
</tr>
</table>
<table>
<tr>
    <td bgcolor="#336699" width="155">
    <b><font color="#FFFF00" face="arial" size="1">For Un-Register Customer</font></b>
    </td>
</tr>
</table>
<font face=arial size=1 COLOR=black>
<img src="MYDOC.GIF" width=15>&nbsp;<a href="newuser.htm" target="main">Cust. Registration</a><br>
<img src="MYDOC.GIF" width=15>&nbsp;<a href="dsupport/dhload.htm?DID=952&UID=57&FADD=False&FDEL=False&FVIEW=False&FEDIT=False&FILTER=False&FOFFHOLD=False&FREJECT=False&CID=1&CNO=1>" target=main>Product DSS</a><br>
<img src="MYDOC.GIF" width=15>&nbsp;<a href="shoppingcart.asp?DID=952&UID=57&FADD=False&FDEL=False&FVIEW=False&FEDIT=False&FILTER=False&FOFFHOLD=False&FREJECT=False&CID=1&CNO=3" target=main>Online Shopping</a><br>
<img src="MYDOC.GIF" width=15>&nbsp;<a href="profile.asp?DID=1022&UID=57&FADD=False&FDEL=False&FVIEW=False&FEDIT=False&FILTER=False&FOFFHOLD=False&FREJECT=False&CID=1" target=main>Company Profile</a><br>
<img src="MYDOC.GIF" width=15>&nbsp;<a href=homepage.htm target=main>Company News</a><br>
<hr>
<img src="MYDOC.GIF" width=15>&nbsp;<a href=mailto:sales@erpweb>Contact us</a><br>
</FONT>
<%
Conn.Close
Set Conn=nothing
%>
</BODY>
</HTML>
