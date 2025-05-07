<%
USR=REQUEST("USR"): IF USR="" OR ISNULL(USR) THEN USR="MyUserName"
PWD=REQUEST("PWD"): IF PWD="" OR ISNULL(PWD) THEN PWD="MyPassword"
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
	<SELECT   name=CID ID="Select1">
	<%WHILE not lst.eof%>
	<OPTION VALUE=<%=lst("CALENDERID")%>><%=lst("THISYEAR")%>
	<%lst.MoveNext%>
	<%wend%>
	</SELECT>
	<%lst.Close%>
<%
End Sub
%>
<% 
Sub GenerateTable( rs )
  
  while not rs.EOF
  
      v = rs(4)
      if isnull(v) then v = ""
      Response.Write( "<img src=arrow.gif><a href=webpage.asp?webpageid=" + Cstr(rs(3)) + "><span class=hh1>" + Cstr(v) + "</span></a><br>" )
 
    rs.MoveNext
   wend 
     
End Sub
%>
<%
DID=REQUEST("DID")
UID=REQUEST("UID")
CID=REQUEST("CID")
%>
<html>
<head>
<title>ERPWEB Home Page</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="main.css" type="text/css">
</head>
<body bgcolor=black leftmargin=0 topmargin=0 marginwidth=0 marginheight=0 link=black vlink=red alink=blue>

<table width=100% border=0 cellpadding=0 cellspacing=0 ID="Table1">
<tr><td align=left background=images/bg8.gif>
<OBJECT classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"
 codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0"
 WIDTH="770" HEIGHT="240" id="flash" ALIGN="">
 <PARAM NAME=movie VALUE="flash.swf"> <PARAM NAME=quality VALUE=high> <PARAM NAME=bgcolor VALUE=#999999> <EMBED src="flash.swf" quality=high bgcolor=#999999  WIDTH="770" HEIGHT="240" NAME="flash" ALIGN=""
 TYPE="application/x-shockwave-flash" PLUGINSPAGE="http://www.macromedia.com/go/getflashplayer"></EMBED>
</OBJECT>
</td></tr>
</table>
<table width=100% border=0 cellpadding=0 cellspacing=0 >
<tr>
<td align=left background=images/bg2.gif>
<table width=267 border=0 cellpadding=0 cellspacing=0 >
<tr>
<td background=images/bg1.gif align=center width=249>

<!-- Search_Form -->
<table border=0 cellpadding=0 cellspacing=0 >
      <td><font face=arial size=2 color=black>eFRM - Finance Resource Management</font></td>
         
</table>
<!-- /Search_Form -->

</td>
<td><img src=images/tr1.gif width=18 height=40><br></td>
</tr>
</table>
</td>
</tr>
</table>

<!-- /HEAD -->
    


<!-- MIDDLE -->

<table width=100% border=0 cellpadding=0 cellspacing=0 >
      
<tr valign=top>
<td background=images/bg3.gif>

<!-- Left_Column -->

<table width=250 border=0 cellpadding=0 cellspacing=0 background="" ID="Table6">
<tr>

<table width=250 border=0 cellpadding=15 cellspacing=0 background="" ID="Table7">

<tr><td align=left class=news>
<form method=post action=usercockpit.asp >

            <table width="130" height="100%" border="0" cellpadding="0" cellspacing="0" class="text" >
                <tr> 
                  <td height="40" valign="top" ><img src="images/members.gif" width="203" height="40"></td>
                </tr>
                <tr> 
                  <td height="130" align="center" background="images/member_bg.gif">
                  <table width="180" border="0" cellpadding="2" cellspacing="0" class="text" >
                      <tr> 
                        <td width="71"><img src="images/login.gif" width="53" height="17"></td>
                        <td width="101"><font color="#3C3C3C"> 
                          <INPUT TYPE=TEXT NAME=UNAME  VALUE="" size=12 >
                          </font></td>
                      </tr>
                      <tr> 
                        <td><img src="images/password.gif" width="64" height="15"></td>
                        <td><font color="#3C3C3C"> 
                          <INPUT TYPE=PASSWORD NAME=PWD  VALUE="" size=12 >
                          </font></td>
                      </tr>
                      <tr> 
                         <td><input type=image src="images/enter.gif" width="68" height="19"  ></td>
                         <td><%List1 Conn%></td>
                      </tr>
                      </tr>
         </table>
         </form>

<span class=hh1>Explore</span><br>
<font color=black>Our website ... ...</b><br><br>
<%
SQL="SELECT * FROM WEBSITE"
Set rs=Server.CreateObject("ADODB.Recordset")
rs.Open SQL , Conn
GenerateTable rs
rs.Close
%>
</td></tr>
</table>

</td></tr>
</table>

<!-- /Left_Column -->

<br>
</td>
<td width=100% bgcolor=#82A6F3>

<!-- Right_Column -->

<table width=100% border=0 cellpadding=0 cellspacing=0 bgcolor=#5787EF ID="Table8">
<tr align=left valign=top>
<td><img src=images/pic1.gif width=134 height=89><br><img src=images/pic2.gif width=134 height=128><br></td>
<td width=100% background=images/bg6.gif class=bg1><h3>Select Roles:</h3><br>
<table width=100% >
<td><img src=images/pic06.gif > <br>Accounts Manager <br><BR>User: iLEDGER<br>PWD: iLEDGER</td>
<td><img src=images/pic07.gif > <br>Billing Manager <br><BR>User: iINVOICE<br>PWD: iINVOICE</td>
<td><img src=images/pic09.gif > <br>Payments Manager <br><BR>User: iPAYMENTS<br>PWD: iPAYMENTS</td>
<td><img src=images/pic08.gif > <br>Costing Manager <br><BR>User: iCOSTING<br>PWD: iCOSTING</td>
</table>
</td>
</tr>
</table>



<!-- /Right_Column -->

</td>
</tr>
</table>

<!-- /MIDDLE -->


<!-- BOTTOM -->

<table width=100% border=0 cellpadding=0 cellspacing=0 ID="Table10">
<tr><td align=right background=images/bg5.gif><img src=images/bot1.gif width=273 height=62><br></td></tr>
</table>

<!-- /BOTTOM -->


<%
Conn.Close
Set Conn=nothing
%>
</body>
</html>