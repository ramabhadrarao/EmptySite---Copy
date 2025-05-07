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
<HTML>
<HEAD><TITLE>ERPWEB USER LOGIN SERVER(sales@ERPWEB.com)</TITLE>
</HEAD>
<BODY  topmargin=0 leftmargin=0 background="images/member_bg.gif">
<img src=logo.gif>
<img src="images/members.gif" >
<form method=post action=startmenu.asp ID="Form1">
            <table width="130" border="0" cellpadding="0" cellspacing="0" class="text" >
                 <tr> 
                  <td height="110" valign="top" background="images/member_bg.gif">
                  <table width="180" border="0" cellpadding="2" cellspacing="0" class="text" >
                      <tr> 
                        <td width="71"><img src="images/login.gif" width="53" height="17"></td>
                        <td width="101"><font color="#3C3C3C"> 
                          <INPUT TYPE=TEXT NAME="TEXT1"  VALUE="" size=12 ID="Text1">
                          </font></td>
                      </tr>
                      <tr> 
                        <td><img src="images/password.gif" width="64" height="15"></td>
                        <td><font color="#3C3C3C"> 
                          <INPUT TYPE=PASSWORD NAME="TEXT2"  VALUE="" size=12 ID="TEXT2">
                          </font></td>
                      </tr>
                      <tr> 
                         <td><input type=image src="images/enter.gif" width="68" height="19" ></td>
                         <td><%List1 Conn%></td>
                      </tr>
				</table>
       </td>
   </tr>
   <tr>
    <TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080 width="180">
    <b><font face="arial" size="2" color=gray>For Un-registered Users</font></b>
    </td>
  </tr>
  <tr>
    <td width="155">
<font face=arial size=1 COLOR=black>
  <a href=FORMADD.ASP?UID=85&DID=1448 target="main">Step1. New User - Click here</a><br>
  Step2: Provide your email address in ADDRESS BOX<BR>
  Step3: Sign as GUEST to submit application for approval. <br>
  Step4: After approval you will receive email from administrator about password.<BR>
  Step5: Sign in MobileERP, ERPWEB to work as Employee, Customer, Supplier, Others etc.
  </td>
  </tr>
  <tr>
     <tr>
    <TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080>
    <b><font face="arial" size="2" color=gray>For Free Trial Users</font></b>
    </td>
  </tr>
  <td>
  <a href=login2.asp target=main><img src=images/pic03.gif border=0 alt="Click here to see pre-defined roles" width=200></a><br>
<font face=ARIAL size=1><b>Step 1. Click above or  <a href=login2.asp target=main>here</a> to see Pre-Defined Users Roles.</b></font><br>
<font face=ARIAL size=1><b>Step 2. Identify Roles that matches your Corporate identity.</b></font><br>
<font face=ARIAL size=1><b>Step 3. Login system with Username, Password for Roles matching you.</b></font><br>
<font face=ARIAL size=1><b>Step 4. Note: This roles will not work in Live System.</b></font>
<hr>
<a href=iagree.asp target=main><font face=ARIAL size=1><b>Read License Agreement</b></font></a><br>
<font face=ARIAL size=1><b>Copyright 2006. All Rights Reserved. MobileERP.net, ERPWEB.com, SoftRobot.net</b></font>
</td></tr>
</table>
</form>
<%
Conn.Close
Set Conn=nothing
%>
</BODY>
</HTML>
