<HTML>
<HEAD>
<TITLE>
IRP SoftRobot Application GENERATOR (sales@erpweb)
</TITLE>
</HEAD>
<BODY topmargin=0>
<%
Set Conn=Server.CreateObject("ADODB.Connection")
Conn.Open Session("erp_ConnectionString")
%>
<%
Sub USERLIST ( n )

	Set rs1=server.createobject("ADODB.Recordset")
	rs1.open "SELECT * FROM USERS ORDER BY UNAME", Conn
	%>
	<SELECT   name=<%="UID" & n%>>
	<%WHILE not rs1.eof%>
	<OPTION VALUE=<%=rs1("UID")%>><%=rs1("UNAME")%> 
	<%rs1.MoveNext%>
	<%wend%>
	</select>
	<%rs1.Close%>

<%
End Sub
%>
<%
Sub SMLIST ( )

	Set rs1=server.createobject("ADODB.Recordset")
	rs1.open "SELECT * FROM SUBMODULE ORDER BY SUBMODULENM", Conn
	%>
	<SELECT name=SUBMODULEID>
	<%WHILE not rs1.eof%>
	<OPTION VALUE=<%=rs1("SUBMODULEID")%>><%=rs1("SUBMODULENM")%> 
	<%rs1.MoveNext%>
	<%wend%>
	</select>
	<%rs1.Close%>

<%
End Sub
%>
<%DDID=REQUEST("DDID")%>

<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF >
<TR>
<TD valign=middle bgcolor=#D4D4D4 bordercolordark=#FFFFFF bordercolor=#808080 bordercolorlight=#808080>&nbsp;&nbsp;<FONT FACE=ARIAL SIZE=2><img src=ofolder.gif border=0 alt='SoftRobot Document Server'><B> SoftRobot Developer: Create New Application - Documents, Folders and Workflow</B></FONT></TD>
</TR>
</table>
<TABLE WIDTH=100% >

<FORM NAME=FORMX ACTION=createdoc.asp METHOD=POST>
<TABLE WIDTH=100% >
<TR bgcolor=#d1d2d3>
<TD COLSPAN=2><FONT FACE=ARIAL SIZE=2><B>NEW DOCUMENT DESIGN:</B></FONT></TD>
<TD><FONT FACE=ARIAL SIZE=2>SUBMODULE/PROCESS</FONT></TD>
</TR>
<TR bgcolor=#E1F2FD>
<TD><FONT FACE=ARIAL SIZE=2 >Document Name:<INPUT TYPE=TEXT NAME=MASTER SIZE=15 maxsize=15></FONT></TD>
<TD><FONT FACE=ARIAL SIZE=2 >Master/Details:<INPUT TYPE=checkbox NAME=DETAILS></FONT></TD>
<TD><%SMLIST%></TD>
</TR>
<TR bgcolor=#d1d2d3>
<TD COLSPAN=2><FONT FACE=ARIAL SIZE=2 ><B>FOLDER/FUNCTION DESIGN:</B></FONT></TD>
<TD><FONT FACE=ARIAL SIZE=2>USERS AND WORKFLOW</FONT></TD>
</TR>
<TR bgcolor=#E1F2FD>
<TD COLSPAN=2><FONT FACE=ARIAL SIZE=1 >USE OPEN REGISTER TO CREATE/MODIFY/DELETE DOCUMENT:</FONT></TD>
<TD><%USERLIST 1%></TD>
</TR>
<TR bgcolor=#E1F2FD>
<TD COLSPAN=2><FONT FACE=ARIAL SIZE=1 >USE APPROVE REGISTER TO APPROVE/ONHOLD/REJECT DOCUMENT:</FONT></TD>
<TD><%USERLIST 2%><INPUT TYPE=TEXT SIZE=2 NAME=NOOFAPPROVALS VALUE=1></TD>
</TR>
<TR bgcolor=#E1F2FD>
<TD COLSPAN=2><FONT FACE=ARIAL SIZE=1 >USE ONHOLD REGISTER TO OFFHOLD DOCUMENT:</FONT></TD>
<TD><%USERLIST 3%></TD>
</TR>
<TR bgcolor=#E1F2FD>
<TD COLSPAN=2><FONT FACE=ARIAL SIZE=1 >USE CLOSED REGISTER AND PRINT DOCUMENT:</FONT></TD>
<TD><%USERLIST 4%></TD>
</TR>
<TR bgcolor=#E1F2FD>
<TD COLSPAN=2><FONT FACE=ARIAL SIZE=1 >USE REJECTED REGISTER AND PRINT DOCUMENT:</FONT></TD>
<TD><%USERLIST 5%></TD>
</TR>
</TABLE>
<textarea name=VALIDATIONCHECK rows=4 cols=80>
Insert your VBScript or JavaScript Client-end scripting code here.
</textarea><hr>
<INPUT TYPE=SUBMIT NAME=CREATE VALUE="Create Document, Folders & Workflow">
</FORM> 
<hr>
<font face="Arial" size="1">
&#169; Copyright 2005 . All rights reserved. IRP
is registered trademark of ERPWEB.
<a href=home.html>Home</a> | <a href=mailto:sales@erpweb>Contact</a> | <a href=softrobot/customize.htm> Help</a>
</font>
</BODY>
</HTML>