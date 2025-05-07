<%@ LANGUAGE="VBScript" %>
<HTML>
<HEAD>
<TITLE>
IRP SoftRobot DELETE REPORT (sales@erpweb)
</TITLE>
</HEAD>
<BODY topmargin=0>
<%
Set Conn=Server.CreateObject("ADODB.Connection")
Conn.Open Session("erp_ConnectionString")
%>
<%
Sub SMLIST ( )

	Set rs1=server.createobject("ADODB.Recordset")
	rs1.open "SELECT DID, NAME FROM DOCNAMELIST WHERE SYSTEMGEN=0", Conn
	%>
	<SELECT name=DID>
	<%WHILE not rs1.eof%>
	<OPTION VALUE=<%=rs1(0)%>><%=rs1(1)%> 
	<%rs1.MoveNext%>
	<%wend%>
	</select>
	<%rs1.Close%>

<%
End Sub
%>
<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF >
<TR>
<TD valign=middle bgcolor=#D4D4D4 bordercolordark=#FFFFFF bordercolor=#808080 bordercolorlight=#808080>&nbsp;&nbsp;<FONT FACE=ARIAL SIZE=2><img src=ofolder.gif border=0 alt='SoftRobot Document Server'><B> Delete Report/Function Manager</B></FONT></TD>
</TR>
</table>
<%MODE=REQUEST("MODE")%>
<%IF MODE="" THEN%>
<FORM NAME=FORMX ACTION=deletereport.asp METHOD=POST>
<TABLE WIDTH=100% >
<TR bgcolor=#E1F2FD>
<TD><FONT FACE=ARIAL SIZE=1 COLOR=WHITE>Report Name:</FONT></TD>
<TD><%SMLIST%></TD><INPUT TYPE=HIDDEN NAME=MODE VALUE=1>
<td><INPUT TYPE=SUBMIT NAME=CREATE VALUE="Delete Report Document & Workflow">
</td>
</TR>
</table>
</FORM>
<%ELSE%>
<%DID=REQUEST("DID")%>
<hr>
<%
SQLINSRT="DELETE FROM WORKFLOW WHERE DID=" & DID
Conn.Execute SQLINSRT
Response.Write SQLINSRT
%>
<hr>
<%
SQLINSRT="DELETE FROM DOCUMENTS1 WHERE DID=" & DID
Conn.Execute SQLINSRT
Response.Write SQLINSRT
%>
<%END IF%>
<hr>
<font face="Arial" size="1">
&#169; Copyright 2005 . All rights reserved. IRP
is registered trademark of ERPWEB.
<a href=home.html>Home</a> | <a href=mailto:sales@erpweb>Contact</a>
</font>
</BODY>
</HTML>