<%@ LANGUAGE="VBScript" %>
<HTML>
<HEAD>
<TITLE>
IRP SoftRobot DELETE DOCUMENT (sales@erpweb)
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
	rs1.open "SELECT * FROM DOCMODULE WHERE SYSTEMGEN=0", Conn
	%>
	<SELECT name=DDID>
	<%WHILE not rs1.eof%>
	<%IF rs1(5)= "Function" or rs1(5)= "Document" or rs1(5)= "Module" or rs1(5)= "SubModule" or rs1(5)= "Users" or rs1(5)= "Workflow" then%>
	<%else%>
	<OPTION VALUE=<%=rs1(3)%>><%=rs1(5)%> 
	<%end if%>
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
<TD valign=middle bgcolor=#D4D4D4 bordercolordark=#FFFFFF bordercolor=#808080 bordercolorlight=#808080>&nbsp;&nbsp;<FONT FACE=ARIAL SIZE=2><img src=ofolder.gif border=0 alt='SoftRobot Document Server'><B> SoftRobot Developer: Delete Application</B></FONT></TD>
</TR>
</table>
<FORM NAME=FORMX ACTION=deletedoc.asp METHOD=POST>
<TABLE WIDTH=100% >
<TR bgcolor=#E1F2FD>
<TD><FONT FACE=ARIAL SIZE=1>Document Name:</FONT></TD>
<TD><%SMLIST%></TD>
<td><INPUT TYPE=SUBMIT NAME=CREATE VALUE="Delete Document & Workflow">
</td>
</TR>

</FORM>
</table>
<hr>
<font face="Arial" size="1">
&#169; Copyright 2001 . All rights reserved. IRP
is registered trademark of ERPWEB.
<a href=home.html>Home</a> | <a href=mailto:sales@erpweb>Contact</a>
</font>
</BODY>
</HTML>