<HTML>
<HEAD>
<TITLE>
IRP SoftRobot VIEW Server(sales@erpweb)
</TITLE>
</HEAD>
<BODY TOPMARGIN=0>
<%
Set Conn=Server.CreateObject("ADODB.Connection")
Conn.Open Session("erp_ConnectionString")
%>
<% 
Sub GenerateTable( rs )

  Response.Write( "<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF>" )

  ' set up column names
  for i = 0 to rs.fields.count - 1
    Response.Write("<TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080><font face=arial size=1>" + rs(i).Name + "</font></TD>")
  next
	Response.Write("<TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080><font face=arial size=1>Select</font></TD>")

  ' write each row
  while not rs.EOF
    Response.Write( "<TR bgcolor=#E1F2FD>" )
    for i = 0 to rs.fields.count - 1
      v = rs(i)
      if isnull(v) then v = ""
      Response.Write( "<TD VALIGN=TOP><font face=arial size=1>" + CStr( v ) + "</font></TD>" )
    next
    IF rs(1)= "TESTREG" or rs(1)= "TESTDETREG" or rs(1)= "DOCMODULEREG" or rs(1)= "DOCUMENTS" or rs(1)= "MODULEREG" or rs(1)= "SUBMODULEREG" or rs(1)= "USERSREG" or rs(1)= "WORKFLOWREG" OR rs(1)= "LISTBOXREG" OR rs(1)="DOCTABLEREG" OR rs(1)="DOCUMENTSREG" OR rs(1)="WFLOWDESIGN" or rs(1)="WORKBENCH" then
	Response.Write("<TD ALIGN=RIGHT><font face=arial size=1><a href=view.asp?MODE=3&TBL=" & rs("NAME") & " target=news>Show</a> / Edit / Delete </a></font></td>")
    ELSE 
    Response.Write("<TD ALIGN=RIGHT><font face=arial size=1><a href=view.asp?MODE=3&TBL=" & rs("NAME") & " target=news>Show</a> / <a href=viewshow.asp?ID=" & rs("ID") & "&TBL=" & rs("NAME") & " target=news>Edit</a> / <a href=view.asp?MODE=2&VIEWNM=" & rs("NAME") & " target=news>Delete</a></font></TD>")
    END IF
    rs.MoveNext
  wend 
  Response.Write( "</TABLE>" )

End Sub
%>
<% 
Sub GenerateTable1( rs )

  Response.Write( "<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF>" )
  ' set up column names
  for i = 0 to rs.fields.count - 1
    Response.Write("<TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080><font face=arial size=1>" + rs(i).Name + "</font></TD>")
  next
	
  ' write each row
  while not rs.EOF
    Response.Write( "<TR bgcolor=#E1F2FD>" )
    for i = 0 to rs.fields.count - 1
      v = rs(i)
      if isnull(v) then v = ""
      Response.Write( "<TD VALIGN=TOP><font face=arial size=1>" + CStr( v ) + "</font></TD>" )
    next
    rs.MoveNext
  wend 
  Response.Write( "</TABLE>" )

End Sub
%>
Help: Press Ctrl+F to find a perticular Query or View
<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF >
<TR>
<TD valign=middle bgcolor=#D4D4D4 bordercolordark=#FFFFFF bordercolor=#808080 bordercolorlight=#808080>&nbsp;&nbsp;<FONT FACE=ARIAL SIZE=2><img src=ofolder.gif border=0 alt='SoftRobot Document Server'><B> SoftRobot  Database Query/Views Server Folder</B></FONT></TD>
</TR>
</table>
<TABLE WIDTH=100% >
<TR bgcolor=#E1F2FD><TD><FONT FACE=ARIAL SIZE=1 > Create / Delete Database Views:</FONT></TD></TR>
</TABLE>
<%
MODE=REQUEST("MODE")
IF MODE="" THEN
	set rs = Server.CreateObject("ADODB.Recordset")
	SQL= "select ID, NAME, CRDATE AS DATE_CREATED from sysobjects where uid=1 and xtype='V' ORDER BY NAME"
	rs.Open SQL, Conn
	GenerateTable rs
	rs.Close
%>
<FORM NAME=UP ACTION=view.asp METHOD=POST>
<TEXTAREA ROWS=10 COLS=70 NAME=VIEW>
CREATE VIEW viewname AS SELECT * FROM tablename
</TEXTAREA>
<INPUT TYPE=HIDDEN VALUE=1 NAME=MODE>
<INPUT TYPE=SUBMIT NAME=S VALUE="ADD NEW VIEW">
</FORM>
<%ELSEIF MODE=1 THEN%>
<%VIEW=REQUEST("VIEW")%>
CREATE VIEW: <%=VIEW%><br>
<%Conn.Execute VIEW%>
<%ELSEIF MODE=2 THEN%>
<%VIEWNM=REQUEST("VIEWNM")%>
DROP VIEW <%=VIEWNM%><BR>
<%Conn.Execute "DROP VIEW " & VIEWNM%>
<%ELSEIF MODE=3 THEN%>
<%
TBL=REQUEST("TBL")
set rs = Server.CreateObject("ADODB.Recordset")
rs.Open TBL, Conn
GenerateTable1 rs
rs.Close
%>
<%END IF%>
<%
Conn.Close
Set Conn=nothing
%>
<hr>
<font face="Arial" size="1">
&#169; Copyright 2005 . All rights reserved. SoftRobot
is registered trademark of ERPWEB.
<a href=home.html>Home</a> | <a href=mailto:sales@erpweb>Contact</a>
</font>
</BODY>
</HTML>