<HTML>
<HEAD>
<TITLE>
IRP DATABASE PROCEDURES SERVER(sales@erpweb)
</TITLE>
</HEAD>
<BODY topmargin=0>
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
    Response.Write("<TD ALIGN=RIGHT><font face=arial size=1><a href=procshow.asp?NAME=" & rs("NAME") & "&ID=" & rs("ID") & " target=news>Show Data</a> / <a href=procedure.asp?mode=1&name=" & rs(1) & " target=news>Delete</a></font></TD>")
    rs.MoveNext
  wend 
  Response.Write ("<TR bgcolor=yellow><TD COLSPAN=3><FORM METHOD=POST ACTION=procedure.asp?mode=2><textarea rows=3 cols=50 NAME=PROCNAME>CREATE PROCEDURE procedurename @varname datatype AS SELECT * FROM tablename WHERE fieldname=@varname</textarea></TD><TD ALIGN=RIGHT><FONT FACE=ARIAL SIZE=1><INPUT TYPE=SUBMIT NAME=S VALUE='Add Procedure'></FORM></FONT></TD></TR>")
  Response.Write( "</TABLE>" )

End Sub
%>
<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF >
<TR>
<TD valign=middle bgcolor=#D4D4D4 bordercolordark=#FFFFFF bordercolor=#808080 bordercolorlight=#808080>&nbsp;&nbsp;<FONT FACE=ARIAL SIZE=2><img src=ofolder.gif border=0 alt='SoftRobot Document Server'><B> SoftRobot Procedures Server</B></FONT></TD>
</TR>
</table>

<TABLE WIDTH=100% >
<TR bgcolor=#E1F2FD><TD><FONT FACE=ARIAL SIZE=2 > Edit Stored Procedures:</FONT></TD></TR>
</TABLE>
<%mode=request("mode")
if mode="" then
	set rs = Server.CreateObject("ADODB.Recordset") 
	SQL= "select ID, NAME, CRDATE AS DATE_CREATED from sysobjects where uid=1 and xtype='P' ORDER BY NAME"
	rs.Open SQL, Conn
	GenerateTable rs
	rs.Close
elseif mode=1 then
	DROPPROC="DROP PROCEDURE " & request("name")
	Response.Write DROPPROC 
	Conn.Execute DROPPROC
elseif mode=2 then
	PROCNAME=REQUEST("PROCNAME")
	Response.Write PROCNAME
	Conn.Execute PROCNAME
end if
%>
<%
Conn.Close
Set Conn=Nothing
%>
<hr>
<font face="Arial" size="1">
&#169; Copyright 2001 . All rights reserved. IRP
is registered trademark of ERPWEB.
<a href=home.html>Home</a> | <a href=mailto:sales@erpweb>Contact</a>
</font>
</BODY>
</HTML>