<HTML>
<HEAD>
<TITLE>
IRP SOFTROBOT DATABASE TABLES (sales@erpweb)
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
    IF rs(1)= "DOCTYPE" or rs(1)= "ALERTS" or rs(1)= "TEST" or rs(1)= "TESTDET" or rs(1)= "DOCMODULE" or rs(1)= "DOCUMENTS1" or rs(1)= "MODULE" or rs(1)= "SUBMODULE" or rs(1)= "USERS" or rs(1)= "WORKFLOW" OR rs(1)= "LISTBOX" OR rs(1)= "USERGROUP" OR rs(1)= "dictionary" OR rs(1)= "dtproperties" OR rs(1)= "DTS" OR rs(1)= "DTSDET" then
	Response.Write("<TD ALIGN=RIGHT><font face=arial size=1><a href=database.asp?mode=3&TBL=" & rs("NAME") & " target=news>Show</a> / Edit / Delete</a></font></td>")
	ELSE
    Response.Write("<TD ALIGN=RIGHT><font face=arial size=1><a href=database.asp?mode=3&TBL=" & rs("NAME") & " target=news>Show</a> / <a href=fields.asp?id=" & rs("NAME") & " target=news>Edit</a> </font></TD>")
    END IF
    rs.MoveNext
  wend 
  Response.Write ("<FORM ACTION=database.asp?mode=2 METHOD=POST><TR bgcolor=#E1F2FD><TD COLSPAN=3><TEXTAREA ROWS=3 COLS=55 NAME=TABLENAME>CREATE TABLE tablename( columnname1 DATATYPE, columnname2 DATATYPE )</TEXTAREA></TD><TD ALIGN=RIGHT><FONT FACE=ARIAL SIZE=1><INPUT TYPE=SUBMIT NAME=TABLE VALUE='Add Table'></FORM></FONT></TD></TR>")
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
	'write each row
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

<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF >
<TR>
<TD valign=middle bgcolor=#D4D4D4 bordercolordark=#FFFFFF bordercolor=#808080 bordercolorlight=#808080>&nbsp;&nbsp;<FONT FACE=ARIAL SIZE=2><img src=ofolder.gif border=0 alt='SoftRobot Document Server'><B> SoftRobot Tables Server</B></FONT></TD>
</TR>
</table>

<TABLE WIDTH=100% >
<TR bgcolor=#E1F2FD><TD><FONT FACE=ARIAL SIZE=1> Create / Delete Database Tables & Fields:</FONT></TD></TR>
</TABLE>
<%
mode=REQUEST("mode")
if mode="" then
	set rs = Server.CreateObject("ADODB.Recordset") 
	SQL= "select ID, NAME, CRDATE AS DATE_CREATED from sysobjects where uid=1 and xtype='U' ORDER BY NAME"
	rs.Open SQL, Conn
	GenerateTable rs
	rs.Close
elseif mode=1 then
table=trim(request("table"))
Response.Write "Deleting Table: DROP TABLE " & table
Conn.Execute "DROP TABLE " & table
elseif mode=2 then
table=request("TABLENAME")
Response.Write "Creating Table using query:" & table
Conn.Execute table
elseif mode=3 then
TBL=REQUEST("TBL")
set rs = Server.CreateObject("ADODB.Recordset") 
SQL= "select * from " & TBL
rs.Open SQL, Conn
GenerateTable1 rs
rs.Close
end if
Conn.Close
Set Conn=nothing
%>
<hr>
<font face="Arial" size="1">
&#169; Copyright 2001 . All rights reserved. IRP
is registered trademark of ERPWEB.
<a href=home.html>Home</a> | <a href=mailto:sales@erpweb>Contact</a>
</font>
</BODY>
</HTML>