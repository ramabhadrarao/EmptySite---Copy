<%Response.Buffer=TRUE%>
<html>
<HEAD>
<TITLE>IRP SoftRobot DATABASE VIEW (sales@erpweb)</TITLE>
</HEAD>
<BODY topmargin=0>
<%ID=REQUEST("ID")%>
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
	Response.Write("<TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080><font face=arial size=1>Action</font></TD>")
 
  ' write each row
  while not rs.EOF
    Response.Write( "<TR bgcolor=#E1F2FD>" )
    for i = 0 to rs.fields.count - 1
      v = rs(i)
      if isnull(v) then v = ""
      Response.Write( "<TD VALIGN=TOP><font face=arial size=1>" + CStr( v ) + "</font></TD>" )
    next
    if rs(0) = "erun" or rs(0) = "IRP" or rs(0) = "master" or rs(0) = "msdb" or rs(0) = "model" or rs(0) = "tempdb" then
    Response.Write( "<TD><font face=arial size=1>Detach</font></TD>" )
    else
    Response.Write("<TD><font face=arial size=1><a href=dbview.asp?MODE=2&DBNAME=" & rs(0) & " target=news>Detach</a></font></TD>")
    end if
    rs.MoveNext
  wend 
  ATTACHDB= "EXECUTE sp_attach_db @dbname = databasename, @filename1 = 'c:\mssql7\data\filename.mdf', @filename2 = 'c:\mssql7\data\filename_log.ldf'"
  LDEVICE= "EXECUTE sp_addumpdevice 'disk', devicename,'c:\MSSQL7\BACKUP\devicename.dat'"
  BKUPDB="BACKUP DATABASE databasename TO devicename"
  RESTORE="RESTORE DATABASE databasename FROM devicename"

  Response.Write ("<TR bgcolor=#E1F2FD><TD COLSPAN=2><FORM METHOD=POST ACTION=dbview.asp?MODE=1 NAME=F1><textarea rows=3 cols=50 NAME=DBNAME>CREATE DATABASE databasename</textarea></TD><TD ALIGN=RIGHT><FONT FACE=ARIAL SIZE=1><INPUT TYPE=SUBMIT NAME=S VALUE='Create Database'></FORM></FONT></TD></TR>")
  Response.Write ("<TR bgcolor=#E1F2FD><TD COLSPAN=2><FORM METHOD=POST ACTION=dbview.asp?MODE=1 NAME=F2><textarea rows=3 cols=50 NAME=DBNAME>" & ATTACHDB & "</textarea></TD><TD ALIGN=RIGHT><FONT FACE=ARIAL SIZE=1><INPUT TYPE=SUBMIT NAME=S VALUE='Attach Database'></FORM></FONT></TD></TR>")
  Response.Write ("<TR bgcolor=#E1F2FD><TD COLSPAN=2><FORM METHOD=POST ACTION=dbview.asp?MODE=1 NAME=F3><textarea rows=3 cols=50 NAME=DBNAME>" & LDEVICE & "</textarea></TD><TD ALIGN=RIGHT><FONT FACE=ARIAL SIZE=1><INPUT TYPE=SUBMIT NAME=S VALUE='Create Device'></FORM></FONT></TD></TR>")
  Response.Write ("<TR bgcolor=#E1F2FD><TD COLSPAN=2><FORM METHOD=POST ACTION=dbview.asp?MODE=1 NAME=F4><textarea rows=3 cols=50 NAME=DBNAME>" & BKUPDB & "</textarea></TD><TD ALIGN=RIGHT><FONT FACE=ARIAL SIZE=1><INPUT TYPE=SUBMIT NAME=S VALUE='Backup Database'></FORM></FONT></TD></TR>")
  Response.Write ("<TR bgcolor=#E1F2FD><TD COLSPAN=2><FORM METHOD=POST ACTION=dbview.asp?MODE=1 NAME=F5><textarea rows=3 cols=50 NAME=DBNAME>" & RESTORE & "</textarea></TD><TD ALIGN=RIGHT><FONT FACE=ARIAL SIZE=1><INPUT TYPE=SUBMIT NAME=S VALUE='Restore Database'></FORM></FONT></TD></TR>")
  Response.Write( "</TABLE>" )

End Sub
%>

<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF >
<TR>
<TD valign=middle bgcolor=#D4D4D4 bordercolordark=#FFFFFF bordercolor=#808080 bordercolorlight=#808080>&nbsp;&nbsp;<FONT FACE=ARIAL SIZE=2><img src=ofolder.gif border=0 alt='SoftRobot Document Server'><B> SoftRobot Database Server</B></FONT></TD>
</TR>
</table>

<TABLE WIDTH=100% >
<TR bgcolor=#E1F2FD>
<TD><FONT FACE=ARIAL SIZE=2 ><B>DATABASE DESIGN:</B></FONT></TD>
</TR>
<TR bgcolor=#d1d2d3><TD><FONT FACE=ARIAL SIZE=1 >Database List</FONT></TD></TR>
</TABLE>
<%
MODE=REQUEST("MODE")
IF MODE="" THEN
set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT CATALOG_NAME, SCHEMA_OWNER FROM INFORMATION_SCHEMA.SCHEMATA"
rs.open sql,conn
GenerateTable rs
rs.Close
Response.Write "<B>Note:</B>  Creating a logical backup device needs to be done only once."
ELSEIF MODE=1 THEN
DBNAME=REQUEST("DBNAME")
Conn.Execute DBNAME
Response.Write "Action on Database " & DBNAME & " is performed. Click on Database menu again."
ELSE
DBNAME=REQUEST("DBNAME")
DETACHDB= "EXECUTE sp_detach_db @dbname = " & DBNAME & ", @skipchecks = TRUE"
Conn.Execute DETACHDB
Response.Write "Action on Database " & DBNAME & " is performed. Click on Database menu again."
END IF
Conn.Close
Set Conn=nothing
%>
<hr>
<font face="Arial" size="1">
&#169; Copyright 2005 . All rights reserved. IRP
is registered trademark of ERPWEB.
<a href=home.html>Home</a> | <a href=mailto:sales@erpweb>Contact</a>
</font>
</BODY>
</HTML>