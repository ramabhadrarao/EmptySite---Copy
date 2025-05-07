
<HTML>
<HEAD>
<TITLE>
IRP REPORT GENERATOR (sales@erpweb)
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
	rs1.open "SELECT * FROM DOCMODULE order by docname", Conn
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
<%
Sub SMLIST1 ( )

	Set rs1=server.createobject("ADODB.Recordset")
	rs1.open "SELECT * FROM USERS order by UNAME", Conn
	%>
	<SELECT name=UID>
	<%WHILE not rs1.eof%>
	<OPTION VALUE=<%=rs1(3)%>><%=rs1(4)%> 
	<%rs1.MoveNext%>
	<%wend%>
	</select>
	<%rs1.Close%>

<%
End Sub
%>
<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF >
<TR>
<TD valign=middle bgcolor=#D4D4D4 bordercolordark=#FFFFFF bordercolor=#808080 bordercolorlight=#808080>&nbsp;&nbsp;<FONT FACE=ARIAL SIZE=2><img src=60-60-60.gif border=0 alt='SoftRobot Document Server'><B> SoftRobot Report/Function Designer</B></FONT></TD>
</TR>
</TABLE>
<%MODE=REQUEST("MODE")%>
<%IF MODE="" THEN%>
<FORM NAME=FORMX ACTION=createreport.asp METHOD=POST>
<font face=arial size=1><B>Basic Data for creating individual linked documents: </B></font><br>

<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF>
<TR>
<TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080>
<font face="ARIAL" size=1>Report/Doc Name</font>
</TD>
<TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080>
<font face="ARIAL" size=1>Document Type</font>
</TD>
<TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080>
<font face="ARIAL" size=1>Document Menu</font>
</TD>
<TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080>
<font face="ARIAL" size=1>User Rights</font>
</TD>
</TR>
<TR>
<TD>
<input type="text" name="REPORT" size="9">
</TD>
<TD>
<SELECT NAME=TYPE>
<OPTION VALUE=1>1.Manual Report/Doc
<OPTION VALUE=2>2.Tabular AutoReport
<OPTION VALUE=3>2.Columnar AutoReport
<OPTION VALUE=4>3.Data Transfer Service
<OPTION VALUE=5>4.Linked Report/Doc
</SELECT>
</TD>
<TD>
<%SMLIST%>
</TD>
<TD>
<%SMLIST1%>
</TD>
<input type=HIDDEN name=MODE VALUE=1>
</TR>
</TABLE>
<hr>
<font face=arial size=1><B>Type1: URL Filename for Manual Reports/Transaction Documents/DTS:</B> </font><br>
<font face=arial size=1>URL Filename :</font><input type="text" name="DADDRESS" size="25" value=filename.asp>
<hr>
<font face=arial size=1><B>Type2: SQL for Columnar or Tabular AutoReports or DTS View: </B></font><br>
<textarea rows=3 cols=50 name=SQL>select fieldnames from viewORtablenames</textarea>
<hr>
<font face=arial size=1><B>Type3: Data Source and Destination for DTS Services: </B></font><br>
<font face=arial size=1>Source Table/View: <input type=text name=SRCTABLE></font>
<font face=arial size=1>Destination Table/View: <input type=text name=DESTTABLE></font>
<hr>
<INPUT TYPE=SUBMIT NAME=CREATE VALUE="Create Report / Documents Function">
</FORM>
<hr><font face=arial size=2>
Sample: <a href=shared/leaveapp/leaveapp1.htm>Manual Report/Doc</a> |
<a href=treport.htm>Tabular Report</a> |
<a href=creport.htm>Columnar Report</a> |
<a href=treport.htm>Data Transfer Service</a> |
<%ELSE%>
<%REPORT=REQUEST("REPORT"): IF REPORT="" THEN Response.Write "Error:REPORTNAME" : Response.END%>
<%DDID=REQUEST("DDID")%>
<%UID=REQUEST("UID")%>
<%DADDRESS=REQUEST("DADDRESS"): IF ISNULL(DADDRESS) THEN DADDRESS=""%>
<%LINKADDRESS=REQUEST("DADDRESS"):IF ISNULL(LINKADRESS) THEN LINKADDRESS=""%>
<%SQL=REQUEST("SQL"): IF REPORT="" THEN Response.Write "Error: SQL" : Response.end%>
<%
RTYPE=REQUEST("TYPE")
IF RTYPE=2 THEN 
DADDRESS="treport.asp" 
DOCTYPEID=3
ELSEIF RTYPE=3 THEN 
DADDRESS="creport.asp"
DOCTYPEID=3
ELSEIF RTYPE=4 THEN 
DADDRESS="dts.asp"
DOCTYPEID=2
ELSEIF RTYPE=5 THEN 
DADDRESS="lreport.asp"
DOCTYPEID=3
ELSE
DOCTYPEID=3
END IF
%>
<hr>
<font face=arial size=1>
<%
SQLINSRT="INSERT INTO DOCUMENTS1 (DDID, DOCTYPEID, DNAME, TITLE, DADDRESS, REPORTNAME, HEADERNOTE, FOOTERNOTE, SEQUENCE, SQLPROGRAM, LINKADDRESS) VALUES (" & DDID & ", " & DOCTYPEID & ", '" & REPORT & "', '" & REPORT & "', '" & DADDRESS & "', '" & REPORT & "', '" & REPORT & "', '" & REPORT & "', " & 7 & ", '" & SQL & "', '" & LINKADDRESS & "')"
Response.Write SQLINSRT
Conn.Execute SQLINSRT
%>
<%
Set rs=Server.CreateObject("ADODB.Recordset")
rs.Open "Select DID from DOCUMENTS1 ORDER BY DID", Conn,1,3
if not rs.EOF then
rs.MoveLast
DID=rs("DID")
end if
rs.Close
%>
<hr>
<%
SQLINSRT="INSERT INTO WORKFLOW (UID, DID) VALUES (" & UID & ", " & DID & ")"
Response.Write SQLINSRT
Conn.Execute SQLINSRT
%>
<HR>
<%
IF RTYPE=4 THEN
SRCTABLE=REQUEST("SRCTABLE"): IF SRCTABLE="" THEN SRCTABLE="ENTERSRCTABLE"
DESTTABLE=REQUEST("DESTTABLE"): IF DESTTABLE="" THEN DESTTABLE="ENTERDESTTABLE"
SQLINSERT="INSERT INTO DTS (DID, SRCTABLE, DESTTABLE) VALUES (" & DID & ", '" & SRCTABLE & "', '" & DESTTABLE & "')"
Response.Write SQLINSERT
Conn.Execute SQLINSERT
END IF
%>
</font>
<%END IF%>
<hr>
<font face=arial size=1>
Rules for creating Linked Reports / Documents:<br>
<br>1. Creating Manual Report or Document: Provide manual file name (*.asp, *.exe, *.jsp, *.doc, *.xls etc.) in URL Texbox.
<br>2. Creating Tabular Report: Provide SQL Statement in SQL Statement textarea through which this report will be build. 
<br>3. Creating Columnar Report: Provide SQL Statement in SQL Statement textarea through which this report will be build.
<br>4. Creating Data Transformation Services: Provide SQL, Source and Destination Tables/Views in respective textboxes. You also
needs to goto DTS document and enter Source fileds and Destination fields. <a href=document.asp?DID=1504&UID=57&FADD=True&FDEL=True&FVIEW=True&FEDIT=True&FILTER=False&FOFFHOLD=False&FREJECT=False&CID=1>Click here to goto DTSRegister</a>.
<BR>5. Creating Linked Reports: Provide File URL, SQL as Register will link to it.
<br>Note: XML based web services documents are automatically created by creating tabular or columnar reports.
</font>
<hr>
<font face="Arial" size="1">
&#169; Copyright 2001 . All rights reserved. IRP & SoftRobot.net
are trademarks of ERPWEB.
<a href=mailto:sales@erpweb>Contact</a>
</font>
</BODY>
</HTML>