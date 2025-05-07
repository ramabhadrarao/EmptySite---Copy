<!--
'************************************************************************
'Pupose						:	This is a SoftRobot document edit server
'Filename					:	formeditnew.asp
'Author						:	Anita Shah
'Created					:	27-Mar-2007
'Project Name				:	ERPWEB
'Contact					:	ashish@erpweb.com
'
'Modification History		:	
'Purpose					:
'Version					:
'Author 					:
'Created					:
'************************************************************************
-->
<%' @TRANSACTION=Required LANGUAGE="VBScript" %>
<%'@ LANGUAGE="VBScript" %>
<%
'Response.Buffer = True
Const adUseClient = 3
Server.ScriptTimeout=3600
%>
<HTML>
<HEAD>
<TITLE>ERPWEB / MobileERP / SoftServer Edit Document Server(sales@erpweb.com)</TITLE>
</HEAD>

<BODY topmargin=0>
<!---------#include file="calendar.js"------------->
<Basefont face=arial size=1>
<%
TOTFLAG=0
Set Conn=Server.CreateObject("ADODB.Connection")
Conn.Open Session("erp_ConnectionString")
%>
<!-- #INCLUDE FILE=addtabdoc.asp -->
<!-- #INCLUDE FILE=acheckdata.asp -->
<!-- #INCLUDE FILE=alistgen.asp -->
<!-- #INCLUDE FILE=agenerateform.asp -->
<!-- #INCLUDE FILE=ageneratedetform.asp -->
<!-- #INCLUDE FILE=ageneratenewform.asp -->
<%
Sub GenerateHeader( rs )
Response.Write( "<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF >" )
Response.Write( "<TR>" )
Response.Write( "<TD valign=middle bgcolor=#D4D4D4 bordercolordark=#FFFFFF bordercolor=#808080 bordercolorlight=#808080>&nbsp;&nbsp;<FONT FACE=ARIAL SIZE=2><img src=60-60-60.gif border=0 alt='SoftRobot Document Server'><B> Edit " & rs("TITLE") & " Document</B></FONT></TD>" )
Response.Write( "</TR>" )
'Response.Write( "<TR>" )
'Response.Write( "<TD BGCOLOR=YELLOW><FONT FACE=ARIAL SIZE=1><B>" & rs("HEADERNOTE")& "</B></FONT></TD>" )
'Response.Write( "</TR>" )
Response.Write( "</TABLE>" )
End Sub
%>
<%
Sub GenerateFooter( rs )
Response.Write( "<TABLE BORDER=0 WIDTH=100% >" )
'Response.Write( "<TR>" )
'Response.Write( "<TD BGCOLOR=yellow><FONT FACE=ARIAL SIZE=1><B>Rules:" & rs("FOOTERNOTE") & "</B></FONT></TD>" )
'Response.Write( "</TR>" )
Response.Write( "<TR>" )
Response.Write( "<TD BGCOLOR=#C0C0C0><FONT FACE=ARIAL SIZE=1><I>" & DATE & "</I></FONT></TD>" )
Response.Write( "</TR>" )
Response.Write( "</TABLE>" )
End Sub
%>

<%
'-----------------------------------------MAIN PROG STARTS
DID=REQUEST("DID")'IDENTIFY DOCUMENT
UID=REQUEST("UID")'IDENTIFY USER
ID=REQUEST("ID")'-----------READ RECORD ID
if DID="" Then Response.Write "ERROR:DID IS NULL":Response.End
if UID="" Then Response.Write "ERROR:UID IS NULL":Response.End
GTOT=0
'-------------------------------------Continue opening the document parameters
set rsDOC = Server.CreateObject("ADODB.Recordset")
rsDOC.Open "Select * from DOCUMENTS WHERE DID=" & DID, Conn 
IF NOT rsDOC.EOF THEN
SQLPROGRAM=rsDOC("MASTERSQL")
DETAILSSQL=rsDOC("SQLDETAILS")
ADDSQL=rsDOC("ADDSQL")
ADDDETAILS=rsDOC("ADDDETAILS")
MSTTBL=rsDOC("MASTERTABLE")
DETTBL=rsDOC("DETAILSTABLE")
DDID=rsDOC("DDID")
if isnull(SQLPROGRAM) THEN SQLPROGRAM=""
if isnull(DETAILSSQL) THEN DETAILSSQL=""
END IF
'-------------------------------------call header subroutine
GenerateHeader rsDoc
'-------------------------------------
MODE=REQUEST("MODE")
'-------------------------ADD MASTER DATA
%>
<!-- #INCLUDE FILE=formeditmodes.asp -->
<FORM METHOD=POST ACTION="formeditmac.asp?MODE=1" name=form1 id=form1> 
<INPUT TYPE=HIDDEN NAME="DID" VALUE=<%=DID%>>
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=UID%>>
<INPUT TYPE=HIDDEN NAME="ID" VALUE=<%=ID%>>
<%
IF SQLPROGRAM <> "" THEN
	set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open SQLPROGRAM & ID, Conn
	if not rs.EOF then
	PKEY=rs(3)
	GenerateForm rs, DDID, CID, UID, Conn ' Display master form for edit function with dictionary mode
	else
	Response.Write "<img src=error.gif> Error: Problems in executing query" & SQLPROGRAM & ID
	END IF
	rs.Close
	Set rs = Nothing	
ELSE
	Response.Write "<img src=error.gif> ERROR: PLEASE ENTER MASTERSQL - Press back button"
END IF
%>
<table width=100% >
<tr>
<TD bgcolor=#f1f1f2 align=right>
<input type=image src=update2.jpg name=Update VALUE="Update Document">
</td>
</tr>
</table>


</form>

<%

'-------------------------------------	
	'GenerateFooter rsDoc
'-------------------------------------
rsDOC.Close
Conn.Close
Set rsDOC = Nothing
SET Conn = nothing
%>
</BODY>
</HTML>


