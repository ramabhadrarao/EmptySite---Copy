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
<FORM METHOD=POST ACTION="formeditnew.asp?MODE=1" name=form1> 
<INPUT TYPE=HIDDEN NAME="DID" VALUE=<%=DID%>>
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=UID%>>
<INPUT TYPE=HIDDEN NAME="ID" VALUE=<%=ID%>>
<%
IF SQLPROGRAM <> "" THEN
	set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open SQLPROGRAM & ID, Conn
	'RESPONSE.WRITE SQLPROGRAM & ID
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
<%
'-------------------------------------
	
IF DETAILSSQL <> "" THEN 
	set rs1 = Server.CreateObject("ADODB.Recordset")
	rs1.Open DETAILSSQL & ID, Conn
	J=0
	Response.Write( "<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF>" )
  '--------------------------------------
  ' set up column names
  for i = 2 to rs1.fields.count - 1
  '--------------------FIND MATCHING FIELD NAME FROM DICTIONARY
	set rsdic = Server.CreateObject("ADODB.Recordset")
	rsdic.Open "SELECT * FROM DICTIONARY WHERE DETFLAG=1 AND DDID=" & DDID & " AND SEQ=" & i, Conn
	IF NOT rsdic.eof then
	FLDNM=rsdic("NEWFLDNM")
	FNT=rsdic("FONT")
	FNTFAMILY="font-family: " & FNT
	else
	FLDNM=rs1(i).Name
	FNT="ARIAL"
	FNTFAMILY="font-family: ARIAL"
	end if
	rsdic.Close
	Set rsdic = Nothing	
 '--------------------------------------
        if rs1(i).Type = 3 then
        Set ls=Server.CreateObject("ADODB.Recordset")
		ls.Open "Select * from LISTBOX where DETFLAG=1 AND DDID=" & DDID & " AND COLUMNNO=" & i, Conn
		if not ls.EOF then
		LISTNAME=ls("LISTNAME")
		Response.Write("<TD bgcolor=#D5EAFF bordercolordark=#FFFFFF bordercolorlight=#808080 ><FONT face='" & FNT & "' SIZE=1 >" + LISTNAME + "</FONT></TD>")
        ELSE
        Response.Write("<TD bgcolor=#D5EAFF bordercolordark=#FFFFFF bordercolorlight=#808080 ><FONT face='" & FNT & "' SIZE=1 >" & FLDNM & "</FONT></TD>")
        end if
        ls.Close
        set ls=nothing
        ELSE
        Response.Write("<TD bgcolor=#D5EAFF bordercolordark=#FFFFFF bordercolorlight=#808080 ><FONT face='" & FNT & "' SIZE=1 >" & FLDNM & "</FONT></TD>")
        end if
    next
    Response.Write("<TD bgcolor=#D5EAFF bordercolordark=#FFFFFF bordercolorlight=#808080 ><FONT FACE=ARIAL SIZE=1 >Action</FONT></TD>")
  '------------------------------------------
	WHILE NOT rs1.EOF 
	Response.Write "<tr>"
	GenerateDetForm rs1, J, DDID, Conn
	rs1.MoveNext
	J=J+1
	Wend
	Response.Write "<input type=hidden name=J value=" & J & ">"
  '-----------------------------------------
%>
<%GTOT=FORMATNUMBER(GTOT,2)%>
<%IF DID=692 AND GTOT<>0 THEN%>
<FONT FACE=ARIAL SIZE=2 COLOR=RED>ERROR: Credit Amount must balance Debit amount. Continue making entry to balance it.</FONT>
<%END IF%>
</FORM> 
<%'---------------------------Detail Add
Response.Write("<form METHOD=POST action=FORMEDITNEW.ASP?MODE=3 name=form2>")
  	Response.Write "<tr>"
	GenerateNewForm rs1, DDID, Conn
	rs1.Close
	Set rs1=Nothing
	%>
  <INPUT TYPE=HIDDEN NAME="DID" VALUE=<%=DID%>>
  <INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=UID%>>
  <INPUT TYPE=HIDDEN NAME="ID" VALUE=<%=ID%>>
<TD bgcolor=#E1F2FD width=40 align=right>
<input type=image src=addnew.jpg name=S Value=AddNew>
</td>
</tr>
</table>
<table width=100% >
<tr>
<TD bgcolor=#f1f1f2 align=right>
<%IF GTOT>0 THEN%>
<FONT FACE=ARIAL SIZE=1> Grand Total: <%=GTOT%></FONT>
<%
'-----------------------update sales order value
IF DID=627 AND GTOT>0 THEN
UPDTSQL="UPDATE SALESORDER SET ORDERVALUE=" & ROUND(GTOT) & " WHERE SALESORDERID=" & ID
'RESPONSE.WRITE UPDTSQL
Conn.Execute UPDTSQL
END IF
'-----------------------update ordervalue in purchase invoice
IF DID=1982 THEN
UPDATESQL="UPDATE PURCHINVOICE SET ORDERVALUE=" & GTOT & " where PURCHINVOICEID=" & ID
'Response.Write UPDATESQL & "<BR>"
Conn.Execute UPDATESQL
END IF
'------------------------CHECK QTY IN PO
IF DID=3369 AND GTOT>0 THEN
	set rsx = Server.CreateObject("ADODB.Recordset")
	SQLPROGRAMX="SELECT PODETID from PODELSCH where PODELSCHID=" 
	rsx.Open SQLPROGRAMX & ID, Conn
	if not rsx.EOF then
	PODETID=rsx(0)
	else
	PODETID=0
	end if
	rsx.close
	'-----------------------------
	set rsx = Server.CreateObject("ADODB.Recordset")
	SQLPROGRAMX="SELECT QTY from PODET where PODETID=" 
	rsx.Open SQLPROGRAMX & PODETID, Conn
	if not rsx.EOF then
	QTY=rsx(0)
	end if
	rsx.close
	'--------------------------------
	IF ROUND(GTOT)>ROUND(QTY) THEN
	'UPDTSQL="UPDATE PODELSCHDET SET POQTY=0 WHERE PODELSCHID=" & ID
	'Conn.Execute UPDTSQL
 	RESPONSE.WRITE("<FONT COLOR=RED>YOU CANNOT ENTER QTY MORE THEN PO QTY</FONT>")
	END IF
	'-------------------------------
END IF
%>
<%END IF%>
</td>
<TD bgcolor=#f1f1f2 align=LEFT width=80>
<FONT FACE=ARIAL SIZE=1><%=date%></font>
</td>
</tr>
</table>
</form>

<%
END IF
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