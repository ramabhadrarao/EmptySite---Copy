<!--
'************************************************************************
'Pupose						:	This is a SoftServer document view server
'Filename					:	formview.asp
'Author						:	Anita Shah
'Created					:	27-Aug-2006
'Project Name					:	ERPWEB
'Contact					:	sales@ERPWEB.com
'
'Modification History				:	
'Purpose					:
'Version					:
'Author 					:
'Created					:
'************************************************************************
-->
<%' @TRANSACTION=Required LANGUAGE="VBScript" %>
<%
'Response.Buffer = True
'Const adUseClient = 3
%>
<HTML>
<HEAD>
<TITLE>SoftServer Print VIEW Server(sales@ERPWEB.com)</TITLE>
</HEAD>


<%
Set Conn=Server.CreateObject("ADODB.Connection")
Conn.Open Session("erp_ConnectionString")
Server.ScriptTimeOut=3600
%>
<% 
Sub GenerateTable1( rs )
  Response.Write( "<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF>" )
  ' set up column names
  for i = 0 to rs.fields.count - 3
    Response.Write("<TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080><font face=arial size=1>" + rs(i).Name + "</font></TD>")
  next
  ' write each row
  
  while not rs.EOF
    Response.Write( "<TR bgcolor=#E1F2FD>" )
     for i = 0 to rs.fields.count - 3
      v = rs(i)
      if isnull(v) then v = ""
      Response.Write( "<TD VALIGN=TOP><font face=arial size=1>" + CStr( v ) + "</font></TD>" )
    next
    rs.MoveNext
   wend 
     Response.Write( "</TABLE>" )
End Sub
%>
<%
'----------------------------------LISTBOX GENERATOR
Sub LIST1 ( LISTTABLE, LISTCOLUMN, LISTVALUE, Conn, name, value )
Set lst=Server.CreateObject("ADODB.Recordset")
lst.Open "Select " & LISTVALUE & ", " & LISTCOLUMN & " FROM " & LISTTABLE & " WHERE " & LISTVALUE & " = " & value, Conn
if not lst.eof then
%><font face=arial size=1>
<%=lst(1)%></font>
<%
end if
lst.Close
End Sub
%>
<% 
'---------------------------------FORM GENERATION STARTS
Sub GenerateForm( rs, DDID, Conn )
' start form
%>
<table WIDTH=100% >
<%
   TOTRECORDS=rs.fields.count - 1
   FIRSTTAB=CINT((TOTRECORDS-3)/2)
'--------------------------
%>
<tr><td valign=top>
<table width=100% >
<%
  ' build input field for each recordset field
  for i = 3 to FIRSTTAB+3
'--------------------
      value = rs(i)
      if isNull(value) then value=""
'--------------------FIND MATCHING FIELD NAME FROM DICTIONARY
	set rsdic = Server.CreateObject("ADODB.Recordset")
	rsdic.Open "SELECT * FROM DICTIONARY WHERE DETFLAG=0 AND DDID=" & DDID & " AND SEQ=" & i, Conn
	IF NOT rsdic.eof then
	FLDNM=rsdic("NEWFLDNM")
	FNT=rsdic("FONT")
	FNTFAMILY="font-family: " & FNT
	else
	FLDNM=rs(i).Name
	FNT="ARIAL"
	FNTFAMILY="font-family: ARIAL"
	end if
	rsdic.Close
	Set rsdic = Nothing	
'----------------------
    if i=3 then
    %><tr bgcolor=#E1F2FD><td><font size=1 face="<%=FNT%>"><IMG SRC=required.gif> <%=FLDNM %><%'= rs(i).Type %></font></td><td><font face="<%=FNT%>" size=1><%=value%></font></td></tr> <%
    elseif (rs(i).Type = 3) and (value > "0") then
		Set ls=Server.CreateObject("ADODB.Recordset")
		ls.Open "Select * from LISTBOX where DETFLAG=0 AND DDID=" & DDID & " AND COLUMNNO=" & i, Conn
		if not ls.EOF then
		LISTNAME=ls("LISTNAME")
		LISTTABLE=ls("LISTTABLE")
		LISTCOLUMN=ls("LISTCOLUMN")
		LISTVALUE=ls("LISTVALUE")
		%><tr bgcolor=#E1F2FD><td><font size=1 face="<%=FNT%>"><IMG SRC=required.gif> <%=LISTNAME %><%'= rs(i).Type %></font></td>
		<td>
		<%LIST1 LISTTABLE, LISTCOLUMN, LISTVALUE, Conn, name, value%> 
		</td></tr> <%
		else
		%><tr bgcolor=#E1F2FD><td><font size=1 face="<%=FNT%>"><IMG SRC=required.gif> <%=FLDNM %><%'= rs(i).Type %></font></td><td><font face="<%=FNT%>" size=1><%=value%></font></td></tr> <%
		end if
		ls.Close
		set ls=nothing
    '-----------------------------------------
    else
		%><tr bgcolor=#E1F2FD><td><font size=1 face="<%=FNT%>"><IMG SRC=required.gif> <%=FLDNM %><%'= rs(i).Type %></font></td><td><font face="<%=FNT%>" size=1><%=value%></font></td></tr> <%
    end if
 next
%>
</table>
</td><td valign=top>
<table width=100%>
<%
  ' build input field for each recordset field
  for i = FIRSTTAB+4 TO TOTRECORDS
'--------------------
      value = rs(i)
      if isNull(value) then value=""
'--------------------FIND MATCHING FIELD NAME FROM DICTIONARY
	set rsdic = Server.CreateObject("ADODB.Recordset")
	rsdic.Open "SELECT * FROM DICTIONARY WHERE DETFLAG=0 AND DDID=" & DDID & " AND SEQ=" & i, Conn
	IF NOT rsdic.eof then
	FLDNM=rsdic("NEWFLDNM")
	FNT=rsdic("FONT")
	FNTFAMILY="font-family: " & FNT
	else
	FLDNM=rs(i).Name
	FNT="ARIAL"
	FNTFAMILY="font-family: ARIAL"
	end if
	rsdic.Close
	Set rsdic = Nothing	
'----------------------
    if i=3 then
    %><tr bgcolor=#E1F2FD><td><font size=1 face="<%=FNT%>"><IMG SRC=required.gif> <%=FLDNM %><%'= rs(i).Type %></font></td><td><font face="<%=FNT%>" size=1><%=value%></font></td></tr> <%
    elseif (rs(i).Type = 3) and (value > "0") then
		Set ls=Server.CreateObject("ADODB.Recordset")
		ls.Open "Select * from LISTBOX where DETFLAG=0 AND DDID=" & DDID & " AND COLUMNNO=" & i, Conn
		if not ls.EOF then
		LISTNAME=ls("LISTNAME")
		LISTTABLE=ls("LISTTABLE")
		LISTCOLUMN=ls("LISTCOLUMN")
		LISTVALUE=ls("LISTVALUE")
		%><tr bgcolor=#E1F2FD><td><font size=1 face="<%=FNT%>"><IMG SRC=required.gif> <%=LISTNAME %><%'= rs(i).Type %></font></td>
		<td>
		<%LIST1 LISTTABLE, LISTCOLUMN, LISTVALUE, Conn, name, value%> 
		</td></tr> <%
		else
		%><tr bgcolor=#E1F2FD><td><font size=1 face="<%=FNT%>"><IMG SRC=required.gif> <%=FLDNM %><%'= rs(i).Type %></font></td><td><font face="<%=FNT%>" size=1><%=value%></font></td></tr> <%
		end if
		ls.Close
		set ls=nothing
    '-----------------------------------------
    else
		%><tr bgcolor=#E1F2FD><td><font size=1 face="<%=FNT%>"><IMG SRC=required.gif> <%=FLDNM %><%'= rs(i).Type %></font></td><td><font face="<%=FNT%>" size=1><%=value%></font></td></tr> <%
    end if
 next
%>  
</table>
</td></tr>
</table>
<%END SUB%>
<%
Sub GenerateHeader( rs )
Response.Write( "<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF >" )
Response.Write( "<TR>" )
Response.Write( "<TD valign=middle bgcolor=#D4D4D4 bordercolordark=#FFFFFF bordercolor=#808080 bordercolorlight=#808080>&nbsp;&nbsp;<FONT FACE=ARIAL SIZE=2><img src=60-60-60.gif border=0 alt='SoftRobot Document Server'><B> " & rs("TITLE") & " Document</B></FONT></TD><td></td>" )
Response.Write( "</TR>" )
Response.Write( "<TR>" )
%><td align=right><a href=formeditnew12.asp?DID=<%=DID%>&ID=<%=ID%>&UID=<%=UID%>><IMG src=editnew.jpg border=0></a></td><%
Response.Write( "</TR>" )
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
Response.Write( "<TD BGCOLOR=#c0c0c0><FONT FACE=ARIAL SIZE=1><I>" & DATE & "</I></FONT></TD>" )
Response.Write( "</TR>" )
Response.Write( "</TABLE>" )
End Sub
%>
<% 
Sub GenerateTable( rs, DDID, Conn )
	
  Response.Write( "<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF>" )
   '--------------------------------------
  ' set up column names
  Response.Write( "<TD bordercolordark=#FFFFFF bordercolorlight=#808080 bgcolor=#D5EAFF><FONT FACE=ARIAL SIZE=1 >Sr#</FONT></TD>" )
  for i = 2 to rs.fields.count - 1
  '--------------------FIND MATCHING FIELD NAME FROM DICTIONARY
	set rsdic = Server.CreateObject("ADODB.Recordset")
	rsdic.Open "SELECT * FROM DICTIONARY WHERE DETFLAG=1 AND DDID=" & DDID & " AND SEQ=" & i, Conn
	IF NOT rsdic.eof then
	FLDNM=rsdic("NEWFLDNM")
	FNT=rsdic("FONT")
	FNTFAMILY="font-family: " & FNT
	else
	FLDNM=rs(i).Name
	FNT="ARIAL"
	FNTFAMILY="font-family: ARIAL"
	end if
	rsdic.Close
	Set rsdic = Nothing	
  '--------------------------------------
        if rs(i).Type = 3 then
        Set ls=Server.CreateObject("ADODB.Recordset")
		ls.Open "Select * from LISTBOX where DETFLAG=1 AND DDID=" & DDID & " AND COLUMNNO=" & i, Conn
		if not ls.EOF then
		LISTNAME=ls("LISTNAME") 
        Response.Write("<TD bgcolor=#D5EAFF bordercolordark=#FFFFFF bordercolorlight=#808080><FONT face='" & FNT & "' SIZE=1 >" + LISTNAME + "</FONT></TD>")
        ELSE
        Response.Write("<TD bgcolor=#D5EAFF bordercolordark=#FFFFFF bordercolorlight=#808080 ><FONT face='" & FNT & "' SIZE=1 >" & FLDNM & "</FONT></TD>")
        end if
        ls.Close
        set ls=nothing
        ELSE
        Response.Write("<TD bgcolor=#D5EAFF bordercolordark=#FFFFFF bordercolorlight=#808080 align=LEFT><FONT face='" & FNT & "' SIZE=1 >" & FLDNM & "</FONT></TD>")
        end if
    next
  srno=1
  TOTAL=0
  '------------------------------------------
  ' write each row
  while NOT rs.EOF 
    Response.Write( "<TR>" )
    Response.Write( "<TD VALIGN=TOP bgcolor=#D5EAFF><FONT FACE=ARIAL SIZE=1 >" & CStr( srno ) & "</FONT></TD>" )
	srno=srno+1
    '--------------------------------
    FLAG=NOT FLAG
    for i = 2 to rs.fields.count - 1
	'--------------------FIND MATCHING FIELD NAME FROM DICTIONARY
	set rsdic = Server.CreateObject("ADODB.Recordset")
	rsdic.Open "SELECT * FROM DICTIONARY WHERE DETFLAG=1 AND DDID=" & DDID & " AND SEQ=" & i, Conn
	IF NOT rsdic.eof then
	FLDNM=rsdic("NEWFLDNM")
	FNT=rsdic("FONT")
	FNTFAMILY="font-family: " & FNT
	else
	FLDNM=rs(i).Name
	FNT="ARIAL"
	FNTFAMILY="font-family: ARIAL"
	end if
	rsdic.Close
	Set rsdic = Nothing	
  '--------------------------------------
      v = rs(i)
      if isnull(v) then v = ""
      '----------------------------
      if (rs(i).Type = 3) and (v > "0") then
		Set ls=Server.CreateObject("ADODB.Recordset")
		ls.Open "Select * from LISTBOX where DETFLAG=1 AND DDID=" & DDID & " AND COLUMNNO=" & i, Conn
		if not ls.EOF then
		LISTNAME=ls("LISTNAME")
		LISTTABLE=ls("LISTTABLE")
		LISTCOLUMN=ls("LISTCOLUMN")
		LISTVALUE=ls("LISTVALUE")
		FFLAG=ls("LINKFLAG"): IF ISNULL(FFLAG) OR FFLAG="" THEN FFLAG=0
		IF FFLAG THEN LISTTABLE=LISTTABLE + "USED"
		if FLAG THEN
		%>
		<TD VALIGN=TOP bgcolor=#E1F2FD ><FONT face="<%=FNT%>" SIZE=1>
		<%LIST1 LISTTABLE, LISTCOLUMN, LISTVALUE, Conn, name, v%> 
		</font></td>
		<%ELSE%>
		<TD VALIGN=TOP bgcolor=#E1F2FF ><FONT face="<%=FNT%>" SIZE=1>
		<%LIST1 LISTTABLE, LISTCOLUMN, LISTVALUE, Conn, name, v%> 
		</font></td>
		<%END IF%>
		<%
	else
		IF FLAG THEN
			Response.Write( "<TD VALIGN=TOP bgcolor=#E1F2FD ><FONT face='" & FNT & "' SIZE=1>" & CStr( v ) & "</FONT></TD>" )
		   ELSE
		    Response.Write( "<TD VALIGN=TOP bgcolor=#E1F2FF ><FONT face='" & FNT & "' SIZE=1>" & CStr( v ) & "</FONT></TD>" )
		END IF
		end if
		ls.Close
		set ls=nothing
    '-----------------------------------------
   		else
   		 IF FLAG THEN
			Response.Write( "<TD VALIGN=TOP bgcolor=#E1F2FD align=LEFT><FONT face='" & FNT & "' SIZE=1>" & CStr( v ) & "</FONT></TD>" )
		   ELSE
		    Response.Write( "<TD VALIGN=TOP bgcolor=#E1F2FF align=LEFT><FONT face='" & FNT & "' SIZE=1>" & CStr( v ) & "</FONT></TD>" )
		END IF  
    end if
    '-------------------------------
    IF rs(i).Name="TOTAL" THEN
    TOTAL=TOTAL+rs("TOTAL")
    END IF
    '-------------------------------
    next
    rs.MoveNext
  wend 
  '-------------------------------------
  Response.Write( "</TABLE>" )
  IF TOTAL>0 THEN
  Response.Write( "<TABLE width=100% >" )
  Response.Write( "<TR><TD ALIGN=RIGHT BGCOLOR=#eeefff><FONT FACE=ARIAL SIZE=1>Grand Total: " & CStr( TOTAL ) & "</FONT></TD>" )
  Response.Write( "</TABLE>" )
  END IF
'-----------------------
End Sub
%>
<BODY topmargin=0>
<font face=arial size=1>

<%
DID=REQUEST("DID")
if DID="" Then Response.Write "ERROR:DID IS NULL":Response.End
ID=REQUEST("ID")
if ID="" Then Response.Write "ERROR:ID IS NULL":Response.End
UID=REQUEST("UID")
if UID="" Then Response.Write "ERROR:UID IS NULL":Response.End
CID=REQUEST("CID")
%>
<!-----------------------------HEADER STRIP--------------------> 
<!--------------------------------------------------------------->
<%
set rsDOC = Server.CreateObject("ADODB.Recordset")
rsDOC.Open "Select * from DOCUMENTS WHERE DID=" & DID, Conn 
IF NOT rsDOC.EOF THEN
SQLPROGRAM=rsDOC("MASTERSQL")
DETAILSSQL=rsDOC("SQLDETAILS")
DDID=rsDOC("DDID")
IF ISNULL(SQLPROGRAM) THEN SQLPROGRAM=""
IF ISNULL(DETAILSSQL) THEN DETAILSSQL=""
END IF
'-------------------------------------
GenerateHeader rsDoc
'-------------------------------------
IF SQLPROGRAM <> "" THEN
	set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open SQLPROGRAM & ID, Conn
	GenerateForm rs, DDID, Conn
	rs.Close
	Set rs = Nothing	
ELSE
	Response.Write "PLEASE ENTER MASTERSQL"
END IF
'-------------------------------------
IF DETAILSSQL <> "" THEN 
	set rs1 = Server.CreateObject("ADODB.Recordset")
	rs1.Open DETAILSSQL & ID, Conn
	GenerateTable rs1, DDID, Conn
	rs1.Close
	Set rs1=Nothing
END IF	
'-------------------------------------	
	GenerateFooter rsDoc
'-------------------------------------
rsDOC.Close
Set rsDOC = Nothing
%>

<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF >
<TR>
<TD valign=middle bgcolor=#D4D4D4 bordercolordark=#FFFFFF bordercolor=#808080 bordercolorlight=#808080>&nbsp;&nbsp;<FONT FACE=ARIAL SIZE=2><img src=60-60-60.gif border=0 alt='SoftRobot Document Server'><B> Document Track</B></FONT></TD>
</TR>
</TABLE>
<%
set rsD = Server.CreateObject("ADODB.Recordset")
rsD.Open "Select * from TrackingLIST WHERE DDID=" & DDID & " AND ID=" & ID, Conn 
	GenerateTable1 rsD
rsD.Close
%>
<%
Conn.Close
SET Conn = nothing
%>
</BODY>
</HTML>
<!-------------------------------------start transaction server---->
<%
'Sub OnTransactionCommit()
    'Response.Write "The Transaction just committed" 
    'Response.Write "This message came from the "
    'Response.Write "OnTransactionCommit() event handler."
'End Sub

'Sub OnTransactionAbort()
    'Response.Write "The Transaction just aborted" 
    'Response.Write "This message came from the "
    'Response.Write "OnTransactionAbort() event handler."
'End Sub
%>