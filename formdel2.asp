<%' @TRANSACTION=Required LANGUAGE="VBScript" %>
<%
Response.Buffer = True
Const adUseClient = 3
%>
<HTML>
<HEAD>
<TITLE>IRP FORM DEL (sales@erpweb)</TITLE>
</HEAD>

<BODY>
<Basefont face=arial size=1>
<%
Set Conn=Server.CreateObject("ADODB.Connection")
Conn.Open Session("erp_ConnectionString")
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
Sub GenerateHeader( rs )
Response.Write( "<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF >" )
Response.Write( "<TR>" )
Response.Write( "<TD valign=middle bgcolor=#D4D4D4 bordercolordark=#FFFFFF bordercolor=#808080 bordercolorlight=#808080>&nbsp;&nbsp;<FONT FACE=ARIAL SIZE=2><img src=60-60-60.gif border=0 alt='SoftRobot Document Server'><B> Delete " & rs("TITLE") & " Document</B></FONT></TD>" )
Response.Write( "</TR>" )
'Response.Write( "<TR>" )
'Response.Write( "<TD BGCOLOR=YELLOW><FONT FACE=ARIAL SIZE=1><B>" & rs("HEADERNOTE") & "</B></FONT></TD>" )
'Response.Write( "</TR>" )
Response.Write( "</TABLE>" )
End Sub
%>
<%
Sub GenerateFooter( rs )
Response.Write( "<TABLE BORDER=0 WIDTH=100% >" )
'Response.Write( "<TR BGCOLOR=YELLOW>" )
'Response.Write( "<TD WIDTH=80% ><FONT FACE=ARIAL SIZE=1><B>Rules:" & rs("FOOTERNOTE") & "</B></FONT></TD><TD ALIGN=RIGHT><FONT FACE=ARIAL SIZE=1><I>" & NOW & "</I></FONT></TD>" )
'Response.Write( "</TR>" )
Response.Write( "<TR BGCOLOR=#d1d2d3>" )
Response.Write("<TD ALIGN=RIGHT>---</TD>")
Response.Write( "</TR>" )
Response.Write( "</TABLE>" )
End Sub
%>

<% 
'---------------------------------FORM GENERATION STARTS
Sub GenerateForm( rs )
' start form
%>
<table WIDTH=100% >
<%
  ' build input field for each recordset field
  for i = 0 to rs.fields.count - 1
  
      value = rs(i)
      if isNull(value) then value=""
    
    '-----------------------------------------
    if rs(i).Type = 3 then
		Set ls=Server.CreateObject("ADODB.Recordset")
		ls.Open "Select * from LISTBOX where DDID=" & DDID & " AND COLUMNNO=" & i, Conn
		if not ls.EOF then
		LISTNAME=ls("LISTNAME")
		LISTTABLE=ls("LISTTABLE")
		LISTCOLUMN=ls("LISTCOLUMN")
		LISTVALUE=ls("LISTVALUE")
		%><tr bgcolor=#E1F2FD><td><font face=arial size=1><img src=required.gif> <%= LISTNAME %><%'= rs(i).Type %></font></td>
		<td>
		<%LIST1 LISTTABLE, LISTCOLUMN, LISTVALUE, Conn, name, value%> 
		</td></tr> <%
		else
		%><tr bgcolor=#E1F2FD><td><font face=arial size=1><img src=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td><font face=arial size=1><%=value%></font></td></tr> <%
		end if
		ls.Close
		set ls=nothing
    '-----------------------------------------
    else
		%><tr bgcolor=#E1F2FD><td><font face=arial size=1><img src=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td><font face=arial size=1><%=value%></font></td></tr> <%
    end if
 next
%> 
</table>
<%END SUB%>
<%
DID=REQUEST("DID")
if DID="" Then Response.Write "ERROR:DID IS NULL":Response.End
ID=REQUEST("ID")
if ID="" Then Response.Write "ERROR:ID IS NULL":Response.End
UID=REQUEST("UID")
if UID="" Then Response.Write "ERROR:UID IS NULL":Response.End
%>
<%
set rsDOC = Server.CreateObject("ADODB.Recordset")
rsDOC.Open "Select * from DOCUMENTS WHERE DID=" & DID, Conn 
IF NOT rsDOC.EOF THEN
SQLPROGRAM=rsDOC("DETAILSSQL")
IF ISNULL(SQLPROGRAM) THEN SQLPROGRAM=""
DDID=rsDOC("DDID")
END IF
'-------------------------------------
GenerateHeader rsDoc
'-------------------------------------
IF SQLPROGRAM <> "" THEN
	set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open SQLPROGRAM & ID, Conn
	GenerateForm rs
	rs.Close
	Set rs = Nothing	
ELSE
	Response.Write "PLEASE ENTER EDITDETAILS"
END IF
'-------------------------------------	
	GenerateFooter rsDoc
'-------------------------------------
rsDOC.Close
Conn.Close
Set rsDOC = Nothing
SET Conn = nothing
%>
<form method=post action=FORMDEL21.ASP?ID=<%=ID%>&DID=<%=DID%>>
<!---------#include file="password.inc"------------->
<INPUT TYPE=HIDDEN NAME=UID VALUE=<%=REQUEST("UID")%>>
<INPUT TYPE=SUBMIT NAME=DELETE VALUE=DELETE>
</FORM>
<%
ID=REQUEST("IDD")
%>
<font face=arial size=2>
<a href=FORMEDIT.ASP?DID=<%=DID%>&ID=<%=ID%>&UID=<%=UID%>><b>GOTO Edit Page</b></a>
</font>
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