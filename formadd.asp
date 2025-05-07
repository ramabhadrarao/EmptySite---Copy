<%' @TRANSACTION=Required LANGUAGE="VBScript" %>
<%
Response.Buffer = True
Const adUseClient = 3
DID=REQUEST("DID")'IDENTIFY DOCUMENT
UID=REQUEST("UID")'IDENTIFY USER
if DID="" Then Response.Write "ERROR:DID" : Response.END
if UID="" Then Response.Write "ERROR:UID" : Response.END
%>
<HTML>
<HEAD>
<TITLE>ERPWEB FORM ADD (sales@ERPWEB.com)</TITLE>
</HEAD>
<BODY topmargin=0>
<!---------#include file="calendar.js"------------->
<Basefont face=arial size=1>
<table width=100% border=0 cellpadding=0 cellspacing=0 ID="Table1">
<tr>
<td align=left background=images/bg2.gif>
<table width=267 border=0 cellpadding=0 cellspacing=0 ID="Table2">
<tr>
<td background=images/bg1.gif align=center width=249>

<!-- Search_Form -->
<table border=0 cellpadding=0 cellspacing=0 ID="Table3">
<form ID="Form1">
<tr valign=middle>
<td align=left><FONT face=ARIAL size=3 color=black><B>IRP Document Server</B></FONT><br></td>
<td></td>
<td><br></td>
</tr>
</form>
</table>
<!-- /Search_Form -->

</td>
<td><img src=images/tr1.gif width=18 height=40><br></td>
</tr>
</table>
</td>
</tr>
</table>
<%
Set Conn=Server.CreateObject("ADODB.Connection")
Conn.Open Session("erp_ConnectionString")
%>
<%
'----------------------------------LISTBOX GENERATOR
Sub LIST1 ( LISTTABLE, LISTCOLUMN, LISTVALUE, Conn, name )
Set lst=Server.CreateObject("ADODB.Recordset")
lst.Open "Select " & LISTVALUE & ", " & LISTCOLUMN & " FROM " & LISTTABLE & " ORDER BY " & LISTCOLUMN & " ASC", Conn
%>
	<SELECT name=<%=name%>>
	<OPTION VALUE=0>NONE
	<%WHILE not lst.eof%>
	<OPTION VALUE=<%=lst(0)%>><%=lst(1)%>
	<%lst.MoveNext%>
	<%wend%>
	</SELECT>
	<%lst.Close%>
<%
End Sub
%>
<%
Sub GenerateHeader( rs )
Response.Write( "<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF >" )
Response.Write( "<TR>" )
Response.Write( "<TD valign=middle bgcolor=#D4D4D4 bordercolordark=#FFFFFF bordercolor=#808080 bordercolorlight=#808080>&nbsp;&nbsp;<FONT FACE=ARIAL SIZE=2><img src=60-60-60.gif border=0 alt='SoftRobot Document Server'><B> Add " & rs("TITLE") & " Document</B></FONT></TD>" )
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
'Response.Write( "<TR>" )
'Response.Write( "<TD BGCOLOR=yellow><FONT FACE=ARIAL SIZE=1><B>Rules:" & rs("FOOTERNOTE") & "</B></FONT></TD>" )
'Response.Write( "</TR>" )
Response.Write( "<TR>" )
Response.Write( "<TD BGCOLOR=#d1d2d3><FONT FACE=ARIAL SIZE=1><I>" & DATE & "</I></FONT></TD>" )
Response.Write( "</TR>" )
Response.Write( "</TABLE>" )
End Sub
%>
<% 
'---------------------------------FORM GENERATION STARTS
Sub GenerateForm( rs, action )

  ' start form
  %> <table WIDTH=100% ><FORM METHOD=POST ACTION="<%= action %>"> <%

  ' build input field for each recordset field
  for i = 3 to rs.fields.count - 1

    ' determine size of input field
    size = rs(i).DefinedSize
    IF size>50 then size=50
    ' determine name of field
    name = "fld"+cstr(i)
    if i=3 then
        %><tr bgcolor=#E1F2FD><td><font size=1><img src=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td>(auto number)</td><td><font face=arial size=1></font></td></tr> <%
    elseif rs(i).name = "TOTAL" then
		%><tr bgcolor=#E1F2FD><td><font size=1><img src=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td>(auto number)</td><td><font face=arial size=1></font></td></tr> <%
    elseif rs(i).Type = 11 then
		%><tr bgcolor=#E1F2FD><td><font size=1><img src=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td><INPUT TYPE=checkbox NAME=<%= name %>></td><td><font face=arial size=1>(boolean only)</font></td></tr> <%
    '-----------------------------------------
    elseif rs(i).Type = 3 then
		Set ls=Server.CreateObject("ADODB.Recordset")
		ls.Open "Select * from LISTBOX where DDID=" & DDID & " AND COLUMNNO=" & i, Conn
		if not ls.EOF then
		LISTNAME=ls("LISTNAME")
		LISTTABLE=ls("LISTTABLE")
		LISTCOLUMN=ls("LISTCOLUMN")
		LISTVALUE=ls("LISTVALUE")
		%><tr bgcolor=#E1F2FD><td><font size=1><img src=required.gif> <%= LISTNAME %><%'= rs(i).Type %></font></td>
		<td>
		<%LIST1 LISTTABLE, LISTCOLUMN, LISTVALUE, Conn, name%> 
		</td><td><font face=arial size=1>(select only)</font></td></tr> <%
		else
		%><tr bgcolor=#E1F2FD><td><font size=1><img src=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td><INPUT TYPE=TEXT SIZE=<%= size %>  name=<%= name %>></td><td><font face=arial size=1>(Int only)</font></td></tr> <%
		end if
		ls.Close
		set ls=nothing
    '-----------------------------------------
    elseif rs(i).Type = 5 or rs(i).Type = 6 or rs(i).Type = 131 or rs(i).Type = 4 then
		%><tr bgcolor=#E1F2FD><td><font size=1><img src=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td><INPUT TYPE=TEXT SIZE=<%= size %>  name=<%= name %>></td><td><font face=arial size=1>(Float only)</font></td></tr> <%
    elseif rs(i).Type = 17 or rs(i).Type = 2 or rs(i).Type = 128 or rs(i).Type = 204 then
		%><tr bgcolor=#E1F2FD><td><font size=1><img src=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td><INPUT TYPE=TEXT SIZE=<%= size %>   name=<%= name %>></td><td><font face=arial size=1>(Integer only)</font></td></tr> <%
    elseif rs(i).Type = 135 then
		%><tr bgcolor=#E1F2FD><td><font size=1><img src=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td><INPUT TYPE=TEXT SIZE=<%= size %>    name=<%= name %>><INPUT onclick="fPopCalendar(<%= name %>,<%= name %>); return false" type=button value=V></td><td><font face=arial size=1>(Date only)</font></td></tr> <%
    elseif rs(i).Type = 72 then
		%><tr bgcolor=#E1F2FD><td><font size=1><img src=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td><INPUT TYPE=HIDDEN SIZE=<%= size %> name=<%= name %> value=0></td><td><font face=arial size=1>(Autono only)</font></td></tr> <%
    elseif rs(i).Type = 203 or rs(i).Type = 201 then
    	%><tr bgcolor=#E1F2FD><td><font size=1><IMG SRC=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td><textarea rows=3 cols=43  name=<%= name %>><%=value%></textarea></td><td><font face=arial size=1>(Text only)</font></td></tr> <%
    else
		%><tr bgcolor=#E1F2FD><td><font size=1><img src=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td><INPUT TYPE=TEXT SIZE=<%= size %>  maxlength=<%=size%> name=<%= name %>></td><td><font face=arial size=1>(Text only)</font></td></tr> <%
    end if
 next
%> 
  </table>
  <BR>
  <!---------#include file="password.inc"-------------> 
  <INPUT TYPE=HIDDEN NAME="i" VALUE=<%=i%>>
  <INPUT TYPE=HIDDEN NAME="DID" VALUE=<%=DID%>>
  <INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=UID%>>
  <INPUT TYPE=SUBMIT VALUE="Insert Document" name=Insert>
  <INPUT TYPE=RESET VALUE="Clear Form" name=Cancel>
   </FORM> 
  <%

End Sub
'-------------------------------------FUNCTION ENDS
%>
<%
set rsDOC = Server.CreateObject("ADODB.Recordset")
rsDOC.Open "Select * from DOCUMENTS WHERE DID=" & DID, Conn 
'--------------------------------------
GenerateHeader rsDoc
'-------------------------------------
SQLPROGRAM=rsDOC("ADDSQL"): IF ISNULL(SQLPROGRAM) THEN SQLPROGRAM=""
DETAILSQL=rsDOC("ADDDETAILS"): IF ISNULL(DETAILSQL) THEN DETAILSQL=""
DDID=rsDOC("DDID"): IF ISNULL(DDID) OR DDID="" THEN DDID=0
'--------------------
set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQLPROGRAM, Conn
GenerateForm rs, "formadd1.asp?MODE=0"
rs.Close
Set rs = Nothing
'---------------------
IF DETAILSQL<>"" THEN
Response.Write( "<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF>" )
set rs = Server.CreateObject("ADODB.Recordset")
rs.Open DETAILSQL, Conn
Response.Write( "<TR>" )
for i = 1 to rs.fields.count - 1
Response.Write ( "<TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080><font face=arial size=1 >" & rs(i).Name & "</font></td>")
next
Response.Write( "</TR>" )
Response.Write( "<TR>" )
for i = 1 to rs.fields.count - 1
Response.Write ( "<td bgcolor=#E1F2FD><font face=arial size=1>Data</font></td>")
next
Response.Write( "</TR>" )
rs.Close
Set rs = Nothing
Response.Write( "</TABLE>" )
END IF
'---------------------------------------
GenerateFooter rsDoc
'----------------------------------------
rsDOC.Close
Conn.Close
Set rsDOC = Nothing
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