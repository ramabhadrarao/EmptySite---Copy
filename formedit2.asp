<%' @TRANSACTION=Required LANGUAGE="VBScript" %>
<%'@ LANGUAGE="VBScript" %>
<%
Response.Buffer = True
Const adUseClient = 3
%>
<HTML>
<HEAD>
<TITLE>IRP FORM Edit (sales@erpweb)</TITLE>
</HEAD>

<BODY>
<!---------#include file="calendar.js"------------->
<Basefont face=arial size=1>
<%
TOTFLAG=0
Set Conn=Server.CreateObject("ADODB.Connection")
Conn.Open Session("erp_ConnectionString")
%>
<%
'----------------------------------LISTBOX GENERATOR
Sub LIST1 ( LISTTABLE, LISTCOLUMN, LISTVALUE, Conn, name, value )
Set lst=Server.CreateObject("ADODB.Recordset")
lst.Open "Select " & LISTVALUE & ", " & LISTCOLUMN & " FROM " & LISTTABLE & " ORDER BY " & LISTCOLUMN & " ASC", Conn
%>
	<SELECT  name=<%=name%>>
	<OPTION VALUE=0>NONE
	<%WHILE not lst.eof%>
	<%if lst(0)=value then%>
	<OPTION VALUE=<%=lst(0)%> selected><%=lst(1)%>
	<%else%>
	<OPTION VALUE=<%=lst(0)%>><%=lst(1)%>
	<%end if%>
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
Response.Write( "<TD BGCOLOR=#D1D2D3><FONT FACE=ARIAL SIZE=1><I>" & DATE & "</I></FONT></TD>" )
Response.Write( "</TR>" )
Response.Write( "</TABLE>" )
End Sub
%>

<% 
'---------------------------------FORM GENERATION STARTS
Sub GenerateForm( rs )
  ' start form
  %> <table WIDTH=100% ><%
  ' build input field for each recordset field
  for i = 1 to rs.fields.count - 1
    value = rs(i)
    if isNull(value) then value=""
    ' determine size of each field
    size=rs(i).DefinedSize
    IF size>50 then size=50
    ' determine name of field
    name = "fld"+cstr(i)
    if i=1 then
		%><tr bgcolor=#E1F2FD><td><font size=1><IMG SRC=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td><%=value%></td><td><font face=arial size=1></font></td></tr> <%
    elseif rs(i).name = "TOTAL" then
		%><tr bgcolor=#E1F2FD><td><font size=1><IMG SRC=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td>AutoComputed</td><td><font face=arial size=1></font></td></tr> <%
    %><INPUT TYPE=HIDDEN  name=<%= name %> value=0 ><%
    elseif rs(i).Type = 11 then
		%><tr bgcolor=#E1F2FD><td><font size=1><IMG SRC=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td><INPUT TYPE=checkbox  name=<%=name %> <%if value then Response.Write " checked " end if%>></td><td><font face=arial size=1>(boolean only)</font></td></tr> <%
    '-----------------------------------------
    elseif rs(i).Type = 3 then
        if value="" or isnull(value) or value<="0" then value=1
		Set ls=Server.CreateObject("ADODB.Recordset")
		ls.Open "Select * from LISTBOX where DETFLAG=1 AND DDID=" & DDID & " AND COLUMNNO=" & i, Conn
		if not ls.EOF then
		LISTNAME=ls("LISTNAME")
		LISTTABLE=ls("LISTTABLE")
		LISTCOLUMN=ls("LISTCOLUMN")
		LISTVALUE=ls("LISTVALUE")
		%><tr bgcolor=#E1F2FD><td><font size=1><IMG SRC=required.gif> <%= LISTNAME %><%'= rs(i).Type %></font></td>
		<td>
		<%LIST1 LISTTABLE, LISTCOLUMN, LISTVALUE, Conn, name, value%> 
		</td><td><font face=arial size=1>(select only)</font></td></tr> <%
		else
		%><tr><td><font size=1><IMG SRC=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td><INPUT TYPE=TEXT SIZE=<%= size %>   name=<%= name %> value=<%=value%>></td><td><font face=arial size=1>(Int only)</font></td></tr> <%
		end if
		ls.Close
		set ls=nothing
    '-----------------------------------------
    elseif rs(i).Type = 5 or rs(i).Type = 6 or rs(i).Type = 131 or rs(i).Type = 4 then
		%><tr bgcolor=#E1F2FD><td><font size=1><IMG SRC=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td><INPUT TYPE=TEXT SIZE=<%= size %>   name=<%= name %> value=<%=value%>></td><td><font face=arial size=1>(Float only)</font></td></tr> <%
    elseif rs(i).Type = 17 or rs(i).Type = 2 or rs(i).Type = 128 or rs(i).Type = 204 then
		%><tr bgcolor=#E1F2FD><td><font size=1><IMG SRC=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td><INPUT TYPE=TEXT SIZE=<%= size %>    name=<%= name %> value=<%=value%>></td><td><font face=arial size=1>(Integer only)</font></td></tr> <%
    elseif rs(i).Type = 135 then
		%><tr bgcolor=#E1F2FD><td><font size=1><IMG SRC=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td><INPUT TYPE=TEXT SIZE=<%= size %>     name=<%= name %> value=<%=value%>><INPUT onclick="fPopCalendar(<%= name %>,<%= name %>); return false" type=button value=V name=button2></td><td><font face=arial size=1>(Date only)</font></td></tr> <%
    elseif rs(i).Type = 72 then
		%><tr bgcolor=#E1F2FD><td><font size=1><IMG SRC=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td><INPUT TYPE=HIDDEN SIZE=<%= size %>  name=<%= name %> value=<%=value%>></td><td><font face=arial size=1>(Autono only)</font></td></tr> <%
    elseif rs(i).Type = 203 or rs(i).Type = 201 then
    	%><tr bgcolor=#E1F2FD><td><font size=1><IMG SRC=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td><textarea rows=3 cols=43  name=<%= name %>><%=value%></textarea></td><td><font face=arial size=1>(Text only)</font></td></tr> <%
    else
		%><tr bgcolor=#E1F2FD><td><font size=1><IMG SRC=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td><INPUT TYPE=TEXT SIZE=<%= size %>  maxlength=<%=size%>  name=<%= name %> value=<%="'" & value & "'"%>></td><td><font face=arial size=1>(Text only)</font></td></tr> <%
    end if
 next
%> 
  </table>
  <INPUT TYPE=HIDDEN NAME="i" VALUE=<%=i%>>
<%
End Sub
'-------------------------------------FUNCTION ENDS
%>
<%
'-----------------------------------------MAIN PROG STARTS
DID=REQUEST("DID")'IDENTIFY DOCUMENT
UID=REQUEST("UID")'IDENTIFY USER
ID=REQUEST("ID")'RECORD ID
if DID="" Then Response.Write "ERROR:DID IS NULL":Response.End
if UID="" Then Response.Write "ERROR:UID IS NULL":Response.End
if ID="" Then Response.Write "ERROR:ID IS NULL":Response.End
%>
<%
'-------------------------------------
set rsDOC = Server.CreateObject("ADODB.Recordset")
rsDOC.Open "Select * from DOCUMENTS WHERE DID=" & DID, Conn 
IF NOT rsDOC.EOF THEN
SQLPROGRAM=rsDOC("DETAILSSQL")
DDID=rsDOC("DDID")
if isnull(SQLPROGRAM) THEN SQLPROGRAM=""
END IF
'-------------------------------------
GenerateHeader rsDoc
'-------------------------------------
%>
<FORM METHOD=POST ACTION="formedit21.asp"> 
<%
IF SQLPROGRAM <> "" THEN
	set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open TRIM(SQLPROGRAM) & ID, Conn
	if not rs.EOF then
	GenerateForm rs
	else
	Response.Write "Error: Problems in executing query" & SQLPROGRAM & ID
	END IF
	rs.Close
	Set rs = Nothing	
ELSE
	Response.Write "PLEASE ENTER EDITDETAILS"
END IF
%>
<BR>
<!---------#include file="password.inc"-------------> 
<INPUT TYPE=HIDDEN NAME="DID" VALUE=<%=DID%>>
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=UID%>>
<INPUT TYPE=HIDDEN NAME="ID" VALUE=<%=ID%>>
<INPUT TYPE=HIDDEN NAME="IDD" VALUE=<%=REQUEST("IDD")%>>
<INPUT TYPE=SUBMIT VALUE="Update Document" name=Insert>
</FORM> 
<%
'-------------------------------------	
	GenerateFooter rsDoc
'-------------------------------------
rsDOC.Close
Conn.Close
Set rsDOC = Nothing
SET Conn = nothing
%>
<%
ID=REQUEST("IDD")
%>
<font face=arial size=2>
<a href=FORMEDIT.ASP?DID=<%=DID%>&ID=<%=ID%>&UID=<%=UID%>><b>GOTO Edit Page</b></a>
</font>
</BODY>
</HTML>
<!-------------------------------------start transaction server---->
