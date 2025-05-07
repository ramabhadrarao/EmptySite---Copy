<%' @TRANSACTION=Required LANGUAGE="VBScript" %>
<%'@ LANGUAGE="VBScript" %>
<%
Response.Buffer = True
Const adUseClient = 3
%>
<HTML>
<HEAD>
<TITLE>ERPWEB Edit (sales@ERPWEB.com)</TITLE>
</HEAD>

<BODY>
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
TOTFLAG=0
Set Conn=Server.CreateObject("ADODB.Connection")
Conn.Open Session("erp_ConnectionString")
%>
<%
'----------------------------------LISTBOX GENERATOR
Sub LIST2 ( LISTTABLE, LISTCOLUMN, LISTVALUE, Conn, name, value )
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
Sub GenerateTable( rs )
	
  Response.Write( "<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF>" )
   '--------------------------------------
  ' set up column names
  Response.Write( "<TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080><FONT FACE=ARIAL SIZE=1 >Sr#</FONT></TD>" )
  for i = 2 to rs.fields.count - 1
        if rs(i).Type = 3 then
        Set ls=Server.CreateObject("ADODB.Recordset")
		ls.Open "Select * from LISTBOX where DETFLAG=1 AND DDID=" & DDID & " AND COLUMNNO=" & i, Conn
		if not ls.EOF then
		LISTNAME=ls("LISTNAME") 
        Response.Write("<TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080 align=right><FONT FACE=ARIAL SIZE=1 >" + LISTNAME + "</FONT></TD>")
        ELSE
        Response.Write("<TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080 align=right><FONT FACE=ARIAL SIZE=1 >" + rs(i).Name + "</FONT></TD>")
        end if
        ls.Close
        set ls=nothing
        ELSE
        Response.Write("<TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080 align=right><FONT FACE=ARIAL SIZE=1 >" + rs(i).Name + "</FONT></TD>")
        end if
    next
    Response.Write("<TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080 align=right><FONT FACE=ARIAL SIZE=1 >Action</FONT></TD>")
        
  srno=1
  TOTAL=0
  '------------------------------------------
  ' write each row
  while NOT rs.EOF 
    Response.Write( "<TR>" )
    Response.Write( "<TD VALIGN=TOP bgcolor=#E1F2FD><FONT FACE=ARIAL SIZE=1 >" + CStr( srno ) + "</FONT></TD>" )
	srno=srno+1
    '--------------------------------
    FLAG=NOT FLAG
    for i = 2 to rs.fields.count - 1
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
		if FLAG THEN
		%>
		<TD VALIGN=TOP bgcolor=#E1F2FD align=right><FONT FACE=ARIAL SIZE=1>
		<%LIST2 LISTTABLE, LISTCOLUMN, LISTVALUE, Conn, name, v%> 
		</font></td>
		<%ELSE%>
		<TD VALIGN=TOP bgcolor=#E1F2FF align=right><FONT FACE=ARIAL SIZE=1>
		<%LIST2 LISTTABLE, LISTCOLUMN, LISTVALUE, Conn, name, v%> 
		</font></td>
		<%END IF%>
		<%
	else
		IF FLAG THEN
			Response.Write( "<TD VALIGN=TOP bgcolor=#E1F2FD align=right><FONT FACE=ARIAL SIZE=1>" + CStr( v ) + "</FONT></TD>" )
		   ELSE
		    Response.Write( "<TD VALIGN=TOP bgcolor=#E1F2FF align=right><FONT FACE=ARIAL SIZE=1>" + CStr( v ) + "</FONT></TD>" )
		END IF
		end if
		ls.Close
		set ls=nothing
    '-----------------------------------------
   		else
   		 IF FLAG THEN
			Response.Write( "<TD VALIGN=TOP bgcolor=#E1F2FD align=right><FONT FACE=ARIAL SIZE=1>" + CStr( v ) + "</FONT></TD>" )
		   ELSE
		    Response.Write( "<TD VALIGN=TOP bgcolor=#E1F2FF align=right><FONT FACE=ARIAL SIZE=1>" + CStr( v ) + "</FONT></TD>" )
		END IF  
    end if
    '-------------------------------
    IF rs(i).Name="TOTAL" THEN
    TOTFLAG=1
    TOTAL=TOTAL+rs("TOTAL")
    END IF
    '-------------------------------
    next
    '-------------------------------
    Response.Write("<TD bgcolor=#E1F2FD width=40 align=right><FONT FACE=ARIAL SIZE=1>")
    Response.Write("<A HREF=FORMEDIT2.ASP?DID=" & DID & "&ID=" & rs(1) & "&UID=" & UID & "&IDD=" & IDD & "><img src=update.gif BORDER=0 WIDTH=20></A>")
    Response.Write("<A HREF=FORMDEL2.ASP?DID=" & DID & "&ID=" & rs(1) & "&UID=" & UID & "&IDD=" & IDD & "><img src=delete.gif BORDER=0 WIDTH=20></A>")
    Response.Write("</FONT></TD></tr>")
    rs.MoveNext
  wend 
  '-------------------------------------
  Response.Write( "</TABLE>" )
  IF TOTFLAG THEN
  Response.Write( "<TABLE width=100% >" )
  Response.Write( "<TR BGCOLOR=#D1D2D3><TD ALIGN=RIGHT><FONT FACE=ARIAL SIZE=1>Total:       " + CStr( TOTAL ) + "      --------------------</FONT></TD>" )
  Response.Write( "</TABLE>" )
  END IF
'-----------------------
End Sub
%>


<% 
'---------------------------------FORM GENERATION STARTS
Sub GenerateForm( rs )

  ' start form
  %> <table WIDTH=100% ><%

  ' build input field for each recordset field
  for i = 3 to rs.fields.count - 1

      value = rs(i)
      if isNull(value) then value=""
    ' determine size of each field
    size=rs(i).DefinedSize
    maxsize=size
    IF size>50 then size=50
    ' determine name of field
    name = "fld"+cstr(i)
    if i=3 then
    	%><tr bgcolor=#E1F2FD><td><font size=1><IMG SRC=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td><%=value%></td><td><font face=arial size=1></font></td></tr> <%
    elseif rs(i).name = "TOTAL" then
		%><tr bgcolor=#E1F2FD><td><font size=1><IMG SRC=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td>AutoComputed</td><td><font face=arial size=1></font></td></tr> <%
    %><INPUT TYPE=HIDDEN name=<%= name %> value=0><%
    elseif rs(i).Type = 11 then
		%><tr bgcolor=#E1F2FD><td><font size=1><IMG SRC=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td><INPUT TYPE=checkbox  name=<%=name %> <%if value then Response.Write " checked " end if%>></td><td><font face=arial size=1>(boolean only)</font></td></tr> <%
    '-----------------------------------------
    elseif (rs(i).Type = 3) then
        if value="" or isnull(value) or value<="0" then value=1
		Set ls=Server.CreateObject("ADODB.Recordset")
		ls.Open "Select * from LISTBOX where DETFLAG=0 AND DDID=" & DDID & " AND COLUMNNO=" & i, Conn
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
		%><tr bgcolor=#E1F2FD><td><font size=1><IMG SRC=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td><INPUT TYPE=TEXT SIZE=<%= size %>   name=<%= name %> value=<%=value%>></td><td><font face=arial size=1>(Int only)</font></td></tr> <%
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
		%><tr bgcolor=#E1F2FD><td><font size=1><IMG SRC=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td><INPUT TYPE=HIDDEN name=<%= name %> value=<%=value%>></td><td><font face=arial size=1>(Autono only)</font></td></tr> <%
    elseif rs(i).Type = 203 or rs(i).Type = 201 then
    	%><tr bgcolor=#E1F2FD><td><font size=1><IMG SRC=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td><textarea rows=3 cols=43 name=<%= name %>><%=value%></textarea></td><td><font face=arial size=1>(Text only)</font></td></tr> <%
    else
		%><tr bgcolor=#E1F2FD><td><font size=1><IMG SRC=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td><INPUT TYPE=TEXT SIZE=<%= size %>  maxlength=<%=maxsize%>   name=<%= name %> value=<%="'" & value & "'"%>></td><td><font face=arial size=1>(Text only)</font></td></tr> <%
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
SQLPROGRAM=rsDOC("MASTERSQL")
DETAILSSQL=rsDOC("SQLDETAILS")
DDID=rsDOC("DDID")
if isnull(SQLPROGRAM) THEN SQLPROGRAM=""
if isnull(DETAILSSQL) THEN DETAILSSQL=""
END IF
'-------------------------------------
	GenerateHeader rsDoc
'-------------------------------------
%>
<FORM METHOD=POST ACTION="formedit1.asp"> 
<%
IF SQLPROGRAM <> "" THEN
	set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open SQLPROGRAM & ID, Conn
	if not rs.EOF then
	IDD=rs(3) 'MAIN ID SUPPLIED
	GenerateForm rs
	else
	Response.Write "Error: Problems in executing query" & SQLPROGRAM & ID
	END IF
	rs.Close
	Set rs = Nothing	
ELSE
	Response.Write "PLEASE ENTER MASTERSQL"
END IF
'-------------------------------------
IF DETAILSSQL <> "" THEN 
	set rs1 = Server.CreateObject("ADODB.Recordset")
	rs1.Open DETAILSSQL & ID, Conn
	IF NOT rs1.EOF then
	GenerateTable rs1
	else
	Response.Write "<HR>NO RECORDS FOUND"
	END IF
	rs1.Close
	Set rs1=Nothing
'-------------------------------------
  Response.Write( "</TABLE>" )
  Response.Write( "<TABLE width=100% >" )
  Response.Write( "<TR><TD ALIGN=RIGHT WIDTH=90% ></TD>" )
  Response.Write("<TD ALIGN=RIGHT><A HREF=FORMADD2.ASP?UID=" & UID & "&DID=" & DID & "&ID=" & ID & "><img src=add.jpg border=0></A></TD></TR>")
  Response.Write( "</TABLE>" )
'-----------------------
END IF
%>
<BR>
  <!---------#include file="password.inc"-------------> 
  <INPUT TYPE=HIDDEN NAME="DID" VALUE=<%=DID%>>
  <INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=UID%>>
  <INPUT TYPE=HIDDEN NAME="ID" VALUE=<%=ID%>>
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
'REDURL="document.asp?DID=" & DID & "&UID=" & UID & "&FADD=" & FADD & "&FDEL=" & FDEL & "&FVIEW=" & FVIEW & "&FEDIT=" & FEDIT & "&FILTER=" & FAPPROVE & "&FOFFHOLD=" & FOFFHOLD & "&FREJECT=" & FREJECT
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