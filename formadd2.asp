
<%
Response.Buffer = True
Const adUseClient = 3
ID=REQUEST("ID") 'identify record number
DID=REQUEST("DID")'IDENTIFY DOCUMENT
UID=REQUEST("UID")'IDENTIFY USER
if DID="" Then Response.Write "ERROR:DID" : Response.END
if UID="" Then Response.Write "ERROR:UID" : Response.END
if ID="" Then Response.Write "ERROR:ID" : Response.END
%>
<HTML>
<HEAD>
<TITLE>IRP FORM ADD (sales@erpweb)</TITLE>
</HEAD>
<BODY>
<!---------#include file="calendar.js"------------->
<Basefont face=arial size=1>
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
Sub GenerateHeader( )
Response.Write( "<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF >" )
Response.Write( "<TR>" )
Response.Write( "<TD valign=middle bgcolor=#D4D4D4 bordercolordark=#FFFFFF bordercolor=#808080 bordercolorlight=#808080>&nbsp;&nbsp;<FONT FACE=ARIAL SIZE=2><img src=60-60-60.gif border=0 alt='SoftRobot Document Server'><B> Add New Document</B></FONT></TD>" )
Response.Write( "</TR>" )
Response.Write( "</TABLE>" )
End Sub
%>
<%
Sub GenerateFooter( )
Response.Write( "<TABLE BORDER=0 WIDTH=100% >" )
Response.Write( "<TR>" )
Response.Write( "<TD BGCOLOR=#6699CC><FONT FACE=ARIAL SIZE=1><I>" & DATE & "</I></FONT></TD>" )
Response.Write( "</TR>" )
Response.Write( "</TABLE>" )
End Sub
%>
<% 
'---------------------------------FORM GENERATION STARTS
Sub GenerateForm1( rs )
' start form
%>
<table WIDTH=100% >
<%
  ' build input field for each recordset field
  for i = 3 to rs.fields.count - 1
  
      value = rs(i)
      if isNull(value) then value=""
    if i=3 then
    %><tr bgcolor=#E1F2FD><td><font face=arial size=1><IMG SRC=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td><font face=arial size=1><%=value%></font></td></tr> <%
    '-----------------------------------------
    elseif (rs(i).Type = 3) and (value > "0") then
		Set ls=Server.CreateObject("ADODB.Recordset")
		ls.Open "Select * from LISTBOX where DETFLAG=0 AND DDID=" & DDID & " AND COLUMNNO=" & i, Conn
		if not ls.EOF then
		LISTNAME=ls("LISTNAME")
		LISTTABLE=ls("LISTTABLE")
		LISTCOLUMN=ls("LISTCOLUMN")
		LISTVALUE=ls("LISTVALUE")
		%><tr bgcolor=#E1F2FD><td><font face=arial size=1><IMG SRC=required.gif> <%= LISTNAME %><%'= rs(i).Type %></font></td>
		<td>
		<%LIST2 LISTTABLE, LISTCOLUMN, LISTVALUE, Conn, name, value%> 
		</td></tr> <%
		else
		%><tr bgcolor=#E1F2FD><td><font face=arial size=1><IMG SRC=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td><font face=arial size=1><%=value%></font></td></tr> <%
		end if
		ls.Close
		set ls=nothing
    '-----------------------------------------
    else
		%><tr bgcolor=#E1F2FD><td><font face=arial size=1><IMG SRC=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td><font face=arial size=1><%=value%></font></td></tr> <%
    end if
 next
%> 
</table>
<%END SUB%>

<% 
'---------------------------------FORM GENERATION STARTS
Sub GenerateForm( rs, action, ID )

  ' start form
  %> <table WIDTH=100% ><FORM METHOD=POST ACTION="<%= action %>"> <%
  ' build input field for each recordset field
  for i = 0 to rs.fields.count - 1

    ' determine size of input field
    size = rs(i).DefinedSize
    if size>50 then size=50
    ' determine name of field
    name = "fld"+cstr(i)
    if i=1 then
		%><tr bgcolor=#E1F2FD><td><font face=arial size=1><IMG SRC=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td>(auto number)</td><td><font face=arial size=1></font></td></tr> <%
    elseif i=0 then%>
        <INPUT TYPE=HIDDEN NAME=<%= name %> value=<%=ID%>><%
    elseif rs(i).Name = "TOTAL" then
		%><tr bgcolor=#E1F2FD><td><font face=arial size=1><IMG SRC=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td>(auto total)</td><td><font face=arial size=1></font></td></tr> <%
    elseif rs(i).Type = 11 then
		%><tr bgcolor=#E1F2FD><td><font face=arial size=1><IMG SRC=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td><INPUT TYPE=checkbox NAME=<%= name %>></td><td><font face=arial size=1>(boolean only)</font></td></tr> <%
    '-----------------------------------------
    elseif rs(i).Type = 3 then
		Set ls=Server.CreateObject("ADODB.Recordset")
		ls.Open "Select * from LISTBOX where DETFLAG=1 AND DDID=" & DDID & " AND COLUMNNO=" & i, Conn
		if not ls.EOF then
		LISTNAME=ls("LISTNAME")
		LISTTABLE=ls("LISTTABLE")
		LISTCOLUMN=ls("LISTCOLUMN")
		LISTVALUE=ls("LISTVALUE")
		%><tr bgcolor=#E1F2FD><td><font face=arial size=1><IMG SRC=required.gif> <%= LISTNAME %><%'= rs(i).Type %></font></td>
		<td>
		<%LIST1 LISTTABLE, LISTCOLUMN, LISTVALUE, Conn, name%> 
		</td><td><font face=arial size=1>(select only)</font></td></tr> <%
		else
		%><tr bgcolor=#E1F2FD><td><font face=arial size=1><IMG SRC=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td><INPUT TYPE=TEXT SIZE=<%= size %>  name=<%= name %>></td><td><font face=arial size=1>(Int only)</font></td></tr> <%
		end if
		ls.Close
		set ls=nothing
    '-----------------------------------------
    elseif rs(i).Type = 5 or rs(i).Type = 6 or rs(i).Type = 131 or rs(i).Type = 4 then
		%><tr bgcolor=#E1F2FD><td><font face=arial size=1><IMG SRC=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td><INPUT TYPE=TEXT SIZE=<%= size %>  name=<%= name %>></td><td><font face=arial size=1>(Float only)</font></td></tr> <%
    elseif rs(i).Type = 17 or rs(i).Type = 2 or rs(i).Type = 128 or rs(i).Type = 204 then
		%><tr bgcolor=#E1F2FD><td><font face=arial size=1><IMG SRC=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td><INPUT TYPE=TEXT SIZE=<%= size %>   name=<%= name %>></td><td><font face=arial size=1>(Integer only)</font></td></tr> <%
    elseif rs(i).Type = 135 then
		%><tr bgcolor=#E1F2FD><td><font face=arial size=1><IMG SRC=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td><INPUT TYPE=TEXT SIZE=<%= size %>    name=<%= name %>><INPUT onclick="fPopCalendar(<%= name %>,<%= name %>); return false" type=button value=V></td><td><font face=arial size=1>(Date only)</font></td></tr> <%
    elseif rs(i).Type = 72 then
		%><tr bgcolor=#E1F2FD><td><font face=arial size=1><IMG SRC=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td><INPUT TYPE=HIDDEN SIZE=<%= size %> name=<%= name %> value=0></td><td><font face=arial size=1>(Autono only)</font></td></tr> <%
    elseif rs(i).Type = 203 or rs(i).Type = 201 then
    	%><tr bgcolor=#E1F2FD><td><font face=arial size=1><IMG SRC=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td><textarea rows=3 cols=43  name=<%= name %>><%=value%></textarea></td><td><font face=arial size=1>(Text only)</font></td></tr> <%
    else
		%><tr bgcolor=#E1F2FD><td><font face=arial size=1><IMG SRC=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td><INPUT TYPE=TEXT SIZE=<%= size %>  maxlength=<%=size%> name=<%= name %>></td><td><font face=arial size=1>(Text only)</font></td></tr> <%
    end if
 next
%> 
</table>
  <BR>
  <!---------#include file="password.inc"-------------> 
  <INPUT TYPE=HIDDEN NAME="i" VALUE=<%=i%>>
  <INPUT TYPE=HIDDEN NAME="DID" VALUE=<%=DID%>>
  <INPUT TYPE=HIDDEN NAME="ID" VALUE=<%=ID%>>
  <INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=UID%>>
  <INPUT TYPE=SUBMIT VALUE="Insert Document" name=Insert>
  <INPUT TYPE=RESET VALUE="Clear Form" name=Cancel>
  </FORM> 
 <%

End Sub
'-------------------------------------FUNCTION ENDS
%>
<% 
Sub GenerateTable2( rs )
	
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
  srno=1
  TOTAL=0
  '------------------------------------------
  ' write each row
  while NOT rs.EOF 
    Response.Write( "<TR>" )
    Response.Write( "<TD VALIGN=TOP bgcolor=#E1F2FD><FONT FACE=ARIAL SIZE=1>" + CStr( srno ) + "</FONT></TD>" )
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
  Response.Write( "<TR><TD ALIGN=RIGHT BGCOLOR=#D1D2D3><FONT FACE=ARIAL SIZE=1>Grand Total: " + CStr( TOTAL ) + "</FONT></TD>" )
  Response.Write( "</TABLE>" )
  END IF
'-----------------------
End Sub
%>
<%
'--------------------------------------
set rsDOC = Server.CreateObject("ADODB.Recordset")
rsDOC.Open "Select * from DOCUMENTS WHERE DID=" & DID, Conn 
SQLPROGRAM1=rsDOC("MASTERSQL")
SQLPROGRAM=rsDOC("ADDDETAILS")
SQLPROGRAM2=rsDOC("SQLDETAILS")
DDID=rsDOC("DDID")
'-------------------------
GenerateHeader
'--------------------------------------
set rs=Server.CreateObject("ADODB.Recordset")
rs.Open SQLPROGRAM1 & ID, Conn
'------------------------------set data after validation  
  GenerateForm1 rs
'-------------------------
rs.Close
'-------------
set rs=Server.CreateObject("ADODB.Recordset")
rs.Open SQLPROGRAM2 & ID, Conn
'------------------------------set data after validation  
  GenerateTable2 rs
'-------------------------
rs.Close
Response.Write "<HR>"
Response.Write( "<TABLE width=100% >" )
  Response.Write( "<TR><TD BGCOLOR=#C0C0C0><FONT FACE=ARIAL SIZE=1>Enter New Details Page</FONT></TD>" )
  Response.Write( "</TABLE>" )
'-------------------------
  set rs = Server.CreateObject("ADODB.Recordset")
  rs.Open SQLPROGRAM, Conn
'---------------------------------------  
  GenerateForm rs, "formadd21.asp", ID
'---------------------------------------
GenerateFooter
'----------------------------------------
%>
<%
rs.Close	
rsDOC.Close
Conn.Close
Set rs = Nothing
Set rsDOC = Nothing
SET Conn = nothing
%>
<font face=arial size=2>
<a href=FORMEDIT.ASP?DID=<%=DID%>&ID=<%=ID%>&UID=<%=UID%>><b>GOTO Secure Edit Form</b></a>
</font>
</BODY>
</HTML>
<!-------------------------------------start transaction server---->

