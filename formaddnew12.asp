<!--
'************************************************************************
'Pupose						:	This is a SoftRobot document add server
'Filename					:	formaddnew.asp
'Author						:	Anita Shah
'Created					:	27-AUG-2006
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
DETID=REQUEST("DETID")'FOR MORE LINKED FORM FEATURE
IF ISNULL(DETID) OR DETID="" THEN DETID=0
UFLD=REQUEST("FLD"): IF ISNULL(UFLD) THEN UFLD=0
%>
<%
'Response.Buffer = True
Const adUseClient = 3
DID=REQUEST("DID")'IDENTIFY DOCUMENT
UID=REQUEST("UID")'IDENTIFY USER
CID=REQUEST("CID")'IDENTIFY YEAR
if DID="" Then Response.Write "ERROR:DID" : Response.END
if UID="" Then Response.Write "ERROR:UID" : Response.END
if CID="" Then Response.Write "ERROR:CID" : Response.END
%>

<HTML>
<HEAD>
<TITLE>ERPWEB FORM ADD SoftServer (sales@ERPWEB.com)</TITLE>
</HEAD>
<BODY topmargin=0 >
<!---------#include file="calendar.js"------------->
<Basefont face=arial size=1>
<!-----------------------------HEADER STRIP--------------------> 


<!--------------------------------------------------------------->
<%
Set Conn=Server.CreateObject("ADODB.Connection")
Conn.Open Session("erp_ConnectionString")
%>

<%
'----------------------------------LISTCOUNT GENERATOR
Function LISTCNT ( LISTTABLE, Conn)
Set lst=Server.CreateObject("ADODB.Recordset")
lst.Open "Select count(*) FROM " & LISTTABLE, Conn
%>
	<%LISTCNT=lst(0):if ISNULL(LISTCNT) or LISTCNT="" THEN LISTCNT=0%>
	<%lst.Close%>
<%
End Function
%>
<%
'----------------------------------LISTBOX GENERATOR
Sub LIST1 ( LISTTABLE, LISTCOLUMN, LISTVALUE, Conn, name )
Set lst=Server.CreateObject("ADODB.Recordset")
lst.Open "Select " & LISTVALUE & ", " & LISTCOLUMN & " FROM " & LISTTABLE & " ORDER BY " & LISTCOLUMN & " ASC", Conn
%>
	<SELECT name=<%=name%> style="background: #E1F2FD">
	<OPTION VALUE=0 SELECTED>NONE
	<%WHILE not lst.eof%>
	<OPTION VALUE=<%=lst(0)%>><%=cstr(lst(1))%>
	<%lst.MoveNext%>
	<%wend%>
	</SELECT>
	<%lst.Close%>
<%
End Sub
%>
<%
'----------------------------------LISTBOX GENERATOR
Sub LIST2 ( LISTTABLE, LISTCOLUMN, LISTVALUE, Conn, name )
Set lst=Server.CreateObject("ADODB.Recordset")
lst.Open "Select " & LISTVALUE & ", " & LISTCOLUMN & " FROM " & LISTTABLE & " ORDER BY " & LISTCOLUMN & " ASC", Conn
FRST=1
%>

    <INPUT TYPE=RADIO name=<%=name%> VALUE=0 style="background: #E1F2FD"><font face=arial size=1>NONE</font>

	<%WHILE not lst.eof%>
	<%IF FRST THEN%>
    <INPUT TYPE=RADIO name=<%=name%> VALUE=<%=lst(0)%> CHECKED style="background: #E1F2FD"><font face=arial size=1><%=cstr(lst(1))%></font>
	<%FRST=0%>
	<%ELSE%>
	<INPUT TYPE=RADIO name=<%=name%> VALUE=<%=lst(0)%> style="background: #E1F2FD"><font face=arial size=1><%=cstr(lst(1))%></font>
	<%END IF%>
	<%lst.MoveNext%>
	<%wend%>
	<%lst.Close%>
<%
End Sub
%>
<%
Sub GenerateHeader( rs )
Response.Write( "<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF >" )
Response.Write( "<TR>" )
Response.Write( "<TD valign=middle bgcolor=#D4D4D4 bordercolordark=#FFFFFF bordercolor=#808080 bordercolorlight=#808080>&nbsp;&nbsp;<FONT FACE=ARIAL SIZE=2><img src=60-60-60.gif border=0 alt='SoftRobot Document Server'><B> " & rs("TITLE") & " Document</B></FONT></TD>" )
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
Response.Write( "<TD BGCOLOR=#C0C0C0><FONT FACE=ARIAL SIZE=1><I>" & DATE & "</I></FONT></TD>" )
Response.Write( "</TR>" )
Response.Write( "</TABLE>" )
End Sub
%>
<% 
'---------------------------------FORM GENERATION STARTS
Sub GenerateForm( rs, action, DDID, Conn, CFIELD, DAUTONO )' start form
%> 
  <table WIDTH=100% >
  <FORM NAME=form1 METHOD=POST ACTION="<%=action%>" > 
  <tr  valign=TOP>
  <%COLCNT=0%>
  <%
  ' build input field for each recordset field
  for i = 3 to rs.fields.count - 1
  '----------------------------------
  '----------------------------------
  IF COLCNT=3 THEN
  Response.Write ("<tr VALIGN=TOP>")
  COLCNT=0
  END IF
'--------------------------------------
    ' determine size of input field
    size = rs(i).DefinedSize
    IF size>20 then size=20
    ' determine name of field
    name = "mastfld"+ cstr(i)
'----------------------------------------
'--------------------FIND MATCHING FIELD NAME FROM DICTIONARY
	set rsdic = Server.CreateObject("ADODB.Recordset")
	rsdic.Open "SELECT * FROM DICTIONARY WHERE DETFLAG=0 AND DDID=" & DDID & " AND SEQ=" & i, Conn
	IF NOT rsdic.eof then
	FLDNM=rsdic("NEWFLDNM")
	FNT=rsdic("FONT")
	VF=rsdic("VALIDFROM")'APPLICABLE TO ONLY FLOAT FIELDS
	VT=rsdic("VALIDTO")'APPLICABLE TO ONLY FLOAT FIELDS
	DV=rsdic("DEFAULTVALUE")'APPLICABLE TO ONLY FLOAT FIELDS
	LK=rsdic("LOOKUP")
	FNTFAMILY="background: #E1F2FD;font-family: " & FNT
	else
	FLDNM=rs(i).Name
	FNT="ARIAL"
	FNTFAMILY="font-family: ARIAL;background: #E1F2FD"
	VF=0
	VT=99999999999999999999999999
	DV=0
	end if
	rsdic.Close
	Set rsdic = Nothing	
'-----------------------------------------
    if i=3 then'--------------------
        %><td colspan=2><font size=1 face="<%=FNT%>"><img src=required.gif> <%=FLDNM %><%'= rs(i).Type %>: </font><FONT FACE=ARIAL SIZE=1>(auto number)</FONT></td>
		<%COLCNT=COLCNT+1%>
	<%elseif i=4 AND DETID>0 then'--------------------
        %><td colspan=2><font size=1 face="<%=FNT%>"><img src=required.gif> <%=FLDNM %><%'= rs(i).Type %>: </font><FONT FACE=ARIAL SIZE=1><INPUT type=TEXT READONLY NAME=<%=name%> value=<%=DETID%> ></FONT></td>
		<%COLCNT=COLCNT+1%>
		<%
	elseif rs(i).name = "UPLOADFILE" then '--------------------------------Uploading File...
		%>
	<td><font size=1 face="<%=FNT%>"><IMG SRC=required.gif> <%=FLDNM  %><%'= rs(i).Type %></font><br>
	<A HREF=Process_File.asp target=uploadshow>Upload</A>
	<br><font face=arial size=1>(Upload File)</font></td>
	<%
    elseif rs(i).name = "CALENDERID" then
		%><INPUT type=hidden NAME=<%=name%> value=<%=CID%>>
		<%
    elseif rs(i).name = "UID" AND DID<>3174 then
		%><INPUT type=hidden NAME=<%=name%> value=<%=UID%>>
	    <%
    elseif rs(i).name = "TOTAL" then
		%><td><font size=1 face="<%=FNT%>"><img src=required.gif> <%=FLDNM %><%'= rs(i).Type %></font><br><FONT FACE=ARIAL SIZE=1>(auto number)</FONT></td> <%
    elseif rs(i).name = "ORDERVALUE" then
		%><td><font size=1 face="<%=FNT%>"><img src=required.gif> <%=FLDNM %><%'= rs(i).Type %></font><br><FONT FACE=ARIAL SIZE=1>(auto number)</FONT></td> <%
   	elseif rs(i).name = CFIELD then
		%><td><font size=1 face="<%=FNT%>"><img src=required.gif> <%=FLDNM %><%'= rs(i).Type %></font><br><INPUT TYPE=TEXT SIZE=<%=size %>  VALUE=<%=DAUTONO%> name=<%= name %>></td> <%
   	elseif rs(i).Type = 11 then
		%><td><font size=1 face="<%=FNT%>"><img src=required.gif> <%=FLDNM %><%'= rs(i).Type %></font><br><INPUT TYPE=checkbox NAME=<%=name %> style="background: #E1F2FD; color: #FF00FF"><br><font face=arial size=1>(boolean only)</font></td> <%
    '-----------------------------------------
    elseif rs(i).Type = 3 then
		Set ls=Server.CreateObject("ADODB.Recordset")
		ls.Open "Select * from LISTBOX where DETFLAG=0 AND DDID=" & DDID & " AND COLUMNNO=" & i, Conn
		if NOT ls.eof then'-----
		LISTID=ls("LISTBOXID")
	    LISTNAME=ls("LISTNAME")
		LISTTABLE=ls("LISTTABLE")
		LISTCOLUMN=ls("LISTCOLUMN")
		LISTVALUE=ls("LISTVALUE")
		'LINKURL=ls("LINKURL"):IF ISNULL(LINKURL) OR LINKURL="" THEN LINKURL="help.htm"
		LINKURL="listbox.asp?LID=" & LISTID & "&NM=" & name
		%>
		<%LCNT=LISTCNT( LISTTABLE, Conn)%> 
		<%
		'IF LCNT<=2 THEN
		%>
		<!--
		<td><font size=1 face="<%=FNT%>"><img src=required.gif> <%= LISTNAME %><%'= rs(i).Type %></font><br>
		<%'LIST2 LISTTABLE, LISTCOLUMN, LISTVALUE, Conn, name%> 
		<br><font face=arial size=1>(select only)</font></td>
		-->
		<%
		IF LCNT<50 THEN '--------------------
		%>
        <td><font size=1 face="<%=FNT%>"><img src=required.gif> <%= LISTNAME %><%'= rs(i).Type %></font><br>
		<%LIST1 LISTTABLE, LISTCOLUMN, LISTVALUE, Conn, name%> 
		<br><font face=arial size=1>(select only)</font></td> 
		<%
		ELSE'--------------------
		%>
		<td><font size=1 face="<%=FNT%>"><img src=required.gif> <%=LISTNAME %><%'= rs(i).Type %></font>
		<br><INPUT TYPE=TEXT SIZE=<%= size %> style="<%=FNTFAMILY%>" name=<%= name %> >
		<a href="#"  onclick="window.open('<%=LINKURL%>','popup','scrollbars=1,width=620,height=620,top=50,left=200')" title="Listbox" class="toplinks1"><img src=search.gif border=0></a>
		
		<!--
		<img src=search.gif border=0 onclick="document.all.<%=name%>1.style.display='inline'" alt="Expand">
		<img src=nosearch.gif onclick="document.all.<%=name%>1.style.display='none'" alt="Compress">
		<IFRAME name=<%=name%>1 src=<%=LINKURL%> STYLE="height:250;width:300;.display:none">
		</IFRAME>
		-->
		<br><font face=arial size=1>(Int only)</font></td> 
		<%
		END IF'-------------------
		else
		%><td><font size=1 face="<%=FNT%>"><img src=required.gif> <%=FLDNM %><%'= rs(i).Type %></font><br><INPUT TYPE=TEXT SIZE=<%=size %> style="<%=FNTFAMILY%>" name=<%= name %> value=<%=DV%>><br><font face=arial size=1>(Int only)</font></td> <%
		end if
		ls.Close
		set ls=nothing
    '-----------------------------------------
    elseif rs(i).Type = 5 or rs(i).Type = 6 or rs(i).Type = 131 or rs(i).Type = 4 then'FLOAT
		%><td><font size=1 face="<%=FNT%>"><img src=required.gif> <%=FLDNM %><%'= rs(i).Type %></font><br><INPUT TYPE=TEXT SIZE=<%=size %> style="<%=FNTFAMILY%>" name=<%=name %> value=<%=DV%>><br><font face=arial size=1>(Float only)</font></td> <%
    elseif rs(i).Type = 17 or rs(i).Type = 2 or rs(i).Type = 128 or rs(i).Type = 204 then'INT
		%><td><font size=1 face="<%=FNT%>"><img src=required.gif> <%=FLDNM %><%'= rs(i).Type %></font><br><INPUT TYPE=TEXT SIZE=<%=size %> style="<%=FNTFAMILY%>"  name=<%=name %> value=<%=DV%>><br><font face=arial size=1>(Integer only)</font></td> <%
    elseif rs(i).Type = 135 AND (DID=767 OR DID=1869) then
	    %><td><font size=1 face="<%=FNT%>"><IMG SRC=required.gif> <%=FLDNM%><%'= rs(i).Type %></font><br><INPUT TYPE=TEXT SIZE=<%=size %>    style="<%=FNTFAMILY%>" name=<%=name %> value=<%="'" & value & "'"%>><br><font face=arial size=1>(Text only)</font></td> <%
	elseif rs(i).Type = 135 AND DID<>767 then
		%><td><font size=1 face="<%=FNT%>"><img src=required.gif> <%=FLDNM %><%'= rs(i).Type %></font><br><INPUT TYPE=TEXT SIZE=<%=size %>  style="<%=FNTFAMILY%>"  name=<%=name %> value=<%=Now()%> ><IMG src=popupcalendar.gif alt="Calendar" onclick="fPopCalendar(<%= name %>,<%= name %>); return false"><br><font face=arial size=1>(Date only)</font></td> <%
    elseif rs(i).Type = 72 then
		%><td><font size=1 face="<%=FNT%>"><img src=required.gif> <%=FLDNM %><%'= rs(i).Type %></font><br><INPUT TYPE=HIDDEN SIZE=<%=size %> name=<%= name %> value=0><br><font face=arial size=1>(Autono only)</font></td><%
    elseif rs(i).Type = 203 or rs(i).Type = 201 then
    	%><td><font size=1 face="<%=FNT%>"><IMG SRC=required.gif> <%=FLDNM %><%'= rs(i).Type %></font><br><textarea rows=3 cols=43 style="<%=FNTFAMILY%>" name=<%= name %>><%=value%></textarea><br><font face=arial size=1>(Text only)</font></td><%
    else
		%><td><font size=1 face="<%=FNT%>"><img src=required.gif> <%=FLDNM %><%'= rs(i).Type %></font><br><INPUT TYPE=TEXT SIZE=<%=size %>  style="<%=FNTFAMILY%>" maxlength=<%=size%> name=<%= name %> ><br><font face=arial size=1>(Text only)</font></td><%
    end if
    COLCNT=COLCNT+1
 next
%> 
	</tr>
  </table>
  <BR>
  <INPUT TYPE=HIDDEN NAME="i" VALUE=<%=i%>>
  <INPUT TYPE=HIDDEN NAME="DID" VALUE=<%=DID%>>
  <INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=UID%>>
  <INPUT TYPE=HIDDEN NAME="CID" VALUE=<%=CID%>>
  <INPUT TYPE=HIDDEN NAME="FLD" VALUE=<%=UFLD%>>
  <INPUT TYPE=HIDDEN NAME="DETID" VALUE=<%=DETID%>>
  <INPUT TYPE=SUBMIT VALUE="Insert Document" name=Insert>
  <INPUT TYPE=RESET VALUE="Clear Form" name=Cancel>
   </FORM> 
  <%

End Sub
'-------------------------------------FUNCTION ENDS
%>
<%'--------------------------------------------------MAIN PROGRAM STARTS HERE
set rsDOC = Server.CreateObject("ADODB.Recordset")
rsDOC.Open "Select * from DOCUMENTS WHERE DID=" & DID, Conn 
'--------------------------------------
GenerateHeader rsDoc
'-------------------------------------
SQLPROGRAM=rsDOC("ADDSQL"): IF ISNULL(SQLPROGRAM) THEN SQLPROGRAM=""
DETAILSQL=rsDOC("ADDDETAILS"): IF ISNULL(DETAILSQL) THEN DETAILSQL=""
DDID=rsDOC("DDID"): IF ISNULL(DDID) OR DDID="" THEN DDID=0
'------------------------------GENERATE AUTO NUMBERING LOGIC FOR DOCUMENTS
CSQL=rsDOC("COUNTSQL")
'RESPONSE.WRITE CSQL
CFIELD=rsDOC("COUNTFIELD")
PCODE=rsDOC("PREFIX_CODE")
STARTNO=rsDOC("START_NO")
SCODE=rsDOC("SUFFIX_CODE")
IF SCODE="YEAR" THEN SCODE=YEAR(DATE)
'--------------------
IF CSQL="" OR ISNULL(CSQL) THEN
ELSE
CSQL=rsDOC("COUNTSQL") & CID
set rs = Server.CreateObject("ADODB.Recordset")
rs.Open CSQL, Conn
if not rs.EOF then
DAUTONO=CSTR(PCODE) + "/" + CSTR(rs(0)) + "/" + CSTR(SCODE)
end if
rs.Close
Set rs = Nothing
END IF
'--------------------
set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQLPROGRAM, Conn
GenerateForm rs, "formeditnew.asp?MODE=4", DDID, Conn, CFIELD, DAUTONO
rs.Close
Set rs = Nothing
'---------------------
IF DETAILSQL<>"" THEN
Response.Write( "<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF>" )
set rs = Server.CreateObject("ADODB.Recordset")
rs.Open DETAILSQL, Conn
Response.Write( "<TR>" )
for i = 1 to rs.fields.count - 1
Response.Write ( "<td bgcolor=#D5EAFF bordercolordark=#FFFFFF bordercolorlight=#808080><font face=arial size=1 >" & rs(i).Name & "</font></td>")
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