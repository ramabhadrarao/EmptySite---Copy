<HTML>
<HEAD>
<TITLE>
IRP USERS MANAGEMENT (sales@erpweb)
</TITLE>
</HEAD>
<BODY topmargin=0>
<%
Set Conn=Server.CreateObject("ADODB.Connection")
Conn.Open Session("erp_ConnectionString")
%>
<%
Sub MODULES ( )
	Set rs1=server.createobject("ADODB.Recordset")
	rs1.open "SELECT DISTINCT MODULENAME, SUBMODULENM, DOCNAME, DID, DNAME, DOC_TYPE FROM WORKBENCH ORDER BY MODULENAME", Conn
	%>
	<SELECT name=DID>
	<%WHILE not rs1.eof%>
	<OPTION VALUE=<%=rs1("DID")%>><%=rs1("MODULENAME")%>-<%=rs1("SUBMODULENM")%>-<%=rs1("DOCNAME")%>-<%=rs1("DNAME")%>-<%=rs1("DOC_TYPE")%>
	<%rs1.MoveNext%>
	<%wend%>
	</SELECT>
	<%rs1.Close%>
<%
End Sub
%>
<% 
Sub GenerateTable( rs )

  Response.Write( "<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF>" )

  ' set up column names
  for i = 0 to rs.fields.count - 2
    Response.Write("<TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080><font face=arial size=1>" + rs(i).Name + "</font></TD>")
  next
	Response.Write("<TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080><font face=arial size=1>Select</font></TD>")

  ' write each row
  while not rs.EOF
    Response.Write( "<TR bgcolor=#E1F2FD>" )
    for i = 0 to rs.fields.count - 2
      v = rs(i)
      if isnull(v) then v = ""
      Response.Write( "<TD VALIGN=TOP><font face=arial size=1>" + CStr( v ) + "</font></TD>" )
    next
    Response.Write("<TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080><font face=arial size=1><a href=users.asp?mode=1&workid=" & rs("WORKFLOWID") & ">Delete User Right</a></font></TD>")
    rs.MoveNext
  wend 
  Response.Write ("<FORM METHOD=POST ACTION=users.asp?mode=2>")
  Response.Write ("<TR><TD COLSPAN=4>")
  MODULES 
  Response.Write("</TD><INPUT TYPE=HIDDEN NAME=UID VALUE=" & UID & "><TD>")
  %>
  <SELECT name=DOCTYPEID>
	<OPTION VALUE=1>OPEN
	<OPTION VALUE=2>APPROVE
	<OPTION VALUE=3>ONHOLD
	<OPTION VALUE=4>CLOSED
	<OPTION VALUE=5>REJECT 
  </SELECT>
  <%
  Response.Write("</TD><TD ALIGN=RIGHT><input type=submit name=s value='Add Rights'></form></TD></TR>")
  Response.Write( "</TABLE>" )
End Sub
%>
<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF >
<TR>
<TD valign=middle bgcolor=#D4D4D4 bordercolordark=#FFFFFF bordercolor=#808080 bordercolorlight=#808080>&nbsp;&nbsp;<FONT FACE=ARIAL SIZE=2><img src=60-60-60.gif border=0 alt='SoftRobot Document Server'><B> User Workbench Designing Screen</B></FONT></TD>
</TR>
</TABLE>
<hr>
<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF>
<TR><TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080><FONT FACE=ARIAL SIZE=2> Create / Delete Document Access rights:</FONT></TD></TR>
</TABLE>
<%UID=REQUEST("UID")%>
<%mode=REQUEST("mode")%>
<%
if mode="" then
	set rs = Server.CreateObject("ADODB.Recordset") 
	SQL= "SELECT DISTINCT WORKFLOWID, MODULENAME, SUBMODULENM, DOCNAME, DNAME, DOC_TYPE, UNAME FROM WORKBENCH WHERE UID=" & UID
	rs.Open SQL, Conn
	if not rs.EOF then
	if isnull(rs("UNAME")) and rs("UNAME")="" then
	else
	Response.Write "<font face=arial size=1>User access rights for USER NAME: " & rs("UNAME") & "</font>"
	end if
	GenerateTable rs
	end if
	rs.Close
elseif mode=1 then
Response.Write("Delete")
WORKID=REQUEST("WORKID")
DSQL="DELETE FROM WORKFLOW WHERE WORKID=" & WORKID
ON ERROR RESUME NEXT
Conn.Execute DSQL
elseif mode=2 then
Response.Write("Add New")
DID=REQUEST("DID")
UID=REQUEST("UID")
DOCTYPEID=REQUEST("DOCTYPEID")
IF DOCTYPEID=1 THEN 'CREATE
FADD=1: FDEL=1: FVIEW=1: FEDIT=1: FFILTER=0: FOFFHOLD=0: FREJECT=0
ELSEIF DOCTYPEID=2 THEN 'APPROVE
FADD=0: FDEL=0: FVIEW=0: FEDIT=0: FFILTER=1: FOFFHOLD=0: FREJECT=0
ELSEIF DOCTYPEID=3 THEN 'ONHOLD
FADD=0: FDEL=0: FVIEW=0: FEDIT=0: FFILTER=0: FOFFHOLD=1: FREJECT=0
ELSEIF DOCTYPEID=4 THEN 'PRINT
FADD=1: FDEL=0: FVIEW=1: FEDIT=0: FFILTER=0: FOFFHOLD=0: FREJECT=0
ELSEIF DOCTYPEID=5 THEN 'REJECT
FADD=0: FDEL=0: FVIEW=0: FEDIT=0: FFILTER=0: FOFFHOLD=0: FREJECT=1
END IF
INSSQL="INSERT INTO WORKFLOW (APPROVE, ONHOLD, REJECT, UID, DID, FADD, FDEL, FVIEW, FEDIT, FILTER, FOFFHOLD, FREJECT ) VALUES (0,0,0," & UID & ", " & DID & ", " & FADD & ", " &  FDEL & ", " &  FVIEW & ", " &  FEDIT & ", " &  FFILTER & ", " & FOFFHOLD & ", " &  FREJECT & " )"
ON ERROR RESUME NEXT
Conn.Execute INSSQL
end if	
%>
</BODY>
</HTML>