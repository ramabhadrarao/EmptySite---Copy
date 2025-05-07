<%@ Language=VBScript %>
<HTML>
<HEAD>
<TITLE>IRP WIZARDS (sales@erpweb)</TITLE>
</HEAD>
<BODY topmargin=10  leftmargin=1>
<%
Set Conn=Server.CreateObject("ADODB.Connection")
Conn.Open Session("erp_ConnectionString")
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
	ID=rs("WizardsID")
	Response.Write( "<TD VALIGN=TOP><font face=arial size=1><a href=wizards.asp?MODE=1&N=0&ID=" & ID & " TARGET=startmenu><img src=wizard.gif></A></font></TD>" )
    rs.MoveNext
   wend 
     Response.Write( "</TABLE>" )
End Sub
%>

<% 
Sub GenerateTable1( rs )
  Response.Write( "<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF>" )
  ' set up column names
  for i = 2 to rs.fields.count - 3
    Response.Write("<TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080><font face=arial size=1>" + rs(i).Name + "</font></TD>")
  next
  Response.Write("<TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080 COLSPAN=2><font face=arial size=1 >ACTION</font></TD>")

  ' write each row
  
  IF not rs.EOF THEN
    Response.Write( "<TR bgcolor=#E1F2FD>" )
     for i = 2 to rs.fields.count - 3
      v = rs(i)
      if isnull(v) then v = ""
      Response.Write( "<TD VALIGN=TOP><font face=arial size=1>" + CStr( v ) + "</font></TD>" )
	  next
	  N=rs("NEXTID")
	  P=rs("PREVID")
END IF
    IF P>0 THEN
	Response.Write( "<TD VALIGN=TOP width=10><font face=arial size=1><a href=wizards.asp?MODE=2&P=" & P & "&ID=" & ID & " TARGET=startmenu><img src=gryleft.gif border=0 width=20></A></font></TD>" )
   ELSE
   Response.Write( "<TD VALIGN=TOP width=10></TD>")
   END IF 
   IF N>0 THEN
	Response.Write( "<TD VALIGN=TOP width=10><font face=arial size=1><a href=wizards.asp?MODE=1&N=" & N & "&ID=" & ID & " TARGET=startmenu><img src=gryright.gif border=0 width=20></A></font></TD>" )
   ELSE
   Response.Write( "<TD VALIGN=TOP width=10></TD>")
   END IF

End Sub
%>


<%
DID=REQUEST("DID")
UID=REQUEST("UID")
CID=REQUEST("CID")
%>
<%
MODE=REQUEST("MODE")
IF MODE="" THEN
%>
<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF >
<TR>
<TD valign=middle bgcolor=#D4D4D4 bordercolordark=#FFFFFF bordercolor=#808080 bordercolorlight=#808080>&nbsp;&nbsp;<FONT FACE=ARIAL SIZE=2><img src=ofolder.gif border=0 alt='SoftRobot Document Server'><B> IRP Wizards</B></FONT></TD>
</TR>
</TABLE>
<%
SQL="SELECT WizardsID, WizardName, CreatedBy FROM wizards"
Set rs=Server.CreateObject("ADODB.Recordset")
rs.Open SQL , Conn
GenerateTable rs
rs.Close
%>
<HR>
<%ELSEIF MODE=4 THEN%>
<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF >
<TR>
<TD valign=middle bgcolor=#D4D4D4 bordercolordark=#FFFFFF bordercolor=#808080 bordercolorlight=#808080>&nbsp;&nbsp;<FONT FACE=ARIAL SIZE=2><img src=ofolder.gif border=0 alt='SoftRobot Document Server'><B> IRP Wizards</B></FONT></TD>
</TR>
</TABLE>
<%
SQL="SELECT * FROM wizardusers where UID=" & UID
Set rs=Server.CreateObject("ADODB.Recordset")
rs.Open SQL , Conn
GenerateTable rs
rs.Close
%>
<hr>
<%ELSEIF MODE=1 THEN%>
<%
ID=REQUEST("ID")
N=REQUEST("N")
IF N=0 THEN
SQL="SELECT * FROM WizardsDET where WizardsID=" & ID
ELSE
SQL="SELECT * FROM WizardsDET where WizardsID=" & ID & " AND WizardsDETID=" & N
END IF
Set rs=Server.CreateObject("ADODB.Recordset")
rs.Open SQL , Conn
GenerateTable1 rs
rs.Close
%>
<%ELSEIF MODE=2 THEN%>
<%
ID=REQUEST("ID")
P=REQUEST("P")
SQL="SELECT * FROM WizardsDET where WizardsID=" & ID & " AND WizardsDETID=" & P
Set rs=Server.CreateObject("ADODB.Recordset")
rs.Open SQL , Conn
GenerateTable1 rs
rs.Close
%>
<%
END IF
Conn.Close
Set Conn=nothing
%>
</BODY>
</HTML>
