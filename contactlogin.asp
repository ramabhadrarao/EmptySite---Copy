<%@ Language=VBScript %>
<HTML>
<HEAD>
<TITLE>IRP Contact Login (sales@erpweb)</TITLE>
</HEAD>
<BODY>

<%
Set Conn=Server.CreateObject("ADODB.Connection")
Conn.Open Session("erp_ConnectionString")
%>
<%
Sub LIST1 ( Conn )
Set lst=Server.CreateObject("ADODB.Recordset")
lst.Open "Select USERID, UNAME FROM USERS", Conn
%>
	<SELECT   name=EID>
	<%WHILE not lst.eof%>
	<OPTION VALUE=<%=lst("EMPLOYEEID")%>><%=lst("EMPLOYEENAME")%>
	<%lst.MoveNext%>
	<%wend%>
	</SELECT>
	<%lst.Close%>
<%
End Sub
%>

<% 
Sub GenerateTable( rs )
  Response.Write( "<TABLE BORDER=0 WIDTH=100% >" )
  ' set up column names
  for i = 0 to rs.fields.count - 3
    Response.Write("<TD bgcolor=yellow><font face=arial size=1>" + rs(i).Name + "</font></TD>")
  next
  ' write each row
  Response.Write("<TD bgcolor=yellow><font face=arial size=1>PHOTO</font></TD>")
   
  while not rs.EOF
    Response.Write( "<TR>" )
     for i = 0 to rs.fields.count - 3
      v = rs(i)
      if isnull(v) then v = ""
      Response.Write( "<TD VALIGN=TOP><font face=arial size=1>" + CStr( v ) + "</font></TD>" )
    next
    PHOTO=rs("IMGPHOTO")
    IF PHOTO="" THEN PHOTO="photo/blank.jpg"
  Response.Write( "<TD VALIGN=TOP><font face=arial size=1><IMG SRC='" & PHOTO & "' WIDTH=80></font></TD>" )
  rs.MoveNext
   wend 
    
  Response.Write( "</TABLE>" )
End Sub
%>
<% 
Sub GenerateTable1( rs )
  Response.Write( "<TABLE BORDER=0 WIDTH=100% >" )
  ' set up column names
  for i = 0 to rs.fields.count - 1
    Response.Write("<TD bgcolor=yellow><font face=arial size=1>" + rs(i).Name + "</font></TD>")
  next
  Response.Write("<TD bgcolor=yellow><font face=arial size=1>PRINT</font></TD>")
 ' write each row
 
  while not rs.EOF
    Response.Write( "<TR>" )
     for i = 0 to rs.fields.count - 1
      v = rs(i)
      if isnull(v) then v = ""
      Response.Write( "<TD VALIGN=TOP><font face=arial size=1>" + CStr( v ) + "</font></TD>" )
    next
    Response.Write( "<TD VALIGN=TOP><font face=arial size=1><a href=printpay.asp?MODE=2&ID=" & rs(0) & " target=new>Print</a></font></TD>" )
    rs.MoveNext
   wend 
  Response.Write( "</TABLE>" )
End Sub
%>
<% 
Sub GenerateTable2( rs )
  Response.Write( "<TABLE BORDER=0 WIDTH=100% >" )
  ' set up column names
  for i = 0 to rs.fields.count - 1
    Response.Write("<TD bgcolor=yellow><font face=arial size=1>" + rs(i).Name + "</font></TD>")
  next
  ' write each row
 
  while not rs.EOF
    Response.Write( "<TR>" )
     for i = 0 to rs.fields.count - 1
      v = rs(i)
      if isnull(v) then v = ""
      Response.Write( "<TD VALIGN=TOP><font face=arial size=1>" + CStr( v ) + "</font></TD>" )
    next
    rs.MoveNext
   wend 
  Response.Write( "</TABLE>" )
End Sub
%>
<TABLE WIDTH=100% >
<TR bgcolor=BLUE><TD><FONT FACE=ARIAL SIZE=2 COLOR=WHITE>Payslip </FONT></TD></TR>
</TABLE>
<%
MODE=REQUEST("MODE")
IF MODE="" THEN '----------------------------------------
DID=REQUEST("DID")
UID=REQUEST("UID")
CID=REQUEST("CALENDERID")
%>
<form action=printpay.asp?MODE=1 Method=post>
<font face=arial size=1>
Employee:<%List1 Conn%><br>
Password:<input type=password name=PWD><br>
<input type=submit name=s value="Log into my Personal Payslip Record">
</font>
</form>
<%
ELSEIF MODE=1 THEN '----------------------------------
EID=REQUEST("EID")
PWD=REQUEST("PWD")
SQL="SELECT EMPLOYEEID, EMPLOYEENAME, PASSWORD, IMGPHOTO FROM EMPLOYEE WHERE EMPLOYEEID=" & EID
Set rs=Server.CreateObject("ADODB.Recordset")
rs.Open SQL , Conn
IF NOT rs.eof then
PASSWORD=rs("PASSWORD")
IF TRIM(PWD) <> TRIM(PASSWORD) THEN Response.Write "Error: Wrong Password! Please try again" : Response.End
else
Response.Write "Error: Employee Record not present"
Response.End
End if
GenerateTable rs
rs.Close
%>
<HR>
<font face=arial size=2><b>My Personal Payslip Record are as follows:</b></font>
<%
SQL="SELECT PAYSLIPID, PAYSLIPDATE, PERIODFROM, PERIODTO, DAYSPRESENT, DAYSABSENT, PAYAMOUNT FROM PAYSLIP WHERE DONEFLAG=1 AND EMPLOYEEID=" & EID
Set rs=Server.CreateObject("ADODB.Recordset")
rs.Open SQL , Conn
GenerateTable1 rs
rs.Close
%>
<HR>
<%ELSEIF MODE=2 THEN%>
<!-- #INCLUDE FILE=letterhead.inc ---------->
<%
ID=REQUEST("ID")
SQL="SELECT * FROM PRINTPAY WHERE PAYSLIPID=" & ID
Set rs=Server.CreateObject("ADODB.Recordset")
rs.Open SQL , Conn
AMT=rs("PAYAMOUNT")
GenerateTable2 rs
rs.Close
SQL="SELECT * FROM PRINTPAY1 WHERE PAYSLIPID=" & ID
Set rs=Server.CreateObject("ADODB.Recordset")
rs.Open SQL , Conn
GenerateTable2 rs
rs.Close
%>
<TABLE WIDTH=100% >
<TR><TD VALIGN=TOP>
<%
SQL="SELECT PAYHEAD, EARNINGS FROM PAYSLIPDET WHERE EARNINGS>0 AND PAYSLIPID=" & ID
Set rs=Server.CreateObject("ADODB.Recordset")
rs.Open SQL , Conn
GenerateTable2 rs
rs.Close
%>
</TD>
<TD VALIGN=TOP>
<%
SQL="SELECT PAYHEAD, DEDUCTIONS FROM PAYSLIPDET WHERE DEDUCTIONS>0 AND PAYSLIPID=" & ID
Set rs=Server.CreateObject("ADODB.Recordset")
rs.Open SQL , Conn
GenerateTable2 rs
rs.Close
%>
</TD></TR>
<%
SQL="SELECT SUM(EARNINGS) FROM PAYSLIPDET WHERE EARNINGS>0 AND PAYSLIPID=" & ID
Set rs=Server.CreateObject("ADODB.Recordset")
rs.Open SQL , Conn
E=rs(0)
rs.Close
%>
<%
SQL="SELECT SUM(DEDUCTIONS) FROM PAYSLIPDET WHERE DEDUCTIONS>0 AND PAYSLIPID=" & ID
Set rs=Server.CreateObject("ADODB.Recordset")
rs.Open SQL , Conn
D=rs(0)
rs.Close
%>
<tr bgcolor=yellow><td><font face=arial size=2>Total Earnings: <%=E%></font></td><td align=right><font face=arial size=2>Total Deduction: <%=D%></font></TD></TR>
<tr bgcolor=yellow><td></td><td align=right><font face=arial size=2>Total Amount Payable: <%=AMT%></font></TD></TR>
</TABLE>

<hr>
<!-- #INCLUDE FILE=footerhead.inc ---------->
<%
END IF '--------------------------------------------
Conn.Close
Set Conn=nothing
%>
</BODY>
</HTML>
