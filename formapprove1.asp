<HTML>
<HEAD>
<TITLE>ERPWEB - SoftRobot / SoftServe APPROVE Server (sales@ERPWEB.com)</TITLE>
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
'---------------------------------FORM GENERATION STARTS
Sub GenerateForm( rs )
' start form
%>
<table WIDTH=100% >
<%
  ' build input field for each recordset field
  for i = 3 to rs.fields.count - 1
  
      value = rs(i)
      if isNull(value) then value=""
    if i=3 then
    %><tr bgcolor=#D5EAFF><td><font face=arial size=1><IMG SRC=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td><font face=arial size=1><%=value%></font></td></tr> <%
    '-----------------------------------------
    elseif rs(i).Type = 3 then
		Set ls=Server.CreateObject("ADODB.Recordset")
		ls.Open "Select * from LISTBOX where DDID=" & DDID & " AND COLUMNNO=" & i, Conn
		if not ls.EOF then
		LISTNAME=ls("LISTNAME")
		LISTTABLE=ls("LISTTABLE")
		LISTCOLUMN=ls("LISTCOLUMN")
		LISTVALUE=ls("LISTVALUE")
		%><tr bgcolor=#D5EAFF><td><font face=arial size=1><IMG SRC=required.gif> <%= LISTNAME %><%'= rs(i).Type %></font></td>
		<td>
		<%LIST1 LISTTABLE, LISTCOLUMN, LISTVALUE, Conn, name, value%> 
		</td></tr> <%
		else
		%><tr bgcolor=#D5EAFF><td><font face=arial size=1><IMG SRC=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td><font face=arial size=1><%=value%></font></td></tr> <%
		end if
		ls.Close
		set ls=nothing
    '-----------------------------------------
    else
		%><tr bgcolor=#D5EAFF><td><font face=arial size=1><IMG SRC=required.gif> <%= rs(i).Name %><%'= rs(i).Type %></font></td><td><font face=arial size=1><%=value%></font></td></tr> <%
    end if
 next
%> 
</table>
<%END SUB%>
<%
Sub GenerateHeader( )
Response.Write( "<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF >" )
Response.Write( "<TR>" )
Response.Write( "<TD valign=middle bgcolor=#D4D4D4 bordercolordark=#FFFFFF bordercolor=#808080 bordercolorlight=#808080>&nbsp;&nbsp;<FONT FACE=ARIAL SIZE=3><img src=60-60-60.gif border=0 alt='SoftRobot Document Server'><B> Approve / OnHold / Reject Document Updated - Thankyou</B></FONT></TD>" )
Response.Write( "</TR>" )
Response.Write( "</TABLE>" )
End Sub
%>
<%
Sub GenerateFooter( )
Response.Write( "<TABLE BORDER=0 WIDTH=100% >" )
Response.Write( "<TR>" )
Response.Write( "<TD BGCOLOR=#c0c0c0><FONT FACE=ARIAL SIZE=1><I>Approved on: " & DATE & "</I></FONT></TD>" )
Response.Write( "</TR>" )
Response.Write( "</TABLE>" )
End Sub
%>
<%SIGN=REQUEST("SIGN")%>
<%if SIGN="" Then Response.Write "<BR><BR><IMG SRC=error.gif> ERROR:SIGN IS BLANK - Press back button":Response.End%>
<%UID=REQUEST("UID")%>
<%if UID="" Then Response.Write "<BR><BR><IMG SRC=error.gif> ERROR:UID IS NULL - Press back button":Response.End%>
<%ID=REQUEST("ID")%>
<%if ID="" Then Response.Write "<BR><BR><IMG SRC=error.gif> ERROR:ID IS NULL - Press back button":Response.End%>
<%DID=REQUEST("DID")%>
<%if DID="" Then Response.Write "<BR><BR><IMG SRC=error.gif> ERROR:DID IS NULL - Press back button":Response.End%>
<%STATUS=REQUEST("STATUS")%>
<%if STATUS="" Then Response.Write "<BR><BR><IMG SRC=error.gif> ERROR:STATUS IS NULL - Press back button":Response.End%>
<%
set rsPWD = Server.CreateObject("ADODB.Recordset")
rsPWD.Open "Select PASSWORD from USERS WHERE UID=" & UID, Conn 
'Response.WRITE "Select PASSWORD from USERS WHERE UID=" & UID 
IF NOT rsPWD.EOF THEN
	PWD=TRIM(rsPWD("PASSWORD"))
	IF ISNULL(PWD) THEN PWD=""
else
	Response.Write "<BR><BR><IMG SRC=error.gif> Error: setup signature in user table - Press back button"
	Response.End
END IF
rsPWD.Close
%>
<%
IF SIGN=PWD THEN '------------------------------CHECK SIGNATURE SAME AS PWD
set rsDOC = Server.CreateObject("ADODB.Recordset")
rsDOC.Open "Select MASTERSQL, DDID, MASTERTABLE from DOCUMENTS WHERE DID=" & DID, Conn 
IF NOT rsDOC.EOF THEN
	SQLPROGRAM=TRIM(rsDOC("MASTERSQL"))
	IF ISNULL(SQLPROGRAM) THEN SQLPROGRAM=""
	DDID=rsDOC("DDID")
	MTABLE=rsDOC("MASTERTABLE")
else
	Response.Write "<BR><BR><IMG SRC=error.gif> Error: setup document table - Press back button"
	Response.End
END IF
rsDOC.Close '------------------------------set data after validation  
'------------------------------
MTABLEID=MTABLE + "ID"
IF SQLPROGRAM <> "" THEN
set rs=Server.CreateObject("ADODB.Recordset")
rs.Open SQLPROGRAM & ID, Conn, 1, 3
'Response.Write SQLPROGRAM & ID
ELSE
Response.Write "<BR><BR><IMG SRC=error.gif> PLEASE ENTER MASTERSQL - Press back button"
Response.End
END IF
'-------------------------
NUID=REQUEST("NUID")
IF NUID>0 THEN
NSQL="UPDATE " & MTABLE & " SET UID=" & NUID &  " WHERE " & MTABLEID & "=" & ID
'Response.Write NSQL
Conn.Execute NSQL
END IF
%>
<!-----------------------------HEADER STRIP--------------------> 

<!--------------------------------------------------------------->
<%
'-------------------------
GenerateHeader
'-------------------------
  IF NOT rs.EOF THEN
  GenerateForm rs
  if STATUS="APPROVE" THEN
  IF NUID=0 THEN 'CHECK WHETHER USER ACCEPT FINAL APPROVAL?
  rs("APPROVE")=1
  rs("ONHOLD")=0
  rs("REJECT")=0
  rs.Update
  Response.Write "<HR><B>APPROVE RECORD NO " & ID
  END IF '------------------------- FOR TOTAPPR
  elseif STATUS="ONHOLD" then
  rs("APPROVE")=0
  rs("ONHOLD")=1
  rs("REJECT")=0
  rs.Update
  Response.Write "<HR><B>ONHOLD RECORD NO " & ID
  elseif STATUS="REJECT" then
  rs("APPROVE")=0
  rs("ONHOLD")=0
  rs("REJECT")=1
  rs.Update
  Response.Write "<HR><B>REJECT RECORD NO " & ID
  end if
  ELSE
  Response.Write "<BR><BR><IMG SRC=error.gif> ERROR: CONFIRMING RECORDSET - Press back button"
  END IF
'------------------------------update&close
 GenerateFooter
'-------------------------
rs.Close
'-----------------------ADD IN TRACKING TABLE
FUNC=STATUS
REMARKS=TRIM(REQUEST("REMARKS"))
TSQL="INSERT INTO TRACKING (UID, DID, ID, FUNC, REMARKS) VALUES (" & UID & ", " & DID & ", " & ID & ", '" & FUNC & "', '" & REMARKS & "')"
'Response.Write "<hr>" & TSQL
Conn.EXECUTE TSQL
'-----------------------ADD IN TRACKING TABLE
ELSE
Response.Write "<BR><BR><IMG SRC=error.gif> <font face=arial size=3>Error: Wrong Signature, Please enter correct signature - Press back button.</font>" 
END IF
Conn.Close
Set Conn=nothing
%>
<hr>
<font face=arial size=1>Click on refresh button in register folder window</font>
</body>
</html>