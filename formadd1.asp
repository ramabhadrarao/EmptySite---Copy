<HTML>
<HEAD>
<TITLE>ERPWEB FORM ADD (sales@ERPWEB.com)</TITLE>
</HEAD>
<%count=Request("i")%>
<%ERRORFLAG=0%>
<BODY>
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
Sub GenerateHeader( )
Response.Write( "<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF >" )
Response.Write( "<TR>" )
Response.Write( "<TD valign=middle bgcolor=#D4D4D4 bordercolordark=#FFFFFF bordercolor=#808080 bordercolorlight=#808080>&nbsp;&nbsp;<FONT FACE=ARIAL SIZE=2><img src=60-60-60.gif border=0 alt='SoftRobot Document Server'><B> Add New Document</B></FONT></TD>" )
Response.Write( "</TR>" )
'Response.Write( "<TR>" )
'Response.Write( "<TD BGCOLOR=YELLOW><FONT FACE=ARIAL SIZE=1><B>Adding new Data</B></FONT></TD>" )
'Response.Write( "</TR>" )
Response.Write( "</TABLE>" )
End Sub
%>
<%
Sub GenerateFooter( )
Response.Write( "<TABLE BORDER=0 WIDTH=100% >" )
Response.Write( "<TR>" )
Response.Write( "<TD BGCOLOR=#d1d2d3><FONT FACE=ARIAL SIZE=1><B>Data Added</B></FONT></TD>" )
Response.Write( "</TR>" )
'Response.Write( "<TR>" )
'Response.Write( "<TD BGCOLOR=#6699CC><FONT FACE=ARIAL SIZE=1><I>" & DATE & "</I></FONT></TD>" )
'Response.Write( "</TR>" )
Response.Write( "</TABLE>" )
End Sub
%>
<%
'--------------------------------SUBROUTINE STARTS 
Sub Adddata( rs )
%> 
<table WIDTH=100%>
<%
  '------------------------------gothrough each field
  ' build input field for each recordset field
  for i = 4 to count-1
  '------------------------------display field 
    %><tr bgcolor=#E1F2FD><td><font face=arial size=1><%=rs(i).Name%><%'=rs(i).Type%></font></td><td><%
  '-------------------------------create input field name
      name = "fld"+cstr(i)
      value=Request(name)
  '------------------------------null validation    
    if isNull(value) then value=""
    '-------------------------------
   
  '------------------------------numeric validation
    IF rs(i).Type=5 or rs(i).Type=6 or rs(i).Type=131 or rs(i).Type=4 or rs(i).Type=17 or rs(i).Type=2 or rs(i).Type=128 or rs(i).Type=204 or rs(i).Type=3 THEN
		IF not IsNumeric(value) then
		Response.Write "Error: Enter Numeric Data"
		value=0
		ERRORFLAG=1
	    End If
	End if
	'----------------------------date validation 
	IF rs(i).Type=135 THEN
		IF not IsDate(value) then
		Response.Write "Error: Enter DD/MM/YYYY Date"
		value=Date
		ERRORFLAG=1
	    End If
	End if
	'-----------------------------boolean validation   
	IF rs(i).Type=11 THEN
		IF value="on" then
		value=1
		else
		value=0
	    End If
	End if
	'------------------------------LISTBOX
	if rs(i).Type = 3 then
		Set ls=Server.CreateObject("ADODB.Recordset")
		ls.Open "Select * from LISTBOX where DDID=" & DDID & " AND COLUMNNO=" & i, Conn
		if not ls.EOF then
		LISTNAME=ls("LISTNAME")
		LISTTABLE=ls("LISTTABLE")
		LISTCOLUMN=ls("LISTCOLUMN")
		LISTVALUE=ls("LISTVALUE")
		%>
		
		<%LIST1 LISTTABLE, LISTCOLUMN, LISTVALUE, Conn, name, value%> 
		</td></tr> <%
		else
		%><%=value%></td></tr> <%
		end if
		ls.Close
		set ls=nothing
    '-----------------------------------------
    else
		%><font face=arial size=1><%=value%></font></td></tr><%
    end if
	'-----------------------------display value    
 %><%
  '------------------------------setdata
    IF i = 3 or rs(i).name="TOTAL" then 'AVOID PRIMARY KEY
    ELSE
    rs(i)=value
    end if
  next
  '-------------------------------check errorflag
  IF ERRORFLAG THEN
  Response.Write "Error: Cannot add data because of error in entering data as mentioned below:(Press your back button for correction)"
  Response.End
  END IF
  %> 
  </table>
  <BR>
<%
End Sub
'-------------------------------------SUBROUTINE ENDS
%>

<Basefont face=arial size=1>
<%SIGN=REQUEST("SIGN")%>
<%if SIGN="" Then Response.Write "ERROR:SIGN IS BLANK":Response.End%>
<%UID=REQUEST("UID")%>
<%if UID="" Then Response.Write "ERROR:UID IS NULL":Response.End%>
<%DID=REQUEST("DID")%>
<%if DID="" Then Response.Write "ERROR:DID IS NULL:Response.End"%>
<!-----------------------------HEADER STRIP--------------------> 

<!--------------------------------------------------------------->
<%
set rsPWD = Server.CreateObject("ADODB.Recordset")
rsPWD.Open "Select PASSWORD from USERS WHERE UID=" & UID, Conn 
'Response.WRITE "Select PASSWORD from USERS WHERE UID=" & UID 
IF NOT rsPWD.EOF THEN
	PWD=TRIM(rsPWD("PASSWORD"))
	IF ISNULL(PWD) THEN PWD=""
else
	Response.Write "Error: setup signature in user table"
	Response.End
END IF
rsPWD.Close
%>
<%
IF SIGN=PWD THEN
'------------------------------
set rsDOC = Server.CreateObject("ADODB.Recordset")
rsDOC.Open "Select * from DOCUMENTS WHERE DID=" & DID, Conn 
IF NOT rsDOC.EOF THEN
SQLPROGRAM=rsDOC("ADDSQL")
DDID=rsDOC("DDID")
else
Response.Write "Error: setup document table"
Response.End
end if
rsDOC.Close
'------------------------------
set rs=Server.CreateObject("ADODB.Recordset")
rs.Open SQLPROGRAM, Conn, 1, 3
'-------------------------
GenerateHeader
'------------------------------set add database
  rs.AddNew
'------------------------------set data after validation  
  Adddata rs
'------------------------------update&close
  rs.Update
GenerateFooter
rs.MoveLast
ID=rs(3)
'-------------------------
rs.Close
'-----------------------ADD IN TRACKING TABLE
FUNC="ADD"
REMARKS=TRIM(REQUEST("REMARKS"))
TSQL="INSERT INTO TRACKING (UID, DID, ID, FUNC, REMARKS) VALUES (" & UID & ", " & DID & ", " & ID & ", '" & FUNC & "', '" & REMARKS & "')"
'Response.Write "<hr>" & TSQL
Conn.EXECUTE TSQL
'-----------------------ADD IN TRACKING TABLE
ELSE
Response.Write "Error: Wrong Signature, Please enter correct signature." 
END IF
Conn.Close
Set Conn=nothing
%>
<font face=arial size=2>
<a href=FORMEDIT.ASP?DID=<%=DID%>&ID=<%=ID%>&UID=<%=UID%>><b>Enter Details or Edit Page</b></a>
</font>
</body>
</html>