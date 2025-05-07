<%@ LANGUAGE="VBScript" %>
<HTML>
<HEAD>
<TITLE>IRP Edit Updating (sales@erpweb)</TITLE>
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
Sub GenerateHeader( )
Response.Write( "<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF >" )
Response.Write( "<TR>" )
Response.Write( "<TD valign=middle bgcolor=#D4D4D4 bordercolordark=#FFFFFF bordercolor=#808080 bordercolorlight=#808080>&nbsp;&nbsp;<FONT FACE=ARIAL SIZE=2><img src=60-60-60.gif border=0 alt='SoftRobot Document Server'><B> Edit Document Updated</B></FONT></TD>" )
Response.Write( "</TR>" )
Response.Write( "</TABLE>" )
End Sub
%>
<%
Sub GenerateFooter( )
Response.Write( "<TABLE BORDER=0 WIDTH=100% >" )
Response.Write( "<TR>" )
Response.Write( "<TD BGCOLOR=#D1D2D3><FONT FACE=ARIAL SIZE=1><B>Data Updated</B></FONT></TD>" )
Response.Write( "</TR>" )
Response.Write( "</TABLE>" )
End Sub
%>
<%
'--------------------------------SUBROUTINE STARTS 
Sub Adddata( rs, count, ERRORFLAG, Conn)
%> 
<table WIDTH=100%>
<%
  '------------------------------gothrough each field
  ' build input field for each recordset field
  for i = 0 to count-1
  '------------------------------display field 
    %><tr bgcolor=#E1F2FD><td><font face=arial size=1><img src=required.gif> <%=rs(i).Name%><%'=rs(i).Type%></font></td><td><%
  '-------------------------------create input field name
      name = "fld"+cstr(i)
      value=Request(name)
  '------------------------------null validation    
    if isNull(value) OR isEmpty(value) then value=""
  '-----------------------------
  if i=0 then
  value=rs(0)
  rs(i)=value
  ' avoid primary key
  ELSEIF i=1 then
  value=ID
  ELSEIF rs(i).NAME = "TOTAL" THEN
  ELSEIF rs(i).Type=11 then
  ELSE
  rs(i)=value
  end if
  
  '------------------------------numeric validation
    IF rs(i).Type=5 or rs(i).Type=6 or rs(i).Type=131 or rs(i).Type=4 or rs(i).Type=17 or rs(i).Type=2 or rs(i).Type=128 or rs(i).Type=204 THEN
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
		rs(i)=1
		else
		value=0
		rs(i)=0
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
		%><font face=arial size=1><%=value%></font></td></tr> <%
		end if
		ls.Close
		set ls=nothing
    '-----------------------------------------
    else
		%><font face=arial size=1> <%=value%></font></td></tr><%
    end if
	'-----------------------------display value    
 %><%
  '------------------------------setdata
  
   next
 '-------------------------------check errorflag
  IF ERRORFLAG THEN
  Response.Write "Error: Cannot update data because of error in entering data as mentioned below:(Press your back button for correction)"
  Response.End
  END IF
  rs.Update
  %> 
  </table>
  <BR>
<%
End Sub
'-------------------------------------SUBROUTINE ENDS
%>
<%SIGN=REQUEST("SIGN")%>
<%if SIGN="" Then Response.Write "ERROR:SIGN IS BLANK":Response.End%>
<%UID=REQUEST("UID")%>
<%if UID="" Then Response.Write "ERROR:UID IS NULL":Response.End%>
<%ID=REQUEST("ID")%>
<%if ID="" Then Response.Write "ERROR:ID IS NULL":Response.End%>
<%DID=REQUEST("DID")%>
<%if DID="" Then Response.Write "ERROR:DID IS NULL":Response.End%>
<%count=Request("i") 'HOW MANY FIELDS%>
<%if count="" Then Response.Write "ERROR:count is null":Response.End%>
<%ERRORFLAG=0 'ERROR FLAG DISABLED%>
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
rsDOC.Open "Select DETAILSSQL, DDID from DOCUMENTS WHERE DID=" & DID, Conn 
IF NOT rsDOC.EOF THEN
SQLPROGRAM=TRIM(rsDOC("DETAILSSQL"))
DDID=rsDOC("DDID")
else
Response.Write "Error: setup document table"
Response.End
END IF
rsDOC.Close
'------------------------------
IF SQLPROGRAM <> "" THEN
set rs=Server.CreateObject("ADODB.Recordset")
rs.Open SQLPROGRAM & ID, Conn, 1, 3
'Response.Write SQLPROGRAM & ID
ELSE
Response.Write "PLEASE ENTER DETAILSSQL"
Response.End
END IF
'-------------------------
GenerateHeader
'------------------------------set data after validation  
  IF NOT rs.EOF THEN
  Adddata rs, count, ERRORFLAG, Conn
  ELSE
  Response.Write "ERROR: OPENING EDIT RECORDSET"
  END IF
'------------------------------update&close
 GenerateFooter
'-------------------------
rs.Close
ELSE
Response.Write "Error: Wrong Signature, Please enter correct signature." 
END IF
Conn.Close
Set Conn=nothing
%>
<%
ID=REQUEST("IDD")
%>
<font face=arial size=2>
<a href=FORMEDIT.ASP?DID=<%=DID%>&ID=<%=ID%>&UID=<%=UID%>><b>Goto Edit Page</b></a>
</font>
</body>
</html>