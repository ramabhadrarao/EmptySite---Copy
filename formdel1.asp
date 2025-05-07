<!--
'************************************************************************
'Pupose						:	This is a SoftServer delete update server
'Filename					:	formdel1.asp
'Author						:	Anita Shah
'Created					:	27-Mar-2001
'Project Name				:	ERPWEB
'Contact					:	sales@ERPWEB.com
'
'Modification History		:	
'Purpose					:
'Version					:
'Author 					:
'Created					:
'************************************************************************
-->
<HTML>
<HEAD>
<TITLE>SoftServer FORM DELETE update(sales@ERPWEB.com)</TITLE>
</HEAD>
<BODY>
<Basefont face=arial size=1>
<%
Set Conn=Server.CreateObject("ADODB.Connection")
Conn.Open Session("erp_ConnectionString")
%>
<%
Sub GenerateHeader( )
Response.Write( "<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF >" )
Response.Write( "<TR>" )
Response.Write( "<TD valign=middle bgcolor=#D4D4D4 bordercolordark=#FFFFFF bordercolor=#808080 bordercolorlight=#808080>&nbsp;&nbsp;<FONT FACE=ARIAL SIZE=3><img src=60-60-60.gif border=0 alt='SoftRobot Document Server'><B> Deleted Document - Press Refresh button in register</B></FONT></TD>" )
Response.Write( "</TR>" )
Response.Write( "</TABLE>" )
End Sub
%>
<%
Sub GenerateFooter( )
Response.Write( "<TABLE BORDER=0 WIDTH=100% >" )
Response.Write( "<TR>" )
Response.Write( "<TD BGCOLOR=#c0c0c0><FONT FACE=ARIAL SIZE=1><I>" & DATE & "</I></FONT></TD>" )
Response.Write( "</TR>" )
Response.Write( "</TABLE>" )
End Sub
%>
<%GenerateHeader%>
<%SIGN=REQUEST("SIGN")%>
<%if SIGN="" Then Response.Write "<br><br><img src=error.gif> ERROR:SIGN IS BLANK - Press back button":Response.End%>
<%UID=REQUEST("UID")%>
<%if UID="" Then Response.Write "<br><br><img src=error.gif> ERROR:UID IS NULL - User access denied - Press back button":Response.End%>
<%ID=REQUEST("ID")%>
<%if ID="" Then Response.Write "<br><br><img src=error.gif> ERROR:ID IS NULL - Record access denied - Press back button":Response.End%>
<%DID=REQUEST("DID")%>
<%if DID="" Then Response.Write "<br><br><img src=error.gif> ERROR:DID IS NULL - Document access denied - Press back button:Response.End"%>
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
	Response.Write "<br><br><img src=error.gif> Error: setup signature in user table - Press back button"
	Response.End
END IF
rsPWD.Close
%>
<%
IF SIGN=PWD THEN
'------------------------------
set rsDOC = Server.CreateObject("ADODB.Recordset")
rsDOC.Open "Select DELETESQL, DELDETSQL from DOCUMENTS WHERE DID=" & DID, Conn 
IF NOT rsDOC.EOF THEN
	'M=TRIM(rsDOC("MASTERTABLE"))
	'IF NOT ISNULL(M) THEN SQLPROGRAM="DELETE FROM " & M & " WHERE " & M & "ID="
	'D=TRIM(rsDOC("DETAILSTABLE"))
	'IF NOT ISNULL(D) THEN DETAILSSQL="DELETE FROM " & D & " WHERE " & M & "ID="
	SQLPROGRAM=rsDOC("DELETESQL")
	DETAILSSQL=rsDOC("DELDETSQL")
ELSE
	Response.Write "<br><br><img src=error.gif> Error: setup document table"
	Response.End
END IF
rsDOC.Close
'----------------------------
IF NOT (ISNULL(DETAILSSQL) OR DETAILSSQL = "") THEN
ON ERROR RESUME NEXT
Conn.Execute DETAILSSQL & ID
Response.Write "<BR>DELETING DETAILS TABLE ENTRY:<BR>" & DETAILSSQL & ID
END IF
'------------------------------set data after validation  
IF NOT (ISNULL(SQLPROGRAM) OR SQLPROGRAM = "") THEN
ON ERROR RESUME NEXT
Conn.Execute SQLPROGRAM & ID
Response.Write "<BR>DELETING MASTER TABLE ENTRY:<BR>" & SQLPROGRAM & ID
ELSE
Response.Write "<br><br><img src=error.gif> PLEASE ENTER MASTERSQL"
Response.End
END IF
'------------------------------update&close
 GenerateFooter
'-----------------------ADD IN TRACKING TABLE
FUNC="DELETE"
REMARKS=TRIM(REQUEST("REMARKS"))
TSQL="INSERT INTO TRACKING (UID, DID, ID, FUNC, REMARKS) VALUES (" & UID & ", " & DID & ", " & ID & ", '" & FUNC & "', '" & REMARKS & "')"
'Response.Write "<hr>" & TSQL
Conn.EXECUTE TSQL
'-----------------------ADD IN TRACKING TABLE
ELSE
Response.Write "<br><br><img src=error.gif> Error: Wrong Signature, Please enter correct signature.Press back button" 
END IF
Conn.Close
Set Conn=nothing
%>
</body>
</html>