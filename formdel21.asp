<%@ LANGUAGE="VBScript" %>
<HTML>
<HEAD>
<TITLE>IRP Delete (sales@erpweb)</TITLE>
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
Response.Write( "<TD valign=middle bgcolor=#D4D4D4 bordercolordark=#FFFFFF bordercolor=#808080 bordercolorlight=#808080>&nbsp;&nbsp;<FONT FACE=ARIAL SIZE=2><img src=60-60-60.gif border=0 alt='SoftRobot Document Server'><B> Deleting Document</B></FONT></TD>" )
Response.Write( "</TR>" )
Response.Write( "</TABLE>" )
End Sub
%>
<%
Sub GenerateFooter( )
Response.Write( "<TABLE BORDER=0 WIDTH=100% >" )
Response.Write( "<TR>" )
Response.Write( "<TD BGCOLOR=#d1d2d3><FONT FACE=ARIAL SIZE=1><B>Data Deleted</B></FONT></TD>" )
Response.Write( "</TR>" )
Response.Write( "</TABLE>" )
End Sub
%>
<%SIGN=REQUEST("SIGN")%>
<%if SIGN="" Then Response.Write "ERROR:SIGN IS BLANK":Response.End%>
<%UID=REQUEST("UID")%>
<%if UID="" Then Response.Write "ERROR:UID IS NULL":Response.End%>
<%ID=REQUEST("ID")%>
<%if ID="" Then Response.Write "ERROR:ID IS NULL":Response.End%>
<%DID=REQUEST("DID")%>
<%if DID="" Then Response.Write "ERROR:DID IS NULL:Response.End"%>
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
rsDOC.Open "Select DETAILSTABLE from DOCUMENTS WHERE DID=" & DID, Conn 
IF NOT rsDOC.EOF THEN
    D=TRIM(rsDOC("DETAILSTABLE"))
	IF NOT (ISNULL(D) OR ISNULL(M)) THEN SQLPROGRAM="DELETE FROM " & D & " WHERE " & D & "ID="
else
	Response.Write "Error: setup document table"
	Response.End
END IF
rsDOC.Close
'------------------------------

'-------------------------
GenerateHeader
'------------------------------set data after validation  
IF NOT (ISNULL(SQLPROGRAM) OR SQLPROGRAM = "") THEN
ON ERROR RESUME NEXT
Conn.Execute SQLPROGRAM & ID
Response.Write "Deleted entry of details table"
ELSE
Response.Write "PLEASE ENTER DETAILSSQL"
Response.End
END IF
'------------------------------update&close
 GenerateFooter
'-------------------------
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
<a href=FORMEDIT.ASP?DID=<%=DID%>&ID=<%=ID%>&UID=<%=UID%>><b>GOTO Edit Page</b></a>
</font>
</body>
</html>