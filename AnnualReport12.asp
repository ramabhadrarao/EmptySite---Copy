<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft FrontPage 4.0">
<TITLE>I-Business ANNUAL REPORT</TITLE>
</HEAD>
<BODY>
<%
Set Conn=Server.CreateObject("ADODB.Connection")
Conn.Open Session("erp_ConnectionString")
Set rs=Server.CreateObject("ADODB.RecordSet")
Set rs1=Server.CreateObject("ADODB.RecordSet")
Set rs2=Server.CreateObject("ADODB.RecordSet")
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
<TABLE BORDER=1 width=650 bordercolordark=#FFFFFF bordercolorlight=#FFFFFF >
<TR>
<TD valign=middle bgcolor=#D4D4D4 bordercolordark=#FFFFFF bordercolor=#808080 bordercolorlight=#808080>&nbsp;&nbsp;<FONT FACE=ARIAL SIZE=2><img src=OFOLDER.GIF border=0 alt='SoftRobot Document Server'><B>Annual Report</B></FONT></TD>
</TR>
</TABLE>
<%
SQL="SELECT * FROM FINANCIALREPORT"
rs.Open SQL,Conn
%>
<FORM NAME="FRMANNUALREPORT" METHOD="POST" ACTION="ANNUALREPORT12.ASP?MODE=1" TARGET=NEWS>
<font face=arial Size=2><B>Select AnnaulReport:</B>
<Select name=FRID>
<% While Not rs.Eof %>
<option value=<%=rs("FinancialReportID")%>><%=rs("ReportName")%></Option>
<% rs.MoveNext %>
<% Wend %>
<% rs.Close %>
<INPUT TYPE="HIDDEN" VALUE=<%=CID%> NAME="CID">
<INPUT TYPE="HIDDEN" VALUE=<%=DID%> NAME="DID">
<INPUT TYPE="HIDDEN" VALUE=<%=UID%> NAME="UID">
<INPUT TYPE="SUBMIT" VALUE="ShowAnnualReport" name=cmdannualReport>
</FORM>
<%
ELSE   '----------------MODE=1
%>
<TABLE BORDER=1 width=650 bordercolordark=#FFFFFF bordercolorlight=#FFFFFF >
<TR>
<TD valign=middle bgcolor=#D4D4D4 bordercolordark=#FFFFFF bordercolor=#808080 bordercolorlight=#808080>&nbsp;&nbsp;<FONT FACE=ARIAL SIZE=2><img src=60-60-60.GIF border=0 alt='SoftRobot Document Server'><B>Annual Report</B></FONT></TD>
</TR>
</TABLE>
<%
FRID=REQUEST("FRID")
CID=REQUEST("CID")
DID=REQUEST("DID")
UID=REQUEST("UID")
IF FRID=2 THEN
Response.redirect "ANNUALREPORT123.ASP?MODE=1&DID=" & DID & "&CID=" & CID & "&UID=" & UID & "&FRID=" & FRID
END IF

SQL="SELECT COMPANYNAME FROM COMPANY"
Set rs=Server.CreateObject("ADODB.RecordSet")
rs.Open SQL,Conn
IF Not rs.EOF then
Companyname=rs("Companyname")
rs.Close
End If
SQL="SELECT FYSTART,FYEND,THISYEAR FROM CALENDER WHERE CALENDERID=" & CID
rs.Open SQL,CONN
IF Not rs.EOF then
IF ISNULL(rs("FYSTART")) THEN FYSTART="" ELSE FYSTART=rs("FYSTART")
IF ISNULL(rs("FYEND")) THEN FYEND="" ELSE FYEND=rs("FYEND")
IF ISNULL(rs("THISYEAR")) THEN THISYEAR="" ELSE THISYEAR=rs("THISYEAR")
End IF
rs.close
PCID=CID-1
SQL="SELECT FYSTART,FYEND,THISYEAR FROM CALENDER WHERE CALENDERID=" & PCID
rs.Open SQL,CONN
IF Not rs.EOF then
IF ISNULL(rs("FYSTART")) THEN PREVFYSTART="" ELSE PREVFYSTART=rs("FYSTART")
IF ISNULL(rs("FYEND")) THEN PREVFYEND="" ELSE PREVFYEND=rs("FYEND")
IF ISNULL(rs("THISYEAR")) THEN PREVTHISYEAR="" ELSE PREVTHISYEAR=rs("THISYEAR")
End IF
rs.close
SQL="SELECT REPORTNAME FROM FINANCIALREPORT WHERE FINANCIALREPORTID=" & FRID
rs.open SQL,Conn
IF Not rs.Eof Then
ReportName=rs("ReportName")
End IF
rs.Close

'-----------------------FIND PROFIT OR LOSS FOR THE CUREENT YEAR-----------------------

Set rs=Server.CreateObject("ADODB.Recordset")
SQL="SELECT * FROM LEDGERLIST WHERE HEADNAME='Expense'"
rs.Open SQL , Conn
DRTOT=0
CRTOT=0
PREVGROUP=""
while not rs.EOF
AID=rs("ACCGROUPDETID")
SQL1="SELECT SUM(DEBIT), SUM(CREDIT) FROM LEDGERSCH WHERE ACCGROUPDETID=" & AID & " AND CALENDERID=" & CID
Set rs1=Server.CreateObject("ADODB.Recordset")
rs1.Open SQL1 , Conn
IF NOT rs1.EOF THEN'------------------1
'---------------------
OPDEBIT=rs1(0): IF ISNULL(OPDEBIT) OR OPDEBIT="" THEN OPDEBIT=0
OPCREDIT=rs1(1): IF ISNULL(OPCREDIT) OR OPCREDIT="" THEN OPCREDIT=0
ELSE
OPDEBIT=0
OPCREDIT=0
END IF
rs1.Close
'---------------------
SQL1="SELECT SUM(DEBIT), SUM(CREDIT) FROM LEDGERSCHF WHERE ACCGROUPDETID=" & AID & " AND CALENDERID=" & CID
Set rs1=Server.CreateObject("ADODB.Recordset")
rs1.Open SQL1 , Conn
IF NOT rs1.EOF THEN'------------------1
'---------------------
DEBIT=rs1(0): IF ISNULL(DEBIT) OR DEBIT="" THEN DEBIT=0
CREDIT=rs1(1): IF ISNULL(CREDIT) OR CREDIT="" THEN CREDIT=0
'-----------------
DEBIT=OPDEBIT+DEBIT
CREDIT=OPCREDIT+CREDIT
'-----------------
IF DEBIT>CREDIT THEN
DEBIT=DEBIT-CREDIT
CREDIT=0
ELSE
CREDIT=CREDIT-DEBIT
DEBIT=0
END IF
ELSE '------------------1
DEBIT=0
CREDIT=0
END IF'---------------------1
rs1.Close
DRTOT=DRTOT+DEBIT
CRTOT=CRTOT+CREDIT
'--------------
rs.MoveNext
wend
EXPENSE=DRTOT-CRTOT
'Response.Write("<FONT FACE=ARIAL SIZE=1><b>Total Expenses=" & formatnumber(EXPENSE,2) & "</Font>")
rs.Close

'----------------------Calculating INCOME

SQL="SELECT * FROM LEDGERLIST WHERE HEADNAME='Income'"
rs.Open SQL , Conn
DRTOT=0
CRTOT=0
PREVGROUP=""
while not rs.EOF
AID=rs("ACCGROUPDETID")
SQL1="SELECT SUM(DEBIT), SUM(CREDIT) FROM LEDGERSCH WHERE ACCGROUPDETID=" & AID & " AND CALENDERID=" & CID
rs1.Open SQL1 , Conn
IF NOT rs1.EOF THEN'------------------1
'---------------------
OPDEBIT=rs1(0): IF ISNULL(OPDEBIT) OR OPDEBIT="" THEN OPDEBIT=0
OPCREDIT=rs1(1): IF ISNULL(OPCREDIT) OR OPCREDIT="" THEN OPCREDIT=0
ELSE
OPDEBIT=0
OPCREDIT=0
END IF
rs1.Close
'---------------------
SQL1="SELECT SUM(DEBIT), SUM(CREDIT) FROM LEDGERSCHF WHERE ACCGROUPDETID=" & AID & " AND CALENDERID=" & CID
Set rs1=Server.CreateObject("ADODB.Recordset")
rs1.Open SQL1 , Conn
IF NOT rs1.EOF THEN'------------------1
'---------------------
DEBIT=rs1(0): IF ISNULL(DEBIT) OR DEBIT="" THEN DEBIT=0
CREDIT=rs1(1): IF ISNULL(CREDIT) OR CREDIT="" THEN CREDIT=0
'-----------------
DEBIT=OPDEBIT+DEBIT
CREDIT=OPCREDIT+CREDIT
'-----------------
IF DEBIT>CREDIT THEN
DEBIT=DEBIT-CREDIT
CREDIT=0
ELSE
CREDIT=CREDIT-DEBIT
DEBIT=0
END IF
ELSE '------------------1
DEBIT=0
CREDIT=0
END IF'---------------------1
rs1.Close
DRTOT=DRTOT+DEBIT
CRTOT=CRTOT+CREDIT
'--------------
rs.MoveNext
wend
rs.Close
INCOME=CRTOT-DRTOT
IF INCOME > EXPENSE THEN
PROFIT=INCOME-EXPENSE
LOSS=0
ELSE
LOSS=EXPENSE-INCOME
PROFIT=0
END IF

'=======================ACTUAL CODE STARTS==========================================

SQL="SELECT * FROM ANNUALREPORT WHERE FINANCIALREPORTID=" & FRID
rs.Open SQL,Conn
IF Not rs.Eof Then
ReportID=rs("AnnualReportId")
rs.Close
End If

'---------------TITLE PRINTING
Response.Write("<Table Border=0 Width=650>")
Response.Write("<TR>")
Response.Write("<TD><Font Face=arial Size=2>(&nbsp;<B><U>" & companyname & "</B></U>&nbsp;)</Font></TD>")
Response.Write("</TR>")
Response.Write("<TR>")
Response.Write("<TD Align=center><Font Face=arial Size=2><B><U>" & ReportName & "</B></U></Font></TD>")
Response.Write("</TR>")
Response.Write("<TR>")
Response.Write("<TD Align=center><Font Face=arial Size=2><B><U>For the Year Ended:" & FYEND & "</B></U></Font></TD>")
Response.Write("</TR>")
Response.Write("<TR>")
Response.Write("<TD Align=center>&nbsp;</TD>")
Response.Write("</TR>")
Response.Write("</Table>")
GroupTot=0.0
CurGroupTot=0.0
OpPLSubGroupTot=0.0
Pound=chr(163)
SQL="SELECT * FROM ANNUALREPORTDET WHERE ANNUALREPORTID=" & REPORTID
rs.Open SQL,Conn
IF Not rs.EOF Then
Response.Write("<TABLE BORDER=0 WIDTH=650>")
Response.Write("<TR>")
Response.Write("<TD colspan=2><Font Face=arial Size=1>&nbsp;</Font></TD>")
Response.Write("<TD colspan=2 align=center><Font Face=arial Size=2><B>" & PREVFYEND & "</B></Font></TD>")
Response.Write("<TD colspan=2 align=center><Font Face=arial Size=2><B>" & FYEND & "</B></Font></TD>")
Response.Write("</TR>")
Response.Write("<TR>")
Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
Response.Write("<TD><Font Face=arial Size=1>Notes</Font></TD>")
Response.Write("<TD align=right><Font Face=arial Size=1>" & Pound & "</Font></TD>")
Response.Write("<TD align=right><Font Face=arial Size=1>" & Pound & "</Font></TD>")
Response.Write("<TD align=right><Font Face=arial Size=1>" & Pound & "</Font></TD>")
Response.Write("<TD align=right><Font Face=arial Size=1>" & Pound & "</Font></TD>")
Response.Write("</TR>")
While Not rs.EOF
I=1
GROUPID=rs("ACCGROUPID")
IF GROUPID > 0 THEN        '----------GroupID > 0 then
Response.Write("<TR>")
Response.Write("<TD><Font Face=arial Size=2><B>" & rs("Alias") & "</B></Font></TD>")
Response.Write("<TD COLSPAN=5><Font Face=arial Size=1>&nbsp;</Font></TD>")
Response.Write("</TR>")
Set rs1=Server.CreateObject("ADODB.Recordset")
SQL1="SELECT * FROM ACCGROUPDET WHERE ACCGROUPID=" & GROUPID
rs1.Open SQL1,Conn
SubGroupTot=0.0
CurYrSubGroupTot=0.0
IF Not rs1.EOF THEN  '----AccGroupDet If Statement 
While Not rs1.EOF  '----------GroupDet Process 
GDETID=rs1("ACCGROUPDETID")
NAME=rs1("SUBGROUPNM")
'----------------------FIND OPENING BALANCE---------------------------- 

SQL2="SELECT SUM(DEBIT), SUM(CREDIT) FROM LEDGERSCH WHERE ACCGROUPDETID=" & GDETID & " AND CALENDERID=" & CID
rs2.Open SQL2,Conn
	If Not rs2.EOF THEN
	OPDEBIT=rs2(0): IF ISNULL(OPDEBIT) OR OPDEBIT="" THEN OPDEBIT=0
	OPCREDIT=rs2(1): IF ISNULL(OPCREDIT) OR OPCREDIT="" THEN OPCREDIT=0
	ELSE
	OPDEBIT=0
	OPCREDIT=0
	END IF
rs2.Close
'-------------------END----------------------------------------------
'-------------------FIND TRANSACTIONS DURING YEAR--------------------
SQL2="SELECT SUM(DEBIT), SUM(CREDIT) FROM LEDGERSCHF WHERE ACCGROUPDETID=" & GDETID & " AND CALENDERID=" & CID
rs2.Open SQL2 , Conn
	IF NOT rs2.EOF THEN
	DEBIT=rs2(0): IF ISNULL(DEBIT) OR DEBIT="" THEN DEBIT=0
	CREDIT=rs2(1): IF ISNULL(CREDIT) OR CREDIT="" THEN CREDIT=0
	DEBIT=OPDEBIT+DEBIT
	CREDIT=OPCREDIT+CREDIT
		IF DEBIT>CREDIT THEN
		DEBIT=DEBIT-CREDIT
		CREDIT=0
		ELSE
		CREDIT=CREDIT-DEBIT
		DEBIT=0
		END IF
	ELSE 
		DEBIT=0
		CREDIT=0
	END IF
	rs2.Close
	Response.Write("<TR>")
	Response.Write("<TD><Font Face=arial Size=1><B>&nbsp;" & rs1("NAMEASPERCLAW") & "</B></Font></TD>")
	Response.Write("<TD><Font Face=arial Size=1><A href=statementagroup.asp?MODE=2&CID=" & CID & "&AID=" & GDETID & ">" & rs1("SCHEDULE") & "</A></Font></TD>")
	IF OPDEBIT=0 THEN
	OpVal=OPCREDIT
	ELSE
	OpVal=OPDEBIT
	END IF
	IF DEBIT=0 THEN
	CloseVal=CREDIT
	ELSE
	CloseVal=DEBIT
	END IF
	IF STRCOMP(NAME,"Profit & Loss A/C.",1)=0 THEN
	IF LOSS=0 THEN
	CloseVal=CloseVal + PROFIT
	ELSE
	CloseVal=CloseVal - LOSS
	END IF
	END IF
	Response.Write("<TD Align=right><Font Face=arial Size=1>" & FormatNumber(OpVal) & "</Font></TD>")
	Response.Write("<TD Align=right><Font Face=arial Size=1>&nbsp;</Font></TD>")
	Response.Write("<TD Align=right><Font Face=arial Size=1>" & FormatNumber(Closeval,2) & "</Font></TD>")
	Response.Write("<TD Align=right><Font Face=arial Size=1>&nbsp;</Font></TD>")
	Response.Write("</TR>")
	DRTOT=DRTOT+DEBIT
	CRTOT=CRTOT+CREDIT
	SubGroupTot=SubGroupTot + opval
	CurYrSubGroupTot=CurYrSubGroupTot + Closeval 
	rs1.MoveNext
	Wend              '----------GroupDet End
	rs1.Close
End IF
IF rs("SubTotal") Then
Response.Write("<TR>")
Response.Write("<TD Colspan=2><Font Face=arial Size=1>&nbsp;</Font></TD>")
Response.Write("<TD align=right><Font Face=arial Size=1>-------------------</Font></TD>")
Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
Response.Write("<TD align=right><Font Face=arial Size=1>-------------------</Font></TD>")
Response.Write("<TD align=right><Font Face=arial Size=1>&nbsp;</Font></TD>")
Response.Write("</TR>")
Response.Write("<TR>")
Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
Response.Write("<TD Colspan=2><Font Face=arial Size=1>&nbsp;</Font></TD>")
Response.Write("<TD align=right><Font Face=arial Size=1>" & FormatNumber(SubGroupTot,2) & "</Font></TD>")
Response.Write("<TD align=right><Font Face=arial Size=1>&nbsp;</Font></TD>")
Response.Write("<TD align=right><Font Face=arial Size=1>" & FormatNumber(CurYrSubGroupTot,2) & "</Font></TD>")
Response.Write("</TR>")
End If
GroupTot=GroupTot + SubGroupTot
CurGroupTot=CurGroupTot + CurYrSubGroupTot
ELSE '--------IF Groupid=0 Then
IF rs("Add") Then
Response.Write("<TR>")
Response.Write("<TD Colspan=3><Font Face=arial Size=1>&nbsp;</Font></TD>")
Response.Write("<TD align=right><Font Face=arial Size=1>-------------------</Font></TD>")
Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
Response.Write("<TD align=right><Font Face=arial Size=1>-------------------</Font></TD>")
Response.Write("<TD align=right><Font Face=arial Size=1>&nbsp;</Font></TD>")
Response.Write("</TR>")
Response.Write("<TR>")
Response.Write("<TD colspan=3><Font Face=arial Size=2><B>&nbsp;" & rs("Alias") & "</B></Font></TD>")
Response.Write("<TD Align=right><Font Face=arial Size=2><B>" & FormatNumber(GroupTot,2) & "</B></Font></TD>")
Response.Write("<TD Align=right><Font Face=arial Size=2>&nbsp;</Font></TD>")
Response.Write("<TD Align=right><Font Face=arial Size=2><B>" & FormatNumber(CurGroupTot,2) & "</B></Font></TD>")
Response.Write("</TR>")
IF STRCOMP(rs("Formula"),"TOTALASSETS",1)=0 THEN
TOTALASSETS=GroupTot
CurTOTALASSETS=CurGroupTot
'Response.Write TOTALASSETS
ELSEIF STRCOMP(rs("Formula"),"TOTALLIABILITIES",1)=0 THEN
TOTALLIABILITIES=GroupTot
CurTOTALLIABILITIES=CurGroupTot
'Response.Write TOTALLIABILITIES
END IF
GroupTot=0.0
CurGroupTot=0.0
ELSE   '================== Groupid=0 and Execute Formula
Formula=rs("Formula")
IF STRCOMP(Formula,"0",1)=0 then    '---------Formula Logic

'------------------------SUBGROUP--------------------------

ADETID=rs("AccGroupDetID")
SQL1="SELECT * FROM ACCGROUPDET WHERE ACCGROUPDETID=" & ADETID
rs1.Open SQL1,Conn
IF NOT rs1.EOF then
SQL2="SELECT * FROM LEDGERANALYSISOPBAL WHERE ACCGROUPDETID=" & ADETID & "AND CALENDERID=" & PCID
'Response.Write SQL2
rs2.Open SQL2,Conn
IF Not rs2.Eof Then
OPDEBIT=rs2("DEBIT"): IF ISNULL(OPDEBIT) OR OPDEBIT="" THEN OPDEBIT=0
OPCREDIT=rs2("CREDIT"): IF ISNULL(OPCREDIT) OR OPCREDIT="" THEN OPCREDIT=0
ELSE
OPDEBIT=0
OPCREDIT=0
END IF
rs2.Close
IF OPDEBIT >= OPCREDIT Then
Opval=OPDEBIT
ELSE
OpVal=OPCREDIT
END IF
Response.Write("<TR>")
IF rs("Bold") Then
Response.Write("<TD><Font Face=arial Size=2><B>&nbsp;" & rs("Alias") & "</B></Font></TD>")
ELSE
Response.Write("<TD><Font Face=arial Size=1>&nbsp;" & rs("Alias") & "</Font></TD>")
End If
Response.Write("<TD><Font Face=arial Size=1><A href=statementagroup.asp?MODE=2&CID=" & CID & "&AID=" & ADETID & ">" & rs1("SCHEDULE") & "</A></Font></TD>")
Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
Response.Write("<TD Align=Right><Font Face=arial Size=1>" & FORMATNUMBER(OpVal,2) & "</Font></TD>")
Response.Write("</TR>")
END IF
IF rs("Sub") Then
OpPLSubGroupTot=OpPLSubGroupTot - OpVal
ELSEIF rs("Add") Then
OpPLSubGroupTot=OpPLSubGroupTot + OpVal
ELSEIF rs("SubTotal") Then
Response.Write("<TR>")
Response.Write("<TD Colspan=3>&nbsp;</TD>")
Response.Write("<TD Align=Right>---------</TD>")
Response.Write("</TR>")
Response.Write("<TR>")
IF rs("Bold") Then
Response.Write("<TD Colspan=3><Font Face=arial Size=2><B>" & rs("Alias") & "</B></Font></TD>")
Else
Response.Write("<TD Colspan=3><Font Face=arial Size=2>" & rs("Alias") & "</Font></TD>")
End IF
Response.Write("<TD Align=Right><Font Face=Arial Size=1>" & FormatNumber(OpPLSubGroupTot,2) & "</Font></TD>")
Response.Write("</TR>")
ELSE
OpPLSubGroupTot=OpPLSubGroupTot + OpVal
END IF
rs1.close


'=====================================ELSE EXECUTE FORMULA===========================================

ELSE
FORMULA=rs("Formula")
FORMULA="Result =" & FORMULA
EXECUTE FORMULA
TOTALASSETS=CurTOTALASSETS
TOTALLIABILITIES=CurTOTALLIABILITIES
FORMULA="CurResult=" & rs("Formula")
EXECUTE FORMULA
Response.Write("<TR>")
Response.Write("<TD colspan=3><Font Face=arial Size=1>&nbsp;</Font></TD>")
Response.Write("<TD Align=right><Font Face=arial Size=1>-------------------</Font></TD>")
Response.Write("<TD ><Font Face=arial Size=1>&nbsp;</Font></TD>")
Response.Write("<TD Align=right><Font Face=arial Size=1>-------------------</Font></TD>")
Response.Write("</TR>")
Response.Write("<TR>")
Response.Write("<TD><Font Face=arial Size=2><B>" & rs("Alias") & "</B></Font></TD>")
Response.Write("<TD colspan=2><Font Face=arial Size=1>&nbsp;</Font></TD>")
Response.Write("<TD Align=right><Font Face=arial Size=2><B>" & FormatNumber(Result,2) & "</B></Font></TD>")
Response.Write("<TD ><Font Face=arial Size=1>&nbsp;</Font></TD>")
Response.Write("<TD Align=right><Font Face=arial Size=2><B>" & FORMATNUMBER(CurResult,2) & "</B></Font></TD>")
Response.Write("</TR>")
Response.Write("<TR>")
Response.Write("<TD COLSPAN=6>&nbsp;</TD>")
Response.Write("</TR>")
END IF								'---------Formula Logic End
End IF
End IF 
'Response.End
rs.MoveNext 
Wend
rs.Close
Response.Write("</TABLE>")
End If
END IF
%>
</BODY>
</HTML>
















