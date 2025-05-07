<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE></TITLE>
</HEAD>
<BODY>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>IRP ANNUAL REPORT</TITLE>
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
<FORM NAME="FRMANNUALREPORT" METHOD="POST" ACTION="ANNUALREPORT123.ASP?MODE=1" TARGET=NEWS>
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
'GroupTot=0.0
'CurGroupTot=0.0
OpPLSubGroupTot=0.0
CurPLSubGroupTot=0.0
Pound=chr(163)
SQL="SELECT * FROM ANNUALREPORTDET WHERE ANNUALREPORTID=" & REPORTID
rs.Open SQL,Conn
IF Not rs.EOF Then
Response.Write("<TABLE BORDER=0 WIDTH=650>")
Response.Write("<TR>")
Response.Write("<TD colspan=2><Font Face=arial Size=1>&nbsp;</Font></TD>")
Response.Write("<TD colspan=2 align=center><Font Face=arial Size=2><B>" & FYEND& "</B></Font></TD>")
Response.Write("<TD colspan=2 align=center><Font Face=arial Size=2><B>" & PREVFYEND & "</B></Font></TD>")
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
GROUPID=rs("ACCGROUPID")
StrAlias=rs("Alias")
'RESPONSE.WRITE "GROUPID=" & GROUPID
IF GROUPID=0 THEN
ADETID=rs("AccGroupDetID")
	'RESPONSE.WRITE "ADETID=" & ADETID
	IF ADETID > 0 THEN
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
		DebitFlag=1
		ELSE
		OpVal=OPCREDIT
		DebitFlag=0
		END IF
		'---------------Find Current Year Data-------------------------
		
		SQL2="SELECT SUM(DEBIT), SUM(CREDIT) FROM LEDGERSCH WHERE ACCGROUPDETID=" & ADETID & " AND CALENDERID=" & CID
		rs2.Open SQL2,Conn 
		IF NOT rs2.EOF THEN'------------------1
		'---------------------
		OPDEBIT=rs2(0): IF ISNULL(OPDEBIT) OR OPDEBIT="" THEN OPDEBIT=0
		OPCREDIT=rs2(1): IF ISNULL(OPCREDIT) OR OPCREDIT="" THEN OPCREDIT=0
		ELSE
		OPDEBIT=0
		OPCREDIT=0
		END IF
		rs2.Close
		'---------------------
		SQL2="SELECT SUM(DEBIT), SUM(CREDIT) FROM LEDGERSCHF WHERE ACCGROUPDETID=" & ADETID & " AND CALENDERID=" & CID
		rs2.Open SQL2,Conn
		
		IF NOT rs2.EOF THEN'------------------1
		'---------------------
		DEBIT=rs2(0): IF ISNULL(DEBIT) OR DEBIT="" THEN DEBIT=0
		CREDIT=rs2(1): IF ISNULL(CREDIT) OR CREDIT="" THEN CREDIT=0
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
		rs2.Close
		IF DEBIT >= CREDIT Then
		CurYrVal=DEBIT
		CDebitFlag=1
		ELSE
		CurYrVal=CREDIT
		CDebitFlag=0
		END IF
		DRTOT=DRTOT+DEBIT
		CRTOT=CRTOT+CREDIT
'--------------


	IF rs("ADD") Then '-------------IF ADD FLAG SELECTED THEN

		OpPLSubGroupTot=OpPLSubGroupTot + OpVal
        CurPLSubGroupTot=CurPLSubGroupTot + CurYrVal
		ELSEIF rs("SUB") Then '------------IF SUB FLAG SELECTED THEN

		OpPLSubGroupTot=OpPLSubGroupTot - OpVal
        CurPLSubGroupTot=CurPLSubGroupTot - CurYrVal
		ELSEIF rs("SUBTOTAL") Then
		Response.Write("<TR>")
		IF rs("Bold") Then
		Response.Write("<TD COLSPAN=3><Font Face=arial Size=2><B>&nbsp;" & rs("Alias") & "</B></Font></TD>")
		ELSE
		Response.Write("<TD COLSPAN=3><Font Face=arial Size=1>&nbsp;" & rs("Alias") & "</Font></TD>")
		End If
		Response.Write("<TD><Font Face=arial Size=1>" & FORMATNUMBER(CurPLSubGroupTot,2) & "</Font></TD>")
		Response.Write("<TD Align=Right><Font Face=arial Size=1>" & FORMATNUMBER(OpPLSubGroupTot,2) & "</Font></TD>")
		Response.Write("</TR>")
	END IF
		Response.Write("<TR>")
		IF rs("Bold") Then
		Response.Write("<TD><Font Face=arial Size=2><B>&nbsp;" & rs("Alias") & "</B></Font></TD>")
		ELSE
		Response.Write("<TD><Font Face=arial Size=1>&nbsp;" & rs("Alias") & "</Font></TD>")
		End If
		Response.Write("<TD><Font Face=arial Size=1><A href=statementagroup.asp?MODE=2&CID=" & CID & "&AID=" & ADETID & ">" & rs1("SCHEDULE") & "</A></A></Font></TD>")
		Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
		IF Cdebitflag=1 Then
		Response.Write("<TD Align=Right><Font Face=arial Size=1>(" & FORMATNUMBER(CurYrVal,2) & ")</Font></TD>")
		Else
		Response.Write("<TD Align=Right><Font Face=arial Size=1>" & FORMATNUMBER(CurYrVal,2) & "</Font></TD>")
		End if
		Response.Write("<TD ><Font Face=arial Size=1>&nbsp;</Font></TD>")
		IF debitflag=1 Then
		Response.Write("<TD Align=Right><Font Face=arial Size=1>(" & FORMATNUMBER(OpVal,2) & ")</Font></TD>")
		Else
		Response.Write("<TD Align=Right><Font Face=arial Size=1>" & FORMATNUMBER(OpVal,2) & "</Font></TD>")
		End if
		Response.Write("</TR>")
  END IF
  rs1.close
'END IF
ELSE  '---------------------ADETID=0 THEN

IF rs("SUBTOTAL") Then
Response.Write("<TD colspan=3><Font Face=arial Size=1>&nbsp;</Font></TD>")
Response.Write("<TD align=right>----------</TD>")
Response.Write("<TD ><Font Face=arial Size=1>&nbsp;</Font></TD>")
Response.Write("<TD align=right>----------</TD>")
Response.Write("<TR>")
IF rs("Bold") Then
Response.Write("<TD colspan=3><Font Face=arial Size=2><B>&nbsp;" & rs("Alias") & "</B></Font></TD>")
ELSE
Response.Write("<TD colspan=3><Font Face=arial Size=1>&nbsp;" & rs("Alias") & "</Font></TD>")
End If
Response.Write("<TD Align=right><Font Face=arial Size=1>" & FORMATNUMBER(CurPLSubGroupTot,2) & "</Font></TD>")
Response.Write("<TD ><Font Face=arial Size=1>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</Font></TD>")
Response.Write("<TD Align=Right><Font Face=arial Size=1>" & FORMATNUMBER(OpPLSubGroupTot,2) & "</Font></TD>")
Response.Write("</TR>")
END IF
END IF
END IF
rs.MoveNext
Wend
rs.Close
End IF
end if
%>
</BODY>
</HTML>



