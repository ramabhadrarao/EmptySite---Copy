<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft FrontPage 4.0">
<TITLE>I-Business Annual Report</TITLE>
</HEAD>
<BODY>
<%
Set Conn=Server.CreateObject("ADODB.Connection")
Conn.Open Session("erp_ConnectionString")
%>
<%
SUB PRINTLINE
RESPONSE.WRITE("<TD ALIGN=RIGHT>-------------</TD>")
END SUB
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
<FORM NAME="FRMANNUALREPORT" METHOD="POST" ACTION="ANNUALREPORT.ASP?MODE=1" TARGET=NEWS>
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
CID=REQUEST("CID")
DID=REQUEST("DID")
UID=REQUEST("UID")
PREVYRDRTOT=0
PREVYRCRTOT=0
YRDRTOT=0
YRCRTOT=0
'pound=chr(163)
pound=Rs.
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
'---------------TITLE PRINTING
Response.Write("<Table Border=0 Width=650>")
Response.Write("<TR>")
Response.Write("<TD><Font Face=arial Size=2>(&nbsp;<B><U>" & companyname & "</B></U>&nbsp;)</Font></TD>")
Response.Write("</TR>")
Response.Write("<TR>")
Response.Write("<TD Align=center><Font Face=arial Size=2><B><U>Profit & Loss Account</B></U></Font></TD>")
Response.Write("</TR>")
Response.Write("<TR>")
Response.Write("<TD Align=center><Font Face=arial Size=2><B><U>For the Year Ended:" & FYEND & "</B></U></Font></TD>")
Response.Write("</TR>")
Response.Write("</Table>")
Response.Write("<br>")
'---------------TITLE PRINTING END
SQL="SELECT * FROM ANNUALLEDGERLIST WHERE SCHEDULE=1 OR SCHEDULE=2 OR SCHEDULE=3 OR SCHEDULE=4  ORDER BY SCHEDULE"
rs.Open SQL,Conn
If Not rs.EOF Then
Response.Write("<Table Border=0 Width=650>")
Response.Write("<TR>")
Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
Response.Write("<TD COLSPAN=2 Align=Center><Font Face=arial Size=1>(&nbsp;" & FYEND & "&nbsp;)</Font></TD>")
Response.Write("<TD COLSPAN=2 Align=Center><Font Face=arial Size=1>(&nbsp;" & FYEND & "&nbsp;)</Font></TD>")
Response.Write("</TR>")
Response.Write("<TR>")
Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
Response.Write("<TD><Font Face=arial Size=1>Notes</Font></TD>")
Response.Write("<TD align=right><Font Face=arial Size=1>" & pound & "</Font></TD>")
Response.Write("<TD align=right><Font Face=arial Size=1>" & pound & "</Font></TD>")
Response.Write("<TD align=right><Font Face=arial Size=1>" & pound & "</Font></TD>")
Response.Write("<TD align=right><Font Face=arial Size=1>" & pound & "</Font></TD>")
Response.Write("</TR>")
WHILE NOT rs.EOF
'------------------Found Privious Year Data.
SQL1="SELECT * FROM ANNUALANALYSISGROUPOPBAL WHERE ACCGROUPDETID=" & rs("ACCGROUPDETID") & " AND CALENDERID=" & CID
'RESPONSE.WRITE SQL1
Set rs1=Server.CreateObject("ADODB.RecordSet")
rs1.Open SQL1,Conn
IF Not rs1.EOF Then
PREVYRDR=rs1("DEBIT")
PREVYRCR=rs1("CREDIT")
ELSE
PREVYRDR=0
PREVYRCR=0
END IF
rs1.Close
'-----------------Found Currentyear Data.
SQL1="SELECT * FROM ANNUALTRANSACTION WHERE ACCGROUPDETID=" & rs("ACCGROUPDETID") & " AND CALENDERID=" & CID
'RESPONSE.WRITE SQL1
rs1.Open SQL1,Conn
IF Not rs1.EOF Then
YRDR=rs1("DEBIT")
YRCR=rs1("CREDIT")
rs1.Close
ELSE
YRDR=0
YRCR=0
END IF
Response.Write("<TR>")
Response.Write("<TD><Font Face=arial Size=1><B>" & rs("NAMEASPERCLAW") & "</B></Font></TD>")
Response.Write("<TD><Font Face=arial Size=1>" & rs("SCHEDULE") & "</Font></TD>")
IF YRDR > YRCR THEN
Response.Write("<TD align=right><Font Face=arial Size=1>" & FORMATCURRENCY(YRDR-YRCR,2)& "</Font></TD>")
Response.Write("<TD align=right><Font Face=arial Size=1>&nbsp;</Font></TD>")
ELSE
Response.Write("<TD align=right><Font Face=arial Size=1>&nbsp;</Font></TD>")
Response.Write("<TD align=right><Font Face=arial Size=1>" & FORMATCURRENCY(YRCR-YRDR,2) & "</Font></TD>")
END IF
IF PREVYRDR=0 OR PREVYRDR="" THEN
Response.Write("<TD align=right><Font Face=arial Size=1>&nbsp;</Font></TD>")
ELSE
Response.Write("<TD align=right><Font Face=arial Size=1>" & FORMATCURRENCY(PREVYRDR,2) & "</Font></TD>")
END IF
IF PREVYRCR=0 OR PREVYRCR="" THEN
Response.Write("<TD align=right><Font Face=arial Size=1>&nbsp;</Font></TD>")
ELSE
Response.Write("<TD align=right><Font Face=arial Size=1>" & FORMATCURRENCY(PREVYRCR,2) & "</Font></TD>")
END IF
Response.Write("</TR>")
IF rs("SCHEDULE")=1 OR rs("SCHEDULE")=4  THEN
PREVYRDRTOT=PREVYRDRTOT + PREVYRDR
PREVYRCRTOT=PREVYRCRTOT + PREVYRCR
YRDRTOT=YRDRTOT + YRDR
YRCRTOT=YRCRTOT + YRCR
ELSEIF rs("SCHEDULE")=2 or  rs("SCHEDULE")=3  THEN 
PREVYRDRTOT=PREVYRDRTOT - PREVYRDR
PREVYRCRTOT=PREVYRCRTOT - PREVYRCR
YRDRTOT=YRDRTOT - YRDR
YRCRTOT=YRCRTOT - YRCR
END IF
IF rs("SCHEDULE")=2  THEN
	Response.Write("<TR>")
	Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
	Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
	IF YRDRTOT=0 THEN
	Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
	ELSE
	PRINTLINE
	'Response.Write("<TD align=right><Font Face=arial Size=1><hr></Font></TD>")
	END IF
	IF YRCRTOT=0 THEN
	Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
	ELSE
	PRINTLINE
	'Response.Write("<TD align=right><Font Face=arial Size=1><hr></Font></TD>")
	END IF
	IF PREVYRDRTOT=0 THEN
	Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
	ELSE
	PRINTLINE
	'Response.Write("<TD align=right><Font Face=arial Size=1><hr></Font></TD>")
	END IF
	IF PREVYRCRTOT=0 THEN
	Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
	ELSE
	PRINTLINE
	'Response.Write("<TD align=right><Font Face=arial Size=1><hr></Font></TD>")
	END IF
	Response.Write("</TR>")
	Response.Write("<TR>")
	Response.Write("<TD><Font Face=arial Size=1><b>GROSSPROFIT:</b></Font></TD>")
	Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
	IF YRDRTOT=0 THEN
	Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
	ELSE
	Response.Write("<TD align=right><Font Face=arial Size=1>" & FORMATCURRENCY(YRDRTOT,2) & "</Font></TD>")
	END IF
	IF YRCRTOT=0 THEN
	Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
	ELSE
	Response.Write("<TD align=right><Font Face=arial Size=1>" & FORMATCURRENCY(YRCRTOT,2) & "</Font></TD>")
	END IF
	IF PREVYRDRTOT=0 THEN
	Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
	ELSE
	Response.Write("<TD align=right><Font Face=arial Size=1>" & FORMATCURRENCY(PREVYRDRTOT,2) & "</Font></TD>")
	END IF
	IF PREVYRCRTOT=0 THEN
	Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
	ELSE
	Response.Write("<TD align=right><Font Face=arial Size=1>" & FORMATCURRENCY(PREVYRCRTOT,2) & "</Font></TD>")
	END IF
	Response.Write("</TR>")
ELSEIF rs("SCHEDULE")=3  THEN
	Response.Write("<TR>")
	Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
	Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
	IF YRDRTOT > YRCRTOT THEN
	PRINTLINE
	'Response.Write("<TD align=right><Font Face=arial Size=1><HR></Font></TD>")
	Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
	ELSE
	Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
	PRINTLINE
	'Response.Write("<TD><Font Face=arial Size=1><HR></Font></TD>")
	END IF
	IF PREVYRDRTOT > PREVYRCRTOT THEN
	PRINTLINE
	'Response.Write("<TD><Font Face=arial Size=1><HR></Font></TD>")
	Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
	ELSE
	Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
	'Response.Write("<TD><Font Face=arial Size=1><HR></Font></TD>")
	PRINTLINE
	END IF
	Response.Write("</TR>")
	Response.Write("<TR>")
	Response.Write("<TD><Font Face=arial Size=1><b>OPERATING PROFIT / LOSS</b></Font></TD>")
	Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
	IF YRDRTOT > YRCRTOT THEN
	YRDRTOT=YRDRTOT+YRCRTOT
	Response.Write("<TD align=right><Font Face=arial Size=1>" & formatcurrency(YRDRTOT,2) & "</Font></TD>")
	Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
	ELSE
	YRCRTOT=YRCRTOT+YRDRTOT
	Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
	Response.Write("<TD align=right><Font Face=arial Size=1>" & formatcurrency(YRCRTOT,2) & "</Font></TD>")
	END IF
	IF PREVYRDRTOT > PREVYRCRTOT THEN
	PREVYRDRTOT=PREVYRDRTOT+PREVYRCRTOT
	Response.Write("<TD align=right><Font Face=arial Size=1>" & formatcurrency(PREVYRDRTOT,2) & "</Font></TD>")
	Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
	ELSE
	PREVYRCRTOT=PREVYRCRTOT+PREVYRDRTOT
	Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
	Response.Write("<TD align=right><Font Face=arial Size=1>" & formatcurrency(PREVYRCRTOT,2) & "</Font></TD>")
	END IF
	Response.Write("</TR>")
ELSEIF rs("SCHEDULE")=5 THEN
	Response.Write("<TR>")
	Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
	Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
	IF YRDRTOT=0 THEN
	Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
	ELSE
	PRINTLINE
	'Response.Write("<TD align=right><Font Face=arial Size=1><hr></Font></TD>")
	END IF
	IF YRCRTOT=0 THEN
	Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
	ELSE
	PRINTLINE
	'Response.Write("<TD align=right><Font Face=arial Size=1><hr></Font></TD>")
	END IF
	IF PREVYRDRTOT=0 THEN
	Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
	ELSE
	PRINTLINE
	'Response.Write("<TD align=right><Font Face=arial Size=1><hr></Font></TD>")
	END IF
	IF PREVYRCRTOT=0 THEN
	Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
	ELSE
	PRINTLINE
	'Response.Write("<TD align=right><Font Face=arial Size=1><hr></Font></TD>")
	END IF
	Response.Write("</TR>")
	Response.Write("<TR>")
	Response.Write("<TD><Font Face=arial Size=1><b>GROSSPROFIT:</b></Font></TD>")
	Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
	IF YRDRTOT=0 THEN
	Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
	ELSE
	Response.Write("<TD align=right><Font Face=arial Size=1>" & FORMATCURRENCY(YRDRTOT,2) & "</Font></TD>")
	END IF
	IF YRCRTOT=0 THEN
	Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
	ELSE
	Response.Write("<TD align=right><Font Face=arial Size=1>" & FORMATCURRENCY(YRCRTOT,2) & "</Font></TD>")
	END IF
	IF PREVYRDRTOT=0 THEN
	Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
	ELSE
	Response.Write("<TD align=right><Font Face=arial Size=1>" & FORMATCURRENCY(PREVYRDRTOT,2) & "</Font></TD>")
	END IF
	IF PREVYRCRTOT=0 THEN
	Response.Write("<TD><Font Face=arial Size=1>&nbsp;</Font></TD>")
	ELSE
	Response.Write("<TD align=right><Font Face=arial Size=1>" & FORMATCURRENCY(PREVYRCRTOT,2) & "</Font></TD>")
	END IF
	Response.Write("</TR>")
END IF
rs.MoveNext
WEND
Response.Write("<TR>")
Response.Write("<Td Colspan=6>&nbsp;</Font></TD>")
Response.Write("</TR>")
Response.Write("<TR>")
Response.Write("<Td Colspan=6><Font Face=arial Size=1>&nbsp;</Font></TD>")
Response.Write("</TR>")
Response.Write("<TR>")
Response.Write("<Td Colspan=6><Font Face=arial Size=1>*(Currency symbol to be modified as appropriate)</Font></TD>")
Response.Write("</TR>")
Response.Write("</Table>")
End IF
END IF  '------------Mode=1 Loop End
%>
</BODY>
</HTML>
