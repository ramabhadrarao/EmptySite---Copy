<%@ Language=VBScript %>
<HTML>
<HEAD>
<TITLE>STOCK(sales@erpweb)</TITLE>
</HEAD>
<BODY TOPMARGIN=0>
<%
Set Conn=Server.CreateObject("ADODB.Connection")
Conn.Open Session("erp_ConnectionString")
%>

<% 
Sub GenerateTable1( rs, ITEMID )
  'Response.Write("<form action=warehousetransfer.asp?mode=2 method=post>")
  Response.Write( "<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF>" )
  ' set up column names
  'Response.Write( "<TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080><font face=arial size=1>Select</font></TD>" )
  for i = 0 to rs.fields.count -3
    Response.Write("<TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080><font face=arial size=1>" + rs(i).Name + "</font></TD>")
  next
  'Response.Write( "<TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080><font face=arial size=1>IssueQTY</font></TD>" )
 ' Response.Write( "<TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080><font face=arial size=1>Warehouse</font></TD>" )
  
  ' write each row
  J=0
  '  Response.Write( "<input type=HIDDEN name=ITEMID value=" & ITEMID & ">" )
 '   Response.Write( "<input type=HIDDEN name=CID value=" & CID & ">" )
	''Response.Write( "<input type=HIDDEN name=STOREDETID value=" & STOREDETID & ">" )
  while not rs.EOF
    Response.Write( "<TR bgcolor=#E1F2FD>" )
    '---------------------------copy data
    Response.Write( "<input type=HIDDEN name=FLD1" & J & " value=" & rs("itemID") & ">" )
    '---------------------------
    'Response.Write( "<TD VALIGN=TOP><font face=arial size=1><input type=checkbox name=COPY" & J & " ></font></TD>" )
    for i = 0 to rs.fields.count -3
      v = rs(i)
      if isnull(v) then v = ""
      Response.Write( "<TD VALIGN=TOP><font face=arial size=1>" + CStr( v ) + "</font></TD>" )
    next
  '  Response.Write( "<TD><input type=TEXT name=FLD2" & J & " value=0 SIZE=5></TD>" )
'	Response.Write( "<TD>") 
'	 LIST2 Conn ,J
'	 Response.Write("</TD>" )
	 
'	Response.Write( "<input type=HIDDEN name=FLD4" & J & " VALUE=" & rs("unitid") & ">" )
'	Response.Write( "<input type=HIDDEN name=FLD5" & J & " VALUE=" & rs("qty") & ">" )
	'Response.Write( "<input type=HIDDEN name=FLD3" & J & " VALUE=" & rs("storedetid") & ">" )
	'Response.Write( "<input type=HIDDEN name=FLD6" & J & " VALUE=" & request("storedetid") & ">" )
'	Response.Write( "<input type=HIDDEN name=FLD7" & J & " VALUE=" & rS("fgRECEIVEid") & ">" )
    rs.MoveNext
    J=J+1
  wend 
  'IF J>=1 THEN
  'Response.Write( "</TABLE><hr>" )
  'Response.Write( "<font face=arial size=1>Issue Date:</font><input size=10 type=TEXT name=ISSUEDATE value=" & DATE & ">" )
'  Response.Write( "<input size=10 type=HIDDEN name=PROFINVNO value=FGTOSTORE>" )
 ' Response.Write( "<input type=HIDDEN name=MODE value=2>" )
 ' Response.Write( "<input type=HIDDEN name=J value=" & J & ">" )
 ' Response.Write( "<input type=submit name=s value='Transfer Stock to Warehouse'>" )
 ' Response.Write( "</form>" ) 
 ' ELSE
 ' RESPONSE.WRITE "MESSAGE: NO STOCK SO CANNOT TRANSFER"
 ' END IF
End Sub
%>



<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF >
<TR>
<TD valign=middle bgcolor=#D4D4D4 bordercolordark=#FFFFFF bordercolor=#808080 bordercolorlight=#808080>&nbsp;&nbsp;<FONT FACE=ARIAL SIZE=2><img src=ofolder.gif border=0 alt='SoftRobot Document Server'><B>Stock</B></FONT></TD>
</TR>
</table>

<%
MODE=REQUEST("MODE")

%>
<%
IF cint(MODE)=1 THEN'---------------

ID=REQUEST("ID")
SQL="SELECT * FROM storeitem WHERE storeID=" & ID


Set rs=Server.CreateObject("ADODB.Recordset")
rs.Open SQL,Conn

GenerateTable1 rs,  ID
rs.Close
elseif mode=2 then'------------------------
ID=REQUEST("ID")
SQL="SELECT * FROM wareitem WHERE warehouseID=" & ID


Set rs=Server.CreateObject("ADODB.Recordset")
rs.Open SQL,Conn

GenerateTable1 rs,  ID
rs.Close
END IF

Set Conn=nothing
'Conn.Close
%>
</BODY>
</HTML>
