<!--
'************************************************************************
'Pupose						:	This is a SoftServer FOLDER / document list server
'Filename					:	document.asp
'Author						:	Ashish Shah
'Created					:	30-Aug-2006
'Project Name				:	erpweb
'Contact					:	sales@erpweb.com
'
'Modification History		:	
'Purpose					:
'Version					:
'Author 					:
'Created					:
'************************************************************************
-->
<%' @TRANSACTION=Required LANGUAGE="VBScript" %>
<%'@ LANGUAGE="VBScript" %>
<%

'Option Explicit
'Response.Buffer = True
Const adUseClient = 3
UID=REQUEST("UID")
PROFILECSS="SAMPLE.CSS"
%>



<%
Set Conn=Server.CreateObject("ADODB.Connection")
Conn.Open Session("erp_ConnectionString")
%>
<%
'----------------------------------PROFILE GENERATOR
Set lst=Server.CreateObject("ADODB.Recordset")
lst.Open "Select * FROM GUIPROFILELIST WHERE UID=" & UID, Conn
if not lst.eof then
%>
<%PROFILECSS=lst("GUIPROFILE")%>
<%
end if
lst.Close
%>

<%
' encodes control characters for javascript
function aw_string(s)

    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""") 'replace javascript control characters - ", \
    s = Replace(s, vbCr, "\r")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, "'", " ")
    
    aw_string = s

end function
%>

<% 
Sub GenerateTablemac( rs, DID, UID )
  Response.Write( "var myHeaders = [" )
  Response.Write("'" + "Tools" + "', ")
  ' set up column names
  for i = 1 to rs.fields.count - 1
    '-------------------------
        if (rs(i).Type = 3) and (i > 2)then
        Set ls=Server.CreateObject("ADODB.Recordset")
		ls.Open "Select * from LISTBOX where DETFLAG=1 AND DDID=" & DDID & " AND COLUMNNO=" & i, Conn
		if not ls.EOF then
		LISTNAME=ls("LISTNAME") 
        Response.Write("'" + LISTNAME + "', ")
        ELSE
        Response.Write("'" + rs(i).Name + "', ")
        end if
        ls.Close
        set ls=nothing
        ELSE
        Response.Write("'" + rs(i).Name + "', ")
        end if
  '-----------------------
  next
  Response.Write( "];" )
  
  ' write each row
 Response.Write( "var myCells = [" )
 
  ' write each row
 
  WHILE NOT rs.EOF
    Response.Write( "[" )
    IDOC=rs(0)
  IDD=rs(1)
  Response.Write( "'" )
  Response.Write("<A HREF=FORMEDIT2.ASP?DID=" & DID & "&ID=" & IDOC & "&UID=" & UID & "&IDD=" & IDD & " target=new>Edit</A>" + " ")
  Response.Write("<A HREF=FORMDEL2.ASP?DID=" & DID & "&ID=" & IDOC & "&UID=" & UID & "&IDD=" & IDD & " target=new>Delete</A>" + " ")
  Response.Write( "'," )
    for i = 1 to rs.fields.count - 1
      v = rs(i):if isnull(v) then v = ""
      v=aw_string(v)
      '------------------------------------------------

		'--------------------------------------------------
      if (rs(i).Type = 3) and (v > "0") and i>2 then
		Set ls=Server.CreateObject("ADODB.Recordset")
		ls.Open "Select * from LISTBOX where DETFLAG=1 AND DDID=" & DDID & " AND COLUMNNO=" & i, Conn
		'-----------------------------------------------
		if not ls.EOF then
		LISTNAME=ls("LISTNAME")
		LISTTABLE=ls("LISTTABLE")
		LISTCOLUMN=ls("LISTCOLUMN")
		LISTVALUE=ls("LISTVALUE")
		
		'-----------------------------------------------
			Set lst=Server.CreateObject("ADODB.Recordset")
			lst.Open "Select " & LISTVALUE & " , " & LISTCOLUMN & " FROM " & LISTTABLE & " WHERE " & LISTVALUE & " = " & v, Conn
			if not lst.eof then
			v=lst(1)
			end if
			lst.Close
		else
		end if
		ls.Close
		Response.Write( "'" + CStr( v ) + "'," )
      '------------------------------------------------
      else
		Response.Write( "'" + CStr( v ) + "'," )
      end if
      '------------------------------------------------

      '------------------------------------------------
   next
    '------------------------------------
    Response.Write( "'" )
    '------------------------------------
    Response.Write( "']," )
    rs.MoveNext
    
  WEND '----------------MAIN FOR PAGESIZE
  Response.Write( "];" )
End Sub
%>


<HTML>
<HEAD>
<TITLE>MobileERP FOLDER DET SERVER (sales@erpweb.com)</TITLE>
<LINK REL=STYLESHEET TYPE="text/css" HREF="<%=PROFILECSS%>">
	<style>body {font: 12px Tahoma}</style>

<!-- include links to the script and stylesheet files -->
	<script src="macstyle/runtime/lib/aw.js"></script>
	<link href="macstyle/runtime/styles/xp/aw.css" rel="stylesheet" />

<!-- change default styles, set control size and position -->
<style>
	#myGrid {width: 900px; height: 140px;}
	#myGrid .aw-alternate-even {background: #eee;}
</style>
</HEAD>

<BODY topmargin=0 leftmargin=5>
<Basefont face=arial size=1>

<!--------------------------------------------------------------->
<table width=100% >
<tr>
<td valign=top>
<%
'-----------------------------------main program starts here

'------------------------------------find doc sql
CID=REQUEST("CID")
DID=CInt(REQUEST("DID"))
ID=REQUEST("ID")'-----------READ RECORD ID
if DID="" Then Response.Write "ERROR:DID":Response.END
set rsDOC = Server.CreateObject("ADODB.Recordset")
rsDOC.Open "Select * from DOCUMENTS WHERE DID=" & DID, Conn 
'SQLPROGRAM=trim(rsDOC("SQLPROGRAM"))
DETAILSSQL=rsDOC("SQLDETAILS")
if isnull(DETAILSSQL) OR DETAILSSQL="" THEN Response.Write ("NO Details Table Exists"):Response.End
DDID=rsDOC("DDID")
'------------------------------------
'GenerateHeader rsDoc

%>
<!-- insert control tag -->
	<span id="myGrid"></span>
<%
'-------------------------------------
    set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open DETAILSSQL & ID, Conn		
 '-----------------------
if not rs.EOF then
'--------------------------------------
	'Response.Write( "<table width=100% border=1 bordercolordark=#FFFFFF bordercolorlight=#FFFFFF><td align=left bgcolor=#f1f1f2><font face=arial SIZE=1></td>" )
	'Response.Write("<td align=right bgcolor=#f1f1f2><font face=arial SIZE=1></td></table>" )
'-------------------------------------
	%>
<script>
	<%
	GenerateTablemac rs, DID, UID
	%>
	//    create grid control
    var grid = new AW.UI.Grid;

//    assign the grid id (same as placeholder tag above)
    grid.setId("myGrid");

//    set grid text
    grid.setHeaderText(myHeaders);
    grid.setCellText(myCells);

//    set number of columns/rows
    grid.setColumnCount(myHeaders.length);
    grid.setRowCount(myCells.length);

//    write grid to the page
    grid.refresh();

</script>
	<%
'------------------------------------	
else'-----if rs not found
	Response.Write "<font face=arial size=2 color=red>No Documents Found</font><br>"
end if
'------------------------------------
'GenerateFooter rsDoc
'------------------------------------
%>
</td>
</tr>
</table>

<%
rs.Close	
rsDOC.Close
Conn.Close
Set rs = Nothing
Set rsDOC = Nothing
SET Conn = nothing
%>
</BODY>
</HTML>