<!--
'************************************************************************
'Pupose						:	This is a SoftServer FOLDER / document list server
'Filename					:	document.asp
'Author						:	Ashish Shah
'Created					:	30-Aug-2007
'Project Name				:	erpweb
'Contact					:	sales@erpweb.com
'
'Modification History		:	
'Purpose					:
'Version					: 7
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
FADD=REQUEST("FADD"): IF FADD THEN FADD=1 ELSE FADD=0
FDEL=REQUEST("FDEL"): IF FDEL THEN FDEL=1 ELSE FDEL=0
FVIEW=REQUEST("FVIEW"): IF FVIEW THEN FVIEW=1 ELSE FVIEW=0
FEDIT=REQUEST("FEDIT"): IF FEDIT THEN FEDIT=1 ELSE FEDIT=0
FFILTER=REQUEST("FILTER"): : IF FFILTER THEN FFILTER=1 ELSE FFILTER=0
FOFFHOLD=REQUEST("FOFFHOLD"): IF FOFFHOLD THEN FOFFHOLD=1 ELSE FOFFHOLD=0
FREJECT=REQUEST("FREJECT"): IF FREJECT THEN FREJECT=1 ELSE FREJECT=0
UID=REQUEST("UID")
PAGE=TRIM((REQUEST("PAGE"))):IF PAGE="" OR ISNULL(PAGE) THEN PAGE=1
FLAG=0
COLUMNNAME=REQUEST("ls")
SEARCHVALUE=REQUEST("txtsearch")

'IF NOT SEARCHVALUE="" THEN
	'RESPONSE.WRITE "<BR><FONT FACE=ARIAL SIZE=3><b>Searched Upon ('" & columnname & "') : " & cstr(searchvalue) & "</b></FONT><br><BR>"
'END IF

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
'----------------------------------LISTBOX GENERATOR
Sub LIST1 ( LISTTABLE, LISTCOLUMN, LISTVALUE, Conn, name, value )
Set lst=Server.CreateObject("ADODB.Recordset")
lst.Open "Select " & LISTVALUE & " , " & LISTCOLUMN & " FROM " & LISTTABLE & " WHERE " & LISTVALUE & " = " & value, Conn
if not lst.eof then
%>
<%=lst(1)%>
<%
end if
lst.Close
End Sub
%>

<%
Sub LISTFIELDS( rs )
for i = 3 to rs.fields.count - 1
    Response.Write("<Option value=" & rs(i).Name & ">" & rs(i).Name & "</option>")
next
End Sub
%>
<%
Sub LISTFIELDSPAGE( rs )
for i = 3 to rs.fields.count - 1
    IF (rs(i).Type=129 OR rs(i).Type=130 OR rs(i).Type=204 OR rs(i).Type=200) then
    Response.Write("<Option value=" & rs(i).Name & ">" & rs(i).Name & "</option>")
    END IF
next
End Sub
%>
<%
Sub LISTFIELDSRANGE( rs )
for i = 3 to rs.fields.count - 1
    IF NOT (rs(i).Type=129 OR rs(i).Type=130 OR rs(i).Type=204 OR rs(i).Type=200) then
    Response.Write("<Option value=" & rs(i).Name & ">" & rs(i).Name & "</option>")
    END IF
next
End Sub
%>
<% 
Sub GenerateTable( rs, pagesize )
	
  Response.Write( "<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF background=images/bg6.gif>" )
  Response.Write( "<TR bgcolor=#D5EAFF >" )
   '--------------------------------------
  ' set up column names
  for i = 3 to rs.fields.count - 1
        if (rs(i).Type = 3) and (i > 3)then
        Set ls=Server.CreateObject("ADODB.Recordset")
		ls.Open "Select * from LISTBOX where DETFLAG=0 AND DDID=" & DDID & " AND COLUMNNO=" & i, Conn
		if not ls.EOF then
		LISTNAME=ls("LISTNAME") 
        Response.Write("<TD bordercolordark=#FFFFFF bordercolorlight=#808080 ><FONT FACE=ARIAL SIZE=1 >" + LISTNAME + "</FONT></TD>")
        ELSE
        Response.Write("<TD bordercolordark=#FFFFFF bordercolorlight=#808080 ><FONT FACE=ARIAL SIZE=1 >" + rs(i).Name + "</FONT></TD>")
        end if
        ls.Close
        set ls=nothing
        ELSE
        Response.Write("<TD bordercolordark=#FFFFFF bordercolorlight=#808080 ><FONT FACE=ARIAL SIZE=1 >" + rs(i).Name + "</FONT></TD>")
        end if
    'Response.Write("<TD><FONT FACE=ARIAL SIZE=1 >" + rs(i).Name + "</FONT></TD>")
  next
  Response.Write("<TD align=right bordercolordark=#FFFFFF bordercolorlight=#808080><FONT FACE=ARIAL SIZE=1 ><B>Tools</B></FONT></TD>")
  
  '------------------------------------------
  ' write each row
  FOR K = 1 TO pagesize
  If NOT rs.EOF Then
    Response.Write( "<TR>" )
    
    '--------------------------------
    FLAG=NOT FLAG
    '--------------------------------
    
    '-------------------------------
    for i = 3 to rs.fields.count - 1
      v = rs(i)
      if isnull(v) then v = ""
        '----------------------------
      if i=3 then%>
        <TD VALIGN=TOP bgcolor=#E1F2FD ><FONT FACE=ARIAL SIZE=1><%=v%></font></td>
	  <%elseif (rs(i).Type = 3) and (v > "0") then
		Set ls=Server.CreateObject("ADODB.Recordset")
		ls.Open "Select * from LISTBOX where DETFLAG=0 AND DDID=" & DDID & " AND COLUMNNO=" & i, Conn
		if not ls.EOF then
		LISTNAME=ls("LISTNAME")
		LISTTABLE=ls("LISTTABLE")
		LISTCOLUMN=ls("LISTCOLUMN")
		LISTVALUE=ls("LISTVALUE")
		if FLAG THEN
		%>
		<TD VALIGN=TOP bgcolor=#E1F2FD ><FONT FACE=ARIAL SIZE=1>
		<%LIST1 LISTTABLE, LISTCOLUMN, LISTVALUE, Conn, name, v%> 
		</font></td>
		<%ELSE%>
		<TD VALIGN=TOP bgcolor=#E1F2FD ><FONT FACE=ARIAL SIZE=1>
		<%LIST1 LISTTABLE, LISTCOLUMN, LISTVALUE, Conn, name, v%> 
		</font></td>
		<%END IF%>
		<%
	else
		IF FLAG THEN
			Response.Write( "<TD VALIGN=TOP bgcolor=#E1F2FD ><FONT FACE=ARIAL SIZE=1>" + CStr( v ) + "</FONT></TD>" )
		   ELSE
		    Response.Write( "<TD VALIGN=TOP bgcolor=#E1F2FD ><FONT FACE=ARIAL SIZE=1>" + CStr( v ) + "</FONT></TD>" )
		END IF
		end if
		ls.Close
		set ls=nothing
		'-----------------------------------------
    elseif rs(i).name = "EMAILNO" then
    '-----------------------
           IF FLAG THEN
			Response.Write( "<TD VALIGN=TOP bgcolor=#E1F2FD><FONT FACE=ARIAL SIZE=1><a href=mailto:" & CStr( v ) & ">" & CStr( v ) & "</a></FONT></TD>" )
		   ELSE
		    Response.Write( "<TD VALIGN=TOP bgcolor=#E1F2FD><FONT FACE=ARIAL SIZE=1><a href=mailto:" & CStr( v ) & ">" & CStr( v ) & "</a></FONT></TD>" )
		   END IF
	 '---------------------
	 elseif rs(i).name = "WEBSITE" then
     '-----------------------
           IF FLAG THEN
			Response.Write( "<TD VALIGN=TOP bgcolor=#E1F2FD><FONT FACE=ARIAL SIZE=1><a href=http://" & CStr( v ) & " target=news>" & CStr( v ) & "</a></FONT></TD>" )
		   ELSE
		    Response.Write( "<TD VALIGN=TOP bgcolor=#E1F2FD><FONT FACE=ARIAL SIZE=1><a href=http://" & CStr( v ) & " target=news>" & CStr( v ) & "</a></FONT></TD>" )
		   END IF
     '--------------------
     elseif rs(i).name = "UPLOADFILE" then
            v=rs(i).value: if v="" or isnull(v) then v=""
	        Response.Write( "<TD VALIGN=TOP bgcolor=#E1F2FD><FONT FACE=ARIAL SIZE=1><a href=" & CStr( v ) & " target=new>ShowFile</a></FONT></TD>" )

	 '---------------------
	 elseif rs(i).name = "PHONENO" then
     '-----------------------
           IF FLAG THEN
			Response.Write( "<TD VALIGN=TOP bgcolor=#E1F2FD><FONT FACE=ARIAL SIZE=1><a href=file:///C:/WINDOWS/DIALER.EXE target=news>" & CStr( v ) & "</a></FONT></TD>" )
		   ELSE
		    Response.Write( "<TD VALIGN=TOP bgcolor=#E1F2FD><FONT FACE=ARIAL SIZE=1><a href=file:///C:/WINDOWS/DIALER.EXE target=news>" & CStr( v ) & "</a></FONT></TD>" )
		   END IF
    '-----------------------------------------
   		else
   		 IF FLAG THEN
			Response.Write( "<TD VALIGN=TOP bgcolor=#E1F2FD ><FONT FACE=ARIAL SIZE=1>" + CStr( v ) + "</FONT></TD>" )
		   ELSE
		    Response.Write( "<TD VALIGN=TOP bgcolor=#E1F2FD ><FONT FACE=ARIAL SIZE=1>" + CStr( v ) + "</FONT></TD>" )
		END IF  
    end if
    
        '-----------------------
           'IF FLAG THEN
			'Response.Write( "<TD VALIGN=TOP bgcolor=#E1F2FD><FONT FACE=ARIAL SIZE=1>" + CStr( v ) + "</FONT></TD>" )
		   'ELSE
		    'Response.Write( "<TD VALIGN=TOP bgcolor=#E1F2FD><FONT FACE=ARIAL SIZE=1>" + CStr( v ) + "</FONT></TD>" )
		   'END IF
	   '---------------------	   
    next
      '-----------------------------
      Response.Write("<TD bgcolor=#D5EAFF align=right><FONT FACE=ARIAL SIZE=1>")
  'IF DID=3433 OR DID=747 OR DID = 2740 OR DID=2570 OR DID=2783 OR DID=2815 OR DID=2769 OR DID=2794 OR DID=2816 OR DID=2836 OR DID=2844 or did=2881 OR DID=2706 OR DID=2708 OR DID=2709 or DID=2626 OR DID=2932 or DID=3041 OR DID=3135 or did=3258 or did=3269 or did=3344 or did=3280 or did=3401 or did=3298 or did=3316 or did=3323 or did=3330 or did=3366 or did=3373 or did=3394 THEN
  
  'Response.Write("<A HREF=TABWINDOWS.ASP?DID=" & DID & "&ID=" & rs(3) & "&UID=" & UID & " target=startmenu><img src=info.gif border=0 width=20 alt='Tabbed Document'></A>")
 ' END IF
  
  'IF DID=2740 THEN
  'Response.Write("<A HREF=TABWINDOWS.ASP?DID=" & DID & "&ID=" & rs(3) & "&UID=" & UID & " target=startmenu><img src=info.gif border=0 width=20 alt='Tabbed Document'></A>")
  'END IF
  
  IF FVIEW THEN
  Response.Write("<A HREF=FORMVIEW.ASP?DID=" & DID & "&ID=" & rs(3) & "&UID=" & UID & " target=news><img src=print.gif border=0 width=20 alt='Show Document' ></A>")
  END IF
  IF FEDIT THEN
  Response.Write("<A HREF=FORMEDITNEW.ASP?DID=" & DID & "&ID=" & rs(3) & "&UID=" & UID & " target=news><img src=update.gif border=0 width=20 alt='Edit Document Style 2' ></A>")
  END IF
  IF FDEL THEN
  Response.Write("<A HREF=FORMDEL.ASP?DID=" & DID & "&ID=" & rs(3) & "&UID=" & UID & " target=news><img src=delete.gif border=0 width=20 alt='Delete Document' ></A>")
  END IF
  IF FFILTER THEN
  'Response.Write("<A HREF=FORMCHECK.ASP?DID=" & DID & "&ID=" & rs(3) & "&UID=" & UID & " ><img src=info.gif border=0 width=20 alt='Validate Document'></A>")
  Response.Write("<A HREF=FORMAPPROVE.ASP?DID=" & DID & "&ID=" & rs(3) & "&UID=" & UID & " target=news><img src=approve.gif border=0 width=20 alt='Approve Document' ></A>")
  END IF
  IF FOFFHOLD THEN
  Response.Write("<A HREF=FORMOFFHOLD.ASP?DID=" & DID & "&ID=" & rs(3) & "&UID=" & UID & " target=news><img src=approve.gif border=0 width=20 alt='OffHold Document' ></A>")
  END IF
  Response.Write("</FONT></TD>")
  '----------------------------------
  rs.MoveNext
  end if 
  '-------------------------------------
  NEXT 
  Response.Write( "</TABLE>" )
'-----------------------
End Sub
%>
<%
' encodes control characters for javascript
function aw_string(s)

    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""") 'replace javascript control characters - ", \
    s = Replace(s, vbCr, "\r")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, "'", " ")
    s = Replace(s, ",", " ")

    aw_string = s

end function
%>

<% 
Sub GenerateTablemac( rs, DID, UID, pagesize )
'--------------------------------------------------
  Response.Write( "var myHeaders = [" )
  ' set up column names
  Response.Write("'" + "Tools" + "', ")
  for i = 3 to rs.fields.count - 1
  '-------------------------
        if (rs(i).Type = 3) and (i > 3)then
        Set ls=Server.CreateObject("ADODB.Recordset")
		ls.Open "Select * from LISTBOX where DETFLAG=0 AND DDID=" & DDID & " AND COLUMNNO=" & i, Conn
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
 FOR K = 1 TO pagesize
  If NOT rs.EOF Then
    Response.Write( "[" )
    IDOC=rs(3)
  Response.Write( "'" )
  Response.Write("<A HREF=documentdet.asp?DID=" & DID & "&ID=" & IDOC & "&UID=" & UID & " target=news>Detail</A>" + " ")
  IF FVIEW THEN
  Response.Write("<A HREF=FORMVIEW.ASP?DID=" & DID & "&ID=" & IDOC & "&UID=" & UID & " >View</A>" + " ")
  END IF
  IF FEDIT THEN
  'Response.Write("<A HREF=FORMEDITmac.ASP?DID=" & DID & "&ID=" & IDOC & "&UID=" & UID & " >Edit</A>" + " ")
  Response.Write("<A HREF=FORMEDITNEW.ASP?DID=" & DID & "&ID=" & IDOC & "&UID=" & UID & " >Edit</A>" + " ")
  END IF
  IF FDEL THEN
  Response.Write("<A HREF=FORMDEL.ASP?DID=" & DID & "&ID=" & IDOC & "&UID=" & UID & " >Del</A>" + " ")
  END IF
  IF FFILTER THEN
  Response.Write("<A HREF=FORMAPPROVE.ASP?DID=" & DID & "&ID=" & IDOC & "&UID=" & UID & " >Approve</A>" + " ")
  END IF
  IF FOFFHOLD THEN
  Response.Write("<A HREF=FORMOFFHOLD.ASP?DID=" & DID & "&ID=" & IDOC & "&UID=" & UID & " >Offhold</A>" + " ")
  END IF
  Response.Write( "'," )
  '----------------------------------
    for i = 3 to rs.fields.count - 1
      v = rs(i): if isnull(v) then v = ""
      v=aw_string(v)
      '--------------------------------------------------
      if i=3 then
        Response.Write( "'" + CStr( v ) + "'," )
      elseif (rs(i).Type = 3) and (v > "0") then
		Set ls=Server.CreateObject("ADODB.Recordset")
		ls.Open "Select * from LISTBOX where DETFLAG=0 AND DDID=" & DDID & " AND COLUMNNO=" & i, Conn
		'-----------------------------------------------
		if not ls.EOF then
		LISTNAME=ls("LISTNAME")
		LISTTABLE=ls("LISTTABLE")
		LISTCOLUMN=ls("LISTCOLUMN")
		LISTVALUE=ls("LISTVALUE")
		
		'-----------------------------------------------
		Set lst=Server.CreateObject("ADODB.Recordset")
		lst.Open "Select " & LISTVALUE & " , " & LISTCOLUMN & " FROM " & LISTTABLE & " WHERE " & LISTVALUE & " = " & v, Conn
		'SQL= "Select " & LISTVALUE & " , " & LISTCOLUMN & " FROM " & LISTTABLE & " WHERE " & LISTVALUE 
		'RESPONSE.Write SQL
		'response.end
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
   next
    '------------------------------------
  Response.Write( "'" )
    '------------------------------------
    Response.Write( "']," )
    rs.MoveNext
    END IF '-----------------MAIN IF
  NEXT '----------------MAIN FOR PAGESIZE
  Response.Write( "];" )
End Sub
%>

<%
Sub RLISTFIELDS( rs ) 
for i = 3 to rs.fields.count - 1
	'----------------Date TypesOnly
	if rs(i).type=135 then
	    Response.Write("<Option value=" & rs(i).Name & ">" & rs(i).Name & "</option>")
   	end if
next
End Sub
%>

<%
Sub RLISTNUMFIELDS( rs )

for i = 3 to rs.fields.count - 1
	'----------------Numeric Data Types only
	if (rs(i).Type=3 OR rs(i).Type=4 OR rs(i).Type=5 OR rs(i).Type=6 OR rs(i).Type=2 OR rs(i).Type=131 OR rs(i).Type=17 OR rs(i).Type=128 OR rs(i).Type=204) THEN
	    Response.Write("<Option value=" & rs(i).Name & ">" & rs(i).Name &  "</option>")
   	end if

next
End Sub
%>

<%
Sub RLISTTXTFIELDS( rs )
for i = 3 to rs.fields.count - 1
	'----------------Text Data Types only
	if (rs(i).Type=129 OR rs(i).Type=130 OR rs(i).Type=204 OR rs(i).Type=200) THEN
	    Response.Write("<Option value=" & rs(i).Name & ">" & rs(i).Name & "</option>")
   	end if
next
End Sub
%>

<%
Sub RTLISTFIELDS( rs )
for i = 3 to rs.fields.count - 1
	'----------------Date TypesOnly
	if not (rs(i).type=135 or rs(i).type=72 or rs(i).type=11) then
	    Response.Write("<Option value=" & rs(i).Name & ">" & rs(i).Name & "</option>")
   	end if
next
End Sub
%>

<%
Sub GenerateHeader( rs, PAGE )
%>
<%
Response.Write( "<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF >" )
Response.Write( "<TR>" )
Response.Write( "<TD valign=middle bgcolor=#DCDDDE bordercolordark=#FFFFFF bordercolor=#808080 bordercolorlight=#808080>&nbsp;&nbsp;<FONT FACE=ARIAL SIZE=2><img src=ofolder.gif border=0 alt='SoftRobot Document Server'><B> " & rs("TITLE") & " Folder</B></FONT></TD>" )
%>
<form name=form321 action=document.asp method=post ID="Form1">
<INPUT TYPE=HIDDEN NAME=DID VALUE=<%=DID%> ID="Hidden1">
<INPUT TYPE=HIDDEN NAME=UID VALUE=<%=UID%> ID="Hidden2">
<INPUT TYPE=HIDDEN NAME=FADD VALUE=<%=FADD%> ID="Hidden3">
<INPUT TYPE=HIDDEN NAME=FDEL VALUE=<%=FDEL%> ID="Hidden4">
<INPUT TYPE=HIDDEN NAME=FEDIT VALUE=<%=FEDIT%> ID="Hidden5">
<INPUT TYPE=HIDDEN NAME=FVIEW VALUE=<%=FVIEW%> ID="Hidden6">
<INPUT TYPE=HIDDEN NAME=FILTER VALUE=<%=FFILTER%> ID="Hidden7">
<INPUT TYPE=HIDDEN NAME=FOFFHOLD VALUE=<%=FOFFHOLD%> ID="Hidden8">
<INPUT TYPE=HIDDEN NAME=FREJECT VALUE=<%=FREJECT%> ID="Hidden9">
<INPUT TYPE=HIDDEN NAME=CID VALUE=<%=CID%> ID="Hidden10">
<INPUT TYPE=HIDDEN NAME=YEARFLAG VALUE=<%=YEARFLAG%> ID="Hidden11">
<%IF PAGE<=1 THEN P=PAGE ELSE P=PAGE-1%>
<input type=hidden name=PAGE value=<%=P%>>
<TD valign=middle width=20 align=right bgcolor=#DCDDDE bordercolordark=#FFFFFF bordercolor=#808080 bordercolorlight=#808080>
<input type=image height=25 src=gryleft.gif name=Prev>
</td>
</form>
<form name=form322 action=document.asp method=post ID="Form2">
<INPUT TYPE=HIDDEN NAME=DID VALUE=<%=DID%> ID="Hidden12">
<INPUT TYPE=HIDDEN NAME=UID VALUE=<%=UID%> ID="Hidden13">
<INPUT TYPE=HIDDEN NAME=FADD VALUE=<%=FADD%> ID="Hidden14">
<INPUT TYPE=HIDDEN NAME=FDEL VALUE=<%=FDEL%> ID="Hidden15">
<INPUT TYPE=HIDDEN NAME=FEDIT VALUE=<%=FEDIT%> ID="Hidden16">
<INPUT TYPE=HIDDEN NAME=FVIEW VALUE=<%=FVIEW%> ID="Hidden17">
<INPUT TYPE=HIDDEN NAME=FILTER VALUE=<%=FFILTER%> ID="Hidden18">
<INPUT TYPE=HIDDEN NAME=FOFFHOLD VALUE=<%=FOFFHOLD%> ID="Hidden19">
<INPUT TYPE=HIDDEN NAME=FREJECT VALUE=<%=FREJECT%> ID="Hidden20">
<INPUT TYPE=HIDDEN NAME=CID VALUE=<%=CID%> ID="Hidden21">
<INPUT TYPE=HIDDEN NAME=YEARFLAG VALUE=<%=YEARFLAG%> ID="Hidden22">
<TD valign=middle width=20 align=right bgcolor=#DCDDDE bordercolordark=#FFFFFF bordercolor=#808080 bordercolorlight=#808080>
<%'IF PAGE<numpages THEN P=PAGE ELSE P=PAGE+1%>
<input type=hidden name=PAGE value=<%=PAGE+1%>>
<input type=image height=25 src=gryright.gif name=Next>
</td>
</form>
<%
Response.Write( "</TR>" )
Response.Write( "</TABLE>" )
%>
<%
End Sub
%>
<%
Sub GenerateFooter( rs )
Response.Write( "<TABLE BORDER=0 WIDTH=100% >" )
Response.Write( "<TR bgcolor=#F1F1F2>" )
IF FADD THEN
  Response.Write("<TD WIDTH=80% ></TD><TD ALIGN=RIGHT><FONT FACE=ARIAL SIZE=2><A HREF=FORMADDNEW.ASP?UID=" & UID & "&DID=" & DID & "&CID=" & CID & " ><img height=15 src=add.jpg border=0 alt='Add Document'></A></FONT></TD>")
ELSE
  Response.Write("<TD WIDTH=80% ></TD><TD></TD>" )
END IF
Response.Write( "</TR>" )
Response.Write( "</TABLE>" )
End Sub
%>

<HTML>
<HEAD>
<TITLE>MobileERP FOLDER SERVER (sales@erpweb.com)</TITLE>
<script type="text/javascript" src="js/tabpane.js"></script>
<link type="text/css" rel="StyleSheet" href="css/tab.webfx.css" />
<LINK REL=STYLESHEET TYPE="text/css" HREF="<%=PROFILECSS%>">
	<style>body {font: 12px Tahoma}</style>

<!-- include links to the script and stylesheet files -->
	<script src="macstyle/runtime/lib/aw.js"></script>
	<link href="macstyle/runtime/styles/xp/aw.css" rel="stylesheet" />

<!-- change default styles, set control size and position -->
<style>
	#myGrid {width: 970px; height: 300px;}
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
Dim mypage, numpages
Dim numrecs, pagesize
'pagesize = CInt( Request("recs") )
'If pagesize = 0 Then pagesize = 2
pagesize=50
mypage = trim(Request("PAGE")):if not isnumeric(mypage) then mypage=1
If mypage="" Then mypage=1
'------------------------------------find doc sql
CID=REQUEST("CID")
DID=CInt(REQUEST("DID"))
SORT=REQUEST("SORT")
if DID="" Then Response.Write "ERROR:DID":Response.END
set rsDOC = Server.CreateObject("ADODB.Recordset")
rsDOC.Open "Select * from DOCUMENTS WHERE DID=" & DID, Conn 
SQLPROGRAM=trim(rsDOC("SQLPROGRAM"))
DDID=rsDOC("DDID")
TABLENM=rsDOC("MASTERTABLE")
DOCID=TRIM(TABLENM) & "ID"
YEARFLAG=0
IF rsDOC("YEARFILTER") THEN YEARFLAG=1 ELSE YEARFLAG=0
'------------------------------------
GenerateHeader rsDoc, PAGE
%>

<%
'-------------------------------------
    set rs = Server.CreateObject("ADODB.Recordset")
	rs.PageSize = pagesize
	rs.CacheSize = pagesize
	rs.CursorLocation = adUseClient
	'------------------------FILTER DATA
	IF SEARCHVALUE = "" THEN
        	IF DID=385 OR DID=384 OR DID=390 OR DID=391 THEN
				rs.Open SQLPROGRAM , Conn
			ELSE
				IF YEARFLAG THEN
					IF SORT=1 THEN
					rs.Open SQLPROGRAM & " AND UID=" & UID & " AND CALENDERID=" & CID & " ORDER BY " & COLUMNNAME & " ASC", Conn
					ELSEIF SORT=2 THEN
					rs.Open SQLPROGRAM & " AND UID=" & UID & " AND CALENDERID=" & CID & " ORDER BY " & COLUMNNAME & " DESC", Conn
					ELSE 
'					'RESPONSE.WRITE SQLPROGRAM & " AND CALENDERID=" & CID
'					'RESPONSE.END
					rs.Open SQLPROGRAM & " AND UID=" & UID & " AND CALENDERID=" & CID, Conn
					END IF
				ELSE
					IF SORT=1 THEN
					rs.Open SQLPROGRAM & " ORDER BY " & COLUMNNAME & " ASC", Conn
					ELSEIF SORT=2 THEN
					rs.Open SQLPROGRAM & " ORDER BY " & COLUMNNAME & " DESC", Conn
					ELSE 
					rs.Open SQLPROGRAM, Conn
					END IF
				END IF
        	END IF
	ELSE
					SQLFIND= SQLPROGRAM
					rs.Open SQLFIND, Conn
					if NOT rs.EOF THEN
						rsTYPEID=rs(COLUMNNAME).Type
						'RESPONSE.WRITE TYPEID
					END IF
					rs.Close
					'-----------------------------
					seachdt = trim(searchvalue)
					IF rsTYPEID=135 THEN
						'IF NOT ISDATE( searchdt ) THEN RESPONSE.WRITE "<FONT FACE=ARIAL SIZE=5>ERROR ENTERING DATE</FONT>": RESPONSE.END
					END IF
					'-----------------------------
	
	
			IF YEARFLAG THEN
					if rsTYPEID=135 then
		  					rs.Open SQLPROGRAM & " AND datepart(dd, " & COLUMNNAME & ") = datepart(dd, '" & TRIM(SEARCHVALUE) & "')"  & " AND datepart(mm, " & COLUMNNAME & ") = datepart(mm, '" & TRIM(SEARCHVALUE) & "')"  & " AND datepart(yy, " & COLUMNNAME & ") = datepart(yy, '" & TRIM(SEARCHVALUE) & "')" & " AND CALENDERID=" & CID, Conn
					else
							rs.Open SQLPROGRAM & " AND " & COLUMNNAME & " LIKE '%" & TRIM(SEARCHVALUE) & "%'" & " AND CALENDERID=" & CID, Conn
					end if
			ELSE
'					'response.write "<font size=5>XXXXX</font>" & rsTYPEID & columnname & searchvalue
					if rsTYPEID=135 then
						rs.Open SQLPROGRAM & " AND datepart(dd, " & COLUMNNAME & ") = datepart(dd, '" & TRIM(SEARCHVALUE) & "')"  & " AND datepart(mm, " & COLUMNNAME & ") = datepart(mm, '" & TRIM(SEARCHVALUE) & "')"  & " AND datepart(yy, " & COLUMNNAME & ") = datepart(yy, '" & TRIM(SEARCHVALUE) & "')", Conn
					else
						rs.Open SQLPROGRAM & " AND " & COLUMNNAME & " LIKE '%" & TRIM(SEARCHVALUE) & "%'", Conn
					end if
			END IF
	END IF
    '-----------------------
if not rs.EOF then
	numpages = rs.PageCount
	numrecs = rs.RecordCount
	rs.AbsolutePage = mypage
'--------------------------------------

'-------------------------------------
	'GenerateTable rs, pagesize
	%>

	<!-- insert control tag -->
	<span id="myGrid"></span>
<script>
	<%
	GenerateTablemac rs, DID, UID, pagesize
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
Response.Write( "<table width=100% border=0 >" )

IF FADD THEN
  Response.Write("<TD ALIGN=LEFT><A HREF=FORMADDNEW.ASP?UID=" & UID & "&DID=" & DID & "&CID=" & CID & " ><img height=15 src=add.jpg border=0 alt='Add Document'></A></TD>")
ELSE
  Response.Write("<TD></TD>" )
END IF
	Response.Write("<td align=right><font face=arial SIZE=1>" & numrecs & " Documents found.</td><td align=right ><font face=arial SIZE=1><i>Register Index Page " & mypage & " of " & numpages & " </i></td></table>" )
%>
</td>
</tr><tr>
<td valign=top bgcolor=white width=40% >
<%
Response.Write( "<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF >" )
Response.Write( "<TR>" )
Response.Write( "<TD valign=middle bgcolor=#DCDDDE bordercolordark=#FFFFFF bordercolor=#808080 bordercolorlight=#808080>&nbsp;&nbsp;<FONT FACE=ARIAL SIZE=2><img src=60-60-60.gif border=0 alt='SoftRobot Folder Dashboard'><B> Folder Dashboard</B></FONT></TD>" )
Response.Write( "</TR>" )
Response.Write( "</TABLE>" )
%>
<!---------------------Folder Reports Starts ------------------->
<div class="tab-pane" id="Div1">
<div class="tab-page"><h2 class="tab">Details</h2>
<iframe height=100% width=100% src=macstyle/blank.htm name=news>

</iframe>
</div>
<!--------------------------------------------------------------------------------->
<div class="tab-page"><h2 class="tab">Search</h2>
<form name=form1234 action=document.asp method=post>
<font face=arial size=1>
<INPUT TYPE=HIDDEN NAME=DID VALUE=<%=DID%> >
<INPUT TYPE=HIDDEN NAME=UID VALUE=<%=UID%> >
<INPUT TYPE=HIDDEN NAME=FADD VALUE=<%=FADD%> >
<INPUT TYPE=HIDDEN NAME=FDEL VALUE=<%=FDEL%> >
<INPUT TYPE=HIDDEN NAME=FEDIT VALUE=<%=FEDIT%> >
<INPUT TYPE=HIDDEN NAME=FVIEW VALUE=<%=FVIEW%> >
<INPUT TYPE=HIDDEN NAME=FILTER VALUE=<%=FFILTER%> >
<INPUT TYPE=HIDDEN NAME=FOFFHOLD VALUE=<%=FOFFHOLD%> >
<INPUT TYPE=HIDDEN NAME=FREJECT VALUE=<%=FREJECT%> >
<INPUT TYPE=HIDDEN NAME=CID VALUE=<%=CID%> >
<INPUT TYPE=HIDDEN NAME=YEARFLAG VALUE=<%=YEARFLAG%> >
<b>&nbsp;&nbsp;1. Search Reports: Select<br>
Text :<input type=text name=txtsearch size=10 value=<%=searchvalue%> ><br>
or Sort
<select name=SORT >
<OPTION Value=0>None
<OPTION Value=1>Asc
<OPTION Value=2>Desc
</select><br>
or Goto page no:
<select name=PAGE >
<%for pg=1 to numpages%>
<OPTION VALUE=<%=pg%>><%=pg%>
<%next%>
</select>
<br>
on field:<Select name=ls >
<%LISTFIELDS rs%>
</select> 
</font>
</b>
<br><input type=submit src=go.gif name=go Value="Search" >
</form>
<font face=arial size=1>
<b>(NOTE: In case of Date search provide with whole date: eg. 13/08/2005)</b></font>
<hr>
</div>
<!--------------------------------------------------------------------------------->
<div class="tab-page"><h2 class="tab">Matrix</h2>
<form name=form138 action=RGrid.asp?DID=<%=DID%>  method=post target=newwidow>
<INPUT TYPE=HIDDEN NAME=UID VALUE=<%=UID%> >
<font face=arial size=1>
<b>2. Matrix Reports: Select</b><br>
Date :&nbsp;
<Select name=RDate >
<%RLISTFIELDS rs%>
</select>
<br> Variable:&nbsp;
<Select name=RFld >
<%RTLISTFIELDS rs%>
</select>
On basis of:&nbsp;<br>
<input type=RADIO name="Rfilter" checked value="W" >Weekly
<input type=RADIO name="Rfilter" value="M" >Monthly
&nbsp;&nbsp;&nbsp;
<br><input type=submit src=go.gif name=graphreg Value="Search" >
</font>
</form>
<font face=arial size=1><b>(NOTE: In case of non availability of date field you cannot do this search)</b></font>
<hr>
</div>
<!--------------------------------------------------------------------------------->
<div class="tab-page"><h2 class="tab">Range</h2>
<form name=form11 action=printreg.asp method=post target=newwindow>
<font face=arial size=1>
<INPUT TYPE=HIDDEN NAME=DID VALUE=<%=DID%> >
<INPUT TYPE=HIDDEN NAME=UID VALUE=<%=UID%> >
<INPUT TYPE=HIDDEN NAME=FADD VALUE=<%=FADD%> >
<INPUT TYPE=HIDDEN NAME=FDEL VALUE=<%=FDEL%> >
<INPUT TYPE=HIDDEN NAME=FEDIT VALUE=<%=FEDIT%> >
<INPUT TYPE=HIDDEN NAME=FVIEW VALUE=<%=FVIEW%> >
<INPUT TYPE=HIDDEN NAME=FILTER VALUE=<%=FFILTER%> >
<INPUT TYPE=HIDDEN NAME=FOFFHOLD VALUE=<%=FOFFHOLD%> >
<INPUT TYPE=HIDDEN NAME=FREJECT VALUE=<%=FREJECT%> >
<INPUT TYPE=HIDDEN NAME=CID VALUE=<%=CID%> >
<INPUT TYPE=HIDDEN NAME=YEARFLAG VALUE=<%=YEARFLAG%> >
<b>3.Range Reports: Select<br>
&nbsp;&nbsp;From:<input type=text name=FROMTXT size=10 ><br>
&nbsp;&nbsp;To:<input type=text name=TOTXT size=10 ><br>
on:<Select name=ls1 ID="Select6">
<%LISTFIELDSRANGE rs%>
</select><br>
&nbsp;&nbsp;PageBy:<Select name=ls2 ID="Select7"><%LISTFIELDSPAGE rs%></select><br>
<br><input type=submit src=go.gif name=printreg Value="Search" >
</FORM>
<br>
<br>
(NOTE: In case of Date search provide whole dates e.g. 11/2/2003 OR 11 Feb 2003)
<hr>
</div>
<!--------------------------------------------------------------------------------->
<div class="tab-page"><h2 class="tab">Graphic</h2>
<!---
<form name=form2 action=tchart/tchart2.asp?mode=2 method=post target=newwindow>
-->

<form name=form2 action=http://localhost/dotnetcharting/mycharts/folderchart.aspx method=post target=newwindow >
<font face=arial size=1>
<b>4. Graphical Reports: Select<br>
&nbsp;&nbsp;X=:<Select name=XAXIS ID="XAXIS"><%RLISTTXTFIELDS rs%></select>(Name)<br>
&nbsp;&nbsp;Y=:<Select name=YAXIS ID="YAXIS"><%RLISTNUMFIELDS rs%></select>(Nos)<br>
<INPUT TYPE=HIDDEN NAME=TABLENM VALUE=<%=TABLENM%> >
<INPUT TYPE=HIDDEN NAME=CN VALUE=<%=Session("erp_ConnectionString")%> >
<br><input type=Submit src=go.gif name=graphreg Value="Search" >
</form>
<br>
<br>
<br>
<br>
(NOTE: Only select Numberic and String Fields. Avoid Linked Fields, Dates & Bits to generate graphs)
</font>
</b>
<hr>
</div>
<!--------------------------------------------------------------------------------->
<div class="tab-page"><h2 class="tab">Custom</h2>
<form name=form111 action=printreg1.asp  method=post target=newwindow>
<font face=arial size=1>
<INPUT TYPE=HIDDEN NAME=DID VALUE=<%=DID%> >
<INPUT TYPE=HIDDEN NAME=UID VALUE=<%=UID%> >
<INPUT TYPE=HIDDEN NAME=FADD VALUE=<%=FADD%> >
<INPUT TYPE=HIDDEN NAME=FDEL VALUE=<%=FDEL%> >
<INPUT TYPE=HIDDEN NAME=FEDIT VALUE=<%=FEDIT%> >
<INPUT TYPE=HIDDEN NAME=FVIEW VALUE=<%=FVIEW%> >
<INPUT TYPE=HIDDEN NAME=FILTER VALUE=<%=FFILTER%> >
<INPUT TYPE=HIDDEN NAME=FOFFHOLD VALUE=<%=FOFFHOLD%> >
<INPUT TYPE=HIDDEN NAME=FREJECT VALUE=<%=FREJECT%> >
<INPUT TYPE=HIDDEN NAME=CID VALUE=<%=CID%> >
<INPUT TYPE=HIDDEN NAME=YEARFLAG VALUE=<%=YEARFLAG%> >
<B>5.Customized Reports: 
Select </B> <br>
<Select name=ls1 >
<%LISTFIELDS rs%>
</select>
&nbsp;&nbsp;
<select name=BOOLEANID1 >
<OPTION Value='='>=
<OPTION Value='>'>>
<OPTION Value='<'><
<OPTION Value='>='>=
<OPTION Value='<='><=
<OPTION Value='<>'><>
</SELECT>
&nbsp;&nbsp;
<input type=text name=VALUE1 VALUE=0 size=10 > <BR>

<select name=LOGICALID1 >
<OPTION Value=0>NONE
<OPTION Value=1>AND
<OPTION Value=2>OR
</select>


<Select name=ls2 >
<%LISTFIELDS rs%>
</select>
<select name=BOOLEANID2 >
<OPTION Value='='>=
<OPTION Value='>'>>
<OPTION Value='<'><
<OPTION Value='>='>=
<OPTION Value='<='><=
<OPTION Value='<>'><>
</SELECT>
&nbsp;&nbsp;
<input type=text name=VALUE2 size=10 VALUE=0 > <BR>

<select name=LOGICALID2 >
<OPTION Value=0>NONE
<OPTION Value=1>AND
<OPTION Value=2>OR
</select>
on:<Select name=ls3 >
<%LISTFIELDS rs%>
</select>
<select name=BOOLEANID3 >
<OPTION Value='='>=
<OPTION Value='>'>>
<OPTION Value='<'><
<OPTION Value='>='>=
<OPTION Value='<='><=
<OPTION Value='<>'><>
</SELECT>
&nbsp;&nbsp;
<input type=text name=VALUE3 size=10 VALUE=0 > 
<br><input type=SUBMIT src=go.gif name=printreg Value="Search" >
</form>
<br>
<br>
(NOTE: In case of Date search provide whole dates e.g. 11/2/2003 OR 11 Feb 2003)
<hr>
</font>
</b>
</div>
<!--------------------------------------------------------------------------------->
<div class="tab-page"><h2 class="tab">Excel</h2>
<font face=arial size=1>
<b>6. Special Tools: Select<br>
A. <a href=reports/pivot.asp?TABLENM=<%=TABLENM%>REG WIDTH=99% height=400 target=newwindow>Create Pivot Reports</a><br>
B. <a href=reports/exceldata.asp?TABLENM=<%=TABLENM%>REG WIDTH=99% height=400 target=newwindow>Apply Excel Filters</a><br>
<%IF UID=83 OR UID=104 OR UID=57 THEN%>
C. <a href=reports/summary.asp?TABLENM=<%=TABLENM%>&DID=<%=DID%>&UID=<%=UID%> WIDTH=99% target=newindow>Folder Summary Report</a><br>
D. <a href=utilities/deletefolder.asp?TABLENM=<%=TABLENM%>&DID=<%=DID%>&UID=<%=UID%> WIDTH=99% target=newindow>Delete All Data in Folder</a><br>
<%END IF%>
<br>
(NOTE: In case of Delete Function all data from main and detail tables will be deleted)
<hr>
</font>
</div>
<!--------------------------------------------------------------------------------->

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
<hr>
<font face="Arial" size="1">
&#169; Copyright 2007 . All rights reserved. IRP, ERPWEB, MOBILEERP, SoftRobot, SoftServer
is property of MobileERP Softech P Ltd. India, Malaysia, UK, USA</font>
<script type="text/javascript">
setupAllTabs();
</script>
</BODY>
</HTML>