<!--
'************************************************************************
'Pupose						:	This is a SoftRobot secure document server
'Filename					:	sdocument.asp
'Author						:	Anita Shah
'Created					:	27-Mar-2001
'Project Name				:	IRP
'Contact					:	anita@IRP
'
'Modification History		:	There is a non secure document as well
'Purpose					:	Refer quick document entry at document.asp
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
FADD=REQUEST("FADD"): IF FADD THEN FADD=1 ELSE FADD=0
FDEL=REQUEST("FDEL"): IF FDEL THEN FDEL=1 ELSE FDEL=0
FVIEW=REQUEST("FVIEW"): IF FVIEW THEN FVIEW=1 ELSE FVIEW=0
FEDIT=REQUEST("FEDIT"): IF FEDIT THEN FEDIT=1 ELSE FEDIT=0
FFILTER=REQUEST("FILTER"): : IF FFILTER THEN FFILTER=1 ELSE FFILTER=0
FOFFHOLD=REQUEST("FOFFHOLD"): IF FOFFHOLD THEN FOFFHOLD=1 ELSE FOFFHOLD=0
FREJECT=REQUEST("FREJECT"): IF FREJECT THEN FREJECT=1 ELSE FREJECT=0
UID=REQUEST("UID")
FLAG=0
COLUMNNAME=REQUEST("ls")
SEARCHVALUE=REQUEST("txtsearch")
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
<HTML>
<HEAD>
<TITLE>IRP DOCUMENT (sales@erpweb)</TITLE>
<LINK REL=STYLESHEET TYPE="text/css" HREF="<%=PROFILECSS%>">
</HEAD>
<BODY topmargin=0>
<Basefont face=arial size=1>

<%
'----------------------------------LISTBOX GENERATOR
Sub LIST1 ( LISTTABLE, LISTCOLUMN, LISTVALUE, Conn, name, value )
Set lst=Server.CreateObject("ADODB.Recordset")
lst.Open "Select " & LISTVALUE & ", " & LISTCOLUMN & " FROM " & LISTTABLE & " WHERE " & LISTVALUE & " = " & value, Conn
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
Sub GenerateTable( rs, pagesize )
	
  Response.Write( "<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF>" )
   '--------------------------------------
  ' set up column names
  for i = 3 to rs.fields.count - 1
        if (rs(i).Type = 3) and (i > 3)then
        Set ls=Server.CreateObject("ADODB.Recordset")
		ls.Open "Select * from LISTBOX where DETFLAG=0 AND DDID=" & DDID & " AND COLUMNNO=" & i, Conn
		if not ls.EOF then
		LISTNAME=ls("LISTNAME") 
        Response.Write("<TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080><FONT FACE=ARIAL SIZE=1 >" + LISTNAME + "</FONT></TD>")
        ELSE
        Response.Write("<TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080><FONT FACE=ARIAL SIZE=1 >" + rs(i).Name + "</FONT></TD>")
        end if
        ls.Close
        set ls=nothing
        ELSE
        Response.Write("<TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080 align=right><FONT FACE=ARIAL SIZE=1 COLOR=WHITE>" + rs(i).Name + "</FONT></TD>")
        end if
    'Response.Write("<TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080><FONT FACE=ARIAL SIZE=1 >" + rs(i).Name + "</FONT></TD>")
  next
  Response.Write("<TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080><FONT FACE=ARIAL SIZE=1 ><B>Tools</B></FONT></TD>")
  
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
        <TD VALIGN=TOP BGCOLOR=#cccccc align=right><FONT FACE=ARIAL SIZE=1><%=v%></font></td>
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
		<TD VALIGN=TOP BGCOLOR=#cccccc align=right><FONT FACE=ARIAL SIZE=1>
		<%LIST1 LISTTABLE, LISTCOLUMN, LISTVALUE, Conn, name, v%> 
		</font></td>
		<%ELSE%>
		<TD VALIGN=TOP BGCOLOR=#eeeeee align=right><FONT FACE=ARIAL SIZE=1>
		<%LIST1 LISTTABLE, LISTCOLUMN, LISTVALUE, Conn, name, v%> 
		</font></td>
		<%END IF%>
		<%
	else
		IF FLAG THEN
			Response.Write( "<TD VALIGN=TOP BGCOLOR=#cccccc align=right><FONT FACE=ARIAL SIZE=1>" + CStr( v ) + "</FONT></TD>" )
		   ELSE
		    Response.Write( "<TD VALIGN=TOP BGCOLOR=#eeeeee align=right><FONT FACE=ARIAL SIZE=1>" + CStr( v ) + "</FONT></TD>" )
		END IF
		end if
		ls.Close
		set ls=nothing
		'-----------------------------------------
    elseif rs(i).name = "EMAILNO" then
    '-----------------------
           IF FLAG THEN
			Response.Write( "<TD VALIGN=TOP BGCOLOR=#cccccc><FONT FACE=ARIAL SIZE=1><a href=mailto:" & CStr( v ) & ">" & CStr( v ) & "</a></FONT></TD>" )
		   ELSE
		    Response.Write( "<TD VALIGN=TOP BGCOLOR=#eeeeee><FONT FACE=ARIAL SIZE=1><a href=mailto:" & CStr( v ) & ">" & CStr( v ) & "</a></FONT></TD>" )
		   END IF
	 '---------------------
	 elseif rs(i).name = "WEBSITE" then
     '-----------------------
           IF FLAG THEN
			Response.Write( "<TD VALIGN=TOP BGCOLOR=#cccccc><FONT FACE=ARIAL SIZE=1><a href=http://" & CStr( v ) & " target=new>" & CStr( v ) & "</a></FONT></TD>" )
		   ELSE
		    Response.Write( "<TD VALIGN=TOP BGCOLOR=#eeeeee><FONT FACE=ARIAL SIZE=1><a href=http://" & CStr( v ) & " target=new>" & CStr( v ) & "</a></FONT></TD>" )
		   END IF
	 '---------------------
	 elseif rs(i).name = "PHONENO" then
     '-----------------------
           IF FLAG THEN
			Response.Write( "<TD VALIGN=TOP BGCOLOR=#cccccc><FONT FACE=ARIAL SIZE=1><a href=file:///C:/WINDOWS/DIALER.EXE target=new>" & CStr( v ) & "</a></FONT></TD>" )
		   ELSE
		    Response.Write( "<TD VALIGN=TOP BGCOLOR=#eeeeee><FONT FACE=ARIAL SIZE=1><a href=file:///C:/WINDOWS/DIALER.EXE target=new>" & CStr( v ) & "</a></FONT></TD>" )
		   END IF
    '-----------------------------------------
   		else
   		 IF FLAG THEN
			Response.Write( "<TD VALIGN=TOP BGCOLOR=#cccccc align=right><FONT FACE=ARIAL SIZE=1>" + CStr( v ) + "</FONT></TD>" )
		   ELSE
		    Response.Write( "<TD VALIGN=TOP BGCOLOR=#eeeeee align=right><FONT FACE=ARIAL SIZE=1>" + CStr( v ) + "</FONT></TD>" )
		END IF  
    end if
    
        '-----------------------
           'IF FLAG THEN
			'Response.Write( "<TD VALIGN=TOP BGCOLOR=#cccccc><FONT FACE=ARIAL SIZE=1>" + CStr( v ) + "</FONT></TD>" )
		   'ELSE
		    'Response.Write( "<TD VALIGN=TOP BGCOLOR=#eeeeee><FONT FACE=ARIAL SIZE=1>" + CStr( v ) + "</FONT></TD>" )
		   'END IF
	   '---------------------	   
    next
      '-----------------------------
      Response.Write("<TD BGCOLOR=#6699CC><FONT FACE=ARIAL SIZE=1>")
  IF FVIEW THEN
  Response.Write("<A HREF=FORMVIEW.ASP?DID=" & DID & "&ID=" & rs(3) & "&UID=" & UID & " target=news><img src=print.gif border=0 width=20 alt='Show Document'></A>")
  END IF
  IF FEDIT THEN
  Response.Write("<A HREF=FORMEDIT.ASP?DID=" & DID & "&ID=" & rs(3) & "&UID=" & UID & " target=news><img src=update.gif border=0 width=20 alt='Edit Document Style 1'></A>")
  END IF
  IF FDEL THEN
  Response.Write("<A HREF=FORMDEL.ASP?DID=" & DID & "&ID=" & rs(3) & "&UID=" & UID & " target=news><img src=delete.gif border=0 width=20 alt='Delete Document'></A>")
  END IF
  IF FFILTER THEN
  Response.Write("<A HREF=FORMAPPROVE.ASP?DID=" & DID & "&ID=" & rs(3) & "&UID=" & UID & " target=news><img src=approve.gif border=0 width=20 alt='Approve Document'></A>")
  END IF
  IF FOFFHOLD THEN
  Response.Write("<A HREF=FORMOFFHOLD.ASP?DID=" & DID & "&ID=" & rs(3) & "&UID=" & UID & " target=news><img src=approve.gif border=0 width=20 alt='OffHold Document'></A>")
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
Sub GenerateHeader( rs )
Response.Write( "<TABLE BORDER=0 WIDTH=100% >" )
Response.Write( "<TR BGCOLOR=#6699cc>" )
Response.Write( "<TD><FONT FACE=ARIAL SIZE=3><img src=ofolder.gif width=20 border=0 alt='SoftRobot Document Server'><B> " & rs("TITLE") & " Register</B></FONT></TD>" )
Response.Write( "</TR>" )
Response.Write( "<TR BGCOLOR=YELLOW>" )
Response.Write( "<TD><FONT FACE=ARIAL SIZE=1><B>" & rs("HEADERNOTE") & "</B></FONT></TD>" )
Response.Write( "</TR>" )
Response.Write( "</TABLE>" )
End Sub
%>
<%
Sub GenerateFooter( rs )
Response.Write( "<TABLE BORDER=0 WIDTH=100% >" )
Response.Write( "<TR BGCOLOR=YELLOW>" )
Response.Write( "<TD WIDTH=80% ><FONT FACE=ARIAL SIZE=1><B>Rules:" & rs("FOOTERNOTE") & "</B></FONT></TD><TD ALIGN=RIGHT><FONT FACE=ARIAL SIZE=1><I>" & NOW & "</I></FONT></TD>" )
Response.Write( "</TR>" )
Response.Write( "<TR BGCOLOR=#6699CC>" )
IF FADD THEN
  Response.Write("<TD WIDTH=80% ></TD><TD ALIGN=RIGHT><FONT FACE=ARIAL SIZE=2><A HREF=FORMADD.ASP?UID=" & UID & "&DID=" & DID & " target=news><img src=add.gif border=0 alt='Add Document'></A></FONT></TD>")
ELSE
  Response.Write("<TD WIDTH=80% ></TD><TD></TD>" )
END IF
Response.Write( "</TR>" )
Response.Write( "</TABLE>" )
End Sub
%>

<%
'-----------------------------------main program starts here
Dim mypage, numpages
Dim numrecs, pagesize
'pagesize = CInt( Request("recs") )
'If pagesize = 0 Then pagesize = 2
pagesize=15
mypage = TRIM(Request("PAGE") )
If mypage="" Then mypage=1 
'------------------------------------find doc sql
CID=REQUEST("CID")
DID=CInt(REQUEST("DID"))
SA=REQUEST("sa")
SD=REQUEST("sd")
'RESPONSE.WRITE "VAL=" & SA
if DID="" Then Response.Write "ERROR:DID":Response.END
set rsDOC = Server.CreateObject("ADODB.Recordset")
rsDOC.Open "Select * from DOCUMENTS WHERE DID=" & DID, Conn 
SQLPROGRAM=trim(rsDOC("SQLPROGRAM"))
DDID=rsDOC("DDID")
DOCID=TRIM(rsDOC("MASTERTABLE")) & "ID"
YEARFLAG=0
IF rsDOC("YEARFILTER") THEN YEARFLAG=1 ELSE YEARFLAG=0
'------------------------------------
GenerateHeader rsDoc
'-------------------------------------
    set rs = Server.CreateObject("ADODB.Recordset")
	rs.PageSize = pagesize
	rs.CacheSize = pagesize
	rs.CursorLocation = adUseClient
	'------------------------FILTER DATA
	IF SEARCHVALUE = "" THEN
        	IF DID=385 OR DID=384 OR DID=390 OR DID=391 THEN
				rs.Open SQLPROGRAM , Conn
			ELSEIF SA="SortASC" THEN
				'RESPONSE.WRITE "FILTER DATA"
				rs.Open SQLPROGRAM & " ORDER BY " & COLUMNNAME & " ASC", Conn
			ELSEIF SD="SortDESC" THEN
				'RESPONSE.WRITE "FILTER DATA"
				rs.Open SQLPROGRAM & " ORDER BY " & COLUMNNAME & " DESC", Conn
			ELSE
				IF YEARFLAG THEN
					IF SD="SortDESC" THEN
					rs.Open SQLPROGRAM & " AND CALENDERID=" & CID & " ORDER BY " & COLUMNNAME & " DESC", Conn
					ELSEIF SA="SortASC" THEN
					rs.Open SQLPROGRAM & " AND CALENDERID=" & CID & " ORDER BY " & COLUMNNAME & " ASC", Conn
					ELSE 
					rs.Open SQLPROGRAM & " AND CALENDERID=" & CID, Conn
					END IF
				ELSE
				rs.Open SQLPROGRAM & " ORDER BY " & DOCID, Conn
				END IF
        	END IF
	ELSE
	rs.Open SQLPROGRAM & " AND " & COLUMNNAME & " LIKE '%" & TRIM(SEARCHVALUE) & "%'", Conn
    'Response.WRITE SQLPROGRAM & " AND " & COLUMNNAME & " LIKE '%" & TRIM(SEARCHVALUE) & "%'"
    END IF
    '-----------------------
if not rs.EOF then
	numpages = rs.PageCount
	numrecs = rs.RecordCount
	rs.AbsolutePage = mypage
'--------------------------------------
	Response.Write( "<table width=100% ><td align=left><font face=arial SIZE=1>" & numrecs & " Documents found.</td>" )
	Response.Write("<td align=right><font face=arial SIZE=1><i>Register Index Page " & mypage & " of " & numpages & " </i></td></table>" )
'-------------------------------------
	GenerateTable rs, pagesize
'------------------------------------	
else'-----if rs not found
	Response.Write "<font face=arial size=2 color=red>No Documents Found</font><br>"
end if
'------------------------------------
GenerateFooter rsDoc
'------------------------------------
%>
<form name=form1 action=sdocument.asp method=post>
<table width=100%>
<tr bgcolor=blue><td><font size=3 color=yellow><b>Search</b></font></td></tr>
<tr><td>
<font face=arial size=2>
<INPUT TYPE=HIDDEN NAME=DID VALUE=<%=DID%>>
<INPUT TYPE=HIDDEN NAME=UID VALUE=<%=UID%>>
<INPUT TYPE=HIDDEN NAME=FADD VALUE=<%=FADD%>>
<INPUT TYPE=HIDDEN NAME=FDEL VALUE=<%=FDEL%>>
<INPUT TYPE=HIDDEN NAME=FEDIT VALUE=<%=FEDIT%>>
<INPUT TYPE=HIDDEN NAME=FVIEW VALUE=<%=FVIEW%>>
<INPUT TYPE=HIDDEN NAME=FILTER VALUE=<%=FFILTER%>>
<INPUT TYPE=HIDDEN NAME=FOFFHOLD VALUE=<%=FOFFHOLD%>>
<INPUT TYPE=HIDDEN NAME=FREJECT VALUE=<%=FREJECT%>>
<INPUT TYPE=HIDDEN NAME=CID VALUE=<%=CID%>>
<INPUT TYPE=HIDDEN NAME=YEARFLAG VALUE=<%=YEARFLAG%>>
<img src=info.gif border=0 width=20 alt="Document info">Search:<input type=text name=txtsearch>
in Column:<Select name=ls>
<%LISTFIELDS rs%>
</select>
</font>
<input type=submit name=go Value="Search">
</td></tr>
<tr bgcolor=blue><td><b><font size=3 color=yellow>Sort & Navigate:</font></b></td></tr>
<tr><td>
<img src=info.gif border=0 width=20 alt="Document info"><font face=arial size=2>Select Page Number:</font>
<select name=PAGE>
<%for pg=1 to numpages%>
<OPTION VALUE=<%=pg%>><%=pg%>
<%next%>
</SELECT>
<input type=submit name=go1 VALUE=Go>
<img src=grydown.gif border=0 width=20 alt="Document info">
<input type=SUBMIT name=sa VALUE="SortASC" src=gryup.gif border=0 width=20 alt="Sort Ascending">
<img src=gryup.gif border=0 width=20 alt="Document info">
<input type=SUBMIT name=sd VALUE="SortDESC" src=grydown.gif border=0 width=20 alt="Sort Descending">
<input type=image name=refresh src=home.gif border=0 width=20 alt="Refresh after every add, edit & delete">
<img src=hrefresh.gif border=0 alt="Refresh after every add, edit & delete">
</td></tr>
</table>
</form>
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
&#169; Copyright 2001 . SoftRobot Secure Document Server
is property of ERPWEB.</font>
</BODY>
</HTML>
