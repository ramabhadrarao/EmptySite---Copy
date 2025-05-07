<%' @TRANSACTION=Required LANGUAGE="VBScript" %>
<%'@ LANGUAGE="VBScript" %>
<%

'Lastly Modified by :: Gaurav (3rd March, 2006)
'Lastly Modified by :: Gaurav (3rd Sept, 2006)

'Option Explicit
Response.Buffer = True
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
%>

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
Sub GenerateTable( rs, pagesize, LINKADDRESS, CID )
	
  Response.Write( "<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF background=images/bg6.gif>" )
   '--------------------------------------
  ' set up column names
  for i = 3 to rs.fields.count - 1
        if (rs(i).Type = 3) and (i > 3)then
        Set ls=Server.CreateObject("ADODB.Recordset")
		ls.Open "Select * from LISTBOX where DETFLAG=0 AND DDID=" & DDID & " AND COLUMNNO=" & i, Conn
		if not ls.EOF then
		LISTNAME=ls("LISTNAME") 
        Response.Write("<TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080 align=middle><FONT FACE=ARIAL SIZE=1 >" + LISTNAME + "</FONT></TD>")
        ELSE
        Response.Write("<TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080 align=middle><FONT FACE=ARIAL SIZE=1 >" + rs(i).Name + "</FONT></TD>")
        end if
        ls.Close
        set ls=nothing
        ELSE
        Response.Write("<TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080 align=middle><FONT FACE=ARIAL SIZE=1 >" + rs(i).Name + "</FONT></TD>")
        end if
    'Response.Write("<TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080><FONT FACE=ARIAL SIZE=1 >" + rs(i).Name + "</FONT></TD>")
  next
  Response.Write("<TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080 align=middle><FONT FACE=ARIAL SIZE=1 ><B>Tools</B></FONT></TD>")
  
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
        <TD VALIGN=TOP bgcolor=#E1F2FD align=left><FONT FACE=ARIAL SIZE=1><%=v%></font></td>
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
		<TD VALIGN=TOP bgcolor=#E1F2FD align=left><FONT FACE=ARIAL SIZE=1>
		<%LIST1 LISTTABLE, LISTCOLUMN, LISTVALUE, Conn, name, v%> 
		</font></td>
		<%ELSE%>
		<TD VALIGN=TOP bgcolor=#E1F2Ff align=left><FONT FACE=ARIAL SIZE=1>
		<%LIST1 LISTTABLE, LISTCOLUMN, LISTVALUE, Conn, name, v%> 
		</font></td>
		<%END IF%>
		<%
	else
		IF FLAG THEN
			Response.Write( "<TD VALIGN=TOP bgcolor=#E1F2FD align=left><FONT FACE=ARIAL SIZE=1>" + CStr( v ) + "</FONT></TD>" )
		   ELSE
		    Response.Write( "<TD VALIGN=TOP bgcolor=#E1F2Ff align=left><FONT FACE=ARIAL SIZE=1>" + CStr( v ) + "</FONT></TD>" )
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
		    Response.Write( "<TD VALIGN=TOP bgcolor=#E1F2Ff><FONT FACE=ARIAL SIZE=1><a href=mailto:" & CStr( v ) & ">" & CStr( v ) & "</a></FONT></TD>" )
		   END IF
	 '---------------------
	 elseif rs(i).name = "WEBSITE" then
     '-----------------------
           IF FLAG THEN
			Response.Write( "<TD VALIGN=TOP bgcolor=#E1F2FD><FONT FACE=ARIAL SIZE=1><a href=http://" & CStr( v ) & " target=news>" & CStr( v ) & "</a></FONT></TD>" )
		   ELSE
		    Response.Write( "<TD VALIGN=TOP bgcolor=#E1F2Ff><FONT FACE=ARIAL SIZE=1><a href=http://" & CStr( v ) & " target=news>" & CStr( v ) & "</a></FONT></TD>" )
		   END IF
	 '---------------------
	 elseif rs(i).name = "PHONENO" then
     '-----------------------
           IF FLAG THEN
			Response.Write( "<TD VALIGN=TOP bgcolor=#E1F2FD><FONT FACE=ARIAL SIZE=1><a href=file:///C:/WINDOWS/DIALER.EXE target=news>" & CStr( v ) & "</a></FONT></TD>" )
		   ELSE
		    Response.Write( "<TD VALIGN=TOP bgcolor=#E1F2Ff><FONT FACE=ARIAL SIZE=1><a href=file:///C:/WINDOWS/DIALER.EXE target=news>" & CStr( v ) & "</a></FONT></TD>" )
		   END IF
    '-----------------------------------------
   		else
   		 IF FLAG THEN
			Response.Write( "<TD VALIGN=TOP bgcolor=#E1F2FD align=left><FONT FACE=ARIAL SIZE=1>" + CStr( v ) + "</FONT></TD>" )
		   ELSE
		    Response.Write( "<TD VALIGN=TOP bgcolor=#E1F2Ff align=left><FONT FACE=ARIAL SIZE=1>" + CStr( v ) + "</FONT></TD>" )
		END IF  
    end if
    
        '-----------------------
           'IF FLAG THEN
			'Response.Write( "<TD VALIGN=TOP bgcolor=#E1F2FD><FONT FACE=ARIAL SIZE=1>" + CStr( v ) + "</FONT></TD>" )
		   'ELSE
		    'Response.Write( "<TD VALIGN=TOP bgcolor=#E1F2Ff><FONT FACE=ARIAL SIZE=1>" + CStr( v ) + "</FONT></TD>" )
		   'END IF
	   '---------------------	   
    next
      '-----------------------------
      Response.Write("<TD bgcolor=#E1F2FD align=left><FONT FACE=ARIAL SIZE=1>")
      
      IF ISNULL(LINKADDRESS) OR LINKADDRESS="" THEN
      ELSE
  
  Response.Write("<A HREF=" & LINKADDRESS & "?CID=" & CID & "&DID=" & DID & "&ID=" & rs(3) & "&UID=" & UID & " target=news1><img src=print.gif border=0 width=20 alt='Show Document'></A>")
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
Response.Write( "<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF >" )
Response.Write( "<TR BGCOLOR=#6699cc>" )
Response.Write( "<TD valign=middle bgcolor=#D4D4D4 bordercolordark=#FFFFFF bordercolor=#808080 bordercolorlight=#808080><FONT FACE=ARIAL SIZE=2><img src=ofolder.gif width=20 border=0 alt='SoftRobot Document Server'><B> " & rs("TITLE") & " Register</B></FONT></TD>" )
Response.Write( "</TR>" )
'Response.Write( "<TR BGCOLOR=YELLOW>" )
'Response.Write( "<TD><FONT FACE=ARIAL SIZE=1><B>" & rs("HEADERNOTE") & "</B></FONT></TD>" )
'Response.Write( "</TR>" )
Response.Write( "</TABLE>" )
End Sub
%>
<%
Sub GenerateFooter( rs, CID )
Response.Write( "<TABLE BORDER=0 WIDTH=100% >" )
'Response.Write( "<TR BGCOLOR=YELLOW>" )
'Response.Write( "<TD WIDTH=80% ><FONT FACE=ARIAL SIZE=1><B>Rules:" & rs("FOOTERNOTE") & "</B></FONT></TD><TD ALIGN=RIGHT><FONT FACE=ARIAL SIZE=1><I>" & NOW & "</I></FONT></TD>" )
'Response.Write( "</TR>" )
'Response.Write( "<TR BGCOLOR=#6699CC>" )
'IF FADD THEN
  'Response.Write("<TD WIDTH=80% ></TD><TD ALIGN=RIGHT><FONT FACE=ARIAL SIZE=2><A HREF=add" & LINKADDRESS & "?CID=" & CID & "&UID=" & UID & "&DID=" & DID & " target=news><img src=add.gif border=0 alt='Add Document'></A></FONT></TD>")
'ELSE
  'Response.Write("<TD WIDTH=80% ></TD><TD></TD>" )
'END IF
Response.Write( "</TR>" )
Response.Write( "</TABLE>" )
End Sub
%>
<HTML>
<HEAD>
<TITLE>SoftServer Linked DOCUMENT Server (sales@ERPWEB.com)</TITLE>
<script type="text/javascript" src="js/tabpane.js"></script>
<link type="text/css" rel="StyleSheet" href="css/tab.webfx.css" />
</HEAD>
<BODY topmargin=0 leftmargin=5>
<Basefont face=arial size=1>
<!-----------------------------HEADER STRIP--------------------> 

<!--------------------------------------------------------------->
<table width=100% border=0 ID="Table4">
<tr>
<td valign=top>
<%
'-----------------------------------main program starts here
Dim mypage, numpages
Dim numrecs, pagesize
'pagesize = CInt( Request("recs") )
'If pagesize = 0 Then pagesize = 2
pagesize=10
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
SQLPROGRAM=rsDOC("SQLPROGRAM"):IF ISNULL(SQLPROGRAM) OR SQLPROGRAM="" THEN RESPONSE.END
'RESPONSE.WRITE SQLPROGRAM
LINKADDRESS=rsDOC("LINKADDRESS")
DDID=rsDOC("DDID")
TABLENM=rsDOC("MASTERTABLE")
DOCID=TRIM(TABLENM) & "ID"
YEARFLAG=0
SORT=REQUEST("SORT")
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
			ELSE
						
				IF YEARFLAG THEN
					IF SORT=1 THEN
					rs.Open SQLPROGRAM & " AND CALENDERID=" & CID & " ORDER BY " & COLUMNNAME & " ASC", Conn
					ELSEIF SORT=2 THEN
					rs.Open SQLPROGRAM & " AND CALENDERID=" & CID & " ORDER BY " & COLUMNNAME & " DESC", Conn
					ELSE 
					rs.Open SQLPROGRAM & " AND CALENDERID=" & CID, Conn
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
					IF rsTYPEID=135 THEN
						IF NOT ISDATE(searchvalue) THEN RESPONSE.WRITE "<FONT FACE=ARIAL SIZE=1>ERROR ENTERING DATE</FONT>": RESPONSE.END
					END IF
					'-----------------------------
	
			'IF YEARFLAG THEN
			'	rs.Open SQLPROGRAM & " AND " & COLUMNNAME & " LIKE '%" & TRIM(SEARCHVALUE) & "%'" & " AND CALENDERID=" & CID, Conn
			'ELSE
			'	rs.Open SQLPROGRAM & " AND " & COLUMNNAME & " LIKE '%" & TRIM(SEARCHVALUE) & "%'", Conn
			'END IF 
			
			IF YEARFLAG THEN
					if rsTYPEID=135 then
		  					rs.Open SQLPROGRAM & " AND datepart(dd, " & COLUMNNAME & ") = datepart(dd, '" & TRIM(SEARCHVALUE) & "')"  & " AND datepart(mm, " & COLUMNNAME & ") = datepart(mm, '" & TRIM(SEARCHVALUE) & "')"  & " AND datepart(yy, " & COLUMNNAME & ") = datepart(yy, '" & TRIM(SEARCHVALUE) & "')" & " AND CALENDERID=" & CID, Conn
					else
							rs.Open SQLPROGRAM & " AND " & COLUMNNAME & " LIKE '%" & TRIM(SEARCHVALUE) & "%'" & " AND CALENDERID=" & CID, Conn
					end if
			ELSE
'					response.write "<font size=5>Rajat</font>" & rsTYPEID & columnname & searchvalue
					if rsTYPEID=135 then
						rs.Open SQLPROGRAM & " AND datepart(dd, " & COLUMNNAME & ") = datepart(dd, '" & TRIM(SEARCHVALUE) & "')"  & " AND datepart(mm, " & COLUMNNAME & ") = datepart(mm, '" & TRIM(SEARCHVALUE) & "')"  & " AND datepart(yy, " & COLUMNNAME & ") = datepart(yy, '" & TRIM(SEARCHVALUE) & "')" & " AND CALENDERID=" & CID, Conn
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
	Response.Write( "<table width=100% ><td align=left><font face=arial SIZE=1>" & numrecs & " Documents found.</td>" )
	Response.Write("<td align=right><font face=arial SIZE=1><i>Register Index Page " & mypage & " of " & numpages & " </i></td></table>" )
'-------------------------------------
	GenerateTable rs, pagesize, LINKADDRESS, CID
'------------------------------------	
else'-----if rs not found
	Response.Write "<font face=arial size=2 color=red>No Documents Found</font><br>"
end if
'------------------------------------
GenerateFooter rsDoc, CID
'------------------------------------
%>
</td>
</tr><tr>
<td width=40% valign=top>
<%
Response.Write( "<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF >" )
Response.Write( "<TR>" )
Response.Write( "<TD valign=middle bgcolor=#DCDDDE bordercolordark=#FFFFFF bordercolor=#808080 bordercolorlight=#808080>&nbsp;&nbsp;<FONT FACE=ARIAL SIZE=2><img src=60-60-60.gif border=0 alt='SoftRobot Document Server'><B> Folder Dashboard</B></FONT></TD>" )
Response.Write( "</TR>" )
Response.Write( "</TABLE>" )
%>

<% '---------->changes document.asp<--------------------------%>
<% '---------->New Tab Wise Search----------------------------%>
<div class="tab-pane" id="Div1">
<div class="tab-page"><h2 class="tab">Search</h2>

<form name=form123 action=lreport.asp method=post >
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
<!---------------------SEARCH STRING 2------------------->
</div>
<div class="tab-page"><h2 class="tab">Matrix</h2>

<form name=form138 action=RGrid.asp?DID=<%=DID%>  method=post target=news>
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
<div class="tab-page"><h2 class="tab">Range</h2>
<!---------------------SEARCH STRING 3------------------->
<form name=form11 action=printreg.asp method=post >
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
<%LISTFIELDS rs%>
</select>
<br><input type=submit src=go.gif name=printreg Value="Search" >
</FORM>
<br>
<br>
(NOTE: In case of Date search provide whole dates e.g. 11/2/2003 OR 11 Feb 2003)
<hr>
</div>
<div class="tab-page"><h2 class="tab">Graphic</h2>
<!---------------------SEARCH STRING------------------->
<form name=form2 action=tchart/tchart2.asp?mode=2 method=post target=news>
<font face=arial size=1>
<b>4. Graphical Reports: Select<br>
&nbsp;&nbsp;X=:<Select name=XAXIS ID="Select7"><%RLISTTXTFIELDS rs%></select>(Name)<br>
&nbsp;&nbsp;Y=:<Select name=YAXIS ID="Select8"><%RLISTNUMFIELDS rs%></select>(Nos)<br>
<INPUT TYPE=HIDDEN NAME=TABLENM VALUE=<%=TABLENM%> >
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
<div class="tab-page"><h2 class="tab">Custom</h2>
<!---------------------SEARCH STRING------------------->
<form name=form111 action=printreg1.asp  method=post >
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
A. <a href=reports/pivot.asp?TABLENM=<%=TABLENM%> WIDTH=99% height=400 target=news>Create Pivot Reports</a><br>
B. <a href=reports/exceldata.asp?TABLENM=<%=TABLENM%> WIDTH=99% height=400 target=news>Apply Excel Filters</a><br>
<br>
(NOTE: In case of error check whether you have installed MS Office Web Components or not?.)
<hr>
</font>
</div>
<!--------------------------------------------------------------------------------->
</td>
</tr>
</table>
<% '-----------------------------------End------------------------------------------%>
<!--form name=form123 action=lreport.asp method=post>
<table width=100% height=45 border=0>
<tr><td background=blueband.jpg valign=middle>
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
<b>&nbsp;&nbsp;Search:<input type=text name=txtsearch size=10 value=<%=searchvalue%>>
OR Sort
<select name=SORT>
<OPTION Value=0>None
<OPTION Value=1>Asc
<OPTION Value=2>Desc
</select>
on:<Select name=ls>
<%LISTFIELDS rs%>
</select>
Goto page:
<select name=PAGE>
<%for pg=1 to numpages%>
<OPTION VALUE=<%=pg%>><%=pg%>
<%next%>
</select>
</font>
</b>
<wbr><input type=image src=go.gif name=go Value="Search">
</td>
</tr>
</table>
</form>
<font face=arial size=2><b>(NOTE: In case of Date search provide with whole date: eg. 13/08/2005)</b></font>


<form name=form138 action=RGrid.asp?DID=<%=DID%> target=news method=post>
<INPUT TYPE=HIDDEN NAME=UID VALUE=<%=UID%>>

<table width=100% height=45 border=0>
<tr><td  bgcolor=#D5EAFF valign=middle>
<font face=arial size=2>
<b>1. Date Wise Scheduled Reports:<br>
</font>
</b>
Date Field:&nbsp;
<Select name=RDate>
<%RLISTFIELDS rs%>
</select>&nbsp;&nbsp;&nbsp;
Select Variable:&nbsp;
<Select name=RFld>
<%RTLISTFIELDS rs%>
</select>
<br>
On basis of:&nbsp;
<input type=RADIO name="Rfilter" checked value="W">Weekly
<input type=RADIO name="Rfilter" value="M">Monthly
&nbsp;&nbsp;&nbsp;
<wbr><input type=image src=go.gif name=graphreg Value="Show">

<hr>

</form>
<form name=form11 action=printreg.asp target=news method=post>
<table width=100% height=45 border=0>
<tr><td  bgcolor=#D5EAFF valign=middle>
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
<b>2.Range Reports:
&nbsp;&nbsp;From:<input type=text name=FROMTXT size=10>
&nbsp;&nbsp;To:<input type=text name=TOTXT size=10>
on:<Select name=ls1>
<%LISTFIELDS rs%>
</select>
<input type=image src=go.gif name=printreg Value="Search">
</FORM>
(NOTE: In case of Date search provide whole dates e.g. 11/2/2003 OR 11 Feb 2003)
<hr>

<form name=form23 action=tchart/tchart2.asp?mode=2 target=news method=post>
<table width=100% height=45 border=0>
<tr><td  bgcolor=#D5EAFF valign=middle>
<font face=arial size=2>
<b>3. Graphical Reports:
&nbsp;&nbsp;X=:<Select name=XAXIS><%RLISTTXTFIELDS rs%></select>(Name)
&nbsp;&nbsp;Y=:<Select name=YAXIS><%RLISTNUMFIELDS rs%></select>(Nos)
</font>
</b>
<INPUT TYPE=HIDDEN NAME=TABLENM VALUE=<%=TABLENM%>>
<input type=hidden name=RDID value=<%=DID%>>
<wbr><input type=image src=go.gif name=graphreg Value="Show">
</td>
</tr>
</table>
</form>
(NOTE: Only select Numberic and String Fields. Avoid Linked Fields, Dates & Bits to generate graphs)
<hr-->

<!---------------------SEARCH STRING------------------->
<!--form name=form111 action=printreg1.asp target=news method=post>
<table width=100% height=45 border=0>
<tr><td  bgcolor=#D5EAFF valign=middle>
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
<B>4.Customized Reports: Select </B> 
<Select name=ls1>
<%LISTFIELDS rs%>
</select>
&nbsp;&nbsp;
<select name=BOOLEANID1>
<OPTION Value='='>=
<OPTION Value='>'>>
<OPTION Value='<'><
<OPTION Value='>='>=
<OPTION Value='<='><=
<OPTION Value='<>'><>
</SELECT>
&nbsp;&nbsp;
<input type=text name=VALUE1 VALUE=0 size=10> <BR>

<select name=LOGICALID1>
<OPTION Value=0>NONE
<OPTION Value=1>AND
<OPTION Value=2>OR
</select>


<Select name=ls2>
<%LISTFIELDS rs%>
</select>
<select name=BOOLEANID2>
<OPTION Value='='>=
<OPTION Value='>'>>
<OPTION Value='<'><
<OPTION Value='>='>=
<OPTION Value='<='><=
<OPTION Value='<>'><>
</SELECT>
&nbsp;&nbsp;
<input type=text name=VALUE2 size=10 VALUE=0> <BR>

<select name=LOGICALID2>
<OPTION Value=0>NONE
<OPTION Value=1>AND
<OPTION Value=2>OR
</select>
on:<Select name=ls3>
<%LISTFIELDS rs%>
</select>
<select name=BOOLEANID3>
<OPTION Value='='>=
<OPTION Value='>'>>
<OPTION Value='<'><
<OPTION Value='>='>=
<OPTION Value='<='><=
<OPTION Value='<>'><>
</SELECT>
&nbsp;&nbsp;
<input type=text name=VALUE3 size=10 VALUE=0> 

</font>
</b>
<wbr><input type=SUBMIT src=go.gif name=printreg Value="Search">
</td>
</tr>
</table>
</form>
(NOTE: In case of Date search provide whole dates e.g. 11/2/2003 OR 11 Feb 2003)-->


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
&#169; Copyright 2006 . All rights reserved. SoftRobot Document Server
is property of ERPWEB.</font>
<script type="text/javascript">
setupAllTabs();
</script>
</BODY>

</HTML>
<!-------------------------------------start transaction server---->
<%
'Sub OnTransactionCommit()
    'Response.Write "The Transaction just committed" 
    'Response.Write "This message came from the "
    'Response.Write "OnTransactionCommit() event handler."
'End Sub

'Sub OnTransactionAbort()
    'Response.Write "The Transaction just aborted" 
    'Response.Write "This message came from the "
    'Response.Write "OnTransactionAbort() event handler."
'End Sub
%>