<!--
'************************************************************************
'Pupose						:	This is a SoftServer document list server
'Filename					:	document.asp
'Author						:	Rajat Taheem
'Created					:	07-Jan-2006
'Project Name				:	IRP
'Contact					:	rajat_b_taheem@yahoo.co.in
'
'Modification History		:	
'Purpose					:
'Version					:
'Author 					:
'Created					:
'************************************************************************
-->

<%
Set Conn=Server.CreateObject("ADODB.Connection")
Conn.Open Session("erp_ConnectionString")
%>
<%' @TRANSACTION=Required LANGUAGE="VBScript" %>
<%'@ LANGUAGE="VBScript" %>
<%
'Option Explicit
'Response.Buffer = True
Const adUseClient = 3
'FADD=REQUEST("FADD"): IF FADD THEN FADD=1 ELSE FADD=0
'FDEL=REQUEST("FDEL"): IF FDEL THEN FDEL=1 ELSE FDEL=0
'FVIEW=REQUEST("FVIEW"): IF FVIEW THEN FVIEW=1 ELSE FVIEW=0
'FEDIT=REQUEST("FEDIT"): IF FEDIT THEN FEDIT=1 ELSE FEDIT=0
'FFILTER=REQUEST("FILTER"): : IF FFILTER THEN FFILTER=1 ELSE FFILTER=0
'FOFFHOLD=REQUEST("FOFFHOLD"): IF FOFFHOLD THEN FOFFHOLD=1 ELSE FOFFHOLD=0
'FREJECT=REQUEST("FREJECT"): IF FREJECT THEN FREJECT=1 ELSE FREJECT=0

'FADD=1
'FDEL=1
FVIEW=1
'FEDIT=1
'FFILTER=1
'FOFFHOLD=1
'FREJECT=1

'SQLPROG=REQUEST("MYSQL")

'SQLPROG=REQUEST("MYSQLQRY")

MASTERTAB=REQUEST("MASTERTAB")
RFLD=REQUEST("RFLD")
FLDVAL=REQUEST("FLDVAL")
RDATE=REQUEST("RDATE")
ii=REQUEST("ii")
DID=REQUEST("DID")
'response.write did
'response.end

RFILTER=REQUEST("RFILTER")

RFLDD=REPLACE(RFLD," ","%20")
FLDVALL=REPLACE(FLDVAL," ","%20")

'RFLDD=server.urlencode(RFLD)
'FLDVALL=server.urlencode(FLDVAL)


%>
<!----
<BR>
<%=RFLD%>
<BR>
<%=RFLDD%>
<BR>
<%=FLDVAL%>
<BR>
<%=FLDVALL%>
<BR>
--->
<%

'RESPONSE.WRITE "<BR>" & RFLD
'RFLD=REPLACE(RFLD, "__", " ")
'RESPONSE.WRITE "<BR>" & RFLD
'RESPONSE.WRITE "<BR>" & FLDVAL
'RFLDVAL=REPLACE(RFLDVAL, "__", " ")
'RESPONSE.WRITE "<BR>" & FLDVAL

dim DocType, LinkAddr

SQL123="select * from Documents1 where DID=" & DID
Set rsSQL123=Server.CreateObject("ADODB.Recordset")
rsSQL123.Open SQL123 , Conn
	SQLProg = rsSQL123("SQLPROGRAM")
	DDID = rsSQL123("DDID")
	DocType=rsSQL123("DOCTYPEID")
	LinkAddr=rsSQL123("LINKADDRESS")
rsSQL123.close


dim mypos1, mypos2
'searchstring=" from "

'mypos1=InStr(sqlprog,searchstring)

'mypos2=InStr(mypos1,sqlprog," ")

'sqlstr=mid(sqlprog,mypos1,(mypos2-mypos1))

dim Myarr, whereGOT

whereGOT=False

wherestat=cstr(" ")

Myarr=split(sqlprog, " ", -1)

i=0
for each v in Myarr

'	response.write "<br>" & myarr(i)
	
	thisv=ucase(myarr(i))
	
	if thisv="FROM" then
		tab_view=myarr(i+1)	
	end if
	
	if thisv="WHERE" then
		remi=i
		whereGOT=TRUE
	end if
	
	if whereGOT then
		wherestat=cstr(wherestat & " " & myarr(i))
	end if
	
	
	i=i+1
next

'response.write RFLD & "<br>rajat" & FLDVAL & "<br>"

if RFILTER="W" then
	SQLPROGRAM="select * from " & tab_view & wherestat & " and " & RFLD & "='" & cstr(FLDVAL) & "' AND datepart(ww," & RDATE & ")=" & ii
else
	SQLPROGRAM="select * from " & tab_view & wherestat & " and " & RFLD & "='" & cstr(FLDVAL) & "' AND month(" & RDATE & ")=" & ii
end if


'IF ISNULL(SQLPROG) OR SQLPROG="" THEN
'	RESPONSE.WRITE "ERROR EXECUTING FURTHER..."
'	RESPONSE.END
'END IF

'RESPONSE.WRITE "<BR>RAJAT" &  SQLPROGRAM
'RESPONSE.END


UID=REQUEST("UID")

'RESPONSE.WRITE UID
'RESPONSE.END

FLAG=0
COLUMNNAME=REQUEST("ls")
SEARCHVALUE=REQUEST("txtsearch")





PROFILECSS="SAMPLE.CSS"
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
<TITLE>IRP FOLDER SERVER (sales@erpweb)</TITLE>
<LINK REL=STYLESHEET TYPE="text/css" HREF="<%=PROFILECSS%>">
</HEAD>
<BODY topmargin=0>
<Basefont face=arial size=1>

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
Sub GenerateTable( rs, pagesize )
	
  Response.Write( "<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF>" )
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
	
   		else
   		 IF FLAG THEN
			Response.Write( "<TD VALIGN=TOP bgcolor=#E1F2FD ><FONT FACE=ARIAL SIZE=1>" + CStr( v ) + "</FONT></TD>" )
		   ELSE
		    Response.Write( "<TD VALIGN=TOP bgcolor=#E1F2FD ><FONT FACE=ARIAL SIZE=1>" + CStr( v ) + "</FONT></TD>" )
		END IF  
    end if
    
   next
      '-----------------------------
  Response.Write("<TD bgcolor=#D5EAFF align=right><FONT FACE=ARIAL SIZE=1>")
  IF DID=3433 OR DID=747 OR DID = 2740 OR DID=2570 OR DID=2783 OR DID=2815 OR DID=2769 OR DID=2794 OR DID=2816 OR DID=2836 OR DID=2844 or did=2881 OR DID=2706 OR DID=2708 OR DID=2709 or DID=2626 OR DID=2932 or DID=3041 OR DID=3135 or did=3258 or did=3269 or did=3344 or did=3280 or did=3401 or did=3298 or did=3316 or did=3323 or did=3330 or did=3366 or did=3373 or did=3394 THEN
   Response.Write("<A HREF=TABWINDOWS.ASP?DID=" & DID & "&ID=" & rs(3) & "&UID=" & UID & " target=startmenu><img src=info.gif border=0 width=20 alt='Tabbed Document'></A>")
  END IF
  
  'IF DID=2740 THEN
  'Response.Write("<A HREF=TABWINDOWS.ASP?DID=" & DID & "&ID=" & rs(3) & "&UID=" & UID & " target=startmenu><img src=info.gif border=0 width=20 alt='Tabbed Document'></A>")
  'END IF
  
  IF FVIEW THEN
	  	if (cint(doctype)=3) then
	  	         IF NOT (ISNULL(LinkAddr) OR LinkAddr="") THEN
	  			  Response.Write("<A HREF=" & LinkAddr & "?DID=" & DID & "&ID=" & rs(3) & "&UID=" & UID & " target=news><img src=print.gif border=0 width=20 alt='Show Document'></A>")
		         END IF
		else
				  Response.Write("<A HREF=FORMVIEW.ASP?DID=" & DID & "&ID=" & rs(3) & "&UID=" & UID & " target=news><img src=print.gif border=0 width=20 alt='Show Document'></A>")
		end if


  END IF
  IF FEDIT THEN
  Response.Write("<A HREF=FORMEDITNEW.ASP?DID=" & DID & "&ID=" & rs(3) & "&UID=" & UID & " target=news><img src=update.gif border=0 width=20 alt='Edit Document Style 2'></A>")
  END IF
  IF FDEL THEN
  Response.Write("<A HREF=FORMDEL.ASP?DID=" & DID & "&ID=" & rs(3) & "&UID=" & UID & " target=news><img src=delete.gif border=0 width=20 alt='Delete Document'></A>")
  END IF
  IF FFILTER THEN
  Response.Write("<A HREF=FORMCHECK.ASP?DID=" & DID & "&ID=" & rs(3) & "&UID=" & UID & " target=news><img src=info.gif border=0 width=20 alt='Validate Document'></A>")
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
Sub GenerateHeader( rs )
Response.Write( "<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF >" )
Response.Write( "<TR>" )
Response.Write( "<TD valign=middle bgcolor=#DCDDDE bordercolordark=#FFFFFF bordercolor=#808080 bordercolorlight=#808080>&nbsp;&nbsp;<FONT FACE=ARIAL SIZE=2><img src=ofolder.gif border=0 alt='SoftRobot Document Server'><B> " & rs("TITLE") & " Register</B></FONT></TD>" )
Response.Write( "</TR>" )
Response.Write( "</TABLE>" )



End Sub
%>
<%
Sub GenerateFooter( rs )
Response.Write( "<TABLE BORDER=0 WIDTH=100% >" )
Response.Write( "<TR bgcolor=#F1F1F2>" )
IF FADD THEN
  Response.Write("<TD WIDTH=80% ></TD><TD ALIGN=RIGHT><FONT FACE=ARIAL SIZE=2><A HREF=FORMADDNEW.ASP?UID=" & UID & "&DID=" & DID & "&CID=" & CID & " target=news><img height=15 src=add.jpg border=0 alt='Add Document'></A></FONT></TD>")
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
FVIEW=TRUE
If mypage="" Then mypage=1 
'------------------------------------find doc sql
CID=REQUEST("CID")
DID=CInt(REQUEST("DID"))
SORT=REQUEST("SORT")
if DID="" Then Response.Write "ERROR:DID":Response.END
set rsDOC = Server.CreateObject("ADODB.Recordset")
rsDOC.Open "Select * from DOCUMENTS WHERE DID=" & DID, Conn 
'SQLPROGRAM=trim(rsDOC("SQLPROGRAM"))
DDID=rsDOC("DDID")
TABLENM=rsDOC("MASTERTABLE")
DOCID=TRIM(TABLENM) & "ID"
YEARFLAG=0
IF rsDOC("YEARFILTER") THEN YEARFLAG=1 ELSE YEARFLAG=0
'------------------------------------

if not searchvalue="" then
	response.write "<br><font face=arial size=2><b>Searched Upon:&nbsp;" & "&nbsp;(&nbsp;" & cstr(columnname) & "&nbsp;)&nbsp;:&nbsp;'&nbsp;"  & cstr(searchvalue) & "&nbsp;'</b></font><br><br>"
end if

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
'					RESPONSE.WRITE SQLPROGRAM & " AND CALENDERID=" & CID
'					RESPONSE.END
					rs.Open SQLPROGRAM & " AND CALENDERID=" & CID, Conn
					END IF
				ELSE
					IF SORT=1 THEN
					rs.Open SQLPROGRAM & " ORDER BY " & COLUMNNAME & " ASC", Conn
					ELSEIF SORT=2 THEN
					rs.Open SQLPROGRAM & " ORDER BY " & COLUMNNAME & " DESC", Conn
					ELSE 
'RESPONSE.WRITE "<br>SQL: " & SQLPROGRAM & "<br>"
'RESPONSE.END					
					rs.Open SQLPROGRAM, Conn
					END IF
				END IF
        	END IF
	ELSE
			IF YEARFLAG THEN
				rs.Open SQLPROGRAM & " AND " & COLUMNNAME & " LIKE '%" & TRIM(SEARCHVALUE) & "%'" & " AND CALENDERID=" & CID, Conn
			ELSE
				rs.Open SQLPROGRAM & " AND " & COLUMNNAME & " LIKE '%" & TRIM(SEARCHVALUE) & "%'", Conn
			END IF
	END IF
    '-----------------------
if not rs.EOF then
	numpages = rs.PageCount
	numrecs = rs.RecordCount
	rs.AbsolutePage = mypage
'--------------------------------------
	Response.Write( "<table width=100% border=1 bordercolordark=#FFFFFF bordercolorlight=#FFFFFF><td align=left bgcolor=#f1f1f2><font face=arial SIZE=1>" & numrecs & " Documents found.</td>" )
	Response.Write("<td align=right bgcolor=#f1f1f2><font face=arial SIZE=1><i>Register Index Page " & mypage & " of " & numpages & " </i></td></table>" )
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
<form name=form1 action=RGRIDdocument.asp method=post>
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
<input type=hidden name="MASTERTAB" value=<%=MasterTab%>>
<input type=hidden name="RFLD" VALUE=<%=RFLDD%>>
<input type=hidden name="FLDVAL" VALUE=<%=FLDVALL%>>
<input type=hidden name="RDATE" VALUE=<%=RDATE%>>
<input type=hidden name="RFILTER" value=<%=RFILTER%>>
<input type=hidden name="ii" VALUE=<%=ii%>>
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
<font face=arial size=2><b>(NOTE: In case of Date search provide Day e.g. 11, Month e.g. Nov or Year e.g. 2003 seperately)</b></font>
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
&#169; Copyright 2007 . All rights reserved. IRP SoftServer
is property of ERPWEB.</font>
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





























































































































