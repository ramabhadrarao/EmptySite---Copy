<%' Copyright MobileERP Softech P Ltd. India.
IF MODE=4 THEN'------------This mode is called by FormAddNew.asp file to add data into system
	IF ADDSQL <> "" THEN 
		set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open ADDSQL, Conn, 1, 3
		MSTTBLID=rs(3).name
		rs.AddNew
		'--------------------------------------------for starts
			FOR K = 4 TO rs.fields.count-1
			DATA=TRIM(REQUEST("mastfld" & K )):IF ISNULL(DATA) OR DATA="" THEN DATA=""
			checkdata rs, DATA, K, 1
				if rs(K).name="TOTAL" OR rs(K).name="ORDERVALUE" then

				elseif rs(K).name="UPLOADFILE" then

'-----------------------------UPLOAD FILE CODE FOR RESERVE FIELD UPLOADFILE-------------------------------

					Set fs=Server.CreateObject("Scripting.FileSystemObject")
					Set f=fs.OpenTextFile(Server.MapPath("upload\hi.txt"), 1)

					do while f.AtEndOfStream = false
						file = f.ReadLine
					loop
					f.Close
					Set f=Nothing
					if fs.FileExists(Server.MapPath("upload\hi.txt")) then
						fs.DeleteFile(Server.MapPath("upload\hi.txt"))
					end if
					Set fs=Nothing
					rs(K) = "upload\" & file
'-------------------------------------------------End-----------------------------------------------------		
				else
					'if DATA="" OR ISNULL(DATA) THEN
					rs(K)=DATA
					'end if
				end if
'--------------------------------------------------------
				if (rs(k).Type=135 and not rs(k).name="PAYDUEDATE") then
					rs(k) = Cdate(cstr(Data))	
				end if
			NEXT
			'--------------------------------------for ends
	rs.Update
	END IF
	rs.Close
	Set rs = Nothing
'---------------------------DETERMINE RECORD ID
	set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorType = adOpenStatic '1 ADDED BY HARSHA .. tO SOLVED THE ERROR 'Rowset does not support fetching backward'
  ' Use client cursor to enable AbsolutePosition property.
	rs.CursorLocation = adUseClient'2
	rs.Open ADDSQL & " ORDER BY " & MSTTBLID, Conn, 1, 3
	rs.MoveLast
	ID=rs(3)
	rs.Close
	Set rs = Nothing
'---------------------------------------------------EDITABLE ZONE FOR EXTERNAL CODE SPECIFIC TO EACH DOCUMENT AS PER PROJECT REQUIREMENTS
IF DID=1989 THEN'------------------------------------------AUTOENTER IQC DOC FROM DO
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.Open "select DORECEIVEID FROM IQCAPPROVE WHERE IQCAPPROVEID=" & ID, Conn
	IF NOT rs.EOF THEN
	DOID=rs(0):IF DOID=0 OR ISNULL(DOID) THEN RESPONSE.Write ("SELECT DO"):RESPONSE.END
	END IF
	rs.Close
	Set rs = Nothing
	'--------------------------------------
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.Open "select ITEMID FROM DORECEIVEDET WHERE DORECEIVEID=" & DOID, Conn
	While NOT rs.EOF
	ITEMID=rs(0):IF ISNULL(ITEMID) THEN ITEMID=0
	'----------------------------------------------------------------
	ISQL="INSERT INTO IQCAPPROVEDET (IQCAPPROVEID, ITEMID) VALUES (" & ID & ", " & ITEMID & ")"
	'RESPONSE.Write ISQL
	Conn.Execute ISQL
	rs.MoveNext
	wend
	rs.Close
	Set rs = Nothing
END IF
IF DID=1982 THEN'------------------------------------------AUTOENTER IQC DOC FROM DO
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.Open "select IQCAPPROVEID FROM PURCHINVOICE WHERE PURCHINVOICEID=" & ID, Conn
	IF NOT rs.EOF THEN
	DOID=rs(0):IF DOID=0 OR ISNULL(DOID) THEN RESPONSE.Write ("SELECT DO"):RESPONSE.END
	END IF
	rs.Close
	Set rs = Nothing
	'--------------------------------------
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.Open "select ITEMID, QTY, RATE FROM COPYGRNPI WHERE ID=" & DOID, Conn
	While NOT rs.EOF
	ITEMID=rs(0):IF ISNULL(ITEMID) THEN RESPONSE.END
	QTY=rs(1):IF ISNULL(QTY) THEN RESPONSE.END
	RATE=rs(2):IF ISNULL(RATE) THEN RESPONSE.END
	'----------------------------------------------------------------
	ISQL="INSERT INTO PURCHINVOICEDET (PURCHINVOICEID, ITEMID, QTY, RATE) VALUES (" & ID & ", " & ITEMID & ", " & QTY & ", " & RATE & ")"
	'RESPONSE.Write ISQL
	Conn.Execute ISQL
	rs.MoveNext
	wend
	rs.Close
	Set rs = Nothing 
	'-------------------------
	USQL="UPDATE IQCAPPROVE SET ONHOLD=1 WHERE IQCAPPROVEID=" & DOID
	Conn.Execute USQL
	'----------------------------
END IF
'IF DID=657 THEN '-----------------------------------------enter podet automatically
	'set rs=Server.CreateObject("ADODB.Recordset")
	'rs.Open "select PRID, POQUOTEREFID, SUPPLIERID FROM PO WHERE POID=" & ID, Conn
	'PRID=rs(0):IF PRID=0 OR ISNULL(PRID) THEN RESPONSE.Write ("SELECT PR"):RESPONSE.END
	'POQUOTEREFID=rs(1):IF POQUOTEREFID=0 OR ISNULL(POQUOTEREFID) THEN RESPONSE.Write ("SELECT POQUOTEREF"):RESPONSE.END
	'SUPPLIERID=rs(2):IF SUPPLIERID=0 OR ISNULL(SUPPLIERID) THEN RESPONSE.Write ("SELECT SUPPLIER"):RESPONSE.END
	'rs.Close
	'Set rs = Nothing
	'IF POQUOTEREFID=1 THEN'-------------------------------------------IF LAST PO REF
	'set rs=Server.CreateObject("ADODB.Recordset")
	'rs.Open "select ITEMID, QTY FROM PRDET WHERE PRID=" & PRID, Conn
	'While NOT rs.EOF
	'ITEMID=rs(0):IF ISNULL(ITEMID) THEN ITEMID=0
	'QTY=rs(1):IF ISNULL(QTY) THEN QTY=0
	'---------------------------------------------------------------find rate from last po
	'set rss=Server.CreateObject("ADODB.Recordset")
	'rss.Open "select RATE FROM PODETFIND WHERE ITEMID=" & ITEMID & " AND SUPPLIERID=" & SUPPLIERID, Conn
	'IF NOT rss.EOF THEN
	'RATE=rss(0):IF ISNULL(RATE) THEN RATE=0
	'ELSE
	'RATE=0
	'END IF
	'rss.Close
	'Set rss = Nothing
	'----------------------------------------------------------------
	'ISQL="INSERT INTO PODET (POID, ITEMID, RATE, QTY) VALUES (" & ID & ", " & ITEMID & ", " & RATE & ", " & QTY & ")"
	'RESPONSE.Write ISQL
	'Conn.Execute ISQL
	'rs.MoveNext
	'wend
	'rs.Close
	'Set rs = Nothing
	'ELSE'-----------------------------------------------------------update from quote
	'set rs=Server.CreateObject("ADODB.Recordset")
	'rs.Open "select ITEMID, QTY FROM PRDET WHERE PRID=" & PRID, Conn
	'While NOT rs.EOF
	'ITEMID=rs(0):IF ISNULL(ITEMID) THEN ITEMID=0
	'QTY=rs(1):IF ISNULL(QTY) THEN QTY=0
	''---------------------------------------------------------------find rate from quote
	'set rss=Server.CreateObject("ADODB.Recordset")
	'rss.Open "select RATE FROM PODETQUOTEFIND WHERE ITEMID=" & ITEMID & " AND SUPPLIERID=" & SUPPLIERID, Conn
	'IF NOT rss.EOF THEN
	'RATE=rss(0):IF ISNULL(RATE) THEN RATE=0
	'ELSE
	'RATE=0
	'END IF
	'rss.Close
	'Set rss = Nothing
	'----------------------------------------------------------------
	'ISQL="INSERT INTO PODET (POID, ITEMID, RATE, QTY) VALUES (" & ID & ", " & ITEMID & ", " & RATE & ", " & QTY & ")"
	'RESPONSE.Write ISQL
	'Conn.Execute ISQL
	'rs.MoveNext
	'wend
	'rs.Close
	'Set rs = Nothing
    'END IF
'END IF
IF DID=3383 THEN '-------------------------------------------------AUTO ENTER IDN/PKG SLIP BASED ON SRNO ENTRY.
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.Open "select FGRECEIVEDETID, FROMSRNO, TOSRNO FROM FGSERIALNO WHERE FGSERIALNOID=" & ID, Conn
	FID=rs(0):if FID<=0 THEN Response.end
	FROMSRNO=rs(1):IF FROMSRNO<=0 OR ISNULL(FROMSRNO) OR NOT ISNUMERIC(FROMSRNO) THEN RESPONSE.Write ("SELECT FROMSRNO"):RESPONSE.END
	TOSRNO=rs(2):IF TOSRNO<=0 OR ISNULL(TOSRNO) OR NOT ISNUMERIC(TOSRNO) THEN RESPONSE.Write ("SELECT TOSRNO"):RESPONSE.END
	rs.Close
	Set rs = Nothing
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.Open "select QTY FROM FGRECEIVEDET WHERE FGRECEIVEDETID=" & FID, Conn
	QTY=rs(0)
	rs.Close
	Set rs = Nothing
	Q1=TOSRNO-FROMSRNO
    IF TOSRNO<FROMSRNO THEN RESPONSE.Write ("YOU CANNOT ENTER TOSRNO LESS THEN FROMSRNO"):RESPONSE.END
    IF Q1 > (QTY-1) THEN RESPONSE.Write ("YOU CANNOT ENTER SRNOs MORE THEN QTY"):RESPONSE.END
	FOR C=FROMSRNO TO TOSRNO
	ISQL="INSERT INTO FGSERIALNODET (FGSERIALNOID, SRNO) VALUES (" & ID & ", " & C & ")"
	'RESPONSE.Write ISQL
	Conn.Execute ISQL
    NEXT
END IF
'----------------------------ENTER MORE FIELD DATA
	DETID=REQUEST("DETID")'FOR MORE LINKED FORM FEATURE
	IF ISNULL(DETID) OR DETID="" THEN DETID=0
	UFLD=REQUEST("FLD"): IF ISNULL(UFLD) THEN UFLD=0
	IF DETID > 0 THEN
		TBL=Mid(trim(UFLD),1,(LEN(UFLD)-2))
		M="formeditnew.asp?DID=" & TRIM(DID) & "&ID=" & TRIM(ID)
		USQL="UPDATE " & TBL & " SET MORE='" & CSTR(M) & "' Where " & UFLD & "=" & DETID
		USQL=CSTR(USQL)
		'RESPONSE.Write USQL
		Conn.EXECUTE USQL 
	END IF
'----------------------------Enter all records in all linked masters
	SQL="SELECT * FROM TABWINDOWITEMS WHERE ID=" & DID
	Set rs=Server.CreateObject("ADODB.Recordset")
	rs.Open SQL , Conn
	IF NOT rs.EOF THEN
		GenerateTable rs, ID, UID   ' call addtabdoc.asp
	END IF
	rs.Close
	Set rs = Nothing
'----------------------------
	TSQL="INSERT INTO TRACKING (UID, DID, ID, FUNC) VALUES (" & UID & ", " & DID & ", " & ID & ", 'ADD')"
	Conn.EXECUTE TSQL
ELSE'-------------------------0
	if ID="" Then Response.Write "ERROR:ID IS NULL":Response.End
END IF'-----------------------MODE 4 ENDS HERE
'-------------------------------------mode to add new document ends here this mode is called by addformnew.asp
%>
<%
IF MODE=1 THEN '----------THIS MODE WILL UPDATE MASTER & DETAILS DATA
'------------------------------
	IF SQLPROGRAM <> "" THEN 
		set rs=Server.CreateObject("ADODB.Recordset")
		rs.Open SQLPROGRAM & ID, Conn, 1, 3
	ELSE
		Response.Write "PLEASE ENTER MASTERSQL"
		Response.End
	END IF
'------------------------------update MASTER data after validation  
  IF NOT rs.EOF THEN
	For i=4 to rs.fields.count-1
		DATA=REQUEST("fld" + CSTR(i)):IF ISNULL(DATA) OR DATA="" THEN DATA=""
		checkdata rs, DATA, i, 0
	
	if rs(i).name="TOTAL" OR rs(i).name="PAYDUEDATE" then
	
	elseif rs(i).name="UPLOADFILE" then
	'---------UPLOAD FILE CODE FOR RESERVE FIELD UPLOADFILE---------------------
		data = rs(i)
     	Set fs=Server.CreateObject("Scripting.FileSystemObject")
		if fs.FileExists(Server.MapPath("upload\hi.txt")) then
    		Set f=fs.OpenTextFile(Server.MapPath("upload\hi.txt"), 1)
			file = f.ReadLine
			f.Close
			Set f=Nothing
			fs.DeleteFile(Server.MapPath("upload\hi.txt"))
  			flnm = "upload\" & file
  			rs(i)=cstr(flnm)
  			Set fs=Nothing
  		else
  			rs(i)=data      
		end if	
	else
		rs(i)=DATA
	end if
	
	if (rs(i).Type=135 and not rs(i).name="PAYDUEDATE") then
		rs(i)=cdate(cstr(DATA))
	end if
	
	Next
	rs.Update
  END IF
rs.Close
'------------------------------set DETAILS data after validation  
	IF DETAILSSQL <> "" THEN 
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.Open DETAILSSQL & ID, Conn, 1, 3
	J=0
  WHILE NOT rs.EOF 
	FOR i=2 TO rs.fields.count-1
		DATA=REQUEST("flds" + CSTR(i) + "_" +  CSTR(J)):IF ISNULL(DATA) OR DATA="" THEN DATA=""
		checkdata rs, DATA, i, 0
		if rs(i).name="TOTAL" then
		else
			rs(i)=DATA
		end if
  
		if (rs(i).Type=135  and not rs(i).name="PAYDUEDATE") then
			rs(i)=cdate(cstr(DATA))
		end if
  
	NEXT
	J=J+1
	rs.Update
	'-----------------------------VALIDATION IN DETAILS EDIT TO CHECK DO QTY NOT GREATER THEN PO QTY
	IF DID=617 THEN
		QTY=rs("QTY")
		ITEMID=rs("ITEMID")
		DOID=rs("DORECEIVEID")
		set rsS=Server.CreateObject("ADODB.Recordset")
		rsS.Open "select POID FROM DORECEIVE WHERE DORECEIVEID=" & DOID, Conn
		POID=rsS(0)
		rsS.Close
		Set rsS = Nothing
		'---------------------------
		set rsS=Server.CreateObject("ADODB.Recordset")
		rsS.Open "select * FROM PODOLINK WHERE POID=" & POID & " AND ITEMID=" & ITEMID, Conn
		POQTY=rsS("POQTY"):IF ISNULL(POQTY) THEN POQTY=0
		DOQTY=rsS("DOQTY"):IF ISNULL(DOQTY) THEN DOQTY=0
		PENDING=POQTY-DOQTY
		rsS.Close
		Set rsS = Nothing
		'---------------------------
		'---------------------------
			IF PENDING >= 0 THEN
				
			ELSE
	%>
			<SCRIPT LANGUAGE=VBSCRIPT>
			MSGBOX "DELIVERY QTY GREATER THEN PO QTY NOT ALLOWED. Kindly update right QTY."
			</SCRIPT>
	<%
			
			END IF
		
	END IF
	rs.MoveNext
	Wend
	rs.Close
	END IF
'-----------------------------------------
	TSQL="INSERT INTO TRACKING (UID, DID, ID, FUNC) VALUES (" & UID & ", " & DID & ", " & ID & ", 'EDIT')"
	Conn.EXECUTE TSQL
END IF
'-------------------------------------------MODE=1 ENDS HERE
IF MODE=2 THEN '----------DELETE DETAIL DATA
	IDD=REQUEST("IDD")
	Conn.Execute "DELETE FROM " & DETTBL & " WHERE " & DETTBL & "ID=" & IDD
'------------------------------------------
	TSQL="INSERT INTO TRACKING (UID, DID, ID, FUNC) VALUES (" & UID & ", " & DID & ", " & ID & ", 'DETDELETE')"
	Conn.EXECUTE TSQL
END IF'------------------------------------2,3
IF MODE=3 THEN '----------ADD DETAIL DATA
	J=REQUEST("J")
	IF ADDDETAILS <> "" THEN 
		set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open ADDDETAILS, Conn, 1, 3
		rs.AddNew
		rs(0)=ID
		FOR P = 2 TO rs.fields.count-1
			DATA=REQUEST("newfld" & P ):IF ISNULL(DATA) OR DATA="" THEN DATA=""
				checkdata rs, DATA, P, 0 '--------------VALIDATE ALL DATA EXTRACTED FROM FORM
			if rs(P).name="TOTAL" OR rs(P).name="ORDERVALUE" OR rs(P).name="MORE" then
     '-------------------------------------------------------------AUTO ENTER DETAILS FIELD IN ADD DETAILS
			elseif (rs(P).name="UNITCOST"  and DID=627 ) then
				SSQL="SELECT BASEPRICE FROM PRICELIST WHERE ITEMID=" & rs("ITEMID")
				'response.Write SSQL
				set rsS=Server.CreateObject("ADODB.Recordset")
				rsS.Open SSQL, Conn
				BP=0
				IF NOT rsS.EOF then
				BP=rsS(0):if ISNULL(BP) OR NOT ISNUMERIC(BP) THEN BP=0
				END IF
				rsS.Close
				rs(P)=BP
			elseif rs(P).name="UPLOADFILE" then
	'-----------------------------UPLOAD FILE CODE FOR RESERVE FIELD UPLOADFILE
			DATA=rs(P)
				IF DATA="" or ISNULL(DATA) THEN
				ELSE
					FILENAME=DATA
					SHOWLINK="<A HREF=" & cstr(FILNAM) & " target=fileshow>Show</A>"
					rs(P)=SHOWLINK
				END IF

			else
				rs(P)=DATA
			end if 
    '-----------------------------------------
			if (rs(P).Type=135 and not rs(P).name="PAYDUEDATE") then
				rs(P) = cdate(cstr(DATA))
			end if

		NEXT
		rs.Update
	END IF
	rs.Close
	Set rs = Nothing
	
	'------------------------------
	TSQL="INSERT INTO TRACKING (UID, DID, ID, FUNC) VALUES (" & UID & ", " & DID & ", " & ID & ", 'DETADD')"
	Conn.EXECUTE TSQL
END IF '------------------------------------END OF MODE 3
IF MODE=5 THEN
EID=REQUEST("EID")
DSQL="DELETE FROM RFQDETAILS WHERE ITEMID=" & EID & " AND RFQID=" & ID
Conn.Execute DSQL
END IF
%>