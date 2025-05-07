<%
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
					rs(K)=DATA
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
'----------------------------------------------SPECIAL CODE FOR AUTO ENTERING RFQ FROM PR
IF DID=3276 THEN
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.Open "select PRID FROM RFQ WHERE RFQID=" & ID, Conn
	PRID=rs(0):IF PRID="" OR ISNULL(PRID) THEN RESPONSE.Write ("SELECT PR"):RESPONSE.END
	rs.Close
	Set rs = Nothing
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.Open "select PRDETID, REQDATE FROM PRDET WHERE PRID=" & PRID, Conn
	While NOT rs.EOF
	PRDETID=rs(0)
	REQDATE=rs(1)
	ISQL="INSERT INTO RFQDET (RFQID, PRDETID, REQDATE) VALUES (" & ID & ", " & PRDETID & ", '" & REQDATE & "')"
	'RESPONSE.Write ISQL
	Conn.Execute ISQL
	rs.MoveNext
	wend
	rs.Close
	Set rs = Nothing
END IF
'----------------------------------------------SPECIAL CODE FOR AUTO ENTERING RFQ MATERIAL FROM BOM REFEIING PR
IF DID=3314 THEN
    set rs=Server.CreateObject("ADODB.Recordset")
	rs.Open "select RFQDETID, RFQDETAILSID FROM RFQDETAILS WHERE RFQDETAILSID=" & ID, Conn
	RID=rs(0)
	RRID=rs(1)
	rs.Close
	Set rs = Nothing
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.Open "select RFQID, ITEMID FROM RFQDETITEM WHERE RFQDETID=" & RID, Conn
	RFQID=rs(0)
	ITEMID=rs(1)
	rs.Close
	Set rs = Nothing
	'----------------------------------
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.Open "select SPECS1ID FROM RFQ WHERE RFQID=" & RFQID, Conn
	TID=rs(0)
	IF TID=1 OR TID=4 THEN
	SSQL="select END_ITEMID FROM RFQBOMDET WHERE ITEMTYPEID=3 AND START_ITEMID=" & ITEMID
	ELSEIF TID=2 THEN
	SSQL="select END_ITEMID FROM RFQBOMDET WHERE (ITEMTYPEID=5 OR ITEMTYPEID=6) AND START_ITEMID=" & ITEMID
	ELSEIF TID=3 THEN
	SSQL="select END_ITEMID FROM RFQBOMDET WHERE START_ITEMID=" & ITEMID
	END IF
	rs.Close
	Set rs = Nothing
	'------------------------------------
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.Open SSQL, Conn
	While NOT rs.EOF
	ENDITEMID=rs(0)
	ISQL="INSERT INTO RFQDETAILSDET (RFQDETAILSID, ITEMID) VALUES (" & RRID & ", " & ENDITEMID & ")"
	RESPONSE.Write ISQL
	Conn.Execute ISQL
	rs.MoveNext
	wend
	rs.Close
	Set rs = Nothing
END IF
'----------------------------ENTER MORE FIELD DATA
	DETID=REQUEST("DETID")'FOR MORE LINKED FORM FEATURE
	IF ISNULL(DETID) OR DETID="" THEN DETID=0
	UFLD=REQUEST("FLD"): IF ISNULL(UFLD) THEN UFLD=0
	IF DETID>0 THEN
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
%>