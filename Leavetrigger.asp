<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE> New Document </TITLE>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
</HEAD>
<BODY>
<TABLE BORDER=1 width=100% bordercolordark=#FFFFFF bordercolorlight=#FFFFFF>
<tr>
<TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080><font face=arial size=1>EmployeeId</td>
<TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080><font face=arial size=1>Employee Name</td>
<TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080><font face=arial size=1>Action</td><tr>
<%
Set Conn=Server.CreateObject("ADODB.Connection")
Conn.Open Session("erp_ConnectionString")
set rsdist = server.createobject("adodb.recordset")
sqldist ="select distinct(employeeid)Employeeid from atndcard where status<>'Present' and month(indatetime)= month('"& date() &"') order by employeeid"
rsdist.open sqldist,conn
'response.write sqldist
while not rsdist.eof 
	'response.write("<br>"&rsdist("employeeid"))
	count = 1
	sqlattn ="select * from atndcard where status<>'Present' and month(indatetime)= month('"& date() &"') and employeeid="&rsdist("Employeeid")&" order by indatetime"
	set rsattn=server.createobject("adodb.recordset")
	rsattn.open sqlattn,conn,1,1
	
	if rsattn.recordcount >=3 then
		'response.write("<br>greater than 3")
		dt1=rsattn("indatetime")
		rsattn.movenext
		while not rsattn.eof 
			if (abs(datediff("d",dt1,rsattn("indatetime")))=1) then 
				'response.write ("<br>dt1"&dt1)
				'response.write ("<br>rsattn"&rsattn("indatetime"))
				'response.write("incremented")
				count=count+1
				'response.write(count)
				dt1=rsattn("indatetime")
				sqlleaveapp ="select * from leaveapp where employeeid = "& rsdist("employeeid") &" and '" & rsattn("indatetime") &"' between fromdate and todate "
				'response.write "<br>"&sqlleaveapp
				set rsleaveapp = server.createobject("adodb.recordset")
				rsleaveapp.open sqlleaveapp,conn
				if not rsleaveapp.eof then
					'response.write("ergsdgd")
					'response.write ("<br>count reset eof")
					'response.write(count)
					count =1
				end if
			elseif count <3 then
				'response.write ("<br>count reset datediff")
				'response.write(count)
				dt1=rsattn("indatetime")
				count=1	
			end if
			rsattn.movenext
		wend
	end if
	'response.write "<br>count"& count
	if count >= 2 	then	
		
		sqlemp= "select * from employee where employeeid ="&rsdist("employeeid")&""
		set rsemp= server.createobject("adodb.recordset")
		rsemp.open sqlemp,conn
		Response.Write( "<TR bgcolor=#E1F2FD>" )
		response.write("<TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080><font face=arial size=1>" & rsemp("EmployeeId") &"</td>")
		response.write("<TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080><font face=arial size=1>" & rsemp("EmployeeName") &"</td>")
		response.write("<TD bgcolor=#D5EAFF bordercolordark=#BADFFE bordercolorlight=#808080><font face=arial size=1><a href=Leavedefaulter1.asp?ID=" & rsdist("EmployeeId") &" target=news>Select</a></td></font>")
	end if		


rsdist.movenext
wend
%>
</table>
</BODY>
</HTML>

