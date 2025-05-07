<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE> Dwell Report </TITLE>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
</HEAD>

<BODY>


<%sub displayverticalgraph(strtitle,strytitle,strxtitle,avalues,alabels)


'************************************************************************
'           user customizeable values for different formats
'			of the graph
' 			FREEWARE!!
'	Just tell me if you are going to use it - info@cool.co.za
'************************************************************************

const GRAPH_HEIGHT 	= 300 	 		'set up the graph height
const GRAPH_WIDTH 	= 400   		'set up the graph width
const GRAPH_SPACING 	= 0	
const GRAPH_BORDER 	= 0	 		'if you would like to see the borders to align things differently	
const GRAPH_BARS 	= 2     		'loops through different colored bars e.g. 2 bars = gold,blue,gold,blue
const USELOWVALUE 	= FALSE 		'this uses the low value of the array as the value of origin (default = 0)
const SHOWLABELS 	= TRUE  		'set this to toggle whether or not the labels are shown
const L_LABEL_SEPARATOR = "|"			' |Label
const R_LABEL_SEPARATOR = "|"			' Label|
const LABELSIZE 	= -4			
const GRAPHBORDERSIZE 	= 1
const INTIMGBORDER 	= 1  	 		'border around the bars
Const ALT_TEXT		= 3			'Changes the format of the alternate text of the bar image
						'1 = Labels ,2 = Values , 3 = Labels + Values , 4 = Percent

'************************************************************************
'array of different bars to loop through
'you can change the order of these 
'Count = 10 
'"dark_green","red","gold","blue","pink","light_blue","light_gold","orange","green","purple"
'cut and paste from here and insert into the agraph_bars array below Make sure the 
'number specified in the const GRAPH_BARS is the same as or less than that in the array 
' 7 graph_bars <= 7 elements in array
'************************************************************************

agraph_bars = array("dark_green","red","gold","blue","pink","light_blue","light_gold","orange","green","purple")


intmax = 0

'find the maximum value of the values array
for i = 0 to ubound(avalues)
if  cint(intmax) < cint(avalues(i)) then  intmax = cint(avalues(i)) 
next
if uselowvalue then 
intmin = avalues(0)
for i = 0 to ubound(avalues)
if  cint(intmin) > cint(avalues(i)) then  intmin = cint(avalues(i)) 
next
end if
'establish the graph multiplier
graphmultiplier = round(graph_height-100/intmax)

imgwidth = round(300/(ubound(avalues)+1))
if imgwidth > 16 then imgwidth = 16 


%>

<table border =<%=GRAPH_BORDER%> width:100% height=<%=graph_height%>>
  
  <tr>
    <td rowspan=3 valign="middle"><%=strytitle%> </td>
    <td colspan=<%=ubound(avalues)+2%> height=50 align="center">
	
      <h4><%=strtitle%></h4></td>
    </tr>
	<% count = 0%>
     <tr>
       <td>
	 <table border=<%=graph_border%> cellpadding = 0 cellspacing = 15<%'=graph_spacing%>><tr>
	    <tr>
	     <TD height="100%">
	      <table border="<%=graph_border%>" height="100%">
	         <tr>
	         <td height="50%" valign="top" align=right><%=intmax%></td>
	         </tr>
	         <tr>
	         <td height="50%" valign="bottom" align=right>
		 
		 <%if uselowvalue then
	             response.write cstr(intmin)
		   else
		     response.write "0"
		   end if
		  %>
		 </td>
	         </tr>
	      </table>
	     </td>
	     <td valign="bottom" align="right"><img src="leftbord.gif" width="2" height="<%=graphmultiplier+8%>">
	     </td>
             
             <%
             '*******************MAIN PART OF THE CHART************************************
	     for i = 0 to ubound(avalues)
	       strgraph = agraph_bars(count)
	       	if alt_text = 1 then 
		   stralt = alabels(i)
		  elseif alt_text = 2 then 
		    stralt = avalues(i)
		  elseif alt_text = 3 then 
		    stralt = alabels(i) &" - "  &avalues(i)
		  elseif alt_text = 4 then   
		    stralt = round(avalues(i) /intmax  *100,2) &"%"
		 end if     
	        
	        if uselowvalue then  %>
                  <td valign="bottom" align="center">
	          <img src="<%=strgraph%>.gif" height="<%=round((avalues(i)-intmin)/intmax*graphmultiplier,0)%>" width="<%=imgwidth%>" alt="<%=strAlt%>" border="<%=intimgborder%>"></td>
	       <%else%>
	          <td valign="bottom" align="center">
	          <img src="<%=strgraph%>.gif" height="<%=round(avalues(i)/intmax*graphmultiplier,0)%>" width="<%=imgwidth %>" alt="<%=strAlt%>" border="<%=intimgborder%>"></td>
	       <%end if 
		
	        if count = graph_bars-1 then 
	          count = 0 
	        else
	          count = count + 1
	        end if		
	      next  
	      
	        'write out the border at the bottom of the bars also leave a blank cell for spacing on the right
	         response.write "<td width='50'>&nbsp;</td></tr><tr><td width=8>&nbsp;</td><td>&nbsp;</td><td colspan=" &(ubound(avalues)+1) &" valign='top'>" _
	         &"<img src='botbord.gif' width='100%' height='2'</td></tr>"
	     if showlabels then %>
	         <tr><td width=8 height=1>&nbsp;</td><td>&nbsp;</td>
	            <%for i = 0 to ubound(avalues)%>
	               <td valign="bottom" width=<%=imgwidth%> ><font size=
	               <%=labelsize &">" &l_label_separator &alabels(i) &r_label_separator %></font></td>
	            <%next%>
	         </tr>
	     <%end if%>
    	<tr><td colspan=<%=ubound(avalues)+3%> height=50 align="center"><%=strxtitle%></td>
    	</tr>
	</table>
	</td>
	</tr>
	<tr>
	<td></td></tr>
	</table>

<%end sub %>









<% 
ID = request("ID")
'Response.ContentType = "application/vnd.ms-excel"
set conn=server.createobject("adodb.connection")
conn.Open Session("erp_ConnectionString")
set rsCust = server.createobject("adodb.recordset")
sqlCust="select * from Customer where Customerid ="&ID&""
rsCust.open sqlCust,conn
if not rsCust.eof then
	response.write("<b><Center>" &rsCust("CUSTOMERNAME")& "</Center></b><br>")
end if
set rsCust = nothing
%>
<TABLE BORDER=1>
<TR>
<TD align=center><b>Category</b></TD>
<TD align=center><b>Category Name</b></TD>
<TD align=center><b>Reports Name</b></TD>
<TD align=center><b>Number Of Days</b></TD>
</TR>
<TR>
<TD>A.</TD>
<TD>Party's Order Date to Our Despatch Date</TD>
<td>Customers Dwell Report</td>
<%
po_Deldays=0
po_Invdays=0
sql = "select abs(datediff(dd,podate,DeliveryDATE)),abs(datediff(dd,SalesOrderdate,DeliveryDATE)) from SalesOrder where customerid = "&ID&" and month(podate)="&month(date)&" and year(podate)="&year(date)&" "
'response.write sql
'response.end
set rs= conn.execute(sql)
while not rs.eof 
	if isnull(rs(0)) then
		po_Deldays= po_Deldays
		po_Invdays = po_Invdays
	else
		po_Deldays = po_Deldays+rs(0)
		po_Invdays=po_Invdays+rs(1)
	end if
rs.movenext
wend
response.write("<td>"&po_Deldays&"</td>")
%>
</TR>
<TR>
<TD>B.</TD>
<TD>Party's Order Date to Our Invoice Date</TD>
<td>Sales/Mktg Tem</td>
<%response.write("<td>"&po_Deldays&"</td>")%>
</TR>   
<TR>
<TD>C.</TD>
<TD>Our Internal Order Date to JobCard Isuue Date</TD>
<td>Design and Master Dept.</td>
<td>
<%sqlwo = "select abs(datediff(dd,wodate,SalesOrderdate)),abs(datediff(dd,wodate,Deliverydate)) from SalesOrder,LotMgmt where customerid = "&ID&" and month(podate)="&month(date)&" and year(podate)="&year(date)&" and salesorder.salesorderid = lotmgmt.salesorderid"
days_wo = 0
days_del=0
'response.write sqlwo
'response.end
set rswo= conn.execute(sqlwo)
while not rswo.eof 
	if isnull(rswo(0)) then
		days_wo = days_wo
		days_del=days_del
	else
		days_del=days_del+rswo(1)
		days_wo  = days_wo +rswo(0)
	end if
rswo.movenext
wend
response.write days_wo
%>
<td>
</TR>   	
<TR>
<TD>D.</TD>
<TD>Job Card Date to Despatch Date</TD>
<td>Production Dept.</td>
<td><%=days_del%></td>
</TR>
<TR>
<TD>E.</TD>
<TD>Delivery Committed to Party as mentioned in PO</TD>
<td>Mktg. Field Team</td>
<td>
<%
del_ATPDays = 0
tot_del = 0
sql1 = "select abs(datediff(dd,podate,DeliveryDATE)), abs(datediff(dd,SalesOrderdate,atpdate)),abs(datediff(dd,Deliverydate,atpdate)) from SalesOrder,salesOrderDet where customerid = "&ID&" and month(podate)="&month(date)&" and year(podate)="&year(date)&" and salesorder.salesOrderId = salesorderdet.salesOrderId"
set rs1 = conn.execute(sql1)
while not rs1.eof 
	if isnull(rs1(0)) then
		del_ATPDays=del_ATPDays
		tot_del=tot_del
	else
		tot_del=tot_del+rs1(1)
		del_ATPDays=del_ATPDays+rs1(0)
	end if
rs1.movenext
wend
response.write(del_ATPDays)
%>
</td>
</TR>
<TR>
<TD>F.</TD>
<TD>Marketing Int Order to Despatch Date</TD>
<td>Order Dwell Period in Co.</td>
<td><%=po_Invdays%></td>
</TR>
<TR>
<TD>G.</TD>
<TD>Despatch delays from Customer committed date</TD>
<td>Overall Delay Account</td>
<td><% if isnull(tot_del) then 
       Tot_del = 0 
		End If %>
	<%=tot_del%></td>
</TR>

</Table>
<br>
<%
aMonthValues = array(po_Deldays,po_Deldays, days_wo,days_del,del_ATPDays,po_Invdays,tot_del)
aMonthNames = array("A","B","C","D","E","F","G")
displayverticalgraph "DWELL REPORT","[Y-Axis]<br>DWELL TIME RECORD","[X-Axis]<br>NO. OF DAYS",aMonthValues,aMonthNames 
%>
</BODY>
<% set rs = nothing
set rs1 = nothing
set rswo = nothing
%>
</HTML>
