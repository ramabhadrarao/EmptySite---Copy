
<HTML>
<HEAD>
<TITLE>IRP TABBED WINDOWS (sales@erpweb)</TITLE>
<script type="text/javascript" src="js/tabpane.js"></script>
<link type="text/css" rel="StyleSheet" href="css/tab.webfx.css" />
</HEAD>
<BODY topmargin=0>
<%
UID=REQUEST("UID")
CID=REQUEST("CID")
DID=REQUEST("DID")
%>
<%
	UID=REQUEST("UID")
	CID=REQUEST("CID"): IF CID="" THEN CID=1
	ENO=REQUEST("ENO"): IF ENO="" OR ISNULL(ENO) THEN ENO=0
	CNO=REQUEST("CNO"): IF CNO="" OR ISNULL(CNO) THEN CNO=0
	SNO=REQUEST("SNO"): IF SNO="" OR ISNULL(SNO) THEN SNO=0
%>
<div class="tab-pane" id="tab-pane-1">

<div class="tab-page"><h2 class="tab"><b>SCM:Cockpit</b></h2>
<iframe src=scmcockpit.asp?CID=<%=CID%>&UID=<%=UID%>&ENO=<%=ENO%>&CNO=<%=CNO%>&SNO=<%=SNO%>  name=news1 width=100% height=600 target=news></iframe>
</div>

<!------
<div class="tab-page"><h2 class="tab"><b>SCM:ATGlance</b></h2>
<iframe src=crmatglance.asp?CID=<%=CID%>&UID=<%=UID%>&ENO=<%=ENO%>&CNO=<%=CNO%>&SNO=<%=SNO%>  name=news2 width=100% height=600 target=news></iframe>
</div>
--->

<div class="tab-page"><h2 class="tab"><b>SCM:OrgTree</b></h2>
<iframe src=scmtree.asp?DID=2613&UID=<%=UID%>&FADD=False&FDEL=False&FVIEW=False&FEDIT=False&FILTER=False&FOFFHOLD=False&FREJECT=False&CID=<%=CID%> name=news3 width=100% height=600 target=news></iframe>
</div>
<!--
<div class="tab-page"><h2 class="tab"><b>SCM:6-Sigma</b></h2>
<iframe src=sigmaprojectlist.asp?CID=<%=CID%>&UID=<%=UID%>&ENO=<%=ENO%>&CNO=<%=CNO%>&SNO=<%=SNO%> name=news4 width=100% height=600 target=news></iframe>
</div>

<div class="tab-page"><h2 class="tab"><b>SCM:Alerts</b></h2>
<iframe src=todaysalerts.asp?CID=<%=CID%>&UID=<%=UID%>&ENO=<%=ENO%>&CNO=<%=CNO%>&SNO=<%=SNO%> name=news5 width=100% height=600 target=news></iframe>
</div>

<div class="tab-page"><h2 class="tab"><b>SCM:Reports</b></h2>
<iframe src=scmreports.asp?CID=<%=CID%>&UID=<%=UID%>&ENO=<%=ENO%>&CNO=<%=CNO%>&SNO=<%=SNO%> name=news8 width=100% height=600 target=news></iframe>
</div>

<div class="tab-page"><h2 class="tab"><b>SUPPLIER WISE Reports</b></h2>
<iframe src=lreport.asp?DID=3005&UID=<%=UID%>&FADD=False&FDEL=False&FVIEW=False&FEDIT=False&FILTER=False&FOFFHOLD=False&FREJECT=False&CID=<%=CID%> name=news3 width=100% height=600 ></iframe>
</div>

<div class="tab-page"><h2 class="tab"><b>RAWMATERIAL Reports</b></h2>
<iframe src=lreport.asp?DID=3022&UID=<%=UID%>&FADD=False&FDEL=False&FVIEW=False&FEDIT=False&FILTER=False&FOFFHOLD=False&FREJECT=False&CID=<%=CID%> name=news3 width=100% height=600 ></iframe>
</div>


<div class="tab-page"><h2 class="tab">Help</h2>

<div class="tab-pane" id="tab-pane-2">
<div class="tab-page"><h2 class="tab">Training</h2>
<iframe src=helperpdemo/scmhelp/scmreports.htm?UID=<%=UID%>&CID=<%=CID%> name=news62 width=100% height=600 target=news></iframe>
</div>
<div class="tab-page"><h2 class="tab">Theory</h2>
<iframe src=helperpdemo/BLANK.htm?UID=<%=UID%>&CID=<%=CID%> name=news62 width=100% height=600 target=news></iframe>
</div>
<div class="tab-page"><h2 class="tab">References</h2>
<iframe src=helperpdemo/ref/scm/SCMReportsref.asp?UID=<%=UID%>&CID=<%=CID%> name=news62 width=100% height=600 target=news></iframe>
</div>


</div>
</div>
</div>
--->
<script type="text/javascript">
setupAllTabs();
</script>

</BODY>
</HTML>
