
<HTML>
<HEAD>
<TITLE>IRP ERPWEB MOBILEERP TABBED WINDOWS (sales@erpweb.com)</TITLE>
<script type="text/javascript" src="js/tabpane.js"></script>
<link type="text/css" rel="StyleSheet" href="css/tab.webfx.css" />
</HEAD>
<BODY topmargin=0>
<%
UID=REQUEST("UID")
CID=REQUEST("CID")
%>
<div class="tab-pane" id="tab-pane-1">

<div class="tab-page"><h2 class="tab">Bill Entry</h2>
<iframe src=document.asp?DID=1982&UID=<%=uid%>&FADD=True&FDEL=True&FVIEW=True&FEDIT=True&FILTER=True&FOFFHOLD=False&FREJECT=False&CID=<%=cid%> name=news2 width=100% height=600 target=news></iframe>
</div>

<div class="tab-page"><h2 class="tab">Tax Entry</h2>
<iframe src=document.asp?DID=1981&UID=<%=uid%>&FADD=True&FDEL=True&FVIEW=True&FEDIT=True&FILTER=True&FOFFHOLD=False&FREJECT=False&CID=<%=cid%> name=news2 width=100% height=600 target=news></iframe>
</div>

<div class="tab-page"><h2 class="tab">Bill Register</h2>
<iframe src=document.asp?DID=2329&UID=<%=uid%>&FADD=False&FDEL=False&FVIEW=True&FEDIT=False&FILTER=False&FOFFHOLD=False&FREJECT=False&CID=<%=cid%> name=news4 width=100% height=600 target=news></iframe>
</div>

<!----
<div class="tab-page"><h2 class="tab">Help</h2>

<div class="tab-pane" id="tab-pane-2">
<div class="tab-page"><h2 class="tab">Training</h2>
<iframe src=helperpdemo/scmhelp/pi.htm?UID=<%=UID%>&CID=<%=CID%> name=news62 width=100% height=600 target=news></iframe>
</div>
<div class="tab-page"><h2 class="tab">Theory</h2>
<iframe src=helperpdemo/BLANK.htm?UID=<%=UID%>&CID=<%=CID%> name=news62 width=100% height=600 target=news></iframe>
</div>
<div class="tab-page"><h2 class="tab">References</h2>
<iframe src=helperpdemo/ref/scm/PIref.asp?UID=<%=UID%>&CID=<%=CID%> name=news62 width=100% height=600 target=news></iframe>
</div>


</div>
</div>
--->
</div>

<script type="text/javascript">
setupAllTabs();
</script>

</BODY>
</HTML>
