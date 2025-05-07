
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

%>
<div class="tab-pane" id="tab-pane-1">

<div class="tab-page"><h2 class="tab">Clearing Agent</h2>
<iframe src=document.asp?DID=2608&UID=<%=uid%>&FADD=True&FDEL=True&FVIEW=True&FEDIT=True&FILTER=False&FOFFHOLD=False&FREJECT=False&CID=<%=cid%> name=news8 width=100% height=600 target=news></iframe>
</div>

<div class="tab-page"><h2 class="tab">Letter Of Credit</h2>
<iframe src=document.asp?DID=1528&UID=<%=uid%>&FADD=True&FDEL=True&FVIEW=True&FEDIT=True&FILTER=False&FOFFHOLD=False&FREJECT=False&CID=<%=cid%> name=news9 width=100% height=600 target=news></iframe>
</div>
<div class="tab-page"><h2 class="tab">Help</h2>

<div class="tab-pane" id="tab-pane-2">
<div class="tab-page"><h2 class="tab">Training</h2>
<iframe src=helperpdemo/scmhelp/imports.htm?UID=<%=UID%>&CID=<%=CID%> name=news62 width=100% height=600 target=news></iframe>
</div>
<div class="tab-page"><h2 class="tab">Theory</h2>
<iframe src=helperpdemo/BLANK.htm?UID=<%=UID%>&CID=<%=CID%> name=news62 width=100% height=600 target=news></iframe>
</div>
<div class="tab-page"><h2 class="tab">References</h2>
<iframe src=helperpdemo/ref/scm/Importsref.asp?UID=<%=UID%>&CID=<%=CID%> name=news62 width=100% height=600 target=news></iframe>
</div>


</div>
</div>

</div>

<script type="text/javascript">
setupAllTabs();
</script>

</BODY>
</HTML>
