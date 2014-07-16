<%

		' Display elapsed time
		If EW_DEBUG_ENABLED Then Response.Write ew_CalcElapsedTime(StartTimer)
%>
<% If gsExport = "" Then %>
				<p>&nbsp;</p>			
			<!-- right column (end) -->
	    </td>	
		</tr>
	</table>
	<!-- content (end) -->
	<!-- footer (begin) --><!-- *** Note: Only licensed users are allowed to remove or change the following copyright statement. *** -->
	<div class="ewFooterRow ">	
			<div class="muted">&nbsp;<%= Language.ProjectPhrase("FooterText") %></div>
		<!-- Place other links, for example, disclaimer, here -->		
	</div>
	<!-- footer (end) -->	
</div>
<div class="yui-tt" id="ewTooltipDiv" style="visibility: hidden; border: 0px;" name="ewTooltipDivDiv"></div>
<% End If %>
<% If gsExport = "" Or gsExport = "print" Then %>
<script type="text/javascript">
<!--
ewDom.getElementsByClassName(EW_TABLE_CLASS, "TABLE", null, ew_SetupTable); // Init the table
ewDom.getElementsByClassName(EW_GRID_CLASS, "TABLE", null, ew_SetupGrid); // Init grid
ew_InitTooltipDiv(); // init tooltip div
//-->
</script>
<script language="JavaScript" type="text/javascript">
<!--
// Write your global startup script here
// document.write("page loaded");
//-->
</script>
<% End If %>
<!--#include file="Scripts.asp"-->

</body>
</html>
