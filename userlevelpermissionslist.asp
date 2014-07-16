<%@ CodePage="1252" LCID="2057" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="ewcfg9.asp"-->
<!--#include file="userlevelpermissionsinfo.asp"-->
<!--#include file="usersinfo.asp"-->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim UserLevelPermissions_list
Set UserLevelPermissions_list = New cUserLevelPermissions_list
Set Page = UserLevelPermissions_list

' Page init processing
Call UserLevelPermissions_list.Page_Init()

' Page main processing
Call UserLevelPermissions_list.Page_Main()
%>
<!--#include file="header.asp"-->
<% If UserLevelPermissions.Export = "" Then %>
<script type="text/javascript">
<!--
// Create page object
var UserLevelPermissions_list = new ew_Page("UserLevelPermissions_list");
// page properties
UserLevelPermissions_list.PageID = "list"; // page ID
UserLevelPermissions_list.FormID = "fUserLevelPermissionslist"; // form ID
var EW_PAGE_ID = UserLevelPermissions_list.PageID; // for backward compatibility
// extend page with Form_CustomValidate function
UserLevelPermissions_list.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
UserLevelPermissions_list.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
UserLevelPermissions_list.ValidateRequired = true; // uses JavaScript validation
<% Else %>
UserLevelPermissions_list.ValidateRequired = false; // no JavaScript validation
<% End If %>
//-->
</script>
<script type="text/javascript">
<!--
var ew_DHTMLEditors = [];
//-->
</script>
<script language="JavaScript" type="text/javascript">
<!--
// Write your client script here, no need to add script tags.
//-->
</script>
<% End If %>
<% If (UserLevelPermissions.Export = "") Or (EW_EXPORT_MASTER_RECORD And UserLevelPermissions.Export = "print") Then %>
<% End If %>
<% UserLevelPermissions_list.ShowPageHeader() %>
<%

' Load recordset
Set UserLevelPermissions_list.Recordset = UserLevelPermissions_list.LoadRecordset()
	UserLevelPermissions_list.TotalRecs = UserLevelPermissions_list.Recordset.RecordCount
	UserLevelPermissions_list.StartRec = 1
	If UserLevelPermissions_list.DisplayRecs <= 0 Then ' Display all records
		UserLevelPermissions_list.DisplayRecs = UserLevelPermissions_list.TotalRecs
	End If
	If Not (UserLevelPermissions.ExportAll And UserLevelPermissions.Export <> "") Then
		UserLevelPermissions_list.SetUpStartRec() ' Set up start record position
	End If
%>
<p class="aspmaker ewTitle" style="white-space: nowrap;"><%= Language.Phrase("TblTypeTABLE") %><%= UserLevelPermissions.TableCaption %>
&nbsp;&nbsp;<% UserLevelPermissions_list.ExportOptions.Render "body", "" %>
</p>
<% If Security.CanSearch Then %>
<% If UserLevelPermissions.Export = "" And UserLevelPermissions.CurrentAction = "" Then %>
<a href="javascript:ew_ToggleSearchPanel(UserLevelPermissions_list);" style="text-decoration: none;"><img id="UserLevelPermissions_list_SearchImage" src="images/collapse.gif" alt="" width="9" height="9" border="0"></a><span class="aspmaker">&nbsp;<%= Language.Phrase("Search") %></span><br>
<div id="UserLevelPermissions_list_SearchPanel">
<form name="fUserLevelPermissionslistsrch" id="fUserLevelPermissionslistsrch" class="ewForm" action="<%= ew_CurrentPage %>">
<input type="hidden" id="t" name="t" value="UserLevelPermissions">
<div class="ewBasicSearch">
<div id="xsr_1" class="ewCssTableRow">
	<input type="text" name="<%= EW_TABLE_BASIC_SEARCH %>" id="<%= EW_TABLE_BASIC_SEARCH %>" size="20" value="<%= ew_HtmlEncode(UserLevelPermissions.SessionBasicSearchKeyword) %>">
	<input type="Submit" name="Submit" id="Submit" value="<%= ew_BtnCaption(Language.Phrase("QuickSearchBtn")) %>">&nbsp;
	<a href="<%= UserLevelPermissions_list.PageUrl %>cmd=reset"><%= Language.Phrase("ShowAll") %></a>&nbsp;
</div>
</div>
</form>
</div>
<% End If %>
<% End If %>
<% UserLevelPermissions_list.ShowMessage %>
<br>
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<form name="fUserLevelPermissionslist" id="fUserLevelPermissionslist" class="ewForm" action="" method="post">
<input type="hidden" name="t" id="t" value="UserLevelPermissions">
<div id="gmp_UserLevelPermissions" class="ewGridMiddlePanel">
<% If UserLevelPermissions_list.TotalRecs > 0 Then %>
<table cellspacing="0" data-rowhighlightclass="ewTableHighlightRow" data-rowselectclass="ewTableSelectRow" data-roweditclass="ewTableEditRow" class="ewTable ewTableSeparate">
<%= UserLevelPermissions.TableCustomInnerHTML %>
<thead><!-- Table header -->
	<tr class="ewTableHeader">
<%
Call UserLevelPermissions_list.RenderListOptions()

' Render list options (header, left)
UserLevelPermissions_list.ListOptions.Render "header", "left"
%>
<% If UserLevelPermissions.UserLevelID.Visible Then ' UserLevelID %>
	<% If UserLevelPermissions.SortUrl(UserLevelPermissions.UserLevelID) = "" Then %>
		<td><%= UserLevelPermissions.UserLevelID.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= UserLevelPermissions.SortUrl(UserLevelPermissions.UserLevelID) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= UserLevelPermissions.UserLevelID.FldCaption %></td><td style="width: 10px;"><% If UserLevelPermissions.UserLevelID.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf UserLevelPermissions.UserLevelID.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If UserLevelPermissions.zTableName.Visible Then ' TableName %>
	<% If UserLevelPermissions.SortUrl(UserLevelPermissions.zTableName) = "" Then %>
		<td><%= UserLevelPermissions.zTableName.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= UserLevelPermissions.SortUrl(UserLevelPermissions.zTableName) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= UserLevelPermissions.zTableName.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If UserLevelPermissions.zTableName.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf UserLevelPermissions.zTableName.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If UserLevelPermissions.Permission.Visible Then ' Permission %>
	<% If UserLevelPermissions.SortUrl(UserLevelPermissions.Permission) = "" Then %>
		<td><%= UserLevelPermissions.Permission.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= UserLevelPermissions.SortUrl(UserLevelPermissions.Permission) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= UserLevelPermissions.Permission.FldCaption %></td><td style="width: 10px;"><% If UserLevelPermissions.Permission.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf UserLevelPermissions.Permission.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<%

' Render list options (header, right)
UserLevelPermissions_list.ListOptions.Render "header", "right"
%>
	</tr>
</thead>
<tbody><!-- Table body -->
<%
If (UserLevelPermissions.ExportAll And UserLevelPermissions.Export <> "") Then
	UserLevelPermissions_list.StopRec = UserLevelPermissions_list.TotalRecs
Else

	' Set the last record to display
	If UserLevelPermissions_list.TotalRecs > UserLevelPermissions_list.StartRec + UserLevelPermissions_list.DisplayRecs - 1 Then
		UserLevelPermissions_list.StopRec = UserLevelPermissions_list.StartRec + UserLevelPermissions_list.DisplayRecs - 1
	Else
		UserLevelPermissions_list.StopRec = UserLevelPermissions_list.TotalRecs
	End If
End If

' Move to first record
UserLevelPermissions_list.RecCnt = UserLevelPermissions_list.StartRec - 1
If Not UserLevelPermissions_list.Recordset.Eof Then
	UserLevelPermissions_list.Recordset.MoveFirst
	If UserLevelPermissions_list.StartRec > 1 Then UserLevelPermissions_list.Recordset.Move UserLevelPermissions_list.StartRec - 1
ElseIf Not UserLevelPermissions.AllowAddDeleteRow And UserLevelPermissions_list.StopRec = 0 Then
	UserLevelPermissions_list.StopRec = UserLevelPermissions.GridAddRowCount
End If

' Initialize Aggregate
UserLevelPermissions.RowType = EW_ROWTYPE_AGGREGATEINIT
Call UserLevelPermissions.ResetAttrs()
Call UserLevelPermissions_list.RenderRow()
UserLevelPermissions_list.RowCnt = 0

' Output date rows
Do While CLng(UserLevelPermissions_list.RecCnt) < CLng(UserLevelPermissions_list.StopRec)
	UserLevelPermissions_list.RecCnt = UserLevelPermissions_list.RecCnt + 1
	If CLng(UserLevelPermissions_list.RecCnt) >= CLng(UserLevelPermissions_list.StartRec) Then
		UserLevelPermissions_list.RowCnt = UserLevelPermissions_list.RowCnt + 1

	' Set up key count
	UserLevelPermissions_list.KeyCount = UserLevelPermissions_list.RowIndex
	Call UserLevelPermissions.ResetAttrs()
	UserLevelPermissions.CssClass = ""
	If UserLevelPermissions.CurrentAction = "gridadd" Then
	Else
		Call UserLevelPermissions_list.LoadRowValues(UserLevelPermissions_list.Recordset) ' Load row values
	End If
	UserLevelPermissions.RowType = EW_ROWTYPE_VIEW ' Render view
	UserLevelPermissions.RowAttrs.AddAttributes Array(Array("onmouseover", "ew_MouseOver(event, this);"), Array("onmouseout", "ew_MouseOut(event, this);"), Array("onclick", "ew_Click(event, this);"))

	' Render row
	Call UserLevelPermissions_list.RenderRow()

	' Render list options
	Call UserLevelPermissions_list.RenderListOptions()
%>
	<tr<%= UserLevelPermissions.RowAttributes %>>
<%

' Render list options (body, left)
UserLevelPermissions_list.ListOptions.Render "body", "left"
%>
	<% If UserLevelPermissions.UserLevelID.Visible Then ' UserLevelID %>
		<td<%= UserLevelPermissions.UserLevelID.CellAttributes %>>
<div<%= UserLevelPermissions.UserLevelID.ViewAttributes %>><%= UserLevelPermissions.UserLevelID.ListViewValue %></div>
<a name="<%= UserLevelPermissions_list.PageObjName & "_row_" & UserLevelPermissions_list.RowCnt %>" id="<%= UserLevelPermissions_list.PageObjName & "_row_" & UserLevelPermissions_list.RowCnt %>"></a></td>
	<% End If %>
	<% If UserLevelPermissions.zTableName.Visible Then ' TableName %>
		<td<%= UserLevelPermissions.zTableName.CellAttributes %>>
<div<%= UserLevelPermissions.zTableName.ViewAttributes %>><%= UserLevelPermissions.zTableName.ListViewValue %></div>
</td>
	<% End If %>
	<% If UserLevelPermissions.Permission.Visible Then ' Permission %>
		<td<%= UserLevelPermissions.Permission.CellAttributes %>>
<div<%= UserLevelPermissions.Permission.ViewAttributes %>><%= UserLevelPermissions.Permission.ListViewValue %></div>
</td>
	<% End If %>
<%

' Render list options (body, right)
UserLevelPermissions_list.ListOptions.Render "body", "right"
%>
	</tr>
<%
	End If
	If UserLevelPermissions.CurrentAction <> "gridadd" Then
		UserLevelPermissions_list.Recordset.MoveNext()
	End If
Loop
%>
</tbody>
</table>
<% End If %>
</div>
</form>
<%

' Close recordset and connection
UserLevelPermissions_list.Recordset.Close
Set UserLevelPermissions_list.Recordset = Nothing
%>
<% If UserLevelPermissions.Export = "" Then %>
<div class="ewGridLowerPanel">
<% If UserLevelPermissions.CurrentAction <> "gridadd" And UserLevelPermissions.CurrentAction <> "gridedit" Then %>
<form name="ewpagerform" id="ewpagerform" class="ewForm" action="<%= ew_CurrentPage %>">
<table border="0" cellspacing="0" cellpadding="0" class="ewPager">
	<tr>
		<td>
<% If Not IsObject(UserLevelPermissions_list.Pager) Then Set UserLevelPermissions_list.Pager = ew_NewPrevNextPager(UserLevelPermissions_list.StartRec, UserLevelPermissions_list.DisplayRecs, UserLevelPermissions_list.TotalRecs) %>
<% If UserLevelPermissions_list.Pager.RecordCount > 0 Then %>
	<table border="0" cellspacing="0" cellpadding="0"><tr><td><span class="aspmaker"><%= Language.Phrase("Page") %>&nbsp;</span></td>
<!--first page button-->
	<% If UserLevelPermissions_list.Pager.FirstButton.Enabled Then %>
	<td><a href="<%= UserLevelPermissions_list.PageUrl %>start=<%= UserLevelPermissions_list.Pager.FirstButton.Start %>"><img src="images/first.gif" alt="<%= Language.Phrase("PagerFirst") %>" width="16" height="16" border="0"></a></td>
	<% Else %>
	<td><img src="images/firstdisab.gif" alt="<%= Language.Phrase("PagerFirst") %>" width="16" height="16" border="0"></td>
	<% End If %>
<!--previous page button-->
	<% If UserLevelPermissions_list.Pager.PrevButton.Enabled Then %>
	<td><a href="<%= UserLevelPermissions_list.PageUrl %>start=<%= UserLevelPermissions_list.Pager.PrevButton.Start %>"><img src="images/prev.gif" alt="<%= Language.Phrase("PagerPrevious") %>" width="16" height="16" border="0"></a></td>
	<% Else %>
	<td><img src="images/prevdisab.gif" alt="<%= Language.Phrase("PagerPrevious") %>" width="16" height="16" border="0"></td>
	<% End If %>
<!--current page number-->
	<td><input type="text" name="<%= EW_TABLE_PAGE_NO %>" id="<%= EW_TABLE_PAGE_NO %>" value="<%= UserLevelPermissions_list.Pager.CurrentPage %>" size="4"></td>
<!--next page button-->
	<% If UserLevelPermissions_list.Pager.NextButton.Enabled Then %>
	<td><a href="<%= UserLevelPermissions_list.PageUrl %>start=<%= UserLevelPermissions_list.Pager.NextButton.Start %>"><img src="images/next.gif" alt="<%= Language.Phrase("PagerNext") %>" width="16" height="16" border="0"></a></td>
	<% Else %>
	<td><img src="images/nextdisab.gif" alt="<%= Language.Phrase("PagerNext") %>" width="16" height="16" border="0"></td>
	<% End If %>
<!--last page button-->
	<% If UserLevelPermissions_list.Pager.LastButton.Enabled Then %>
	<td><a href="<%= UserLevelPermissions_list.PageUrl %>start=<%= UserLevelPermissions_list.Pager.LastButton.Start %>"><img src="images/last.gif" alt="<%= Language.Phrase("PagerLast") %>" width="16" height="16" border="0"></a></td>	
	<% Else %>
	<td><img src="images/lastdisab.gif" alt="<%= Language.Phrase("PagerLast") %>" width="16" height="16" border="0"></td>
	<% End If %>
	<td><span class="aspmaker">&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= UserLevelPermissions_list.Pager.PageCount %></span></td>
	</tr></table>
	</td>	
	<td>&nbsp;&nbsp;&nbsp;&nbsp;</td>
	<td>
	<span class="aspmaker"><%= Language.Phrase("Record") %>&nbsp;<%= UserLevelPermissions_list.Pager.FromIndex %>&nbsp;<%= Language.Phrase("To") %>&nbsp;<%= UserLevelPermissions_list.Pager.ToIndex %>&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= UserLevelPermissions_list.Pager.RecordCount %></span>
<% Else %>
	<% If Security.CanList Then %>
	<% If UserLevelPermissions_list.SearchWhere = "0=101" Then %>
	<span class="aspmaker"><%= Language.Phrase("EnterSearchCriteria") %></span>
	<% Else %>
	<span class="aspmaker"><%= Language.Phrase("NoRecord") %></span>
	<% End If %>
	<% Else %>
	<span class="aspmaker"><%= Language.Phrase("NoPermission") %></span>
	<% End If %>
<% End If %>
		</td>
	</tr>
</table>
</form>
<% End If %>
<span class="aspmaker">
<% If Security.CanAdd Then %>
<a class="ewGridLink" href="<%= UserLevelPermissions_list.AddUrl %>"><%= Language.Phrase("AddLink") %></a>&nbsp;&nbsp;
<% End If %>
<% If UserLevelPermissions_list.TotalRecs > 0 Then %>
<% If Security.CanDelete Then %>
<a class="ewGridLink" href="" onclick="ew_SubmitSelected(document.fUserLevelPermissionslist, '<%= UserLevelPermissions_list.MultiDeleteUrl %>');return false;"><%= Language.Phrase("DeleteSelectedLink") %></a>&nbsp;&nbsp;
<% End If %>
<% End If %>
</span>
</div>
<% End If %>
</td></tr></table>
<% If UserLevelPermissions.Export = "" And UserLevelPermissions.CurrentAction = "" Then %>
<% End If %>
<%
UserLevelPermissions_list.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If UserLevelPermissions.Export = "" Then %>
<script language="JavaScript" type="text/javascript">
<!--
// Write your table-specific startup script here
// document.write("page loaded");
//-->
</script>
<% End If %>
<!--#include file="footer.asp"-->
<%

' Drop page object
Set UserLevelPermissions_list = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cUserLevelPermissions_list

	' Page ID
	Public Property Get PageID()
		PageID = "list"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "UserLevelPermissions"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "UserLevelPermissions_list"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If UserLevelPermissions.UseTokenInUrl Then PageUrl = PageUrl & "t=" & UserLevelPermissions.TableVar & "&" ' add page token
	End Property

	' Common urls
	Dim AddUrl
	Dim EditUrl
	Dim CopyUrl
	Dim DeleteUrl
	Dim ViewUrl
	Dim ListUrl

	' Export urls
	Dim ExportPrintUrl
	Dim ExportHtmlUrl
	Dim ExportExcelUrl
	Dim ExportWordUrl
	Dim ExportXmlUrl
	Dim ExportCsvUrl

	' Inline urls
	Dim InlineAddUrl
	Dim InlineCopyUrl
	Dim InlineEditUrl
	Dim GridAddUrl
	Dim GridEditUrl
	Dim MultiDeleteUrl
	Dim MultiUpdateUrl

	' Message
	Public Property Get Message()
		Message = Session(EW_SESSION_MESSAGE)
	End Property

	Public Property Let Message(v)
		Dim msg
		msg = Session(EW_SESSION_MESSAGE)
		Call ew_AddMessage(msg, v)
		Session(EW_SESSION_MESSAGE) = msg
	End Property

	Public Property Get FailureMessage()
		FailureMessage = Session(EW_SESSION_FAILURE_MESSAGE)
	End Property

	Public Property Let FailureMessage(v)
		Dim msg
		msg = Session(EW_SESSION_FAILURE_MESSAGE)
		Call ew_AddMessage(msg, v)
		Session(EW_SESSION_FAILURE_MESSAGE) = msg
	End Property

	Public Property Get SuccessMessage()
		SuccessMessage = Session(EW_SESSION_SUCCESS_MESSAGE)
	End Property

	Public Property Let SuccessMessage(v)
		Dim msg
		msg = Session(EW_SESSION_SUCCESS_MESSAGE)
		Call ew_AddMessage(msg, v)
		Session(EW_SESSION_SUCCESS_MESSAGE) = msg
	End Property

	' Show Message
	Public Sub ShowMessage()
		Dim sMessage
		sMessage = Message
		Call Message_Showing(sMessage, "")
		If sMessage <> "" Then Response.Write "<p class=""ewMessage"">" & sMessage & "</p>"
		Session(EW_SESSION_MESSAGE) = "" ' Clear message in Session

		' Success message
		Dim sSuccessMessage
		sSuccessMessage = SuccessMessage
		Call Message_Showing(sSuccessMessage, "success")
		If sSuccessMessage <> "" Then Response.Write "<p class=""ewSuccessMessage"">" & sSuccessMessage & "</p>"
		Session(EW_SESSION_SUCCESS_MESSAGE) = "" ' Clear message in Session

		' Failure message
		Dim sErrorMessage
		sErrorMessage = FailureMessage
		Call Message_Showing(sErrorMessage, "failure")
		If sErrorMessage <> "" Then Response.Write "<p class=""ewErrorMessage"">" & sErrorMessage & "</p>"
		Session(EW_SESSION_FAILURE_MESSAGE) = "" ' Clear message in Session
	End Sub
	Dim PageHeader
	Dim PageFooter

	' Show Page Header
	Public Sub ShowPageHeader()
		Dim sHeader
		sHeader = PageHeader
		Call Page_DataRendering(sHeader)
		If sHeader <> "" Then ' Header exists, display
			Response.Write "<p class=""aspmaker"">" & sHeader & "</p>"
		End If
	End Sub

	' Show Page Footer
	Public Sub ShowPageFooter()
		Dim sFooter
		sFooter = PageFooter
		Call Page_DataRendered(sFooter)
		If sFooter <> "" Then ' Footer exists, display
			Response.Write "<p class=""aspmaker"">" & sFooter & "</p>"
		End If
	End Sub

	' -----------------------
	'  Validate Page request
	'
	Public Function IsPageRequest()
		If UserLevelPermissions.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (UserLevelPermissions.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (UserLevelPermissions.TableVar = Request.QueryString("t"))
			End If
		Else
			IsPageRequest = True
		End If
	End Function

	' -----------------------------------------------------------------
	'  Class initialize
	'  - init objects
	'  - open ADO connection
	'
	Private Sub Class_Initialize()
		If IsEmpty(StartTimer) Then StartTimer = Timer ' Init start time

		' Initialize language object
		If IsEmpty(Language) Then
			Set Language = New cLanguage
			Call Language.LoadPhrases()
		End If

		' Initialize table object
		If IsEmpty(UserLevelPermissions) Then Set UserLevelPermissions = New cUserLevelPermissions
		Set Table = UserLevelPermissions

		' Initialize urls
		ExportPrintUrl = PageUrl & "export=print"
		ExportExcelUrl = PageUrl & "export=excel"
		ExportWordUrl = PageUrl & "export=word"
		ExportHtmlUrl = PageUrl & "export=html"
		ExportXmlUrl = PageUrl & "export=xml"
		ExportCsvUrl = PageUrl & "export=csv"
		AddUrl = "userlevelpermissionsadd.asp"
		InlineAddUrl = PageUrl & "a=add"
		GridAddUrl = PageUrl & "a=gridadd"
		GridEditUrl = PageUrl & "a=gridedit"
		MultiDeleteUrl = "userlevelpermissionsdelete.asp"
		MultiUpdateUrl = "userlevelpermissionsupdate.asp"

		' Initialize other table object
		If IsEmpty(Users) Then Set Users = New cUsers

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "list"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "UserLevelPermissions"

		' Open connection to the database
		If IsEmpty(Conn) Then Call ew_Connect()

		' Initialize list options
		Set ListOptions = New cListOptions

		' Export options
		Set ExportOptions = New cListOptions
		ExportOptions.Tag = "span"
		ExportOptions.Separator = "&nbsp;&nbsp;"
	End Sub

	' -----------------------------------------------------------------
	'  Subroutine Page_Init
	'  - called before page main
	'  - check Security
	'  - set up response header
	'  - call page load events
	'
	Sub Page_Init()
		Set UserProfile = New cUserProfile
		UserProfile.LoadProfile Session(EW_SESSION_USER_PROFILE)
		Set Security = New cAdvancedSecurity
		If Not Security.IsLoggedIn() Then Call Security.AutoLogin()
		If Not Security.IsLoggedIn() Then
			Call Security.SaveLastUrl()
			Call Page_Terminate("login.asp")
		End If

		' Table Permission loading event
		Call Security.TablePermission_Loading()
		Call Security.LoadCurrentUserLevel(TableName)

		' Table Permission loaded event
		Call Security.TablePermission_Loaded()
		If Not Security.CanAdmin Then
			Call Security.SaveLastUrl()
			Call Page_Terminate( "login.asp")
		End If

		' Get grid add count
		Dim gridaddcnt
		gridaddcnt = Request.QueryString(EW_TABLE_GRID_ADD_ROW_COUNT)
		If IsNumeric(gridaddcnt) Then
			If gridaddcnt > 0 Then
				UserLevelPermissions.GridAddRowCount = gridaddcnt
			End If
		End If

		' Set up list options
		SetupListOptions()

		' Global page loading event (in userfn7.asp)
		Call Page_Loading()

		' Page load event, used in current page
		Call Page_Load()
	End Sub

	' -----------------------------------------------------------------
	'  Class terminate
	'  - clean up page object
	'
	Private Sub Class_Terminate()
		Call Page_Terminate("")
	End Sub

	' -----------------------------------------------------------------
	'  Subroutine Page_Terminate
	'  - called when exit page
	'  - clean up ADO connection and objects
	'  - if url specified, redirect to url
	'
	Sub Page_Terminate(url)

		' Page unload event, used in current page
		Call Page_Unload()

		' Global page unloaded event (in userfn60.asp)
		Call Page_Unloaded()
		Dim sRedirectUrl
		sReDirectUrl = url
		Call Page_Redirecting(sReDirectUrl)
		If Not (Conn Is Nothing) Then Conn.Close ' Close Connection
		Set Conn = Nothing
		Set Security = Nothing
		Set UserProfile = Nothing
		Set UserLevelPermissions = Nothing
		Set ListOptions = Nothing
		Set ObjForm = Nothing

		' Go to url if specified
		If sReDirectUrl <> "" Then
			If Response.Buffer Then Response.Clear
			Response.Redirect sReDirectUrl
		End If
	End Sub

	'
	'  Subroutine Page_Terminate (End)
	' ----------------------------------------

	Dim DisplayRecs ' Number of display records
	Dim StartRec, StopRec, TotalRecs, RecRange
	Dim SearchWhere
	Dim RecCnt
	Dim EditRowCnt
	Dim RowCnt, RowIndex
	Dim RecPerRow, ColCnt
	Dim KeyCount
	Dim RowAction
	Dim RowOldKey ' Row old key (for copy)
	Dim DbMasterFilter, DbDetailFilter
	Dim MasterRecordExists
	Dim ListOptions
	Dim ExportOptions
	Dim MultiSelectKey
	Dim RestoreSearch
	Dim Recordset, OldRecordset

	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()
		DisplayRecs = 50
		RecRange = 10
		RecCnt = 0 ' Record count
		KeyCount = 0 ' Key count

		' Search filters
		Dim sSrchAdvanced, sSrchBasic, sFilter
		sSrchAdvanced = "" ' Advanced search filter
		sSrchBasic = "" ' Basic search filter
		SearchWhere = "" ' Search where clause
		sFilter = ""

		' Master/Detail
		DbMasterFilter = "" ' Master filter
		DbDetailFilter = "" ' Detail filter
		If IsPageRequest Then ' Validate request

			' Handle reset command
			ResetCmd()

			' Hide all options
			If UserLevelPermissions.Export <> "" Or UserLevelPermissions.CurrentAction = "gridadd" Or UserLevelPermissions.CurrentAction = "gridedit" Then
				ListOptions.HideAllOptions()
				ExportOptions.HideAllOptions()
			End If

			' Get basic search values
			Call LoadBasicSearchValues()

			' Restore search parms from Session
			Call RestoreSearchParms()

			' Call Recordset SearchValidated event
			Call UserLevelPermissions.Recordset_SearchValidated()

			' Set Up Sorting Order
			SetUpSortOrder()

			' Get basic search criteria
			If gsSearchError = "" Then
				sSrchBasic = BasicSearchWhere()
			End If
		End If ' End Validate Request

		' Restore display records
		If UserLevelPermissions.RecordsPerPage <> "" Then
			DisplayRecs = UserLevelPermissions.RecordsPerPage ' Restore from Session
		Else
			DisplayRecs = 50 ' Load default
		End If

		' Load Sorting Order
		LoadSortOrder()

		' Build search criteria
		Call ew_AddFilter(SearchWhere, sSrchAdvanced)
		Call ew_AddFilter(SearchWhere, sSrchBasic)

		' Call Recordset Searching event
		Call UserLevelPermissions.Recordset_Searching(SearchWhere)

		' Save search criteria
		If SearchWhere <> "" Then
			If sSrchBasic = "" Then Call ResetBasicSearchParms()
			UserLevelPermissions.SearchWhere = SearchWhere ' Save to Session
			If Not RestoreSearch Then
				StartRec = 1 ' Reset start record counter
				UserLevelPermissions.StartRecordNumber = StartRec
			End If
		Else
			SearchWhere = UserLevelPermissions.SearchWhere
		End If
		sFilter = ""
		If Not Security.CanList Then
			sFilter = "(0=1)" ' Filter all records
		End If
		Call ew_AddFilter(sFilter, DbDetailFilter)
		Call ew_AddFilter(sFilter, SearchWhere)

		' Set up filter in Session
		UserLevelPermissions.SessionWhere = sFilter
		UserLevelPermissions.CurrentFilter = ""
	End Sub

	' -----------------------------------------------------------------
	' Return Basic Search sql
	'
	Function BasicSearchSQL(Keyword)
		Dim sWhere
		sWhere = ""
			Call BuildBasicSearchSQL(sWhere, UserLevelPermissions.zTableName, Keyword)
		BasicSearchSQL = sWhere
	End Function

	' -----------------------------------------------------------------
	' Build basic search sql
	'
	Sub BuildBasicSearchSql(Where, Fld, Keyword)
		Dim sFldExpression, lFldDataType
		Dim sWrk
		If Fld.FldVirtualExpression <> "" Then
			sFldExpression = Fld.FldVirtualExpression
		Else
			sFldExpression = Fld.FldExpression
		End If
		lFldDataType = Fld.FldDataType
		If Fld.FldIsVirtual Then lFldDataType = EW_DATATYPE_STRING
		If lFldDataType = EW_DATATYPE_NUMBER Then
			sWrk = sFldExpression & " = " & ew_QuotedValue(Keyword, lFldDataType)
		Else
			sWrk = sFldExpression & ew_Like(ew_QuotedValue("%" & Keyword & "%", lFldDataType))
		End If
		If Where <> "" Then Where = Where & " OR "
		Where = Where & sWrk
	End Sub

	' -----------------------------------------------------------------
	' Return Basic Search Where based on search keyword and type
	'
	Function BasicSearchWhere()
		Dim sSearchStr, sSearchKeyword, sSearchType
		Dim sSearch, arKeyword, sKeyword
		sSearchStr = ""
		If Not Security.CanSearch Then
			BasicSearchWhere = ""
			Exit Function
		End If
		sSearchKeyword = UserLevelPermissions.BasicSearchKeyword
		sSearchType = UserLevelPermissions.BasicSearchType
		If sSearchKeyword <> "" Then
			sSearch = Trim(sSearchKeyword)
			If sSearchType <> "" Then
				While InStr(sSearch, "  ") > 0
					sSearch = Replace(sSearch, "  ", " ")
				Wend
				arKeyword = Split(Trim(sSearch), " ")
				For Each sKeyword In arKeyword
					If sSearchStr <> "" Then sSearchStr = sSearchStr & " " & sSearchType & " "
					sSearchStr = sSearchStr & "(" & BasicSearchSQL(sKeyword) & ")"
				Next
			Else
				sSearchStr = BasicSearchSQL(sSearch)
			End If
		End If
		If sSearchKeyword <> "" then
			UserLevelPermissions.SessionBasicSearchKeyword = sSearchKeyword
			UserLevelPermissions.SessionBasicSearchType = sSearchType
		End If
		BasicSearchWhere = sSearchStr
	End Function

	' -----------------------------------------------------------------
	' Clear all search parameters
	'
	Sub ResetSearchParms()

		' Clear search where
		SearchWhere = ""
		UserLevelPermissions.SearchWhere = SearchWhere

		' Clear basic search parameters
		Call ResetBasicSearchParms()
	End Sub

	' -----------------------------------------------------------------
	' Clear all basic search parameters
	'
	Sub ResetBasicSearchParms()

		' Clear basic search parameters
		UserLevelPermissions.SessionBasicSearchKeyword = ""
		UserLevelPermissions.SessionBasicSearchType = ""
	End Sub

	' -----------------------------------------------------------------
	' Restore all search parameters
	'
	Sub RestoreSearchParms()
		Dim bRestore
		bRestore = True
		If UserLevelPermissions.BasicSearchKeyword & "" <> "" Then bRestore = False
		RestoreSearch = bRestore
		If bRestore Then

			' Restore basic search values
			UserLevelPermissions.BasicSearchKeyword = UserLevelPermissions.SessionBasicSearchKeyword
			UserLevelPermissions.BasicSearchType = UserLevelPermissions.SessionBasicSearchType
		End If
	End Sub

	' -----------------------------------------------------------------
	' Set up Sort parameters based on Sort Links clicked
	'
	Sub SetUpSortOrder()
		Dim sOrderBy
		Dim sSortField, sLastSort, sThisSort
		Dim bCtrl

		' Check for an Order parameter
		If Request.QueryString("order").Count > 0 Then
			UserLevelPermissions.CurrentOrder = Request.QueryString("order")
			UserLevelPermissions.CurrentOrderType = Request.QueryString("ordertype")

			' Field UserLevelID
			Call UserLevelPermissions.UpdateSort(UserLevelPermissions.UserLevelID)

			' Field TableName
			Call UserLevelPermissions.UpdateSort(UserLevelPermissions.zTableName)

			' Field Permission
			Call UserLevelPermissions.UpdateSort(UserLevelPermissions.Permission)
			UserLevelPermissions.StartRecordNumber = 1 ' Reset start position
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load Sort Order parameters
	'
	Sub LoadSortOrder()
		Dim sOrderBy
		sOrderBy = UserLevelPermissions.SessionOrderBy ' Get order by from Session
		If sOrderBy = "" Then
			If UserLevelPermissions.SqlOrderBy <> "" Then
				sOrderBy = UserLevelPermissions.SqlOrderBy
				UserLevelPermissions.SessionOrderBy = sOrderBy
			End If
		End If
	End Sub

	' -----------------------------------------------------------------
	' Reset command based on querystring parameter cmd=
	' - RESET: reset search parameters
	' - RESETALL: reset search & master/detail parameters
	' - RESETSORT: reset sort parameters
	'
	Sub ResetCmd()
		Dim sCmd

		' Get reset cmd
		If Request.QueryString("cmd").Count > 0 Then
			sCmd = Request.QueryString("cmd")

			' Reset search criteria
			If LCase(sCmd) = "reset" Or LCase(sCmd) = "resetall" Then
				Call ResetSearchParms()
			End If

			' Reset Sort Criteria
			If LCase(sCmd) = "resetsort" Then
				Dim sOrderBy
				sOrderBy = ""
				UserLevelPermissions.SessionOrderBy = sOrderBy
				UserLevelPermissions.UserLevelID.Sort = ""
				UserLevelPermissions.zTableName.Sort = ""
				UserLevelPermissions.Permission.Sort = ""
			End If

			' Reset start position
			StartRec = 1
			UserLevelPermissions.StartRecordNumber = StartRec
		End If
	End Sub

	' Set up list options
	Sub SetupListOptions()
		Dim item
		ListOptions.Add("view")
		ListOptions.GetItem("view").CssStyle = "white-space: nowrap;"
		ListOptions.GetItem("view").Visible = Security.CanView
		ListOptions.GetItem("view").OnLeft = False
		ListOptions.Add("edit")
		ListOptions.GetItem("edit").CssStyle = "white-space: nowrap;"
		ListOptions.GetItem("edit").Visible = Security.CanEdit
		ListOptions.GetItem("edit").OnLeft = False
		ListOptions.Add("checkbox")
		ListOptions.GetItem("checkbox").CssStyle = "white-space: nowrap;"
		ListOptions.GetItem("checkbox").Visible = Security.CanDelete
		ListOptions.GetItem("checkbox").OnLeft = False
		ListOptions.GetItem("checkbox").Header = "<input type=""checkbox"" name=""key"" id=""key"" class=""aspmaker"" onclick=""UserLevelPermissions_list.SelectAllKey(this);"">"
		Call ListOptions_Load()
	End Sub

	' Render list options
	Sub RenderListOptions()
		Dim item, links
		ListOptions.LoadDefault()
		If Security.CanView And ListOptions.GetItem("view").Visible Then
			ListOptions.GetItem("view").Body = "<a class=""ewRowLink"" href=""" & ViewUrl & """>" & Language.Phrase("ViewLink") & "</a>"
		End If
		If Security.CanEdit And ListOptions.GetItem("edit").Visible Then
			Set item = ListOptions.GetItem("edit")
			item.Body = "<a class=""ewRowLink"" href=""" & EditUrl & """>" & Language.Phrase("EditLink") & "</a>"
		End If
		If Security.CanDelete And ListOptions.GetItem("checkbox").Visible Then
			ListOptions.GetItem("checkbox").Body = "<input type=""checkbox"" name=""key_m"" id=""key_m"" value=""" & ew_HtmlEncode(UserLevelPermissions.UserLevelID.CurrentValue&EW_COMPOSITE_KEY_SEPARATOR&UserLevelPermissions.zTableName.CurrentValue) & """ class=""aspmaker"" onclick='ew_ClickMultiCheckbox(this);'>"
		End If
		Call RenderListOptionsExt()
		Call ListOptions_Rendered()
	End Sub

	Function RenderListOptionsExt()
	End Function
	Dim Pager

	' -----------------------------------------------------------------
	' Set up Starting Record parameters based on Pager Navigation
	'
	Sub SetUpStartRec()
		Dim PageNo

		' Exit if DisplayRecs = 0
		If DisplayRecs = 0 Then Exit Sub
		If IsPageRequest Then ' Validate request

			' Check for a START parameter
			If Request.QueryString(EW_TABLE_START_REC).Count > 0 Then
				StartRec = Request.QueryString(EW_TABLE_START_REC)
				UserLevelPermissions.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					UserLevelPermissions.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = UserLevelPermissions.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			UserLevelPermissions.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			UserLevelPermissions.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			UserLevelPermissions.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	'  Load basic search values
	'
	Function LoadBasicSearchValues()
		UserLevelPermissions.BasicSearchKeyword = Request.QueryString(EW_TABLE_BASIC_SEARCH)
		UserLevelPermissions.BasicSearchType = Request.QueryString(EW_TABLE_BASIC_SEARCH_TYPE)
	End Function

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = UserLevelPermissions.CurrentFilter
		Call UserLevelPermissions.Recordset_Selecting(sFilter)
		UserLevelPermissions.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = UserLevelPermissions.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call UserLevelPermissions.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = UserLevelPermissions.KeyFilter

		' Call Row Selecting event
		Call UserLevelPermissions.Row_Selecting(sFilter)

		' Load sql based on filter
		UserLevelPermissions.CurrentFilter = sFilter
		sSql = UserLevelPermissions.SQL
		Call ew_SetDebugMsg("LoadRow: " & sSql) ' Show SQL for debugging
		Set RsRow = ew_LoadRow(sSql)
		If RsRow.Eof Then
			LoadRow = False
		Else
			LoadRow = True
			RsRow.MoveFirst
			Call LoadRowValues(RsRow) ' Load row values
		End If
		RsRow.Close
		Set RsRow = Nothing
	End Function

	' -----------------------------------------------------------------
	' Load row values from recordset
	'
	Sub LoadRowValues(RsRow)
		Dim sDetailFilter
		If RsRow.Eof Then Exit Sub

		' Call Row Selected event
		Call UserLevelPermissions.Row_Selected(RsRow)
		UserLevelPermissions.UserLevelID.DbValue = RsRow("UserLevelID")
		UserLevelPermissions.zTableName.DbValue = RsRow("TableName")
		UserLevelPermissions.Permission.DbValue = RsRow("Permission")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If UserLevelPermissions.GetKey("UserLevelID")&"" <> "" Then
			UserLevelPermissions.UserLevelID.CurrentValue = UserLevelPermissions.GetKey("UserLevelID") ' UserLevelID
		Else
			bValidKey = False
		End If
		If UserLevelPermissions.GetKey("zTableName")&"" <> "" Then
			UserLevelPermissions.zTableName.CurrentValue = UserLevelPermissions.GetKey("zTableName") ' TableName
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			UserLevelPermissions.CurrentFilter = UserLevelPermissions.KeyFilter
			Dim sSql
			sSql = UserLevelPermissions.SQL
			Set OldRecordset = ew_LoadRecordset(sSql)
			Call LoadRowValues(OldRecordset) ' Load row values
		Else
			OldRecordset = Null
		End If
		LoadOldRecord = bValidKey
	End Function

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		ViewUrl = UserLevelPermissions.ViewUrl
		EditUrl = UserLevelPermissions.EditUrl("")
		InlineEditUrl = UserLevelPermissions.InlineEditUrl
		CopyUrl = UserLevelPermissions.CopyUrl("")
		InlineCopyUrl = UserLevelPermissions.InlineCopyUrl
		DeleteUrl = UserLevelPermissions.DeleteUrl

		' Call Row Rendering event
		Call UserLevelPermissions.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' UserLevelID
		' TableName
		' Permission
		' -----------
		'  View  Row
		' -----------

		If UserLevelPermissions.RowType = EW_ROWTYPE_VIEW Then ' View row

			' UserLevelID
			UserLevelPermissions.UserLevelID.ViewValue = UserLevelPermissions.UserLevelID.CurrentValue
			UserLevelPermissions.UserLevelID.ViewCustomAttributes = ""

			' TableName
			UserLevelPermissions.zTableName.ViewValue = UserLevelPermissions.zTableName.CurrentValue
			UserLevelPermissions.zTableName.ViewCustomAttributes = ""

			' Permission
			UserLevelPermissions.Permission.ViewValue = UserLevelPermissions.Permission.CurrentValue
			UserLevelPermissions.Permission.ViewCustomAttributes = ""

			' View refer script
			' UserLevelID

			UserLevelPermissions.UserLevelID.LinkCustomAttributes = ""
			UserLevelPermissions.UserLevelID.HrefValue = ""
			UserLevelPermissions.UserLevelID.TooltipValue = ""

			' TableName
			UserLevelPermissions.zTableName.LinkCustomAttributes = ""
			UserLevelPermissions.zTableName.HrefValue = ""
			UserLevelPermissions.zTableName.TooltipValue = ""

			' Permission
			UserLevelPermissions.Permission.LinkCustomAttributes = ""
			UserLevelPermissions.Permission.HrefValue = ""
			UserLevelPermissions.Permission.TooltipValue = ""
		End If

		' Call Row Rendered event
		If UserLevelPermissions.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call UserLevelPermissions.Row_Rendered()
		End If
	End Sub

	' Page Load event
	Sub Page_Load()

		'Response.Write "Page Load"
	End Sub

	' Page Unload event
	Sub Page_Unload()

		'Response.Write "Page Unload"
	End Sub

	' Page Redirecting event
	Sub Page_Redirecting(url)

		'url = newurl
	End Sub

	' Message Showing event
	' typ = ""|"success"|"failure"
	Sub Message_Showing(msg, typ)

		' Example:
		'If typ = "success" Then msg = "your success message"

	End Sub

	' Page Data Rendering event
	Sub Page_DataRendering(header)

		' Example:
		'header = "your header"

	End Sub

	' Page Data Rendered event
	Sub Page_DataRendered(footer)

		' Example:
		'footer = "your footer"

	End Sub

	' Form Custom Validate event
	Function Form_CustomValidate(CustomError)

		'Return error message in CustomError
		Form_CustomValidate = True
	End Function

	' ListOptions Load event
	Sub ListOptions_Load()

		'Example: 
		' Dim opt
		' Set opt = ListOptions.Add("new")
		' opt.OnLeft = True ' Link on left
		' opt.MoveTo 0 ' Move to first column

	End Sub

	' ListOptions Rendered event
	Sub ListOptions_Rendered()

		'Example: 
		'ListOptions.GetItem("new").Body = "xxx"

	End Sub
End Class
%>
