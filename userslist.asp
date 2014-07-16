<%@ CodePage="1252" LCID="2057" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="ewcfg9.asp"-->
<!--#include file="usersinfo.asp"-->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim Users_list
Set Users_list = New cUsers_list
Set Page = Users_list

' Page init processing
Call Users_list.Page_Init()

' Page main processing
Call Users_list.Page_Main()
%>
<!--#include file="header.asp"-->
<% If Users.Export = "" Then %>
<script type="text/javascript">
<!--
// Create page object
var Users_list = new ew_Page("Users_list");
// page properties
Users_list.PageID = "list"; // page ID
Users_list.FormID = "fUserslist"; // form ID
var EW_PAGE_ID = Users_list.PageID; // for backward compatibility
// extend page with Form_CustomValidate function
Users_list.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
Users_list.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
Users_list.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Users_list.ValidateRequired = false; // no JavaScript validation
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
<% If (Users.Export = "") Or (EW_EXPORT_MASTER_RECORD And Users.Export = "print") Then %>
<% End If %>
<% Users_list.ShowPageHeader() %>
<%

' Load recordset
Set Users_list.Recordset = Users_list.LoadRecordset()
	Users_list.TotalRecs = Users_list.Recordset.RecordCount
	Users_list.StartRec = 1
	If Users_list.DisplayRecs <= 0 Then ' Display all records
		Users_list.DisplayRecs = Users_list.TotalRecs
	End If
	If Not (Users.ExportAll And Users.Export <> "") Then
		Users_list.SetUpStartRec() ' Set up start record position
	End If
%>

<% If Security.CanSearch Then %>
<% If Users.Export = "" And Users.CurrentAction = "" Then %>
<a href="javascript:ew_ToggleSearchPanel(Users_list);" style="text-decoration: none;"><img id="Users_list_SearchImage" src="images/collapse.gif" alt="" width="9" height="9" border="0"></a><span class="aspmaker">&nbsp;<%= Language.Phrase("Search") %></span><br>
<div id="Users_list_SearchPanel">
<form name="fUserslistsrch" id="fUserslistsrch" class="form-inline" action="<%= ew_CurrentPage %>">
<input type="hidden" id="t" name="t" value="Users">
<div>
<div id="xsr_1" class="input-append">
	<input type="text" name="<%= EW_TABLE_BASIC_SEARCH %>" id="<%= EW_TABLE_BASIC_SEARCH %>" size="20" value="<%= ew_HtmlEncode(Users.SessionBasicSearchKeyword) %>">
	<button class="btn" type="Submit" name="Submit" id="Submit"><%= ew_BtnCaption(Language.Phrase("QuickSearchBtn")) %></button>
	<a class="btn" href="<%= Users_list.PageUrl %>cmd=reset"><i class="icon-refresh"></i>&nbsp;<%= Language.Phrase("ShowAll") %></a>
</div>
</div>
</form>
</div>
<% End If %>
<% End If %>
<% Users_list.ShowMessage %>
<br>
<table cellspacing="0"><tr><td >
<form name="fUserslist" id="fUserslist" class="form-inline" action="" method="post">
<input type="hidden" name="t" id="t" value="Users">
<div id="gmp_Users" class="ewGridMiddlePanel">
<% If Users_list.TotalRecs > 0 Then %>
<table cellspacing="0" data-rowhighlightclass="ewTableHighlightRow" data-rowselectclass="ewTableSelectRow" data-roweditclass="ewTableEditRow" class="table table-bordered table-striped table-condensed">
<%= Users.TableCustomInnerHTML %>
<thead><!-- Table header -->
	<tr class="ewTableHeader">
<%
Call Users_list.RenderListOptions()

' Render list options (header, left)
Users_list.ListOptions.Render "header", "left"
%>
<% If Users.Username.Visible Then ' Username %>
		<td><%= Users.Username.FldCaption %></td>
<% End If %>		
<% If Users.zEmail.Visible Then ' Email %>
		<td><%= Users.zEmail.FldCaption %></td>
<% End If %>		
<% If Users.Permissions.Visible Then ' Permissions %>
		<td><%= Users.Permissions.FldCaption %></td>
<% End If %>		
<%

' Render list options (header, right)
Users_list.ListOptions.Render "header", "right"
%>
	</tr>
</thead>
<tbody><!-- Table body -->
<%
If (Users.ExportAll And Users.Export <> "") Then
	Users_list.StopRec = Users_list.TotalRecs
Else

	' Set the last record to display
	If Users_list.TotalRecs > Users_list.StartRec + Users_list.DisplayRecs - 1 Then
		Users_list.StopRec = Users_list.StartRec + Users_list.DisplayRecs - 1
	Else
		Users_list.StopRec = Users_list.TotalRecs
	End If
End If

' Move to first record
Users_list.RecCnt = Users_list.StartRec - 1
If Not Users_list.Recordset.Eof Then
	Users_list.Recordset.MoveFirst
	If Users_list.StartRec > 1 Then Users_list.Recordset.Move Users_list.StartRec - 1
ElseIf Not Users.AllowAddDeleteRow And Users_list.StopRec = 0 Then
	Users_list.StopRec = Users.GridAddRowCount
End If

' Initialize Aggregate
Users.RowType = EW_ROWTYPE_AGGREGATEINIT
Call Users.ResetAttrs()
Call Users_list.RenderRow()
Users_list.RowCnt = 0

' Output date rows
Do While CLng(Users_list.RecCnt) < CLng(Users_list.StopRec)
	Users_list.RecCnt = Users_list.RecCnt + 1
	If CLng(Users_list.RecCnt) >= CLng(Users_list.StartRec) Then
		Users_list.RowCnt = Users_list.RowCnt + 1

	' Set up key count
	Users_list.KeyCount = Users_list.RowIndex
	Call Users.ResetAttrs()
	Users.CssClass = ""
	If Users.CurrentAction = "gridadd" Then
	Else
		Call Users_list.LoadRowValues(Users_list.Recordset) ' Load row values
	End If
	Users.RowType = EW_ROWTYPE_VIEW ' Render view
	Users.RowAttrs.AddAttributes Array(Array("onmouseover", "ew_MouseOver(event, this);"), Array("onmouseout", "ew_MouseOut(event, this);"), Array("onclick", "ew_Click(event, this);"))

	' Render row
	Call Users_list.RenderRow()

	' Render list options
	Call Users_list.RenderListOptions()
%>
	<tr<%= Users.RowAttributes %>>
<%

' Render list options (body, left)
Users_list.ListOptions.Render "body", "left"
%>
	<% If Users.Username.Visible Then ' Username %>
		<td<%= Users.Username.CellAttributes %>>
<div<%= Users.Username.ViewAttributes %>><%= Users.Username.ListViewValue %></div>
<a name="<%= Users_list.PageObjName & "_row_" & Users_list.RowCnt %>" id="<%= Users_list.PageObjName & "_row_" & Users_list.RowCnt %>"></a></td>
	<% End If %>
	<% If Users.zEmail.Visible Then ' Email %>
		<td<%= Users.zEmail.CellAttributes %>>
<div<%= Users.zEmail.ViewAttributes %>>
<% If Users.zEmail.LinkAttributes <> "" Then %>
<a<%= Users.zEmail.LinkAttributes %>><%= Users.zEmail.ListViewValue %></a>
<% Else %>
<%= Users.zEmail.ListViewValue %>
<% End If %>
</div>
</td>
	<% End If %>
	<% If Users.Permissions.Visible Then ' Permissions %>
		<td<%= Users.Permissions.CellAttributes %>>
<div<%= Users.Permissions.ViewAttributes %>><%= Users.Permissions.ListViewValue %></div>
</td>
	<% End If %>
<%

' Render list options (body, right)
Users_list.ListOptions.Render "body", "right"
%>
	</tr>
<%
	End If
	If Users.CurrentAction <> "gridadd" Then
		Users_list.Recordset.MoveNext()
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
Users_list.Recordset.Close
Set Users_list.Recordset = Nothing
%>
<% If Users.Export = "" Then %>
<div class="ewGridLowerPanel">
<% If Users.CurrentAction <> "gridadd" And Users.CurrentAction <> "gridedit" Then %>
<form name="ewpagerform" id="ewpagerform" class="form-inline" action="<%= ew_CurrentPage %>">
<table border="0" cellspacing="0" cellpadding="0" >
	<tr>
		<td>
<% If Not IsObject(Users_list.Pager) Then Set Users_list.Pager = ew_NewPrevNextPager(Users_list.StartRec, Users_list.DisplayRecs, Users_list.TotalRecs) %>
<% If Users_list.Pager.RecordCount > 0 Then %>
	<table border="0" cellspacing="0" cellpadding="0"><tr><td><span class="aspmaker"><%= Language.Phrase("Page") %>&nbsp;</span></td>
<!--first page button-->
	<% If Users_list.Pager.FirstButton.Enabled Then %>
	<td><a href="<%= Users_list.PageUrl %>start=<%= Users_list.Pager.FirstButton.Start %>"><img src="images/first.gif" alt="<%= Language.Phrase("PagerFirst") %>" width="16" height="16" border="0"></a></td>
	<% Else %>
	<td><img src="images/firstdisab.gif" alt="<%= Language.Phrase("PagerFirst") %>" width="16" height="16" border="0"></td>
	<% End If %>
<!--previous page button-->
	<% If Users_list.Pager.PrevButton.Enabled Then %>
	<td><a href="<%= Users_list.PageUrl %>start=<%= Users_list.Pager.PrevButton.Start %>"><img src="images/prev.gif" alt="<%= Language.Phrase("PagerPrevious") %>" width="16" height="16" border="0"></a></td>
	<% Else %>
	<td><img src="images/prevdisab.gif" alt="<%= Language.Phrase("PagerPrevious") %>" width="16" height="16" border="0"></td>
	<% End If %>
<!--current page number-->
	<td><input type="text" name="<%= EW_TABLE_PAGE_NO %>" id="<%= EW_TABLE_PAGE_NO %>" value="<%= Users_list.Pager.CurrentPage %>" size="4"></td>
<!--next page button-->
	<% If Users_list.Pager.NextButton.Enabled Then %>
	<td><a href="<%= Users_list.PageUrl %>start=<%= Users_list.Pager.NextButton.Start %>"><img src="images/next.gif" alt="<%= Language.Phrase("PagerNext") %>" width="16" height="16" border="0"></a></td>
	<% Else %>
	<td><img src="images/nextdisab.gif" alt="<%= Language.Phrase("PagerNext") %>" width="16" height="16" border="0"></td>
	<% End If %>
<!--last page button-->
	<% If Users_list.Pager.LastButton.Enabled Then %>
	<td><a href="<%= Users_list.PageUrl %>start=<%= Users_list.Pager.LastButton.Start %>"><img src="images/last.gif" alt="<%= Language.Phrase("PagerLast") %>" width="16" height="16" border="0"></a></td>	
	<% Else %>
	<td><img src="images/lastdisab.gif" alt="<%= Language.Phrase("PagerLast") %>" width="16" height="16" border="0"></td>
	<% End If %>
	<td><span class="aspmaker">&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= Users_list.Pager.PageCount %></span></td>
	</tr></table>
	</td>	
	<td>&nbsp;&nbsp;&nbsp;&nbsp;</td>
	<td>
	<span class="aspmaker"><%= Language.Phrase("Record") %>&nbsp;<%= Users_list.Pager.FromIndex %>&nbsp;<%= Language.Phrase("To") %>&nbsp;<%= Users_list.Pager.ToIndex %>&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= Users_list.Pager.RecordCount %></span>
<% Else %>
	<% If Security.CanList Then %>
	<% If Users_list.SearchWhere = "0=101" Then %>
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
<a class="btn btn-success" href="<%= Users_list.AddUrl %>"><i class="icon-plus icon-white"></i>&nbsp;<%= Language.Phrase("AddLink") %></a>&nbsp;&nbsp;
<% End If %>
<% If Users_list.TotalRecs > 0 Then %>
<% If Security.CanDelete Then %>
<a class="btn btn-danger" href="" onclick="ew_SubmitSelected(document.fUserslist, '<%= Users_list.MultiDeleteUrl %>');return false;"><i class="icon-remove icon-white"></i>&nbsp;<%= Language.Phrase("DeleteSelectedLink") %></a>&nbsp;&nbsp;
<% End If %>
<% End If %>
</span>
</div>
<% End If %>
</td></tr></table>
<% If Users.Export = "" And Users.CurrentAction = "" Then %>
<% End If %>
<%
Users_list.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If Users.Export = "" Then %>
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
Set Users_list = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cUsers_list

	' Page ID
	Public Property Get PageID()
		PageID = "list"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Users"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Users_list"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If Users.UseTokenInUrl Then PageUrl = PageUrl & "t=" & Users.TableVar & "&" ' add page token
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
		If Users.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (Users.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (Users.TableVar = Request.QueryString("t"))
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
		If IsEmpty(Users) Then Set Users = New cUsers
		Set Table = Users

		' Initialize urls
		ExportPrintUrl = PageUrl & "export=print"
		ExportExcelUrl = PageUrl & "export=excel"
		ExportWordUrl = PageUrl & "export=word"
		ExportHtmlUrl = PageUrl & "export=html"
		ExportXmlUrl = PageUrl & "export=xml"
		ExportCsvUrl = PageUrl & "export=csv"
		AddUrl = "usersadd.asp"
		InlineAddUrl = PageUrl & "a=add"
		GridAddUrl = PageUrl & "a=gridadd"
		GridEditUrl = PageUrl & "a=gridedit"
		MultiDeleteUrl = "usersdelete.asp"
		MultiUpdateUrl = "usersupdate.asp"

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "list"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "Users"

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
		If Not Security.IsLoggedIn() Then
			Call Security.SaveLastUrl()
			Call Page_Terminate("login.asp")
		End If
		If Not Security.CanList Then
			Call Security.SaveLastUrl()
			Call Page_Terminate("login.asp")
		End If

		' Get grid add count
		Dim gridaddcnt
		gridaddcnt = Request.QueryString(EW_TABLE_GRID_ADD_ROW_COUNT)
		If IsNumeric(gridaddcnt) Then
			If gridaddcnt > 0 Then
				Users.GridAddRowCount = gridaddcnt
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
		Set Users = Nothing
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
			If Users.Export <> "" Or Users.CurrentAction = "gridadd" Or Users.CurrentAction = "gridedit" Then
				ListOptions.HideAllOptions()
				ExportOptions.HideAllOptions()
			End If

			' Get basic search values
			Call LoadBasicSearchValues()

			' Restore search parms from Session
			Call RestoreSearchParms()

			' Call Recordset SearchValidated event
			Call Users.Recordset_SearchValidated()

			' Set Up Sorting Order
			SetUpSortOrder()

			' Get basic search criteria
			If gsSearchError = "" Then
				sSrchBasic = BasicSearchWhere()
			End If
		End If ' End Validate Request

		' Restore display records
		If Users.RecordsPerPage <> "" Then
			DisplayRecs = Users.RecordsPerPage ' Restore from Session
		Else
			DisplayRecs = 50 ' Load default
		End If

		' Load Sorting Order
		LoadSortOrder()

		' Build search criteria
		Call ew_AddFilter(SearchWhere, sSrchAdvanced)
		Call ew_AddFilter(SearchWhere, sSrchBasic)

		' Call Recordset Searching event
		Call Users.Recordset_Searching(SearchWhere)

		' Save search criteria
		If SearchWhere <> "" Then
			If sSrchBasic = "" Then Call ResetBasicSearchParms()
			Users.SearchWhere = SearchWhere ' Save to Session
			If Not RestoreSearch Then
				StartRec = 1 ' Reset start record counter
				Users.StartRecordNumber = StartRec
			End If
		Else
			SearchWhere = Users.SearchWhere
		End If
		sFilter = ""
		If Not Security.CanList Then
			sFilter = "(0=1)" ' Filter all records
		End If
		Call ew_AddFilter(sFilter, DbDetailFilter)
		Call ew_AddFilter(sFilter, SearchWhere)

		' Set up filter in Session
		Users.SessionWhere = sFilter
		Users.CurrentFilter = ""
	End Sub

	' -----------------------------------------------------------------
	' Return Basic Search sql
	'
	Function BasicSearchSQL(Keyword)
		Dim sWhere
		sWhere = ""
			Call BuildBasicSearchSQL(sWhere, Users.Username, Keyword)
			Call BuildBasicSearchSQL(sWhere, Users.Password, Keyword)
			Call BuildBasicSearchSQL(sWhere, Users.zEmail, Keyword)
			Call BuildBasicSearchSQL(sWhere, Users.Profile, Keyword)
			Call BuildBasicSearchSQL(sWhere, Users.Theme, Keyword)
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
		sSearchKeyword = Users.BasicSearchKeyword
		sSearchType = Users.BasicSearchType
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
			Users.SessionBasicSearchKeyword = sSearchKeyword
			Users.SessionBasicSearchType = sSearchType
		End If
		BasicSearchWhere = sSearchStr
	End Function

	' -----------------------------------------------------------------
	' Clear all search parameters
	'
	Sub ResetSearchParms()

		' Clear search where
		SearchWhere = ""
		Users.SearchWhere = SearchWhere

		' Clear basic search parameters
		Call ResetBasicSearchParms()
	End Sub

	' -----------------------------------------------------------------
	' Clear all basic search parameters
	'
	Sub ResetBasicSearchParms()

		' Clear basic search parameters
		Users.SessionBasicSearchKeyword = ""
		Users.SessionBasicSearchType = ""
	End Sub

	' -----------------------------------------------------------------
	' Restore all search parameters
	'
	Sub RestoreSearchParms()
		Dim bRestore
		bRestore = True
		If Users.BasicSearchKeyword & "" <> "" Then bRestore = False
		RestoreSearch = bRestore
		If bRestore Then

			' Restore basic search values
			Users.BasicSearchKeyword = Users.SessionBasicSearchKeyword
			Users.BasicSearchType = Users.SessionBasicSearchType
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
			Users.CurrentOrder = Request.QueryString("order")
			Users.CurrentOrderType = Request.QueryString("ordertype")

			' Field Username
			Call Users.UpdateSort(Users.Username)

			' Field Email
			Call Users.UpdateSort(Users.zEmail)

			' Field Permissions
			Call Users.UpdateSort(Users.Permissions)
			Users.StartRecordNumber = 1 ' Reset start position
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load Sort Order parameters
	'
	Sub LoadSortOrder()
		Dim sOrderBy
		sOrderBy = Users.SessionOrderBy ' Get order by from Session
		If sOrderBy = "" Then
			If Users.SqlOrderBy <> "" Then
				sOrderBy = Users.SqlOrderBy
				Users.SessionOrderBy = sOrderBy
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
				Users.SessionOrderBy = sOrderBy
				Users.Username.Sort = ""
				Users.zEmail.Sort = ""
				Users.Permissions.Sort = ""
			End If

			' Reset start position
			StartRec = 1
			Users.StartRecordNumber = StartRec
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
		ListOptions.GetItem("checkbox").Header = "<input type=""checkbox"" name=""key"" id=""key"" class=""aspmaker"" onclick=""Users_list.SelectAllKey(this);"">"
		Call ListOptions_Load()
	End Sub

	' Render list options
	Sub RenderListOptions()
		Dim item, links
		ListOptions.LoadDefault()
		If Security.CanEdit And ListOptions.GetItem("edit").Visible Then
			Set item = ListOptions.GetItem("edit")
			item.Body = "<a class=""ewRowLink"" rel=""tooltip"" title=""Edit""" & ew_HtmlEncode(Users.Username.CurrentValue) & """ href=""" & EditUrl & """><i class=""icon-pencil""></i></a>"
		End If
		If Security.CanDelete And ListOptions.GetItem("checkbox").Visible Then
			ListOptions.GetItem("checkbox").Body = "<input type=""checkbox"" name=""key_m"" id=""key_m"" value=""" & ew_HtmlEncode(Users.Username.CurrentValue) & """ class=""aspmaker"" onclick='ew_ClickMultiCheckbox(this);'>"
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
				Users.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					Users.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = Users.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			Users.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			Users.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			Users.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	'  Load basic search values
	'
	Function LoadBasicSearchValues()
		Users.BasicSearchKeyword = Request.QueryString(EW_TABLE_BASIC_SEARCH)
		Users.BasicSearchType = Request.QueryString(EW_TABLE_BASIC_SEARCH_TYPE)
	End Function

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = Users.CurrentFilter
		Call Users.Recordset_Selecting(sFilter)
		Users.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = Users.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call Users.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = Users.KeyFilter

		' Call Row Selecting event
		Call Users.Row_Selecting(sFilter)

		' Load sql based on filter
		Users.CurrentFilter = sFilter
		sSql = Users.SQL
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
		Call Users.Row_Selected(RsRow)
		Users.Username.DbValue = RsRow("Username")
		Users.Password.DbValue = RsRow("Password")
		Users.zEmail.DbValue = RsRow("Email")
		Users.Permissions.DbValue = RsRow("Permissions")
		Users.Active.DbValue = ew_IIf(RsRow("Active"), "1", "0")
		Users.Profile.DbValue = RsRow("Profile")
		Users.Theme.DbValue = RsRow("Theme")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If Users.GetKey("Username")&"" <> "" Then
			Users.Username.CurrentValue = Users.GetKey("Username") ' Username
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			Users.CurrentFilter = Users.KeyFilter
			Dim sSql
			sSql = Users.SQL
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
		ViewUrl = Users.ViewUrl
		EditUrl = Users.EditUrl("")
		InlineEditUrl = Users.InlineEditUrl
		CopyUrl = Users.CopyUrl("")
		InlineCopyUrl = Users.InlineCopyUrl
		DeleteUrl = Users.DeleteUrl

		' Call Row Rendering event
		Call Users.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' Username
		' Password
		' Email
		' Permissions
		' Active
		' Profile
		' Theme
		' -----------
		'  View  Row
		' -----------

		If Users.RowType = EW_ROWTYPE_VIEW Then ' View row

			' Username
			Users.Username.ViewValue = Users.Username.CurrentValue
			Users.Username.ViewCustomAttributes = ""

			' Email
			Users.zEmail.ViewValue = Users.zEmail.CurrentValue
			Users.zEmail.ViewCustomAttributes = ""

			' Permissions
			If (Security.CurrentUserLevel And EW_ALLOW_ADMIN) = EW_ALLOW_ADMIN Then ' System admin
			If Users.Permissions.CurrentValue & "" <> "" Then
				sFilterWrk = "[UserLevelID] = " & ew_AdjustSql(Users.Permissions.CurrentValue) & ""
			sSqlWrk = "SELECT [UserLevelName] FROM [UserLevels]"
			sWhereWrk = ""
			Call ew_AddFilter(sWhereWrk, sFilterWrk)
			If sWhereWrk <> "" Then sSqlWrk = sSqlWrk & " WHERE " & sWhereWrk
				Set RsWrk = Conn.Execute(sSqlWrk)
				If Not RsWrk.Eof Then
					Users.Permissions.ViewValue = RsWrk("UserLevelName")
				Else
					Users.Permissions.ViewValue = Users.Permissions.CurrentValue
				End If
				RsWrk.Close
				Set RsWrk = Nothing
			Else
				Users.Permissions.ViewValue = Null
			End If
			Else
				Users.Permissions.ViewValue = "********"
			End If
			Users.Permissions.ViewCustomAttributes = ""

			' View refer script
			' Username

			Users.Username.LinkCustomAttributes = ""
			Users.Username.HrefValue = ""
			Users.Username.TooltipValue = ""

			' Email
			Users.zEmail.LinkCustomAttributes = ""
			If Not ew_Empty(Users.zEmail.CurrentValue) Then
				Users.zEmail.HrefValue = "mailto:" & ew_IIf(Users.zEmail.ViewValue<>"", Users.zEmail.ViewValue, Users.zEmail.CurrentValue)
				Users.zEmail.LinkAttrs.AddAttribute "target", "", True ' Add target
				If Users.Export <> "" Then Users.zEmail.HrefValue = ew_ConvertFullUrl(Users.zEmail.HrefValue)
			Else
				Users.zEmail.HrefValue = ""
			End If
			Users.zEmail.TooltipValue = ""

			' Permissions
			Users.Permissions.LinkCustomAttributes = ""
			Users.Permissions.HrefValue = ""
			Users.Permissions.TooltipValue = ""
		End If

		' Call Row Rendered event
		If Users.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Users.Row_Rendered()
		End If
	End Sub

	' Write Audit Trail start/end for grid update
	Sub WriteAuditTrailDummy(typ)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		Dim table
		table = "Users"

		' Write Audit Trail
		Dim filePfx, curDateTime, id, user, action
		Dim i
		filePfx = "log"
		curDateTime = ew_StdCurrentDateTime()
		id = Request.ServerVariables("SCRIPT_NAME")
    	user = CurrentUserName
		action = typ
		Call ew_WriteAuditTrail(filePfx, curDateTime, id, user, action, table, "", "", "", "")
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
