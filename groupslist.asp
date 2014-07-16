<%@ CodePage="1252" LCID="2057" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="ewcfg9.asp"-->
<!--#include file="groupsinfo.asp"-->
<!--#include file="usersinfo.asp"-->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim Groups_list
Set Groups_list = New cGroups_list
Set Page = Groups_list

' Page init processing
Call Groups_list.Page_Init()

' Page main processing
Call Groups_list.Page_Main()
%>
<!--#include file="header.asp"-->
<% If Groups.Export = "" Then %>
<script type="text/javascript">
<!--
// Create page object
var Groups_list = new ew_Page("Groups_list");
// page properties
Groups_list.PageID = "list"; // page ID
Groups_list.FormID = "fGroupslist"; // form ID
var EW_PAGE_ID = Groups_list.PageID; // for backward compatibility
// extend page with Form_CustomValidate function
Groups_list.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
<% If EW_CLIENT_VALIDATE Then %>
Groups_list.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Groups_list.ValidateRequired = false; // no JavaScript validation
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
<% If (Groups.Export = "") Or (EW_EXPORT_MASTER_RECORD And Groups.Export = "print") Then %>
<% End If %>
<% Groups_list.ShowPageHeader() %>
<%

' Load recordset
Set Groups_list.Recordset = Groups_list.LoadRecordset()
	Groups_list.TotalRecs = Groups_list.Recordset.RecordCount
	Groups_list.StartRec = 1
	If Groups_list.DisplayRecs <= 0 Then ' Display all records
		Groups_list.DisplayRecs = Groups_list.TotalRecs
	End If
	If Not (Groups.ExportAll And Groups.Export <> "") Then
		Groups_list.SetUpStartRec() ' Set up start record position
	End If
%>
<p<% Groups_list.ExportOptions.Render "body", "" %>
</p>
<% If Security.CanSearch Then %>
<% If Groups.Export = "" And Groups.CurrentAction = "" Then %>
<a href="javascript:ew_ToggleSearchPanel(Groups_list);" style="text-decoration: none;"><img id="Groups_list_SearchImage" src="images/collapse.gif" alt="" width="9" height="9" border="0"></a><span class="aspmaker">&nbsp;<%= Language.Phrase("Search") %></span><br>
<div id="Groups_list_SearchPanel">
<form name="fGroupslistsrch" id="fGroupslistsrch" class="form-inline" action="<%= ew_CurrentPage %>">
<input type="hidden" id="t" name="t" value="Groups">
<div>
<div id="xsr_1" class="input-append">
	<input type="text" name="<%= EW_TABLE_BASIC_SEARCH %>" id="<%= EW_TABLE_BASIC_SEARCH %>" size="20" value="<%= ew_HtmlEncode(Groups.SessionBasicSearchKeyword) %>">
	<button class="btn" type="Submit" name="Submit" id="Submit"><%= ew_BtnCaption(Language.Phrase("QuickSearchBtn")) %></button>
	<a class="btn" href="<%= Groups_list.PageUrl %>cmd=reset"><i class="icon-refresh"></i>&nbsp;<%= Language.Phrase("ShowAll") %></a>
</div>
</div>
</form>
</div>
<% End If %>
<% End If %>
<% Groups_list.ShowMessage %>
<br>
<table>
<form name="fGroupslist" id="fGroupslist" class="ewForm" action="" method="post">
<input type="hidden" name="t" id="t" value="Groups">
<div id="gmp_Groups" class="ewGridMiddlePanel">
<% If Groups_list.TotalRecs > 0 Then %>
<table cellspacing="0" data-rowhighlightclass="ewTableHighlightRow" data-rowselectclass="ewTableSelectRow" data-roweditclass="ewTableEditRow" class="table table-bordered table-striped">
<%= Groups.TableCustomInnerHTML %>
<thead><!-- Table header -->
	<tr class="ewTableHeader">
<%
Call Groups_list.RenderListOptions()

' Render list options (header, left)
Groups_list.ListOptions.Render "header", "left"
%>
<% If Groups.Group.Visible Then ' Group %>
		<td><strong><%= Groups.Group.FldCaption %></strong></td>
<% End If %>		
<%

' Render list options (header, right)
Groups_list.ListOptions.Render "header", "right"
%>
	</tr>
</thead>
<tbody><!-- Table body -->
<%
If (Groups.ExportAll And Groups.Export <> "") Then
	Groups_list.StopRec = Groups_list.TotalRecs
Else

	' Set the last record to display
	If Groups_list.TotalRecs > Groups_list.StartRec + Groups_list.DisplayRecs - 1 Then
		Groups_list.StopRec = Groups_list.StartRec + Groups_list.DisplayRecs - 1
	Else
		Groups_list.StopRec = Groups_list.TotalRecs
	End If
End If

' Move to first record
Groups_list.RecCnt = Groups_list.StartRec - 1
If Not Groups_list.Recordset.Eof Then
	Groups_list.Recordset.MoveFirst
	If Groups_list.StartRec > 1 Then Groups_list.Recordset.Move Groups_list.StartRec - 1
ElseIf Not Groups.AllowAddDeleteRow And Groups_list.StopRec = 0 Then
	Groups_list.StopRec = Groups.GridAddRowCount
End If

' Initialize Aggregate
Groups.RowType = EW_ROWTYPE_AGGREGATEINIT
Call Groups.ResetAttrs()
Call Groups_list.RenderRow()
Groups_list.RowCnt = 0

' Output date rows
Do While CLng(Groups_list.RecCnt) < CLng(Groups_list.StopRec)
	Groups_list.RecCnt = Groups_list.RecCnt + 1
	If CLng(Groups_list.RecCnt) >= CLng(Groups_list.StartRec) Then
		Groups_list.RowCnt = Groups_list.RowCnt + 1

	' Set up key count
	Groups_list.KeyCount = Groups_list.RowIndex
	Call Groups.ResetAttrs()
	Groups.CssClass = ""
	If Groups.CurrentAction = "gridadd" Then
	Else
		Call Groups_list.LoadRowValues(Groups_list.Recordset) ' Load row values
	End If
	Groups.RowType = EW_ROWTYPE_VIEW ' Render view
	Groups.RowAttrs.AddAttributes Array(Array("onmouseover", "ew_MouseOver(event, this);"), Array("onmouseout", "ew_MouseOut(event, this);"), Array("onclick", "ew_Click(event, this);"))

	' Render row
	Call Groups_list.RenderRow()

	' Render list options
	Call Groups_list.RenderListOptions()
%>
	<tr<%= Groups.RowAttributes %>>
<%

' Render list options (body, left)
Groups_list.ListOptions.Render "body", "left"
%>
	<% If Groups.Group.Visible Then ' Group %>
		<td<%= Groups.Group.CellAttributes %>>
<div<%= Groups.Group.ViewAttributes %>><%= Groups.Group.ListViewValue %></div>
<a name="<%= Groups_list.PageObjName & "_row_" & Groups_list.RowCnt %>" id="<%= Groups_list.PageObjName & "_row_" & Groups_list.RowCnt %>"></a></td>
	<% End If %>
<%

' Render list options (body, right)
Groups_list.ListOptions.Render "body", "right"
%>
	</tr>
<%
	End If
	If Groups.CurrentAction <> "gridadd" Then
		Groups_list.Recordset.MoveNext()
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
Groups_list.Recordset.Close
Set Groups_list.Recordset = Nothing
%>
<% If Groups.Export = "" Then %>
<div class="ewGridLowerPanel">
<% If Groups.CurrentAction <> "gridadd" And Groups.CurrentAction <> "gridedit" Then %>
<form name="ewpagerform" id="ewpagerform" class="form-inline" action="<%= ew_CurrentPage %>">
<table border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td>
<% If Not IsObject(Groups_list.Pager) Then Set Groups_list.Pager = ew_NewPrevNextPager(Groups_list.StartRec, Groups_list.DisplayRecs, Groups_list.TotalRecs) %>
<% If Groups_list.Pager.RecordCount > 0 Then %>
	<table border="0" cellspacing="0" cellpadding="0"><tr><td><span><%= Language.Phrase("Page") %>&nbsp;</span></td>
<!--first page button-->
	<% If Groups_list.Pager.FirstButton.Enabled Then %>
	<td><a href="<%= Groups_list.PageUrl %>start=<%= Groups_list.Pager.FirstButton.Start %>"><img src="images/first.gif" alt="<%= Language.Phrase("PagerFirst") %>" width="16" height="16" border="0"></a></td>
	<% Else %>
	<td><img src="images/firstdisab.gif" alt="<%= Language.Phrase("PagerFirst") %>" width="16" height="16" border="0"></td>
	<% End If %>
<!--previous page button-->
	<% If Groups_list.Pager.PrevButton.Enabled Then %>
	<td><a href="<%= Groups_list.PageUrl %>start=<%= Groups_list.Pager.PrevButton.Start %>"><img src="images/prev.gif" alt="<%= Language.Phrase("PagerPrevious") %>" width="16" height="16" border="0"></a></td>
	<% Else %>
	<td><img src="images/prevdisab.gif" alt="<%= Language.Phrase("PagerPrevious") %>" width="16" height="16" border="0"></td>
	<% End If %>
<!--current page number-->
	<td><input type="text" name="<%= EW_TABLE_PAGE_NO %>" id="<%= EW_TABLE_PAGE_NO %>" value="<%= Groups_list.Pager.CurrentPage %>" size="4"></td>
<!--next page button-->
	<% If Groups_list.Pager.NextButton.Enabled Then %>
	<td><a href="<%= Groups_list.PageUrl %>start=<%= Groups_list.Pager.NextButton.Start %>"><img src="images/next.gif" alt="<%= Language.Phrase("PagerNext") %>" width="16" height="16" border="0"></a></td>
	<% Else %>
	<td><img src="images/nextdisab.gif" alt="<%= Language.Phrase("PagerNext") %>" width="16" height="16" border="0"></td>
	<% End If %>
<!--last page button-->
	<% If Groups_list.Pager.LastButton.Enabled Then %>
	<td><a href="<%= Groups_list.PageUrl %>start=<%= Groups_list.Pager.LastButton.Start %>"><img src="images/last.gif" alt="<%= Language.Phrase("PagerLast") %>" width="16" height="16" border="0"></a></td>	
	<% Else %>
	<td><img src="images/lastdisab.gif" alt="<%= Language.Phrase("PagerLast") %>" width="16" height="16" border="0"></td>
	<% End If %>
	<td><span>&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= Groups_list.Pager.PageCount %></span></td>
	</tr></table>
	</td>	
	<td>&nbsp;&nbsp;&nbsp;&nbsp;</td>
	<td>
	<span class="aspmaker"><%= Language.Phrase("Record") %>&nbsp;<%= Groups_list.Pager.FromIndex %>&nbsp;<%= Language.Phrase("To") %>&nbsp;<%= Groups_list.Pager.ToIndex %>&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= Groups_list.Pager.RecordCount %></span>
<% Else %>
	<% If Security.CanList Then %>
	<% If Groups_list.SearchWhere = "0=101" Then %>
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
<a class="btn btn-success" href="<%= Groups_list.AddUrl %>"><i class="icon-plus icon-white"></i>&nbsp;Add New Group</a>&nbsp;&nbsp;
<% End If %>
</span>
</div>
<% End If %>
</td></tr></table>
<% If Groups.Export = "" And Groups.CurrentAction = "" Then %>
<% End If %>
<%
Groups_list.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If Groups.Export = "" Then %>
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
Set Groups_list = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cGroups_list

	' Page ID
	Public Property Get PageID()
		PageID = "list"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Groups"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Groups_list"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If Groups.UseTokenInUrl Then PageUrl = PageUrl & "t=" & Groups.TableVar & "&" ' add page token
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
		If sMessage <> "" Then Response.Write "<p class=""alert alert-info"">" & sMessage & "</p>"
		Session(EW_SESSION_MESSAGE) = "" ' Clear message in Session

		' Success message
		Dim sSuccessMessage
		sSuccessMessage = SuccessMessage
		Call Message_Showing(sSuccessMessage, "success")
		If sSuccessMessage <> "" Then Response.Write "<p class=""alert alert-success"">" & sSuccessMessage & "</p>"
		Session(EW_SESSION_SUCCESS_MESSAGE) = "" ' Clear message in Session

		' Failure message
		Dim sErrorMessage
		sErrorMessage = FailureMessage
		Call Message_Showing(sErrorMessage, "failure")
		If sErrorMessage <> "" Then Response.Write "<p class=""alert alert-danger"">" & sErrorMessage & "</p>"
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
			Response.Write "<p>" & sHeader & "</p>"
		End If
	End Sub

	' Show Page Footer
	Public Sub ShowPageFooter()
		Dim sFooter
		sFooter = PageFooter
		Call Page_DataRendered(sFooter)
		If sFooter <> "" Then ' Footer exists, display
			Response.Write "<p>" & sFooter & "</p>"
		End If
	End Sub

	' -----------------------
	'  Validate Page request
	'
	Public Function IsPageRequest()
		If Groups.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (Groups.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (Groups.TableVar = Request.QueryString("t"))
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
		If IsEmpty(Groups) Then Set Groups = New cGroups
		Set Table = Groups

		' Initialize urls
		ExportPrintUrl = PageUrl & "export=print"
		ExportExcelUrl = PageUrl & "export=excel"
		ExportWordUrl = PageUrl & "export=word"
		ExportHtmlUrl = PageUrl & "export=html"
		ExportXmlUrl = PageUrl & "export=xml"
		ExportCsvUrl = PageUrl & "export=csv"
		AddUrl = "groupsadd.asp"
		InlineAddUrl = PageUrl & "a=add"
		GridAddUrl = PageUrl & "a=gridadd"
		GridEditUrl = PageUrl & "a=gridedit"
		MultiDeleteUrl = "groupsdelete.asp"
		MultiUpdateUrl = "groupsupdate.asp"

		' Initialize other table object
		If IsEmpty(Users) Then Set Users = New cUsers

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "list"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "Groups"

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
				Groups.GridAddRowCount = gridaddcnt
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
		Set Groups = Nothing
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
			If Groups.Export <> "" Or Groups.CurrentAction = "gridadd" Or Groups.CurrentAction = "gridedit" Then
				ListOptions.HideAllOptions()
				ExportOptions.HideAllOptions()
			End If

			' Get basic search values
			Call LoadBasicSearchValues()

			' Restore search parms from Session
			Call RestoreSearchParms()

			' Call Recordset SearchValidated event
			Call Groups.Recordset_SearchValidated()

			' Set Up Sorting Order
			SetUpSortOrder()

			' Get basic search criteria
			If gsSearchError = "" Then
				sSrchBasic = BasicSearchWhere()
			End If
		End If ' End Validate Request

		' Restore display records
		If Groups.RecordsPerPage <> "" Then
			DisplayRecs = Groups.RecordsPerPage ' Restore from Session
		Else
			DisplayRecs = 50 ' Load default
		End If

		' Load Sorting Order
		LoadSortOrder()

		' Build search criteria
		Call ew_AddFilter(SearchWhere, sSrchAdvanced)
		Call ew_AddFilter(SearchWhere, sSrchBasic)

		' Call Recordset Searching event
		Call Groups.Recordset_Searching(SearchWhere)

		' Save search criteria
		If SearchWhere <> "" Then
			If sSrchBasic = "" Then Call ResetBasicSearchParms()
			Groups.SearchWhere = SearchWhere ' Save to Session
			If Not RestoreSearch Then
				StartRec = 1 ' Reset start record counter
				Groups.StartRecordNumber = StartRec
			End If
		Else
			SearchWhere = Groups.SearchWhere
		End If
		sFilter = ""
		If Not Security.CanList Then
			sFilter = "(0=1)" ' Filter all records
		End If
		Call ew_AddFilter(sFilter, DbDetailFilter)
		Call ew_AddFilter(sFilter, SearchWhere)

		' Set up filter in Session
		Groups.SessionWhere = sFilter
		Groups.CurrentFilter = ""
	End Sub

	' -----------------------------------------------------------------
	' Return Basic Search sql
	'
	Function BasicSearchSQL(Keyword)
		Dim sWhere
		sWhere = ""
			Call BuildBasicSearchSQL(sWhere, Groups.Group, Keyword)
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
		sSearchKeyword = Groups.BasicSearchKeyword
		sSearchType = Groups.BasicSearchType
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
			Groups.SessionBasicSearchKeyword = sSearchKeyword
			Groups.SessionBasicSearchType = sSearchType
		End If
		BasicSearchWhere = sSearchStr
	End Function

	' -----------------------------------------------------------------
	' Clear all search parameters
	'
	Sub ResetSearchParms()

		' Clear search where
		SearchWhere = ""
		Groups.SearchWhere = SearchWhere

		' Clear basic search parameters
		Call ResetBasicSearchParms()
	End Sub

	' -----------------------------------------------------------------
	' Clear all basic search parameters
	'
	Sub ResetBasicSearchParms()

		' Clear basic search parameters
		Groups.SessionBasicSearchKeyword = ""
		Groups.SessionBasicSearchType = ""
	End Sub

	' -----------------------------------------------------------------
	' Restore all search parameters
	'
	Sub RestoreSearchParms()
		Dim bRestore
		bRestore = True
		If Groups.BasicSearchKeyword & "" <> "" Then bRestore = False
		RestoreSearch = bRestore
		If bRestore Then

			' Restore basic search values
			Groups.BasicSearchKeyword = Groups.SessionBasicSearchKeyword
			Groups.BasicSearchType = Groups.SessionBasicSearchType
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
			Groups.CurrentOrder = Request.QueryString("order")
			Groups.CurrentOrderType = Request.QueryString("ordertype")

			' Field Group
			Call Groups.UpdateSort(Groups.Group)
			Groups.StartRecordNumber = 1 ' Reset start position
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load Sort Order parameters
	'
	Sub LoadSortOrder()
		Dim sOrderBy
		sOrderBy = Groups.SessionOrderBy ' Get order by from Session
		If sOrderBy = "" Then
			If Groups.SqlOrderBy <> "" Then
				sOrderBy = Groups.SqlOrderBy
				Groups.SessionOrderBy = sOrderBy
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
				Groups.SessionOrderBy = sOrderBy
				Groups.Group.Sort = ""
			End If

			' Reset start position
			StartRec = 1
			Groups.StartRecordNumber = StartRec
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
		ListOptions.Add("delete")
		ListOptions.GetItem("delete").CssStyle = "white-space: nowrap;"
		ListOptions.GetItem("delete").Visible = Security.CanDelete
		ListOptions.GetItem("delete").OnLeft = False
		Call ListOptions_Load()
	End Sub

	' Render list options
	Sub RenderListOptions()
		Dim item, links
		ListOptions.LoadDefault()
		If Security.CanDelete And ListOptions.GetItem("delete").Visible Then
			ListOptions.GetItem("delete").Body = "<a rel='tooltip' title='View this group' " & "" & " href=""" & DeleteUrl & """><i class='icon-trash'></i></a>"
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
				Groups.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					Groups.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = Groups.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			Groups.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			Groups.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			Groups.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	'  Load basic search values
	'
	Function LoadBasicSearchValues()
		Groups.BasicSearchKeyword = Request.QueryString(EW_TABLE_BASIC_SEARCH)
		Groups.BasicSearchType = Request.QueryString(EW_TABLE_BASIC_SEARCH_TYPE)
	End Function

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = Groups.CurrentFilter
		Call Groups.Recordset_Selecting(sFilter)
		Groups.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = Groups.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call Groups.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = Groups.KeyFilter

		' Call Row Selecting event
		Call Groups.Row_Selecting(sFilter)

		' Load sql based on filter
		Groups.CurrentFilter = sFilter
		sSql = Groups.SQL
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
		Call Groups.Row_Selected(RsRow)
		Groups.Group.DbValue = RsRow("Group")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If Groups.GetKey("Group")&"" <> "" Then
			Groups.Group.CurrentValue = Groups.GetKey("Group") ' Group
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			Groups.CurrentFilter = Groups.KeyFilter
			Dim sSql
			sSql = Groups.SQL
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
		ViewUrl = Groups.ViewUrl
		EditUrl = Groups.EditUrl("")
		InlineEditUrl = Groups.InlineEditUrl
		CopyUrl = Groups.CopyUrl("")
		InlineCopyUrl = Groups.InlineCopyUrl
		DeleteUrl = Groups.DeleteUrl

		' Call Row Rendering event
		Call Groups.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' Group
		' -----------
		'  View  Row
		' -----------

		If Groups.RowType = EW_ROWTYPE_VIEW Then ' View row

			' Group
			Groups.Group.ViewValue = Groups.Group.CurrentValue
			Groups.Group.ViewCustomAttributes = ""

			' View refer script
			' Group

			Groups.Group.LinkCustomAttributes = ""
			Groups.Group.HrefValue = ""
			Groups.Group.TooltipValue = ""
		End If

		' Call Row Rendered event
		If Groups.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Groups.Row_Rendered()
		End If
	End Sub

	' Write Audit Trail start/end for grid update
	Sub WriteAuditTrailDummy(typ)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		Dim table
		table = "Groups"

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
