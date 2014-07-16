<%@ CodePage="1252" LCID="2057" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="ewcfg9.asp"-->
<!--#include file="userlevelsinfo.asp"-->
<!--#include file="usersinfo.asp"-->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim UserLevels_list
Set UserLevels_list = New cUserLevels_list
Set Page = UserLevels_list

' Page init processing
Call UserLevels_list.Page_Init()

' Page main processing
Call UserLevels_list.Page_Main()
%>
<!--#include file="header.asp"-->
<% If UserLevels.Export = "" Then %>
<script type="text/javascript">
<!--
// Create page object
var UserLevels_list = new ew_Page("UserLevels_list");
// page properties
UserLevels_list.PageID = "list"; // page ID
UserLevels_list.FormID = "fUserLevelslist"; // form ID
var EW_PAGE_ID = UserLevels_list.PageID; // for backward compatibility
// extend page with Form_CustomValidate function
UserLevels_list.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
UserLevels_list.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
UserLevels_list.ValidateRequired = true; // uses JavaScript validation
<% Else %>
UserLevels_list.ValidateRequired = false; // no JavaScript validation
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
<% If (UserLevels.Export = "") Or (EW_EXPORT_MASTER_RECORD And UserLevels.Export = "print") Then %>
<% End If %>
<% UserLevels_list.ShowPageHeader() %>
<%

' Load recordset
Set UserLevels_list.Recordset = UserLevels_list.LoadRecordset()
	UserLevels_list.TotalRecs = UserLevels_list.Recordset.RecordCount
	UserLevels_list.StartRec = 1
	If UserLevels_list.DisplayRecs <= 0 Then ' Display all records
		UserLevels_list.DisplayRecs = UserLevels_list.TotalRecs
	End If
	If Not (UserLevels.ExportAll And UserLevels.Export <> "") Then
		UserLevels_list.SetUpStartRec() ' Set up start record position
	End If
%>
<% UserLevels_list.ShowMessage %>
<br>
<table cellspacing="0"><tr><td>
<form name="fUserLevelslist" id="fUserLevelslist" class="form-inline" action="" method="post">
<input type="hidden" name="t" id="t" value="UserLevels">
<div id="gmp_UserLevels" class="ewGridMiddlePanel">
<% If UserLevels_list.TotalRecs > 0 Then %>
<table cellspacing="0" data-rowhighlightclass="ewTableHighlightRow" data-rowselectclass="ewTableSelectRow" data-roweditclass="ewTableEditRow" class="table table-bordered table-striped table-condensed">
<%= UserLevels.TableCustomInnerHTML %>
<thead><!-- Table header -->
	<tr class="ewTableHeader">
<%
Call UserLevels_list.RenderListOptions()

' Render list options (header, left)
UserLevels_list.ListOptions.Render "header", "left"
%>
<% If UserLevels.UserLevelName.Visible Then ' UserLevelName %>
		<td><%= UserLevels.UserLevelName.FldCaption %></td>
<% End If %>		
<%

' Render list options (header, right)
UserLevels_list.ListOptions.Render "header", "right"
%>
	</tr>
</thead>
<tbody><!-- Table body -->
<%
If (UserLevels.ExportAll And UserLevels.Export <> "") Then
	UserLevels_list.StopRec = UserLevels_list.TotalRecs
Else

	' Set the last record to display
	If UserLevels_list.TotalRecs > UserLevels_list.StartRec + UserLevels_list.DisplayRecs - 1 Then
		UserLevels_list.StopRec = UserLevels_list.StartRec + UserLevels_list.DisplayRecs - 1
	Else
		UserLevels_list.StopRec = UserLevels_list.TotalRecs
	End If
End If

' Move to first record
UserLevels_list.RecCnt = UserLevels_list.StartRec - 1
If Not UserLevels_list.Recordset.Eof Then
	UserLevels_list.Recordset.MoveFirst
	If UserLevels_list.StartRec > 1 Then UserLevels_list.Recordset.Move UserLevels_list.StartRec - 1
ElseIf Not UserLevels.AllowAddDeleteRow And UserLevels_list.StopRec = 0 Then
	UserLevels_list.StopRec = UserLevels.GridAddRowCount
End If

' Initialize Aggregate
UserLevels.RowType = EW_ROWTYPE_AGGREGATEINIT
Call UserLevels.ResetAttrs()
Call UserLevels_list.RenderRow()
UserLevels_list.RowCnt = 0

' Output date rows
Do While CLng(UserLevels_list.RecCnt) < CLng(UserLevels_list.StopRec)
	UserLevels_list.RecCnt = UserLevels_list.RecCnt + 1
	If CLng(UserLevels_list.RecCnt) >= CLng(UserLevels_list.StartRec) Then
		UserLevels_list.RowCnt = UserLevels_list.RowCnt + 1

	' Set up key count
	UserLevels_list.KeyCount = UserLevels_list.RowIndex
	Call UserLevels.ResetAttrs()
	UserLevels.CssClass = ""
	If UserLevels.CurrentAction = "gridadd" Then
	Else
		Call UserLevels_list.LoadRowValues(UserLevels_list.Recordset) ' Load row values
	End If
	UserLevels.RowType = EW_ROWTYPE_VIEW ' Render view
	UserLevels.RowAttrs.AddAttributes Array(Array("onmouseover", "ew_MouseOver(event, this);"), Array("onmouseout", "ew_MouseOut(event, this);"), Array("onclick", "ew_Click(event, this);"))

	' Render row
	Call UserLevels_list.RenderRow()

	' Render list options
	Call UserLevels_list.RenderListOptions()
%>
	<tr<%= UserLevels.RowAttributes %>>
<%

' Render list options (body, left)
UserLevels_list.ListOptions.Render "body", "left"
%>
	<% If UserLevels.UserLevelName.Visible Then ' UserLevelName %>
		<td<%= UserLevels.UserLevelName.CellAttributes %>>
<div<%= UserLevels.UserLevelName.ViewAttributes %>><%= UserLevels.UserLevelName.ListViewValue %></div>
</td>
	<% End If %>
<%

' Render list options (body, right)
UserLevels_list.ListOptions.Render "body", "right"
%>
	</tr>
<%
	End If
	If UserLevels.CurrentAction <> "gridadd" Then
		UserLevels_list.Recordset.MoveNext()
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
UserLevels_list.Recordset.Close
Set UserLevels_list.Recordset = Nothing
%>
<% If UserLevels.Export = "" Then %>
<div class="ewGridLowerPanel">
<% If UserLevels.CurrentAction <> "gridadd" And UserLevels.CurrentAction <> "gridedit" Then %>
<form name="ewpagerform" id="ewpagerform" class="form-inline" action="<%= ew_CurrentPage %>">
<table border="0" cellspacing="0" cellpadding="0" class="ewPager">
	<tr>
		<td>
<% If Not IsObject(UserLevels_list.Pager) Then Set UserLevels_list.Pager = ew_NewPrevNextPager(UserLevels_list.StartRec, UserLevels_list.DisplayRecs, UserLevels_list.TotalRecs) %>
<% If UserLevels_list.Pager.RecordCount > 0 Then %>
	<table border="0" cellspacing="0" cellpadding="0"><tr><td><span class="aspmaker"><%= Language.Phrase("Page") %>&nbsp;</span></td>
<!--first page button-->
	<% If UserLevels_list.Pager.FirstButton.Enabled Then %>
	<td><a href="<%= UserLevels_list.PageUrl %>start=<%= UserLevels_list.Pager.FirstButton.Start %>"><img src="images/first.gif" alt="<%= Language.Phrase("PagerFirst") %>" width="16" height="16" border="0"></a></td>
	<% Else %>
	<td><img src="images/firstdisab.gif" alt="<%= Language.Phrase("PagerFirst") %>" width="16" height="16" border="0"></td>
	<% End If %>
<!--previous page button-->
	<% If UserLevels_list.Pager.PrevButton.Enabled Then %>
	<td><a href="<%= UserLevels_list.PageUrl %>start=<%= UserLevels_list.Pager.PrevButton.Start %>"><img src="images/prev.gif" alt="<%= Language.Phrase("PagerPrevious") %>" width="16" height="16" border="0"></a></td>
	<% Else %>
	<td><img src="images/prevdisab.gif" alt="<%= Language.Phrase("PagerPrevious") %>" width="16" height="16" border="0"></td>
	<% End If %>
<!--current page number-->
	<td><input type="text" name="<%= EW_TABLE_PAGE_NO %>" id="<%= EW_TABLE_PAGE_NO %>" value="<%= UserLevels_list.Pager.CurrentPage %>" size="4"></td>
<!--next page button-->
	<% If UserLevels_list.Pager.NextButton.Enabled Then %>
	<td><a href="<%= UserLevels_list.PageUrl %>start=<%= UserLevels_list.Pager.NextButton.Start %>"><img src="images/next.gif" alt="<%= Language.Phrase("PagerNext") %>" width="16" height="16" border="0"></a></td>
	<% Else %>
	<td><img src="images/nextdisab.gif" alt="<%= Language.Phrase("PagerNext") %>" width="16" height="16" border="0"></td>
	<% End If %>
<!--last page button-->
	<% If UserLevels_list.Pager.LastButton.Enabled Then %>
	<td><a href="<%= UserLevels_list.PageUrl %>start=<%= UserLevels_list.Pager.LastButton.Start %>"><img src="images/last.gif" alt="<%= Language.Phrase("PagerLast") %>" width="16" height="16" border="0"></a></td>	
	<% Else %>
	<td><img src="images/lastdisab.gif" alt="<%= Language.Phrase("PagerLast") %>" width="16" height="16" border="0"></td>
	<% End If %>
	<td><span class="aspmaker">&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= UserLevels_list.Pager.PageCount %></span></td>
	</tr></table>
	</td>	
	<td>&nbsp;&nbsp;&nbsp;&nbsp;</td>
	<td>
	<span class="aspmaker"><%= Language.Phrase("Record") %>&nbsp;<%= UserLevels_list.Pager.FromIndex %>&nbsp;<%= Language.Phrase("To") %>&nbsp;<%= UserLevels_list.Pager.ToIndex %>&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= UserLevels_list.Pager.RecordCount %></span>
<% Else %>
	<% If Security.CanList Then %>
	<% If UserLevels_list.SearchWhere = "0=101" Then %>
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
<a class="btn btn-success" href="<%= UserLevels_list.AddUrl %>"><i class="icon-plus icon-white"></i>&nbsp;<%= Language.Phrase("AddLink") %></a>&nbsp;&nbsp;
<% End If %>
<% If UserLevels_list.TotalRecs > 0 Then %>
<% If Security.CanDelete Then %>
<a class="btn btn-danger" href="" onclick="ew_SubmitSelected(document.fUserLevelslist, '<%= UserLevels_list.MultiDeleteUrl %>');return false;"><i class="icon-trash icon-white"></i>&nbsp;<%= Language.Phrase("DeleteSelectedLink") %></a>&nbsp;&nbsp;
<% End If %>
<% End If %>
</span>
</div>
<% End If %>
</td></tr></table>
<% If UserLevels.Export = "" And UserLevels.CurrentAction = "" Then %>
<% End If %>
<%
UserLevels_list.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If UserLevels.Export = "" Then %>
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
Set UserLevels_list = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cUserLevels_list

	' Page ID
	Public Property Get PageID()
		PageID = "list"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "UserLevels"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "UserLevels_list"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If UserLevels.UseTokenInUrl Then PageUrl = PageUrl & "t=" & UserLevels.TableVar & "&" ' add page token
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
		If UserLevels.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (UserLevels.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (UserLevels.TableVar = Request.QueryString("t"))
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
		If IsEmpty(UserLevels) Then Set UserLevels = New cUserLevels
		Set Table = UserLevels

		' Initialize urls
		ExportPrintUrl = PageUrl & "export=print"
		ExportExcelUrl = PageUrl & "export=excel"
		ExportWordUrl = PageUrl & "export=word"
		ExportHtmlUrl = PageUrl & "export=html"
		ExportXmlUrl = PageUrl & "export=xml"
		ExportCsvUrl = PageUrl & "export=csv"
		AddUrl = "userlevelsadd.asp"
		InlineAddUrl = PageUrl & "a=add"
		GridAddUrl = PageUrl & "a=gridadd"
		GridEditUrl = PageUrl & "a=gridedit"
		MultiDeleteUrl = "userlevelsdelete.asp"
		MultiUpdateUrl = "userlevelsupdate.asp"

		' Initialize other table object
		If IsEmpty(Users) Then Set Users = New cUsers

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "list"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "UserLevels"

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
				UserLevels.GridAddRowCount = gridaddcnt
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
		Set UserLevels = Nothing
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
			If UserLevels.Export <> "" Or UserLevels.CurrentAction = "gridadd" Or UserLevels.CurrentAction = "gridedit" Then
				ListOptions.HideAllOptions()
				ExportOptions.HideAllOptions()
			End If

			' Get basic search values
			Call LoadBasicSearchValues()

			' Restore search parms from Session
			Call RestoreSearchParms()

			' Call Recordset SearchValidated event
			Call UserLevels.Recordset_SearchValidated()

			' Set Up Sorting Order
			SetUpSortOrder()

			' Get basic search criteria
			If gsSearchError = "" Then
				sSrchBasic = BasicSearchWhere()
			End If
		End If ' End Validate Request

		' Restore display records
		If UserLevels.RecordsPerPage <> "" Then
			DisplayRecs = UserLevels.RecordsPerPage ' Restore from Session
		Else
			DisplayRecs = 50 ' Load default
		End If

		' Load Sorting Order
		LoadSortOrder()

		' Build search criteria
		Call ew_AddFilter(SearchWhere, sSrchAdvanced)
		Call ew_AddFilter(SearchWhere, sSrchBasic)

		' Call Recordset Searching event
		Call UserLevels.Recordset_Searching(SearchWhere)

		' Save search criteria
		If SearchWhere <> "" Then
			If sSrchBasic = "" Then Call ResetBasicSearchParms()
			UserLevels.SearchWhere = SearchWhere ' Save to Session
			If Not RestoreSearch Then
				StartRec = 1 ' Reset start record counter
				UserLevels.StartRecordNumber = StartRec
			End If
		Else
			SearchWhere = UserLevels.SearchWhere
		End If
		sFilter = ""
		If Not Security.CanList Then
			sFilter = "(0=1)" ' Filter all records
		End If
		Call ew_AddFilter(sFilter, DbDetailFilter)
		Call ew_AddFilter(sFilter, SearchWhere)

		' Set up filter in Session
		UserLevels.SessionWhere = sFilter
		UserLevels.CurrentFilter = ""
	End Sub

	' -----------------------------------------------------------------
	' Return Basic Search sql
	'
	Function BasicSearchSQL(Keyword)
		Dim sWhere
		sWhere = ""
			Call BuildBasicSearchSQL(sWhere, UserLevels.UserLevelName, Keyword)
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
		sSearchKeyword = UserLevels.BasicSearchKeyword
		sSearchType = UserLevels.BasicSearchType
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
			UserLevels.SessionBasicSearchKeyword = sSearchKeyword
			UserLevels.SessionBasicSearchType = sSearchType
		End If
		BasicSearchWhere = sSearchStr
	End Function

	' -----------------------------------------------------------------
	' Clear all search parameters
	'
	Sub ResetSearchParms()

		' Clear search where
		SearchWhere = ""
		UserLevels.SearchWhere = SearchWhere

		' Clear basic search parameters
		Call ResetBasicSearchParms()
	End Sub

	' -----------------------------------------------------------------
	' Clear all basic search parameters
	'
	Sub ResetBasicSearchParms()

		' Clear basic search parameters
		UserLevels.SessionBasicSearchKeyword = ""
		UserLevels.SessionBasicSearchType = ""
	End Sub

	' -----------------------------------------------------------------
	' Restore all search parameters
	'
	Sub RestoreSearchParms()
		Dim bRestore
		bRestore = True
		If UserLevels.BasicSearchKeyword & "" <> "" Then bRestore = False
		RestoreSearch = bRestore
		If bRestore Then

			' Restore basic search values
			UserLevels.BasicSearchKeyword = UserLevels.SessionBasicSearchKeyword
			UserLevels.BasicSearchType = UserLevels.SessionBasicSearchType
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
			UserLevels.CurrentOrder = Request.QueryString("order")
			UserLevels.CurrentOrderType = Request.QueryString("ordertype")

			' Field UserLevelID
			Call UserLevels.UpdateSort(UserLevels.UserLevelID)

			' Field UserLevelName
			Call UserLevels.UpdateSort(UserLevels.UserLevelName)
			UserLevels.StartRecordNumber = 1 ' Reset start position
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load Sort Order parameters
	'
	Sub LoadSortOrder()
		Dim sOrderBy
		sOrderBy = UserLevels.SessionOrderBy ' Get order by from Session
		If sOrderBy = "" Then
			If UserLevels.SqlOrderBy <> "" Then
				sOrderBy = UserLevels.SqlOrderBy
				UserLevels.SessionOrderBy = sOrderBy
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
				UserLevels.SessionOrderBy = sOrderBy
				UserLevels.UserLevelID.Sort = ""
				UserLevels.UserLevelName.Sort = ""
			End If

			' Reset start position
			StartRec = 1
			UserLevels.StartRecordNumber = StartRec
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
		ListOptions.Add("userpermission")
		ListOptions.GetItem("userpermission").CssStyle = "white-space: nowrap;"
		ListOptions.GetItem("userpermission").Visible = Security.IsAdmin()
		ListOptions.GetItem("userpermission").OnLeft = False
		ListOptions.Add("checkbox")
		ListOptions.GetItem("checkbox").CssStyle = "white-space: nowrap;"
		ListOptions.GetItem("checkbox").Visible = Security.CanDelete
		ListOptions.GetItem("checkbox").OnLeft = False
		ListOptions.GetItem("checkbox").Header = "<input type=""checkbox"" name=""key"" id=""key"" class=""aspmaker"" onclick=""UserLevels_list.SelectAllKey(this);"">"
		Call ListOptions_Load()
	End Sub

	' Render list options
	Sub RenderListOptions()
		Dim item, links
		ListOptions.LoadDefault()
		If Security.CanEdit And ListOptions.GetItem("edit").Visible Then
			Set item = ListOptions.GetItem("edit")
			item.Body = "<a class=""ewRowLink"" href=""" & EditUrl & """><i class=""icon-pencil""></i></a>"
		End If
		If UserLevels.UserLevelID.CurrentValue < 0 Then
			ListOptions.GetItem("userpermission").Body = "-"
		Else
			ListOptions.GetItem("userpermission").Body = "<a class=""ewRowLink"" href=""" & "userpriv.asp?UserLevelID=" & UserLevels.UserLevelID.CurrentValue & """><i class=""icon-lock""></i></a>"
		End If
		If Security.CanDelete And ListOptions.GetItem("checkbox").Visible Then
			ListOptions.GetItem("checkbox").Body = "<input type=""checkbox"" name=""key_m"" id=""key_m"" value=""" & ew_HtmlEncode(UserLevels.UserLevelID.CurrentValue) & """ class=""aspmaker"" onclick='ew_ClickMultiCheckbox(this);'>"
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
				UserLevels.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					UserLevels.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = UserLevels.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			UserLevels.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			UserLevels.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			UserLevels.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	'  Load basic search values
	'
	Function LoadBasicSearchValues()
		UserLevels.BasicSearchKeyword = Request.QueryString(EW_TABLE_BASIC_SEARCH)
		UserLevels.BasicSearchType = Request.QueryString(EW_TABLE_BASIC_SEARCH_TYPE)
	End Function

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = UserLevels.CurrentFilter
		Call UserLevels.Recordset_Selecting(sFilter)
		UserLevels.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = UserLevels.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call UserLevels.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = UserLevels.KeyFilter

		' Call Row Selecting event
		Call UserLevels.Row_Selecting(sFilter)

		' Load sql based on filter
		UserLevels.CurrentFilter = sFilter
		sSql = UserLevels.SQL
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
		Call UserLevels.Row_Selected(RsRow)
		UserLevels.UserLevelID.DbValue = RsRow("UserLevelID")
		If  IsNull(UserLevels.UserLevelID.CurrentValue) Then
			UserLevels.UserLevelID.CurrentValue = 0
		Else
			UserLevels.UserLevelID.CurrentValue = CLng(UserLevels.UserLevelID.CurrentValue)
		End If
		UserLevels.UserLevelName.DbValue = RsRow("UserLevelName")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If UserLevels.GetKey("UserLevelID")&"" <> "" Then
			UserLevels.UserLevelID.CurrentValue = UserLevels.GetKey("UserLevelID") ' UserLevelID
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			UserLevels.CurrentFilter = UserLevels.KeyFilter
			Dim sSql
			sSql = UserLevels.SQL
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
		ViewUrl = UserLevels.ViewUrl
		EditUrl = UserLevels.EditUrl("")
		InlineEditUrl = UserLevels.InlineEditUrl
		CopyUrl = UserLevels.CopyUrl("")
		InlineCopyUrl = UserLevels.InlineCopyUrl
		DeleteUrl = UserLevels.DeleteUrl

		' Call Row Rendering event
		Call UserLevels.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' UserLevelID
		' UserLevelName
		' -----------
		'  View  Row
		' -----------

		If UserLevels.RowType = EW_ROWTYPE_VIEW Then ' View row

			' UserLevelID
			UserLevels.UserLevelID.ViewValue = UserLevels.UserLevelID.CurrentValue
			UserLevels.UserLevelID.ViewCustomAttributes = ""

			' UserLevelName
			UserLevels.UserLevelName.ViewValue = UserLevels.UserLevelName.CurrentValue
			UserLevels.UserLevelName.ViewCustomAttributes = ""

			' View refer script
			' UserLevelID

			UserLevels.UserLevelID.LinkCustomAttributes = ""
			UserLevels.UserLevelID.HrefValue = ""
			UserLevels.UserLevelID.TooltipValue = ""

			' UserLevelName
			UserLevels.UserLevelName.LinkCustomAttributes = ""
			UserLevels.UserLevelName.HrefValue = ""
			UserLevels.UserLevelName.TooltipValue = ""
		End If

		' Call Row Rendered event
		If UserLevels.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call UserLevels.Row_Rendered()
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
