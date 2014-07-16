<%@ CodePage="1252" LCID="2057" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="ewcfg9.asp"-->
<!--#include file="themesinfo.asp"-->
<!--#include file="usersinfo.asp"-->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim Themes_list
Set Themes_list = New cThemes_list
Set Page = Themes_list

' Page init processing
Call Themes_list.Page_Init()

' Page main processing
Call Themes_list.Page_Main()
%>
<!--#include file="header.asp"-->
<% If Themes.Export = "" Then %>
<script type="text/javascript">
<!--
// Create page object
var Themes_list = new ew_Page("Themes_list");
// page properties
Themes_list.PageID = "list"; // page ID
Themes_list.FormID = "fThemeslist"; // form ID
var EW_PAGE_ID = Themes_list.PageID; // for backward compatibility
// extend page with Form_CustomValidate function
Themes_list.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
Themes_list.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
Themes_list.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Themes_list.ValidateRequired = false; // no JavaScript validation
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
<% If (Themes.Export = "") Or (EW_EXPORT_MASTER_RECORD And Themes.Export = "print") Then %>
<% End If %>
<% Themes_list.ShowPageHeader() %>
<%

' Load recordset
Set Themes_list.Recordset = Themes_list.LoadRecordset()
	Themes_list.TotalRecs = Themes_list.Recordset.RecordCount
	Themes_list.StartRec = 1
	If Themes_list.DisplayRecs <= 0 Then ' Display all records
		Themes_list.DisplayRecs = Themes_list.TotalRecs
	End If
	If Not (Themes.ExportAll And Themes.Export <> "") Then
		Themes_list.SetUpStartRec() ' Set up start record position
	End If
%>
<% If Security.CanSearch Then %>
<% If Themes.Export = "" And Themes.CurrentAction = "" Then %>
<a href="javascript:ew_ToggleSearchPanel(Themes_list);" style="text-decoration: none;"><img id="Themes_list_SearchImage" src="images/collapse.gif" alt="" width="9" height="9" border="0"></a><span class="aspmaker">&nbsp;<%= Language.Phrase("Search") %></span><br>
<div id="Themes_list_SearchPanel">
<form name="fThemeslistsrch" id="fThemeslistsrch" class="form-inline" action="<%= ew_CurrentPage %>">
<input type="hidden" id="t" name="t" value="Themes">
<div>
<div id="xsr_1" class="input-append">
	<input type="text" name="<%= EW_TABLE_BASIC_SEARCH %>" id="<%= EW_TABLE_BASIC_SEARCH %>" size="20" value="<%= ew_HtmlEncode(themes.SessionBasicSearchKeyword) %>">
	<button class="btn" type="Submit" name="Submit" id="Submit"><%= ew_BtnCaption(Language.Phrase("QuickSearchBtn")) %></button>
	<a class="btn" href="<%= themes_list.PageUrl %>cmd=reset"><i class="icon-refresh"></i>&nbsp;<%= Language.Phrase("ShowAll") %></a>
</div>
</div>
</form>
</div>
<% End If %>
<% End If %>
<% Themes_list.ShowMessage %>
<br>
<table cellspacing="0">
<form name="fThemeslist" id="fThemeslist" class="form-inline" action="" method="post">
<input type="hidden" name="t" id="t" value="Themes">
<div id="gmp_Themes" class="ewGridMiddlePanel">
<% If Themes_list.TotalRecs > 0 Then %>
<table cellspacing="0" data-rowhighlightclass="ewTableHighlightRow" data-rowselectclass="ewTableSelectRow" data-roweditclass="ewTableEditRow" class="table table-bordered table-striped table-condensed">
<%= Themes.TableCustomInnerHTML %>
<thead><!-- Table header -->
	<tr class="ewTableHeader">
<%
Call Themes_list.RenderListOptions()

' Render list options (header, left)
Themes_list.ListOptions.Render "header", "left"
%>
<% If Themes.Theme_Name.Visible Then ' Theme_Name %>
		<td><strong>Theme Name</strong></td>
<% End If %>		
<%

' Render list options (header, right)
Themes_list.ListOptions.Render "header", "right"
%>
	</tr>
</thead>
<tbody><!-- Table body -->
<%
If (Themes.ExportAll And Themes.Export <> "") Then
	Themes_list.StopRec = Themes_list.TotalRecs
Else

	' Set the last record to display
	If Themes_list.TotalRecs > Themes_list.StartRec + Themes_list.DisplayRecs - 1 Then
		Themes_list.StopRec = Themes_list.StartRec + Themes_list.DisplayRecs - 1
	Else
		Themes_list.StopRec = Themes_list.TotalRecs
	End If
End If

' Move to first record
Themes_list.RecCnt = Themes_list.StartRec - 1
If Not Themes_list.Recordset.Eof Then
	Themes_list.Recordset.MoveFirst
	If Themes_list.StartRec > 1 Then Themes_list.Recordset.Move Themes_list.StartRec - 1
ElseIf Not Themes.AllowAddDeleteRow And Themes_list.StopRec = 0 Then
	Themes_list.StopRec = Themes.GridAddRowCount
End If

' Initialize Aggregate
Themes.RowType = EW_ROWTYPE_AGGREGATEINIT
Call Themes.ResetAttrs()
Call Themes_list.RenderRow()
Themes_list.RowCnt = 0

' Output date rows
Do While CLng(Themes_list.RecCnt) < CLng(Themes_list.StopRec)
	Themes_list.RecCnt = Themes_list.RecCnt + 1
	If CLng(Themes_list.RecCnt) >= CLng(Themes_list.StartRec) Then
		Themes_list.RowCnt = Themes_list.RowCnt + 1

	' Set up key count
	Themes_list.KeyCount = Themes_list.RowIndex
	Call Themes.ResetAttrs()
	Themes.CssClass = ""
	If Themes.CurrentAction = "gridadd" Then
	Else
		Call Themes_list.LoadRowValues(Themes_list.Recordset) ' Load row values
	End If
	Themes.RowType = EW_ROWTYPE_VIEW ' Render view
	Themes.RowAttrs.AddAttributes Array(Array("onmouseover", "ew_MouseOver(event, this);"), Array("onmouseout", "ew_MouseOut(event, this);"), Array("onclick", "ew_Click(event, this);"))

	' Render row
	Call Themes_list.RenderRow()

	' Render list options
	Call Themes_list.RenderListOptions()
%>
	<tr<%= Themes.RowAttributes %>>
<%

' Render list options (body, left)
Themes_list.ListOptions.Render "body", "left"
%>
	<% If Themes.Theme_Name.Visible Then ' Theme_Name %>
		<td<%= Themes.Theme_Name.CellAttributes %>>
<div<%= Themes.Theme_Name.ViewAttributes %>><%= Themes.Theme_Name.ListViewValue %></div>
<a name="<%= Themes_list.PageObjName & "_row_" & Themes_list.RowCnt %>" id="<%= Themes_list.PageObjName & "_row_" & Themes_list.RowCnt %>"></a></td>
	<% End If %>
<%

' Render list options (body, right)
Themes_list.ListOptions.Render "body", "right"
%>
	</tr>
<%
	End If
	If Themes.CurrentAction <> "gridadd" Then
		Themes_list.Recordset.MoveNext()
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
Themes_list.Recordset.Close
Set Themes_list.Recordset = Nothing
%>
<% If Themes.Export = "" Then %>
<div>
<% If Themes.CurrentAction <> "gridadd" And Themes.CurrentAction <> "gridedit" Then %>
<form name="ewpagerform" id="ewpagerform" class="form-inline" action="<%= ew_CurrentPage %>">
<table border="0" cellspacing="0" cellpadding="0" class="ewPager">
	<tr>
		<td>
<% If Not IsObject(Themes_list.Pager) Then Set Themes_list.Pager = ew_NewPrevNextPager(Themes_list.StartRec, Themes_list.DisplayRecs, Themes_list.TotalRecs) %>
<% If Themes_list.Pager.RecordCount > 0 Then %>
	<table border="0" cellspacing="0" cellpadding="0"><tr><td><span class="aspmaker"><%= Language.Phrase("Page") %>&nbsp;</span></td>
<!--first page button-->
	<% If Themes_list.Pager.FirstButton.Enabled Then %>
	<td><a href="<%= Themes_list.PageUrl %>start=<%= Themes_list.Pager.FirstButton.Start %>"><img src="images/first.gif" alt="<%= Language.Phrase("PagerFirst") %>" width="16" height="16" border="0"></a></td>
	<% Else %>
	<td><img src="images/firstdisab.gif" alt="<%= Language.Phrase("PagerFirst") %>" width="16" height="16" border="0"></td>
	<% End If %>
<!--previous page button-->
	<% If Themes_list.Pager.PrevButton.Enabled Then %>
	<td><a href="<%= Themes_list.PageUrl %>start=<%= Themes_list.Pager.PrevButton.Start %>"><img src="images/prev.gif" alt="<%= Language.Phrase("PagerPrevious") %>" width="16" height="16" border="0"></a></td>
	<% Else %>
	<td><img src="images/prevdisab.gif" alt="<%= Language.Phrase("PagerPrevious") %>" width="16" height="16" border="0"></td>
	<% End If %>
<!--current page number-->
	<td><input type="text" name="<%= EW_TABLE_PAGE_NO %>" id="<%= EW_TABLE_PAGE_NO %>" value="<%= Themes_list.Pager.CurrentPage %>" size="4"></td>
<!--next page button-->
	<% If Themes_list.Pager.NextButton.Enabled Then %>
	<td><a href="<%= Themes_list.PageUrl %>start=<%= Themes_list.Pager.NextButton.Start %>"><img src="images/next.gif" alt="<%= Language.Phrase("PagerNext") %>" width="16" height="16" border="0"></a></td>
	<% Else %>
	<td><img src="images/nextdisab.gif" alt="<%= Language.Phrase("PagerNext") %>" width="16" height="16" border="0"></td>
	<% End If %>
<!--last page button-->
	<% If Themes_list.Pager.LastButton.Enabled Then %>
	<td><a href="<%= Themes_list.PageUrl %>start=<%= Themes_list.Pager.LastButton.Start %>"><img src="images/last.gif" alt="<%= Language.Phrase("PagerLast") %>" width="16" height="16" border="0"></a></td>	
	<% Else %>
	<td><img src="images/lastdisab.gif" alt="<%= Language.Phrase("PagerLast") %>" width="16" height="16" border="0"></td>
	<% End If %>
	<td><span class="aspmaker">&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= Themes_list.Pager.PageCount %></span></td>
	</tr></table>
	</td>	
	<td>&nbsp;&nbsp;&nbsp;&nbsp;</td>
	<td>
	<span class="aspmaker"><%= Language.Phrase("Record") %>&nbsp;<%= Themes_list.Pager.FromIndex %>&nbsp;<%= Language.Phrase("To") %>&nbsp;<%= Themes_list.Pager.ToIndex %>&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= Themes_list.Pager.RecordCount %></span>
<% Else %>
	<% If Security.CanList Then %>
	<% If Themes_list.SearchWhere = "0=101" Then %>
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
<span>
<% If Security.CanAdd Then %>
<a class="btn btn-success" href="<%= Themes_list.AddUrl %>"><i class="icon-plus icon-white"></i>&nbsp;New Theme</a>&nbsp;&nbsp;
<% End If %>
<% If Themes_list.TotalRecs > 0 Then %>
<% If Security.CanDelete Then %>
<a class="btn btn-danger" href="" onclick="ew_SubmitSelected(document.fThemeslist, '<%= Themes_list.MultiDeleteUrl %>');return false;"><i class="icon-remove icon-white"></i>&nbsp;<%= Language.Phrase("DeleteSelectedLink") %></a>&nbsp;&nbsp;
<% End If %>
<% End If %>
</span>
</div>
<% End If %>
</td></tr></table>
<% If Themes.Export = "" And Themes.CurrentAction = "" Then %>
<% End If %>
<%
Themes_list.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If Themes.Export = "" Then %>
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
Set Themes_list = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cThemes_list

	' Page ID
	Public Property Get PageID()
		PageID = "list"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Themes"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Themes_list"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If Themes.UseTokenInUrl Then PageUrl = PageUrl & "t=" & Themes.TableVar & "&" ' add page token
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
		If Themes.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (Themes.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (Themes.TableVar = Request.QueryString("t"))
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
		If IsEmpty(Themes) Then Set Themes = New cThemes
		Set Table = Themes

		' Initialize urls
		ExportPrintUrl = PageUrl & "export=print"
		ExportExcelUrl = PageUrl & "export=excel"
		ExportWordUrl = PageUrl & "export=word"
		ExportHtmlUrl = PageUrl & "export=html"
		ExportXmlUrl = PageUrl & "export=xml"
		ExportCsvUrl = PageUrl & "export=csv"
		AddUrl = "themesadd.asp"
		InlineAddUrl = PageUrl & "a=add"
		GridAddUrl = PageUrl & "a=gridadd"
		GridEditUrl = PageUrl & "a=gridedit"
		MultiDeleteUrl = "themesdelete.asp"
		MultiUpdateUrl = "themesupdate.asp"

		' Initialize other table object
		If IsEmpty(Users) Then Set Users = New cUsers

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "list"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "Themes"

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
				Themes.GridAddRowCount = gridaddcnt
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
		Set Themes = Nothing
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
			If Themes.Export <> "" Or Themes.CurrentAction = "gridadd" Or Themes.CurrentAction = "gridedit" Then
				ListOptions.HideAllOptions()
				ExportOptions.HideAllOptions()
			End If

			' Get basic search values
			Call LoadBasicSearchValues()

			' Restore search parms from Session
			Call RestoreSearchParms()

			' Call Recordset SearchValidated event
			Call Themes.Recordset_SearchValidated()

			' Set Up Sorting Order
			SetUpSortOrder()

			' Get basic search criteria
			If gsSearchError = "" Then
				sSrchBasic = BasicSearchWhere()
			End If
		End If ' End Validate Request

		' Restore display records
		If Themes.RecordsPerPage <> "" Then
			DisplayRecs = Themes.RecordsPerPage ' Restore from Session
		Else
			DisplayRecs = 50 ' Load default
		End If

		' Load Sorting Order
		LoadSortOrder()

		' Build search criteria
		Call ew_AddFilter(SearchWhere, sSrchAdvanced)
		Call ew_AddFilter(SearchWhere, sSrchBasic)

		' Call Recordset Searching event
		Call Themes.Recordset_Searching(SearchWhere)

		' Save search criteria
		If SearchWhere <> "" Then
			If sSrchBasic = "" Then Call ResetBasicSearchParms()
			Themes.SearchWhere = SearchWhere ' Save to Session
			If Not RestoreSearch Then
				StartRec = 1 ' Reset start record counter
				Themes.StartRecordNumber = StartRec
			End If
		Else
			SearchWhere = Themes.SearchWhere
		End If
		sFilter = ""
		If Not Security.CanList Then
			sFilter = "(0=1)" ' Filter all records
		End If
		Call ew_AddFilter(sFilter, DbDetailFilter)
		Call ew_AddFilter(sFilter, SearchWhere)

		' Set up filter in Session
		Themes.SessionWhere = sFilter
		Themes.CurrentFilter = ""
	End Sub

	' -----------------------------------------------------------------
	' Return Basic Search sql
	'
	Function BasicSearchSQL(Keyword)
		Dim sWhere
		sWhere = ""
			Call BuildBasicSearchSQL(sWhere, Themes.Theme_Name, Keyword)
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
		sSearchKeyword = Themes.BasicSearchKeyword
		sSearchType = Themes.BasicSearchType
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
			Themes.SessionBasicSearchKeyword = sSearchKeyword
			Themes.SessionBasicSearchType = sSearchType
		End If
		BasicSearchWhere = sSearchStr
	End Function

	' -----------------------------------------------------------------
	' Clear all search parameters
	'
	Sub ResetSearchParms()

		' Clear search where
		SearchWhere = ""
		Themes.SearchWhere = SearchWhere

		' Clear basic search parameters
		Call ResetBasicSearchParms()
	End Sub

	' -----------------------------------------------------------------
	' Clear all basic search parameters
	'
	Sub ResetBasicSearchParms()

		' Clear basic search parameters
		Themes.SessionBasicSearchKeyword = ""
		Themes.SessionBasicSearchType = ""
	End Sub

	' -----------------------------------------------------------------
	' Restore all search parameters
	'
	Sub RestoreSearchParms()
		Dim bRestore
		bRestore = True
		If Themes.BasicSearchKeyword & "" <> "" Then bRestore = False
		RestoreSearch = bRestore
		If bRestore Then

			' Restore basic search values
			Themes.BasicSearchKeyword = Themes.SessionBasicSearchKeyword
			Themes.BasicSearchType = Themes.SessionBasicSearchType
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
			Themes.CurrentOrder = Request.QueryString("order")
			Themes.CurrentOrderType = Request.QueryString("ordertype")

			' Field Theme_Name
			Call Themes.UpdateSort(Themes.Theme_Name)
			Themes.StartRecordNumber = 1 ' Reset start position
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load Sort Order parameters
	'
	Sub LoadSortOrder()
		Dim sOrderBy
		sOrderBy = Themes.SessionOrderBy ' Get order by from Session
		If sOrderBy = "" Then
			If Themes.SqlOrderBy <> "" Then
				sOrderBy = Themes.SqlOrderBy
				Themes.SessionOrderBy = sOrderBy
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
				Themes.SessionOrderBy = sOrderBy
				Themes.Theme_Name.Sort = ""
			End If

			' Reset start position
			StartRec = 1
			Themes.StartRecordNumber = StartRec
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
		ListOptions.Add("copy")
		ListOptions.GetItem("copy").CssStyle = "white-space: nowrap;"
		ListOptions.GetItem("copy").Visible = Security.CanAdd
		ListOptions.GetItem("copy").OnLeft = False
		ListOptions.Add("checkbox")
		ListOptions.GetItem("checkbox").CssStyle = "white-space: nowrap;"
		ListOptions.GetItem("checkbox").Visible = Security.CanDelete
		ListOptions.GetItem("checkbox").OnLeft = False
		ListOptions.GetItem("checkbox").Header = "<input type=""checkbox"" name=""key"" id=""key"" class=""aspmaker"" onclick=""Themes_list.SelectAllKey(this);"">"
		Call ListOptions_Load()
	End Sub

	' Render list options
	Sub RenderListOptions()
		Dim item, links
		ListOptions.LoadDefault()
		If Security.CanDelete And ListOptions.GetItem("checkbox").Visible Then
			ListOptions.GetItem("checkbox").Body = "<input type=""checkbox"" name=""key_m"" id=""key_m"" value=""" & ew_HtmlEncode(Themes.Theme_Name.CurrentValue) & """ class=""aspmaker"" onclick='ew_ClickMultiCheckbox(this);'>"
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
				Themes.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					Themes.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = Themes.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			Themes.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			Themes.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			Themes.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	'  Load basic search values
	'
	Function LoadBasicSearchValues()
		Themes.BasicSearchKeyword = Request.QueryString(EW_TABLE_BASIC_SEARCH)
		Themes.BasicSearchType = Request.QueryString(EW_TABLE_BASIC_SEARCH_TYPE)
	End Function

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = Themes.CurrentFilter
		Call Themes.Recordset_Selecting(sFilter)
		Themes.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = Themes.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call Themes.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = Themes.KeyFilter

		' Call Row Selecting event
		Call Themes.Row_Selecting(sFilter)

		' Load sql based on filter
		Themes.CurrentFilter = sFilter
		sSql = Themes.SQL
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
		Call Themes.Row_Selected(RsRow)
		Themes.Theme_Name.DbValue = RsRow("Theme_Name")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If Themes.GetKey("Theme_Name")&"" <> "" Then
			Themes.Theme_Name.CurrentValue = Themes.GetKey("Theme_Name") ' Theme_Name
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			Themes.CurrentFilter = Themes.KeyFilter
			Dim sSql
			sSql = Themes.SQL
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
		ViewUrl = Themes.ViewUrl
		EditUrl = Themes.EditUrl("")
		InlineEditUrl = Themes.InlineEditUrl
		CopyUrl = Themes.CopyUrl("")
		InlineCopyUrl = Themes.InlineCopyUrl
		DeleteUrl = Themes.DeleteUrl

		' Call Row Rendering event
		Call Themes.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' Theme_Name
		' -----------
		'  View  Row
		' -----------

		If Themes.RowType = EW_ROWTYPE_VIEW Then ' View row

			' Theme_Name
			Themes.Theme_Name.ViewValue = Themes.Theme_Name.CurrentValue
			Themes.Theme_Name.ViewCustomAttributes = ""

			' View refer script
			' Theme_Name

			Themes.Theme_Name.LinkCustomAttributes = ""
			Themes.Theme_Name.HrefValue = ""
			Themes.Theme_Name.TooltipValue = ""
		End If

		' Call Row Rendered event
		If Themes.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Themes.Row_Rendered()
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
