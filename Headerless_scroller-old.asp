<%@ CodePage="1252" LCID="2057" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="ewcfg9.asp"-->
<!--#include file="signage_noticesinfo.asp"-->
<!--#include file="usersinfo.asp"-->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim Approved_Notices_list
Set Approved_Notices_list = New cApproved_Notices_list
Set Page = Approved_Notices_list

' Page init processing
Call Approved_Notices_list.Page_Init()

' Page main processing
Call Approved_Notices_list.Page_Main()
%>
<!--#include file="headerless.asp"-->
<% If Approved_Notices.Export = "" Then %>
<script type="text/javascript">
<!--
// Create page object
var Approved_Notices_list = new ew_Page("Approved_Notices_list");
// page properties
Approved_Notices_list.PageID = "list"; // page ID
Approved_Notices_list.FormID = "fApproved_Noticeslist"; // form ID
var EW_PAGE_ID = Approved_Notices_list.PageID; // for backward compatibility
// extend page with Form_CustomValidate function
Approved_Notices_list.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
<% If EW_CLIENT_VALIDATE Then %>
Approved_Notices_list.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Approved_Notices_list.ValidateRequired = false; // no JavaScript validation
<% End If %>
//-->
</script>
<script type="text/javascript" src="ckeditor/ckeditor.js"></script>
<script type="text/javascript">
<!--
_width_multiplier = 20;
_height_multiplier = 60;
var ew_DHTMLEditors = [];
// update value from editor to textarea
function ew_UpdateTextArea() {
	if (typeof ew_DHTMLEditors != 'undefined' && typeof CKEDITOR != 'undefined') {			
		var inst;			
		for (inst in CKEDITOR.instances)
			CKEDITOR.instances[inst].updateElement();
	}
}
// update value from textarea to editor
function ew_UpdateDHTMLEditor(name) {
	if (typeof ew_DHTMLEditors != 'undefined' && typeof CKEDITOR != 'undefined') {
		var inst = CKEDITOR.instances[name];		
		if (inst)
			inst.setData(inst.element.value);
	}
}
// focus editor
function ew_FocusDHTMLEditor(name) {
	if (typeof ew_DHTMLEditors != 'undefined' && typeof CKEDITOR != 'undefined') {
		var inst = CKEDITOR.instances[name];	
		if (inst)
			inst.focus();
	}
}
//-->
</script>
<link rel="stylesheet" type="text/css" media="all" href="calendar/calendar-win2k-cold-1.css" title="win2k-1">
<link href='http://fonts.googleapis.com/css?family=Raleway' rel='stylesheet' type='text/css'>
<link href='http://fonts.googleapis.com/css?family=Julius+Sans+One' rel='stylesheet' type='text/css'>
<link href='http://fonts.googleapis.com/css?family=Dosis' rel='stylesheet' type='text/css'>
<script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/lang/calendar-en.js"></script>
<script type="text/javascript" src="calendar/calendar-setup.js"></script>
<script language="JavaScript" type="text/javascript">
<!--
// Write your client script here, no need to add script tags.
//-->
</script>
<% End If %>
<% If (Approved_Notices.Export = "") Or (EW_EXPORT_MASTER_RECORD And Approved_Notices.Export = "print") Then %>
<% End If %>
<% Approved_Notices_list.ShowPageHeader() %>
<%

' Load recordset
Set Approved_Notices_list.Recordset = Approved_Notices_list.LoadRecordset()
	Approved_Notices_list.TotalRecs = Approved_Notices_list.Recordset.RecordCount
	Approved_Notices_list.StartRec = 1
	If Approved_Notices_list.DisplayRecs <= 0 Then ' Display all records
		Approved_Notices_list.DisplayRecs = Approved_Notices_list.TotalRecs
	End If
	If Not (Approved_Notices.ExportAll And Approved_Notices.Export <> "") Then
		Approved_Notices_list.SetUpStartRec() ' Set up start record position
	End If
%>
<% Approved_Notices_list.ShowMessage %>
<div class="span6">
<marquee scrollamount="2" direction="up" height="620">
		  <div>
<% If Approved_Notices_list.TotalRecs > 0 Then %>
<%
If (Approved_Notices.ExportAll And Approved_Notices.Export <> "") Then
	Approved_Notices_list.StopRec = Approved_Notices_list.TotalRecs
Else

	' Set the last record to display
	If Approved_Notices_list.TotalRecs > Approved_Notices_list.StartRec + Approved_Notices_list.DisplayRecs - 1 Then
		Approved_Notices_list.StopRec = Approved_Notices_list.StartRec + Approved_Notices_list.DisplayRecs - 1
	Else
		Approved_Notices_list.StopRec = Approved_Notices_list.TotalRecs
	End If
End If

' Move to first record
Approved_Notices_list.RecCnt = Approved_Notices_list.StartRec - 1
If Not Approved_Notices_list.Recordset.Eof Then
	Approved_Notices_list.Recordset.MoveFirst
	If Approved_Notices_list.StartRec > 1 Then Approved_Notices_list.Recordset.Move Approved_Notices_list.StartRec - 1
ElseIf Not Approved_Notices.AllowAddDeleteRow And Approved_Notices_list.StopRec = 0 Then
	Approved_Notices_list.StopRec = Approved_Notices.GridAddRowCount
End If
Approved_Notices_list.RowCnt = 0
Dim strcount           
	strcount = 0

' Output date rows
Do While CLng(Approved_Notices_list.RecCnt) < CLng(Approved_Notices_list.StopRec)
	Approved_Notices_list.RecCnt = Approved_Notices_list.RecCnt + 1
	If CLng(Approved_Notices_list.RecCnt) >= CLng(Approved_Notices_list.StartRec) Then
		Approved_Notices_list.RowCnt = Approved_Notices_list.RowCnt + 1
		Approved_Notices_list.ColCnt = Approved_Notices_list.ColCnt + 1
		If Approved_Notices_list.ColCnt > Approved_Notices_list.RecPerRow Then Approved_Notices_list.ColCnt = 1

	' Set up key count
	Approved_Notices_list.KeyCount = Approved_Notices_list.RowIndex
	Call Approved_Notices.ResetAttrs()
	Approved_Notices.CssClass = ""
	If Approved_Notices.CurrentAction = "gridadd" Then
	Else
		Call Approved_Notices_list.LoadRowValues(Approved_Notices_list.Recordset) ' Load row values
	End If
	Approved_Notices.RowType = EW_ROWTYPE_VIEW ' Render view
	Approved_Notices.RowAttrs.AddAttributes Array()

	' Render row
	Call Approved_Notices_list.RenderRow()

	' Render list options
	Call Approved_Notices_list.RenderListOptions()
	
%>
		    	<div class="notecard">
				<h2 class="tshadow"><%= Approved_Notices.Title.ListViewValue %></h2>
				<!--<p><span class="label label-success">Start Date</span>&nbsp;<%= Approved_Notices.Sdate.ListViewValue %>&nbsp;&nbsp;&nbsp;
				<span class="label label-important">End Date</span>&nbsp;<%= Approved_Notices.Edate.ListViewValue %></p>
				<P><span class="label label-info"><%= Approved_Notices.Group.FldCaption %></span>&nbsp;<%= Approved_Notices.Group.ListViewValue %></P>-->
				<h3 class="noticecontent"><%= Approved_Notices.Notice.ListViewValue %>
				<span class="pull-right small muted noticeauthor"><%= Approved_Notices.Author.FldCaption %>:&nbsp;<%= Approved_Notices.Author.ListViewValue %></span>
				</h3>
				<br/>
				</div>
	<%
	End If
	If Approved_Notices.CurrentAction <> "gridadd" Then
		Approved_Notices_list.Recordset.MoveNext()
        strcount=strcount+1
		 
	End If
Loop
%>
</div>
</marquee>
</div>
<% End If %>
<%

' Close recordset and connection
Approved_Notices_list.Recordset.Close
Set Approved_Notices_list.Recordset = Nothing
%>

<% If Approved_Notices.Export = "" And Approved_Notices.CurrentAction = "" Then %>
<% End If %>
<%
Approved_Notices_list.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If Approved_Notices.Export = "" Then %>
<script type="text/javascript">
<!--
ew_CreateEditor();  // Create DHTML editor(s)
//-->
</script>
<script language="JavaScript" type="text/javascript">
<!--
// Write your table-specific startup script here
// document.write("page loaded");
//-->
</script>
<% End If %>
<!--#include file="footerless.asp"-->
<%

' Drop page object
Set Approved_Notices_list = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cApproved_Notices_list

	' Page ID
	Public Property Get PageID()
		PageID = "list"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Approved Notices"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Approved_Notices_list"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If Approved_Notices.UseTokenInUrl Then PageUrl = PageUrl & "t=" & Approved_Notices.TableVar & "&" ' add page token
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
		If Approved_Notices.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (Approved_Notices.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (Approved_Notices.TableVar = Request.QueryString("t"))
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
		If IsEmpty(Approved_Notices) Then Set Approved_Notices = New cApproved_Notices
		Set Table = Approved_Notices

		' Initialize urls
		ExportPrintUrl = PageUrl & "export=print"
		ExportExcelUrl = PageUrl & "export=excel"
		ExportWordUrl = PageUrl & "export=word"
		ExportHtmlUrl = PageUrl & "export=html"
		ExportXmlUrl = PageUrl & "export=xml"
		ExportCsvUrl = PageUrl & "export=csv"
		AddUrl = "approved_noticesadd.asp"
		InlineAddUrl = PageUrl & "a=add"
		GridAddUrl = PageUrl & "a=gridadd"
		GridEditUrl = PageUrl & "a=gridedit"
		MultiDeleteUrl = "approved_noticesdelete.asp"
		MultiUpdateUrl = "approved_noticesupdate.asp"

		' Initialize other table object
		If IsEmpty(Users) Then Set Users = New cUsers

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "list"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "Approved Notices"

		' Open connection to the database
		If IsEmpty(Conn) Then Call ew_Connect()

		' Initialize list options
		Set ListOptions = New cListOptions
		ListOptions.Tag = "span"
		ListOptions.Separator = "&nbsp;&nbsp;"

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

		' Table Permission loading event
		Call Security.TablePermission_Loading()
		Call Security.LoadCurrentUserLevel(TableName)

		' Table Permission loaded event
		Call Security.TablePermission_Loaded()

		' Get grid add count
		Dim gridaddcnt
		gridaddcnt = Request.QueryString(EW_TABLE_GRID_ADD_ROW_COUNT)
		If IsNumeric(gridaddcnt) Then
			If gridaddcnt > 0 Then
				Approved_Notices.GridAddRowCount = gridaddcnt
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
		Set Approved_Notices = Nothing
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

		' Multi Column
		RecPerRow = 1
		ColCnt = 0

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
			If Approved_Notices.Export <> "" Or Approved_Notices.CurrentAction = "gridadd" Or Approved_Notices.CurrentAction = "gridedit" Then
				ListOptions.HideAllOptions()
				ExportOptions.HideAllOptions()
			End If

			' Get basic search values
			Call LoadBasicSearchValues()

			' Restore search parms from Session
			Call RestoreSearchParms()

			' Call Recordset SearchValidated event
			Call Approved_Notices.Recordset_SearchValidated()

			' Set Up Sorting Order
			SetUpSortOrder()

			' Get basic search criteria
			If gsSearchError = "" Then
				sSrchBasic = BasicSearchWhere()
			End If
		End If ' End Validate Request

		' Restore display records
		If Approved_Notices.RecordsPerPage <> "" Then
			DisplayRecs = Approved_Notices.RecordsPerPage ' Restore from Session
		Else
			DisplayRecs = 50 ' Load default
		End If

		' Load Sorting Order
		LoadSortOrder()

		' Build search criteria
		Call ew_AddFilter(SearchWhere, sSrchAdvanced)
		Call ew_AddFilter(SearchWhere, sSrchBasic)

		' Call Recordset Searching event
		Call Approved_Notices.Recordset_Searching(SearchWhere)

		' Save search criteria
		If SearchWhere <> "" Then
			If sSrchBasic = "" Then Call ResetBasicSearchParms()
			Approved_Notices.SearchWhere = SearchWhere ' Save to Session
			If Not RestoreSearch Then
				StartRec = 1 ' Reset start record counter
				Approved_Notices.StartRecordNumber = StartRec
			End If
		Else
			SearchWhere = Approved_Notices.SearchWhere
		End If
		sFilter = ""
		Call ew_AddFilter(sFilter, DbDetailFilter)
		Call ew_AddFilter(sFilter, SearchWhere)

		' Set up filter in Session
		Approved_Notices.SessionWhere = sFilter
		Approved_Notices.CurrentFilter = ""
	End Sub

	' -----------------------------------------------------------------
	' Return Basic Search sql
	'
	Function BasicSearchSQL(Keyword)
		Dim sWhere
		sWhere = ""
			Call BuildBasicSearchSQL(sWhere, Approved_Notices.Title, Keyword)
			Call BuildBasicSearchSQL(sWhere, Approved_Notices.Author, Keyword)
			Call BuildBasicSearchSQL(sWhere, Approved_Notices.Group, Keyword)
			Call BuildBasicSearchSQL(sWhere, Approved_Notices.Notice, Keyword)
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
		sSearchKeyword = Approved_Notices.BasicSearchKeyword
		sSearchType = Approved_Notices.BasicSearchType
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
			Approved_Notices.SessionBasicSearchKeyword = sSearchKeyword
			Approved_Notices.SessionBasicSearchType = sSearchType
		End If
		BasicSearchWhere = sSearchStr
	End Function

	' -----------------------------------------------------------------
	' Clear all search parameters
	'
	Sub ResetSearchParms()

		' Clear search where
		SearchWhere = ""
		Approved_Notices.SearchWhere = SearchWhere

		' Clear basic search parameters
		Call ResetBasicSearchParms()
	End Sub

	' -----------------------------------------------------------------
	' Clear all basic search parameters
	'
	Sub ResetBasicSearchParms()

		' Clear basic search parameters
		Approved_Notices.SessionBasicSearchKeyword = ""
		Approved_Notices.SessionBasicSearchType = ""
	End Sub

	' -----------------------------------------------------------------
	' Restore all search parameters
	'
	Sub RestoreSearchParms()
		Dim bRestore
		bRestore = True
		If Approved_Notices.BasicSearchKeyword & "" <> "" Then bRestore = False
		RestoreSearch = bRestore
		If bRestore Then

			' Restore basic search values
			Approved_Notices.BasicSearchKeyword = Approved_Notices.SessionBasicSearchKeyword
			Approved_Notices.BasicSearchType = Approved_Notices.SessionBasicSearchType
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
			Approved_Notices.CurrentOrder = Request.QueryString("order")
			Approved_Notices.CurrentOrderType = Request.QueryString("ordertype")

			' Field Title
			Call Approved_Notices.UpdateSort(Approved_Notices.Title)

			' Field Author
			Call Approved_Notices.UpdateSort(Approved_Notices.Author)

			' Field Sdate
			Call Approved_Notices.UpdateSort(Approved_Notices.Sdate)

			' Field Edate
			Call Approved_Notices.UpdateSort(Approved_Notices.Edate)

			' Field Group
			Call Approved_Notices.UpdateSort(Approved_Notices.Group)

			' Field Notice
			Call Approved_Notices.UpdateSort(Approved_Notices.Notice)
			Approved_Notices.StartRecordNumber = 1 ' Reset start position
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load Sort Order parameters
	'
	Sub LoadSortOrder()
		Dim sOrderBy
		sOrderBy = Approved_Notices.SessionOrderBy ' Get order by from Session
		If sOrderBy = "" Then
			If Approved_Notices.SqlOrderBy <> "" Then
				sOrderBy = Approved_Notices.SqlOrderBy
				Approved_Notices.SessionOrderBy = sOrderBy
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
				Approved_Notices.SessionOrderBy = sOrderBy
				Approved_Notices.Title.Sort = ""
				Approved_Notices.Author.Sort = ""
				Approved_Notices.Sdate.Sort = ""
				Approved_Notices.Edate.Sort = ""
				Approved_Notices.Group.Sort = ""
				Approved_Notices.Notice.Sort = ""
			End If

			' Reset start position
			StartRec = 1
			Approved_Notices.StartRecordNumber = StartRec
		End If
	End Sub

	' Set up list options
	Sub SetupListOptions()
		Dim item
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
		If Security.CanEdit And ListOptions.GetItem("edit").Visible Then
			Set item = ListOptions.GetItem("edit")
			item.Body = "<a class=""ewRowLink"" href=""" & EditUrl & """>" & Language.Phrase("EditLink") & "</a>"
		End If
		If Security.CanDelete And ListOptions.GetItem("delete").Visible Then
			ListOptions.GetItem("delete").Body = "<a class=""ewRowLink""" & "" & " href=""" & DeleteUrl & """>" & Language.Phrase("DeleteLink") & "</a>"
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
				Approved_Notices.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					Approved_Notices.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = Approved_Notices.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			Approved_Notices.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			Approved_Notices.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			Approved_Notices.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	'  Load basic search values
	'
	Function LoadBasicSearchValues()
		Approved_Notices.BasicSearchKeyword = Request.QueryString(EW_TABLE_BASIC_SEARCH)
		Approved_Notices.BasicSearchType = Request.QueryString(EW_TABLE_BASIC_SEARCH_TYPE)
	End Function

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = Approved_Notices.CurrentFilter
		Call Approved_Notices.Recordset_Selecting(sFilter)
		Approved_Notices.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = Approved_Notices.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call Approved_Notices.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = Approved_Notices.KeyFilter

		' Call Row Selecting event
		Call Approved_Notices.Row_Selecting(sFilter)

		' Load sql based on filter
		Approved_Notices.CurrentFilter = sFilter
		sSql = Approved_Notices.SQL
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
		Call Approved_Notices.Row_Selected(RsRow)
		Approved_Notices.Notice_ID.DbValue = RsRow("Notice_ID")
		Approved_Notices.Title.DbValue = RsRow("Title")
		Approved_Notices.Author.DbValue = RsRow("Author")
		Approved_Notices.Sdate.DbValue = RsRow("Sdate")
		Approved_Notices.Edate.DbValue = RsRow("Edate")
		Approved_Notices.Group.DbValue = RsRow("Group")
		Approved_Notices.Notice.DbValue = RsRow("Notice")
		Approved_Notices.Approved.DbValue = ew_IIf(RsRow("Approved"), "1", "0")
		Approved_Notices.Archived.DbValue = ew_IIf(RsRow("Archived"), "1", "0")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If Approved_Notices.GetKey("Notice_ID")&"" <> "" Then
			Approved_Notices.Notice_ID.CurrentValue = Approved_Notices.GetKey("Notice_ID") ' Notice_ID
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			Approved_Notices.CurrentFilter = Approved_Notices.KeyFilter
			Dim sSql
			sSql = Approved_Notices.SQL
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
		ViewUrl = Approved_Notices.ViewUrl
		EditUrl = Approved_Notices.EditUrl("")
		InlineEditUrl = Approved_Notices.InlineEditUrl
		CopyUrl = Approved_Notices.CopyUrl("")
		InlineCopyUrl = Approved_Notices.InlineCopyUrl
		DeleteUrl = Approved_Notices.DeleteUrl

		' Call Row Rendering event
		Call Approved_Notices.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' Notice_ID
		' Title
		' Author
		' Sdate
		' Edate
		' Group
		' Notice
		' Approved
		' Archived
		' -----------
		'  View  Row
		' -----------

		If Approved_Notices.RowType = EW_ROWTYPE_VIEW Then ' View row

			' Notice_ID
			Approved_Notices.Notice_ID.ViewValue = Approved_Notices.Notice_ID.CurrentValue
			Approved_Notices.Notice_ID.ViewCustomAttributes = ""

			' Title
			Approved_Notices.Title.ViewValue = Approved_Notices.Title.CurrentValue
			Approved_Notices.Title.ViewCustomAttributes = ""

			' Author
			Approved_Notices.Author.ViewValue = Approved_Notices.Author.CurrentValue
			Approved_Notices.Author.ViewCustomAttributes = ""

			' Sdate
			Approved_Notices.Sdate.ViewValue = Approved_Notices.Sdate.CurrentValue
			Approved_Notices.Sdate.ViewValue = ew_FormatDateTime(Approved_Notices.Sdate.ViewValue, 7)
			Approved_Notices.Sdate.ViewCustomAttributes = ""

			' Edate
			Approved_Notices.Edate.ViewValue = Approved_Notices.Edate.CurrentValue
			Approved_Notices.Edate.ViewValue = ew_FormatDateTime(Approved_Notices.Edate.ViewValue, 7)
			Approved_Notices.Edate.ViewCustomAttributes = ""

			' Group
			If Approved_Notices.Group.CurrentValue & "" <> "" Then
				sFilterWrk = "[Group] = '" & ew_AdjustSql(Approved_Notices.Group.CurrentValue) & "'"
			sSqlWrk = "SELECT DISTINCT [Group] FROM [Groups]"
			sWhereWrk = ""
			Call ew_AddFilter(sWhereWrk, sFilterWrk)
			If sWhereWrk <> "" Then sSqlWrk = sSqlWrk & " WHERE " & sWhereWrk
			sSqlWrk = sSqlWrk & " ORDER BY [Group] Asc"
				Set RsWrk = Conn.Execute(sSqlWrk)
				If Not RsWrk.Eof Then
					Approved_Notices.Group.ViewValue = RsWrk("Group")
				Else
					Approved_Notices.Group.ViewValue = Approved_Notices.Group.CurrentValue
				End If
				RsWrk.Close
				Set RsWrk = Nothing
			Else
				Approved_Notices.Group.ViewValue = Null
			End If
			Approved_Notices.Group.ViewCustomAttributes = ""

			' Notice
			Approved_Notices.Notice.ViewValue = Approved_Notices.Notice.CurrentValue
			Approved_Notices.Notice.ViewCustomAttributes = ""

			' View refer script
			' Title

			Approved_Notices.Title.LinkCustomAttributes = ""
			Approved_Notices.Title.HrefValue = ""
			Approved_Notices.Title.TooltipValue = ""

			' Author
			Approved_Notices.Author.LinkCustomAttributes = ""
			Approved_Notices.Author.HrefValue = ""
			Approved_Notices.Author.TooltipValue = ""

			' Sdate
			Approved_Notices.Sdate.LinkCustomAttributes = ""
			Approved_Notices.Sdate.HrefValue = ""
			Approved_Notices.Sdate.TooltipValue = ""

			' Edate
			Approved_Notices.Edate.LinkCustomAttributes = ""
			Approved_Notices.Edate.HrefValue = ""
			Approved_Notices.Edate.TooltipValue = ""

			' Group
			Approved_Notices.Group.LinkCustomAttributes = ""
			Approved_Notices.Group.HrefValue = ""
			Approved_Notices.Group.TooltipValue = ""

			' Notice
			Approved_Notices.Notice.LinkCustomAttributes = ""
			Approved_Notices.Notice.HrefValue = ""
			Approved_Notices.Notice.TooltipValue = ""
		End If

		' Call Row Rendered event
		If Approved_Notices.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Approved_Notices.Row_Rendered()
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
