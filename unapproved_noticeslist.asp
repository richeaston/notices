<%@ CodePage="1252" LCID="2057" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="ewcfg9.asp"-->
<!--#include file="unapproved_noticesinfo.asp"-->
<!--#include file="usersinfo.asp"-->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim Unapproved_Notices_list
Set Unapproved_Notices_list = New cUnapproved_Notices_list
Set Page = Unapproved_Notices_list

' Page init processing
Call Unapproved_Notices_list.Page_Init()

' Page main processing
Call Unapproved_Notices_list.Page_Main()
%>
<!--#include file="header.asp"-->
<% If Unapproved_Notices.Export = "" Then %>
<script type="text/javascript">
<!--
// Create page object
var Unapproved_Notices_list = new ew_Page("Unapproved_Notices_list");
// page properties
Unapproved_Notices_list.PageID = "list"; // page ID
Unapproved_Notices_list.FormID = "fUnapproved_Noticeslist"; // form ID
var EW_PAGE_ID = Unapproved_Notices_list.PageID; // for backward compatibility
// extend page with Form_CustomValidate function
Unapproved_Notices_list.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
Unapproved_Notices_list.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
Unapproved_Notices_list.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Unapproved_Notices_list.ValidateRequired = false; // no JavaScript validation
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
<script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/lang/calendar-en.js"></script>
<script type="text/javascript" src="calendar/calendar-setup.js"></script>
<script language="JavaScript" type="text/javascript">
<!--
// Write your client script here, no need to add script tags.
//-->
</script>
<% End If %>
<% If (Unapproved_Notices.Export = "") Or (EW_EXPORT_MASTER_RECORD And Unapproved_Notices.Export = "print") Then %>
<% End If %>
<% Unapproved_Notices_list.ShowPageHeader() %>
<%

' Load recordset
Set Unapproved_Notices_list.Recordset = Unapproved_Notices_list.LoadRecordset()
	Unapproved_Notices_list.TotalRecs = Unapproved_Notices_list.Recordset.RecordCount
	Unapproved_Notices_list.StartRec = 1
	If Unapproved_Notices_list.DisplayRecs <= 0 Then ' Display all records
		Unapproved_Notices_list.DisplayRecs = Unapproved_Notices_list.TotalRecs
	End If
	If Not (Unapproved_Notices.ExportAll And Unapproved_Notices.Export <> "") Then
		Unapproved_Notices_list.SetUpStartRec() ' Set up start record position
	End If
%>
<% If Security.CanSearch Then %>
<% If Unapproved_Notices.Export = "" And Unapproved_Notices.CurrentAction = "" Then %>
<a href="javascript:ew_ToggleSearchPanel(Unapproved_Notices_list);" style="text-decoration: none;"><img id="Unapproved_Notices_list_SearchImage" src="images/collapse.gif" alt="" width="9" height="9" border="0"></a><span class="aspmaker">&nbsp;<%= Language.Phrase("Search") %></span><br>
<div id="Unapproved_Notices_list_SearchPanel" class="input-prepend input-append">
<form name="fUnapproved_Noticeslistsrch" id="fUnapproved_Noticeslistsrch" class="form-search" action="<%= ew_CurrentPage %>">
<input type="hidden" id="t" name="t" value="Unapproved_Notices">
<div class="ewBasicSearch">
<div id="xsr_1" class="input-append">
	<input type="text" name="<%= EW_TABLE_BASIC_SEARCH %>" id="<%= EW_TABLE_BASIC_SEARCH %>" size="20" value="<%= ew_HtmlEncode(unapproved_notices.SessionBasicSearchKeyword) %>">
	<button class="btn" type="Submit" name="Submit" id="Submit"><%= ew_BtnCaption(Language.Phrase("QuickSearchBtn")) %></button>
	<a class="btn" href="<%= unapproved_notices_list.PageUrl %>cmd=reset"><i class="icon-refresh"></i>&nbsp;<%= Language.Phrase("ShowAll") %></a>
</div>

</div>
</form>
</div>
<% End If %>
<% End If %>
<% Unapproved_Notices_list.ShowMessage %>
<br>
<table><tr><td>
<form name="fUnapproved_Noticeslist" id="fUnapproved_Noticeslist" class="ewForm" action="" method="post">
<input type="hidden" name="t" id="t" value="Unapproved_Notices">
<div id="gmp_Unapproved_Notices">
<% If Unapproved_Notices_list.TotalRecs > 0 Then %>
<table cellspacing="0" class="table table-bordered table-striped table-condensed">
<%= Unapproved_Notices.TableCustomInnerHTML %>
<thead><!-- Table header -->
	
<%
Call Unapproved_Notices_list.RenderListOptions()

' Render list options (header, left)
Unapproved_Notices_list.ListOptions.Render "header", "left"
%>
<% If Unapproved_Notices.Title.Visible Then ' Title %>
		<th><%= Unapproved_Notices.Title.FldCaption %></th>
	
<% End If %>		
<% If Unapproved_Notices.Author.Visible Then ' Author %>
		<th><%= Unapproved_Notices.Author.FldCaption %></th>
<% End If %>		
<% If Unapproved_Notices.Sdate.Visible Then ' Sdate %>
		<th>Start</th>
<% End If %>		
<% If Unapproved_Notices.Edate.Visible Then ' Edate %>
		<th>End</th>
<% End If %>		
<% If Unapproved_Notices.Group.Visible Then ' Group %>
		<th><%= Unapproved_Notices.Group.FldCaption %></th>
<% End If %>		
<% If Unapproved_Notices.Notice.Visible Then ' Notice %>
		<th><%= Unapproved_Notices.Notice.FldCaption %></th>
<% End If %>		
<%

' Render list options (header, right)
Unapproved_Notices_list.ListOptions.Render "header", "right"
%>
</thead>
<tbody><!-- Table body -->
<%
If (Unapproved_Notices.ExportAll And Unapproved_Notices.Export <> "") Then
	Unapproved_Notices_list.StopRec = Unapproved_Notices_list.TotalRecs
Else

	' Set the last record to display
	If Unapproved_Notices_list.TotalRecs > Unapproved_Notices_list.StartRec + Unapproved_Notices_list.DisplayRecs - 1 Then
		Unapproved_Notices_list.StopRec = Unapproved_Notices_list.StartRec + Unapproved_Notices_list.DisplayRecs - 1
	Else
		Unapproved_Notices_list.StopRec = Unapproved_Notices_list.TotalRecs
	End If
End If

' Move to first record
Unapproved_Notices_list.RecCnt = Unapproved_Notices_list.StartRec - 1
If Not Unapproved_Notices_list.Recordset.Eof Then
	Unapproved_Notices_list.Recordset.MoveFirst
	If Unapproved_Notices_list.StartRec > 1 Then Unapproved_Notices_list.Recordset.Move Unapproved_Notices_list.StartRec - 1
ElseIf Not Unapproved_Notices.AllowAddDeleteRow And Unapproved_Notices_list.StopRec = 0 Then
	Unapproved_Notices_list.StopRec = Unapproved_Notices.GridAddRowCount
End If

' Initialize Aggregate
Unapproved_Notices.RowType = EW_ROWTYPE_AGGREGATEINIT
Call Unapproved_Notices.ResetAttrs()
Call Unapproved_Notices_list.RenderRow()
Unapproved_Notices_list.RowCnt = 0

' Output date rows
Do While CLng(Unapproved_Notices_list.RecCnt) < CLng(Unapproved_Notices_list.StopRec)
	Unapproved_Notices_list.RecCnt = Unapproved_Notices_list.RecCnt + 1
	If CLng(Unapproved_Notices_list.RecCnt) >= CLng(Unapproved_Notices_list.StartRec) Then
		Unapproved_Notices_list.RowCnt = Unapproved_Notices_list.RowCnt + 1

	' Set up key count
	Unapproved_Notices_list.KeyCount = Unapproved_Notices_list.RowIndex
	Call Unapproved_Notices.ResetAttrs()
	Unapproved_Notices.CssClass = ""
	If Unapproved_Notices.CurrentAction = "gridadd" Then
	Else
		Call Unapproved_Notices_list.LoadRowValues(Unapproved_Notices_list.Recordset) ' Load row values
	End If
	Unapproved_Notices.RowType = EW_ROWTYPE_VIEW ' Render view
	Unapproved_Notices.RowAttrs.AddAttributes Array(Array("onmouseover", "ew_MouseOver(event, this);"), Array("onmouseout", "ew_MouseOut(event, this);"), Array("onclick", "ew_Click(event, this);"))

	' Render row
	Call Unapproved_Notices_list.RenderRow()

	' Render list options
	Call Unapproved_Notices_list.RenderListOptions()
%>
	<tr<%= Unapproved_Notices.RowAttributes %>>
<%

' Render list options (body, left)
Unapproved_Notices_list.ListOptions.Render "body", "left"
%>
	<% If Unapproved_Notices.Title.Visible Then ' Title %>
		<td<%= Unapproved_Notices.Title.CellAttributes %>>
<div<%= Unapproved_Notices.Title.ViewAttributes %>><%= Unapproved_Notices.Title.ListViewValue %></div>
<a name="<%= Unapproved_Notices_list.PageObjName & "_row_" & Unapproved_Notices_list.RowCnt %>" id="<%= Unapproved_Notices_list.PageObjName & "_row_" & Unapproved_Notices_list.RowCnt %>"></a></td>
	<% End If %>
	<% If Unapproved_Notices.Author.Visible Then ' Author %>
		<td<%= Unapproved_Notices.Author.CellAttributes %>>
<div<%= Unapproved_Notices.Author.ViewAttributes %>><%= Unapproved_Notices.Author.ListViewValue %></div>
</td>
	<% End If %>
	<% If Unapproved_Notices.Sdate.Visible Then ' Sdate %>
		<td<%= Unapproved_Notices.Sdate.CellAttributes %>>
<div<%= Unapproved_Notices.Sdate.ViewAttributes %>><%= Unapproved_Notices.Sdate.ListViewValue %></div>
</td>
	<% End If %>
	<% If Unapproved_Notices.Edate.Visible Then ' Edate %>
		<td<%= Unapproved_Notices.Edate.CellAttributes %>>
<div<%= Unapproved_Notices.Edate.ViewAttributes %>><%= Unapproved_Notices.Edate.ListViewValue %></div>
</td>
	<% End If %>
	<% If Unapproved_Notices.Group.Visible Then ' Group %>
		<td<%= Unapproved_Notices.Group.CellAttributes %>>
<div<%= Unapproved_Notices.Group.ViewAttributes %>><%= Unapproved_Notices.Group.ListViewValue %></div>
</td>
	<% End If %>
	<% If Unapproved_Notices.Notice.Visible Then ' Notice %>
		<td<%= Unapproved_Notices.Notice.CellAttributes %>>
<div<%= Unapproved_Notices.Notice.ViewAttributes %>><%= Unapproved_Notices.Notice.ListViewValue %></div>
</td>
	<% End If %>
<%

' Render list options (body, right)
Unapproved_Notices_list.ListOptions.Render "body", "right"
%>
	</tr>
<%
	End If
	If Unapproved_Notices.CurrentAction <> "gridadd" Then
		Unapproved_Notices_list.Recordset.MoveNext()
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
Unapproved_Notices_list.Recordset.Close
Set Unapproved_Notices_list.Recordset = Nothing
%>
<% If Unapproved_Notices.Export = "" Then %>
<div class="ewGridLowerPanel">
<% If Unapproved_Notices.CurrentAction <> "gridadd" And Unapproved_Notices.CurrentAction <> "gridedit" Then %>
<form name="ewpagerform" id="ewpagerform" class="form-inline" action="<%= ew_CurrentPage %>">
<table border="0" cellspacing="0" cellpadding="0" >
	<tr>
		<td>
<% If Not IsObject(Unapproved_Notices_list.Pager) Then Set Unapproved_Notices_list.Pager = ew_NewPrevNextPager(Unapproved_Notices_list.StartRec, Unapproved_Notices_list.DisplayRecs, Unapproved_Notices_list.TotalRecs) %>
<% If Unapproved_Notices_list.Pager.RecordCount > 0 Then %>
	<table border="0" cellspacing="0" cellpadding="0"><tr><td><%= Language.Phrase("Page") %>&nbsp;</td>
<!--first page button-->
	<% If Unapproved_Notices_list.Pager.FirstButton.Enabled Then %>
	<td><a href="<%= Unapproved_Notices_list.PageUrl %>start=<%= Unapproved_Notices_list.Pager.FirstButton.Start %>"><img src="images/first.gif" alt="<%= Language.Phrase("PagerFirst") %>" width="16" height="16" border="0"></a></td>
	<% Else %>
	<td><img src="images/firstdisab.gif" alt="<%= Language.Phrase("PagerFirst") %>" width="16" height="16" border="0"></td>
	<% End If %>
<!--previous page button-->
	<% If Unapproved_Notices_list.Pager.PrevButton.Enabled Then %>
	<td><a href="<%= Unapproved_Notices_list.PageUrl %>start=<%= Unapproved_Notices_list.Pager.PrevButton.Start %>"><img src="images/prev.gif" alt="<%= Language.Phrase("PagerPrevious") %>" width="16" height="16" border="0"></a></td>
	<% Else %>
	<td><img src="images/prevdisab.gif" alt="<%= Language.Phrase("PagerPrevious") %>" width="16" height="16" border="0"></td>
	<% End If %>
<!--current page number-->
	<td><input type="text" name="<%= EW_TABLE_PAGE_NO %>" id="<%= EW_TABLE_PAGE_NO %>" value="<%= Unapproved_Notices_list.Pager.CurrentPage %>" size="4"></td>
<!--next page button-->
	<% If Unapproved_Notices_list.Pager.NextButton.Enabled Then %>
	<td><a href="<%= Unapproved_Notices_list.PageUrl %>start=<%= Unapproved_Notices_list.Pager.NextButton.Start %>"><img src="images/next.gif" alt="<%= Language.Phrase("PagerNext") %>" width="16" height="16" border="0"></a></td>
	<% Else %>
	<td><img src="images/nextdisab.gif" alt="<%= Language.Phrase("PagerNext") %>" width="16" height="16" border="0"></td>
	<% End If %>
<!--last page button-->
	<% If Unapproved_Notices_list.Pager.LastButton.Enabled Then %>
	<td><a href="<%= Unapproved_Notices_list.PageUrl %>start=<%= Unapproved_Notices_list.Pager.LastButton.Start %>"><img src="images/last.gif" alt="<%= Language.Phrase("PagerLast") %>" width="16" height="16" border="0"></a></td>	
	<% Else %>
	<td><img src="images/lastdisab.gif" alt="<%= Language.Phrase("PagerLast") %>" width="16" height="16" border="0"></td>
	<% End If %>
	<td>&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= Unapproved_Notices_list.Pager.PageCount %></span></td>
	</tr></table>
	</td>	
	<td>&nbsp;&nbsp;&nbsp;&nbsp;</td>
	<td>
	<span><%= Language.Phrase("Record") %>&nbsp;<%= Unapproved_Notices_list.Pager.FromIndex %>&nbsp;<%= Language.Phrase("To") %>&nbsp;<%= Unapproved_Notices_list.Pager.ToIndex %>&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= Unapproved_Notices_list.Pager.RecordCount %></span>
<% Else %>
	<% If Security.CanList Then %>
	<% If Unapproved_Notices_list.SearchWhere = "0=101" Then %>
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
<% If Unapproved_Notices_list.TotalRecs > 0 Then %>
<% If Security.CanEdit Then %>
<a class="btn btn-primary" href="" onclick="ew_SubmitSelected(document.fUnapproved_Noticeslist, '<%= Unapproved_Notices_list.MultiUpdateUrl %>');return false;"><i class="icon-check icon-white"></i>&nbsp;Approve Selected Notices</a>&nbsp;&nbsp;
<% End If %>
<% End If %>
</span>
</div>
<% End If %>
</td></tr></table>
<% If Unapproved_Notices.Export = "" And Unapproved_Notices.CurrentAction = "" Then %>
<% End If %>
<%
Unapproved_Notices_list.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If Unapproved_Notices.Export = "" Then %>
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
<!--#include file="footer.asp"-->
<%

' Drop page object
Set Unapproved_Notices_list = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cUnapproved_Notices_list

	' Page ID
	Public Property Get PageID()
		PageID = "list"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Unapproved Notices"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Unapproved_Notices_list"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If Unapproved_Notices.UseTokenInUrl Then PageUrl = PageUrl & "t=" & Unapproved_Notices.TableVar & "&" ' add page token
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
		If sMessage <> "" Then Response.Write "<p class=""ewMessage alert alert-info"">" & sMessage & "</p>"
		Session(EW_SESSION_MESSAGE) = "" ' Clear message in Session

		' Success message
		Dim sSuccessMessage
		sSuccessMessage = SuccessMessage
		Call Message_Showing(sSuccessMessage, "success")
		If sSuccessMessage <> "" Then Response.Write "<p class=""ewSuccessMessage alert alert-success"">" & sSuccessMessage & "</p>"
		Session(EW_SESSION_SUCCESS_MESSAGE) = "" ' Clear message in Session

		' Failure message
		Dim sErrorMessage
		sErrorMessage = FailureMessage
		Call Message_Showing(sErrorMessage, "failure")
		If sErrorMessage <> "" Then Response.Write "<p class=""ewErrorMessage alert alert-important"">" & sErrorMessage & "</p>"
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
		If Unapproved_Notices.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (Unapproved_Notices.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (Unapproved_Notices.TableVar = Request.QueryString("t"))
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
		If IsEmpty(Unapproved_Notices) Then Set Unapproved_Notices = New cUnapproved_Notices
		Set Table = Unapproved_Notices

		' Initialize urls
		ExportPrintUrl = PageUrl & "export=print"
		ExportExcelUrl = PageUrl & "export=excel"
		ExportWordUrl = PageUrl & "export=word"
		ExportHtmlUrl = PageUrl & "export=html"
		ExportXmlUrl = PageUrl & "export=xml"
		ExportCsvUrl = PageUrl & "export=csv"
		AddUrl = "unapproved_noticesadd.asp"
		InlineAddUrl = PageUrl & "a=add"
		GridAddUrl = PageUrl & "a=gridadd"
		GridEditUrl = PageUrl & "a=gridedit"
		MultiDeleteUrl = "approved_noticesdelete.asp"
		MultiUpdateUrl = "unapproved_noticesupdate.asp"

		' Initialize other table object
		If IsEmpty(Users) Then Set Users = New cUsers

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "list"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "Unapproved Notices"

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
				Unapproved_Notices.GridAddRowCount = gridaddcnt
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
		Set Unapproved_Notices = Nothing
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
			If Unapproved_Notices.Export <> "" Or Unapproved_Notices.CurrentAction = "gridadd" Or Unapproved_Notices.CurrentAction = "gridedit" Then
				ListOptions.HideAllOptions()
				ExportOptions.HideAllOptions()
			End If

			' Get basic search values
			Call LoadBasicSearchValues()

			' Restore search parms from Session
			Call RestoreSearchParms()

			' Call Recordset SearchValidated event
			Call Unapproved_Notices.Recordset_SearchValidated()

			' Set Up Sorting Order
			SetUpSortOrder()

			' Get basic search criteria
			If gsSearchError = "" Then
				sSrchBasic = BasicSearchWhere()
			End If
		End If ' End Validate Request

		' Restore display records
		If Unapproved_Notices.RecordsPerPage <> "" Then
			DisplayRecs = Unapproved_Notices.RecordsPerPage ' Restore from Session
		Else
			DisplayRecs = 50 ' Load default
		End If

		' Load Sorting Order
		LoadSortOrder()

		' Build search criteria
		Call ew_AddFilter(SearchWhere, sSrchAdvanced)
		Call ew_AddFilter(SearchWhere, sSrchBasic)

		' Call Recordset Searching event
		Call Unapproved_Notices.Recordset_Searching(SearchWhere)

		' Save search criteria
		If SearchWhere <> "" Then
			If sSrchBasic = "" Then Call ResetBasicSearchParms()
			Unapproved_Notices.SearchWhere = SearchWhere ' Save to Session
			If Not RestoreSearch Then
				StartRec = 1 ' Reset start record counter
				Unapproved_Notices.StartRecordNumber = StartRec
			End If
		Else
			SearchWhere = Unapproved_Notices.SearchWhere
		End If
		sFilter = ""
		If Not Security.CanList Then
			sFilter = "(0=1)" ' Filter all records
		End If
		Call ew_AddFilter(sFilter, DbDetailFilter)
		Call ew_AddFilter(sFilter, SearchWhere)

		' Set up filter in Session
		Unapproved_Notices.SessionWhere = sFilter
		Unapproved_Notices.CurrentFilter = ""
	End Sub

	' -----------------------------------------------------------------
	' Return Basic Search sql
	'
	Function BasicSearchSQL(Keyword)
		Dim sWhere
		sWhere = ""
			Call BuildBasicSearchSQL(sWhere, Unapproved_Notices.Title, Keyword)
			Call BuildBasicSearchSQL(sWhere, Unapproved_Notices.Author, Keyword)
			Call BuildBasicSearchSQL(sWhere, Unapproved_Notices.Group, Keyword)
			Call BuildBasicSearchSQL(sWhere, Unapproved_Notices.Notice, Keyword)
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
		sSearchKeyword = Unapproved_Notices.BasicSearchKeyword
		sSearchType = Unapproved_Notices.BasicSearchType
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
			Unapproved_Notices.SessionBasicSearchKeyword = sSearchKeyword
			Unapproved_Notices.SessionBasicSearchType = sSearchType
		End If
		BasicSearchWhere = sSearchStr
	End Function

	' -----------------------------------------------------------------
	' Clear all search parameters
	'
	Sub ResetSearchParms()

		' Clear search where
		SearchWhere = ""
		Unapproved_Notices.SearchWhere = SearchWhere

		' Clear basic search parameters
		Call ResetBasicSearchParms()
	End Sub

	' -----------------------------------------------------------------
	' Clear all basic search parameters
	'
	Sub ResetBasicSearchParms()

		' Clear basic search parameters
		Unapproved_Notices.SessionBasicSearchKeyword = ""
		Unapproved_Notices.SessionBasicSearchType = ""
	End Sub

	' -----------------------------------------------------------------
	' Restore all search parameters
	'
	Sub RestoreSearchParms()
		Dim bRestore
		bRestore = True
		If Unapproved_Notices.BasicSearchKeyword & "" <> "" Then bRestore = False
		RestoreSearch = bRestore
		If bRestore Then

			' Restore basic search values
			Unapproved_Notices.BasicSearchKeyword = Unapproved_Notices.SessionBasicSearchKeyword
			Unapproved_Notices.BasicSearchType = Unapproved_Notices.SessionBasicSearchType
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
			Unapproved_Notices.CurrentOrder = Request.QueryString("order")
			Unapproved_Notices.CurrentOrderType = Request.QueryString("ordertype")

			' Field Title
			Call Unapproved_Notices.UpdateSort(Unapproved_Notices.Title)

			' Field Author
			Call Unapproved_Notices.UpdateSort(Unapproved_Notices.Author)

			' Field Sdate
			Call Unapproved_Notices.UpdateSort(Unapproved_Notices.Sdate)

			' Field Edate
			Call Unapproved_Notices.UpdateSort(Unapproved_Notices.Edate)

			' Field Group
			Call Unapproved_Notices.UpdateSort(Unapproved_Notices.Group)

			' Field Notice
			Call Unapproved_Notices.UpdateSort(Unapproved_Notices.Notice)
			Unapproved_Notices.StartRecordNumber = 1 ' Reset start position
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load Sort Order parameters
	'
	Sub LoadSortOrder()
		Dim sOrderBy
		sOrderBy = Unapproved_Notices.SessionOrderBy ' Get order by from Session
		If sOrderBy = "" Then
			If Unapproved_Notices.SqlOrderBy <> "" Then
				sOrderBy = Unapproved_Notices.SqlOrderBy
				Unapproved_Notices.SessionOrderBy = sOrderBy
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
				Unapproved_Notices.SessionOrderBy = sOrderBy
				Unapproved_Notices.Title.Sort = ""
				Unapproved_Notices.Author.Sort = ""
				Unapproved_Notices.Sdate.Sort = ""
				Unapproved_Notices.Edate.Sort = ""
				Unapproved_Notices.Group.Sort = ""
				Unapproved_Notices.Notice.Sort = ""
			End If

			' Reset start position
			StartRec = 1
			Unapproved_Notices.StartRecordNumber = StartRec
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
		ListOptions.Add("checkbox")
		ListOptions.GetItem("checkbox").CssStyle = "white-space: nowrap;"
		ListOptions.GetItem("checkbox").Visible = Security.CanEdit
		ListOptions.GetItem("checkbox").OnLeft = False
		ListOptions.GetItem("checkbox").Header = "<input type=""checkbox"" name=""key"" id=""key"" class=""aspmaker"" onclick=""Unapproved_Notices_list.SelectAllKey(this);"">"
		Call ListOptions_Load()
	End Sub

	' Render list options
	Sub RenderListOptions()
		Dim item, links
		ListOptions.LoadDefault()
		If Security.CanEdit And ListOptions.GetItem("edit").Visible Then
			Set item = ListOptions.GetItem("edit")
			item.Body = "<a class=""ewRowLink"" href=""" & EditUrl & """><i class='icon-pencil'></i></a>"
		End If
		<!--
		<!--If Security.CanDelete And ListOptions.GetItem("delete").Visible Then
			<!--ListOptions.GetItem("delete").Body = "<a class=""ewRowLink""" & "" & " href=""" & DeleteUrl & """><i class='icon-remove-circle'></i></a>"
		<!--End If
		-->
		If Security.CanEdit And ListOptions.GetItem("checkbox").Visible Then
			ListOptions.GetItem("checkbox").Body = "<input type=""checkbox"" name=""key_m"" id=""key_m"" value=""" & ew_HtmlEncode(Unapproved_Notices.Notice_ID.CurrentValue) & """ class=""aspmaker"" onclick='ew_ClickMultiCheckbox(this);'>"
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
				Unapproved_Notices.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					Unapproved_Notices.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = Unapproved_Notices.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			Unapproved_Notices.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			Unapproved_Notices.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			Unapproved_Notices.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	'  Load basic search values
	'
	Function LoadBasicSearchValues()
		Unapproved_Notices.BasicSearchKeyword = Request.QueryString(EW_TABLE_BASIC_SEARCH)
		Unapproved_Notices.BasicSearchType = Request.QueryString(EW_TABLE_BASIC_SEARCH_TYPE)
	End Function

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = Unapproved_Notices.CurrentFilter
		Call Unapproved_Notices.Recordset_Selecting(sFilter)
		Unapproved_Notices.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = Unapproved_Notices.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call Unapproved_Notices.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = Unapproved_Notices.KeyFilter

		' Call Row Selecting event
		Call Unapproved_Notices.Row_Selecting(sFilter)

		' Load sql based on filter
		Unapproved_Notices.CurrentFilter = sFilter
		sSql = Unapproved_Notices.SQL
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
		Call Unapproved_Notices.Row_Selected(RsRow)
		Unapproved_Notices.Notice_ID.DbValue = RsRow("Notice_ID")
		Unapproved_Notices.Title.DbValue = RsRow("Title")
		Unapproved_Notices.Author.DbValue = RsRow("Author")
		Unapproved_Notices.Sdate.DbValue = RsRow("Sdate")
		Unapproved_Notices.Edate.DbValue = RsRow("Edate")
		Unapproved_Notices.Group.DbValue = RsRow("Group")
		Unapproved_Notices.Notice.DbValue = RsRow("Notice")
		Unapproved_Notices.Approved.DbValue = ew_IIf(RsRow("Approved"), "1", "0")
		Unapproved_Notices.Archived.DbValue = ew_IIf(RsRow("Archived"), "1", "0")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If Unapproved_Notices.GetKey("Notice_ID")&"" <> "" Then
			Unapproved_Notices.Notice_ID.CurrentValue = Unapproved_Notices.GetKey("Notice_ID") ' Notice_ID
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			Unapproved_Notices.CurrentFilter = Unapproved_Notices.KeyFilter
			Dim sSql
			sSql = Unapproved_Notices.SQL
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
		ViewUrl = Unapproved_Notices.ViewUrl
		EditUrl = Unapproved_Notices.EditUrl("")
		InlineEditUrl = Unapproved_Notices.InlineEditUrl
		CopyUrl = Unapproved_Notices.CopyUrl("")
		InlineCopyUrl = Unapproved_Notices.InlineCopyUrl
		DeleteUrl = Unapproved_Notices.DeleteUrl

		' Call Row Rendering event
		Call Unapproved_Notices.Row_Rendering()

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

		If Unapproved_Notices.RowType = EW_ROWTYPE_VIEW Then ' View row

			' Title
			Unapproved_Notices.Title.ViewValue = Unapproved_Notices.Title.CurrentValue
			Unapproved_Notices.Title.ViewCustomAttributes = ""

			' Author
			Unapproved_Notices.Author.ViewValue = Unapproved_Notices.Author.CurrentValue
			Unapproved_Notices.Author.ViewCustomAttributes = ""

			' Sdate
			Unapproved_Notices.Sdate.ViewValue = Unapproved_Notices.Sdate.CurrentValue
			Unapproved_Notices.Sdate.ViewValue = ew_FormatDateTime(Unapproved_Notices.Sdate.ViewValue, 7)
			Unapproved_Notices.Sdate.ViewCustomAttributes = ""

			' Edate
			Unapproved_Notices.Edate.ViewValue = Unapproved_Notices.Edate.CurrentValue
			Unapproved_Notices.Edate.ViewValue = ew_FormatDateTime(Unapproved_Notices.Edate.ViewValue, 7)
			Unapproved_Notices.Edate.ViewCustomAttributes = ""

			' Group
			If Unapproved_Notices.Group.CurrentValue & "" <> "" Then
				sFilterWrk = "[Group] = '" & ew_AdjustSql(Unapproved_Notices.Group.CurrentValue) & "'"
			sSqlWrk = "SELECT DISTINCT [Group] FROM [Groups]"
			sWhereWrk = ""
			Call ew_AddFilter(sWhereWrk, sFilterWrk)
			If sWhereWrk <> "" Then sSqlWrk = sSqlWrk & " WHERE " & sWhereWrk
			sSqlWrk = sSqlWrk & " ORDER BY [Group] Asc"
				Set RsWrk = Conn.Execute(sSqlWrk)
				If Not RsWrk.Eof Then
					Unapproved_Notices.Group.ViewValue = RsWrk("Group")
				Else
					Unapproved_Notices.Group.ViewValue = Unapproved_Notices.Group.CurrentValue
				End If
				RsWrk.Close
				Set RsWrk = Nothing
			Else
				Unapproved_Notices.Group.ViewValue = Null
			End If
			Unapproved_Notices.Group.ViewCustomAttributes = ""

			' Notice
			Unapproved_Notices.Notice.ViewValue = Unapproved_Notices.Notice.CurrentValue
			Unapproved_Notices.Notice.ViewCustomAttributes = ""

			' View refer script
			' Title

			Unapproved_Notices.Title.LinkCustomAttributes = ""
			Unapproved_Notices.Title.HrefValue = ""
			Unapproved_Notices.Title.TooltipValue = ""

			' Author
			Unapproved_Notices.Author.LinkCustomAttributes = ""
			Unapproved_Notices.Author.HrefValue = ""
			Unapproved_Notices.Author.TooltipValue = ""

			' Sdate
			Unapproved_Notices.Sdate.LinkCustomAttributes = ""
			Unapproved_Notices.Sdate.HrefValue = ""
			Unapproved_Notices.Sdate.TooltipValue = ""

			' Edate
			Unapproved_Notices.Edate.LinkCustomAttributes = ""
			Unapproved_Notices.Edate.HrefValue = ""
			Unapproved_Notices.Edate.TooltipValue = ""

			' Group
			Unapproved_Notices.Group.LinkCustomAttributes = ""
			Unapproved_Notices.Group.HrefValue = ""
			Unapproved_Notices.Group.TooltipValue = ""

			' Notice
			Unapproved_Notices.Notice.LinkCustomAttributes = ""
			Unapproved_Notices.Notice.HrefValue = ""
			Unapproved_Notices.Notice.TooltipValue = ""
		End If

		' Call Row Rendered event
		If Unapproved_Notices.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Unapproved_Notices.Row_Rendered()
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
