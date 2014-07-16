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
Dim UserLevels_view
Set UserLevels_view = New cUserLevels_view
Set Page = UserLevels_view

' Page init processing
Call UserLevels_view.Page_Init()

' Page main processing
Call UserLevels_view.Page_Main()
%>
<!--#include file="header.asp"-->
<% If UserLevels.Export = "" Then %>
<script type="text/javascript">
<!--
// Create page object
var UserLevels_view = new ew_Page("UserLevels_view");
// page properties
UserLevels_view.PageID = "view"; // page ID
UserLevels_view.FormID = "fUserLevelsview"; // form ID
var EW_PAGE_ID = UserLevels_view.PageID; // for backward compatibility
// extend page with Form_CustomValidate function
UserLevels_view.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
UserLevels_view.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
UserLevels_view.ValidateRequired = true; // uses JavaScript validation
<% Else %>
UserLevels_view.ValidateRequired = false; // no JavaScript validation
<% End If %>
//-->
</script>
<script language="JavaScript" type="text/javascript">
<!--
// Write your client script here, no need to add script tags.
//-->
</script>
<% End If %>
<% UserLevels_view.ShowPageHeader() %>
<p class="aspmaker ewTitle"><%= Language.Phrase("View") %>&nbsp;<%= Language.Phrase("TblTypeTABLE") %><%= UserLevels.TableCaption %>
&nbsp;&nbsp;<% UserLevels_view.ExportOptions.Render "body", "" %>
</p>
<% If UserLevels.Export = "" Then %>
<p class="aspmaker">
<a href="<%= UserLevels_view.ListUrl %>"><%= Language.Phrase("BackToList") %></a>&nbsp;
<% If Security.CanAdd Then %>
<a href="<%= UserLevels_view.AddUrl %>"><%= Language.Phrase("ViewPageAddLink") %></a>&nbsp;
<% End If %>
<% If Security.CanEdit Then %>
<a href="<%= UserLevels_view.EditUrl %>"><%= Language.Phrase("ViewPageEditLink") %></a>&nbsp;
<% End If %>
<% If Security.CanDelete Then %>
<a href="<%= UserLevels_view.DeleteUrl %>"><%= Language.Phrase("ViewPageDeleteLink") %></a>&nbsp;
<% End If %>
<% End If %>
</p>
<% UserLevels_view.ShowMessage %>
<p>
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table cellspacing="0" class="ewTable">
<% If UserLevels.UserLevelID.Visible Then ' UserLevelID %>
	<tr id="r_UserLevelID"<%= UserLevels.RowAttributes %>>
		<td class="ewTableHeader"><%= UserLevels.UserLevelID.FldCaption %></td>
		<td<%= UserLevels.UserLevelID.CellAttributes %>>
<div<%= UserLevels.UserLevelID.ViewAttributes %>><%= UserLevels.UserLevelID.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If UserLevels.UserLevelName.Visible Then ' UserLevelName %>
	<tr id="r_UserLevelName"<%= UserLevels.RowAttributes %>>
		<td class="ewTableHeader"><%= UserLevels.UserLevelName.FldCaption %></td>
		<td<%= UserLevels.UserLevelName.CellAttributes %>>
<div<%= UserLevels.UserLevelName.ViewAttributes %>><%= UserLevels.UserLevelName.ViewValue %></div>
</td>
	</tr>
<% End If %>
</table>
</div>
</td></tr></table>
<p>
<%
UserLevels_view.ShowPageFooter()
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
Set UserLevels_view = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cUserLevels_view

	' Page ID
	Public Property Get PageID()
		PageID = "view"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "UserLevels"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "UserLevels_view"
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
		Dim KeyUrl
		KeyUrl = ""
		If Request.QueryString("UserLevelID").Count > 0 Then
			ew_AddKey RecKey, "UserLevelID", Request.QueryString("UserLevelID")
			KeyUrl = KeyUrl & "&UserLevelID=" & Server.URLEncode(Request.QueryString("UserLevelID"))
		End If
		ExportPrintUrl = PageUrl & "export=print" & KeyUrl
		ExportHtmlUrl = PageUrl & "export=html" & KeyUrl
		ExportExcelUrl = PageUrl & "export=excel" & KeyUrl
		ExportWordUrl = PageUrl & "export=word" & KeyUrl
		ExportXmlUrl = PageUrl & "export=xml" & KeyUrl
		ExportCsvUrl = PageUrl & "export=csv" & KeyUrl

		' Initialize other table object
		If IsEmpty(Users) Then Set Users = New cUsers

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "view"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "UserLevels"

		' Open connection to the database
		If IsEmpty(Conn) Then Call ew_Connect()

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
	Dim RecCnt
	Dim RecKey
	Dim ExportOptions ' Export options
	Dim Recordset

	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()
		Dim sReturnUrl
		sReturnUrl = ""
		Dim bMatchRecord
		bMatchRecord = False
		If IsPageRequest Then ' Validate request
			If Request.QueryString("UserLevelID").Count > 0 Then
				UserLevels.UserLevelID.QueryStringValue = Request.QueryString("UserLevelID")
			Else
				sReturnUrl = "userlevelslist.asp" ' Return to list
			End If

			' Get action
			UserLevels.CurrentAction = "I" ' Display form
			Select Case UserLevels.CurrentAction
				Case "I" ' Get a record to display
					If Not LoadRow() Then ' Load record based on key
						If SuccessMessage = "" And FailureMessage = "" Then
							FailureMessage = Language.Phrase("NoRecord") ' Set no record message
						End If
						sReturnUrl = "userlevelslist.asp" ' No matching record, return to list
					End If
			End Select
		Else
			sReturnUrl = "userlevelslist.asp" ' Not page request, return to list
		End If
		If sReturnUrl <> "" Then Call Page_Terminate(sReturnUrl)

		' Render row
		UserLevels.RowType = EW_ROWTYPE_VIEW
		Call UserLevels.ResetAttrs()
		Call RenderRow()
	End Sub
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

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		AddUrl = UserLevels.AddUrl
		EditUrl = UserLevels.EditUrl("")
		CopyUrl = UserLevels.CopyUrl("")
		DeleteUrl = UserLevels.DeleteUrl
		ListUrl = UserLevels.ListUrl

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
End Class
%>
