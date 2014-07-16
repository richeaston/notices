<%@ CodePage="1252" LCID="2057" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="ewcfg9.asp"-->
<!--#include file="audittrailinfo.asp"-->
<!--#include file="usersinfo.asp"-->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim AuditTrail_view
Set AuditTrail_view = New cAuditTrail_view
Set Page = AuditTrail_view

' Page init processing
Call AuditTrail_view.Page_Init()

' Page main processing
Call AuditTrail_view.Page_Main()
%>
<!--#include file="header.asp"-->
<% If AuditTrail.Export = "" Then %>
<script type="text/javascript">
<!--
// Create page object
var AuditTrail_view = new ew_Page("AuditTrail_view");
// page properties
AuditTrail_view.PageID = "view"; // page ID
AuditTrail_view.FormID = "fAuditTrailview"; // form ID
var EW_PAGE_ID = AuditTrail_view.PageID; // for backward compatibility
// extend page with Form_CustomValidate function
AuditTrail_view.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
AuditTrail_view.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
AuditTrail_view.ValidateRequired = true; // uses JavaScript validation
<% Else %>
AuditTrail_view.ValidateRequired = false; // no JavaScript validation
<% End If %>
//-->
</script>
<script language="JavaScript" type="text/javascript">
<!--
// Write your client script here, no need to add script tags.
//-->
</script>
<% End If %>
<% AuditTrail_view.ShowPageHeader() %>
<p class="aspmaker ewTitle"><%= Language.Phrase("View") %>&nbsp;<%= Language.Phrase("TblTypeTABLE") %><%= AuditTrail.TableCaption %>
&nbsp;&nbsp;<% AuditTrail_view.ExportOptions.Render "body", "" %>
</p>
<% If AuditTrail.Export = "" Then %>
<p class="aspmaker">
<a href="<%= AuditTrail_view.ListUrl %>"><%= Language.Phrase("BackToList") %></a>&nbsp;
<% If Security.CanDelete Then %>
<a href="<%= AuditTrail_view.DeleteUrl %>"><%= Language.Phrase("ViewPageDeleteLink") %></a>&nbsp;
<% End If %>
<% End If %>
</p>
<% AuditTrail_view.ShowMessage %>
<p>
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table cellspacing="0" class="ewTable">
<% If AuditTrail.DateTime.Visible Then ' DateTime %>
	<tr id="r_DateTime"<%= AuditTrail.RowAttributes %>>
		<td class="ewTableHeader"><%= AuditTrail.DateTime.FldCaption %></td>
		<td<%= AuditTrail.DateTime.CellAttributes %>>
<div<%= AuditTrail.DateTime.ViewAttributes %>><%= AuditTrail.DateTime.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If AuditTrail.Script.Visible Then ' Script %>
	<tr id="r_Script"<%= AuditTrail.RowAttributes %>>
		<td class="ewTableHeader"><%= AuditTrail.Script.FldCaption %></td>
		<td<%= AuditTrail.Script.CellAttributes %>>
<div<%= AuditTrail.Script.ViewAttributes %>><%= AuditTrail.Script.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If AuditTrail.User.Visible Then ' User %>
	<tr id="r_User"<%= AuditTrail.RowAttributes %>>
		<td class="ewTableHeader"><%= AuditTrail.User.FldCaption %></td>
		<td<%= AuditTrail.User.CellAttributes %>>
<div<%= AuditTrail.User.ViewAttributes %>><%= AuditTrail.User.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If AuditTrail.Action.Visible Then ' Action %>
	<tr id="r_Action"<%= AuditTrail.RowAttributes %>>
		<td class="ewTableHeader"><%= AuditTrail.Action.FldCaption %></td>
		<td<%= AuditTrail.Action.CellAttributes %>>
<div<%= AuditTrail.Action.ViewAttributes %>><%= AuditTrail.Action.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If AuditTrail.zTable.Visible Then ' Table %>
	<tr id="r_zTable"<%= AuditTrail.RowAttributes %>>
		<td class="ewTableHeader"><%= AuditTrail.zTable.FldCaption %></td>
		<td<%= AuditTrail.zTable.CellAttributes %>>
<div<%= AuditTrail.zTable.ViewAttributes %>><%= AuditTrail.zTable.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If AuditTrail.zField.Visible Then ' Field %>
	<tr id="r_zField"<%= AuditTrail.RowAttributes %>>
		<td class="ewTableHeader"><%= AuditTrail.zField.FldCaption %></td>
		<td<%= AuditTrail.zField.CellAttributes %>>
<div<%= AuditTrail.zField.ViewAttributes %>><%= AuditTrail.zField.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If AuditTrail.KeyValue.Visible Then ' KeyValue %>
	<tr id="r_KeyValue"<%= AuditTrail.RowAttributes %>>
		<td class="ewTableHeader"><%= AuditTrail.KeyValue.FldCaption %></td>
		<td<%= AuditTrail.KeyValue.CellAttributes %>>
<div<%= AuditTrail.KeyValue.ViewAttributes %>><%= AuditTrail.KeyValue.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If AuditTrail.OldValue.Visible Then ' OldValue %>
	<tr id="r_OldValue"<%= AuditTrail.RowAttributes %>>
		<td class="ewTableHeader"><%= AuditTrail.OldValue.FldCaption %></td>
		<td<%= AuditTrail.OldValue.CellAttributes %>>
<div<%= AuditTrail.OldValue.ViewAttributes %>><%= AuditTrail.OldValue.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If AuditTrail.NewValue.Visible Then ' NewValue %>
	<tr id="r_NewValue"<%= AuditTrail.RowAttributes %>>
		<td class="ewTableHeader"><%= AuditTrail.NewValue.FldCaption %></td>
		<td<%= AuditTrail.NewValue.CellAttributes %>>
<div<%= AuditTrail.NewValue.ViewAttributes %>><%= AuditTrail.NewValue.ViewValue %></div>
</td>
	</tr>
<% End If %>
</table>
</div>
</td></tr></table>
<p>
<%
AuditTrail_view.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If AuditTrail.Export = "" Then %>
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
Set AuditTrail_view = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cAuditTrail_view

	' Page ID
	Public Property Get PageID()
		PageID = "view"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "AuditTrail"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "AuditTrail_view"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If AuditTrail.UseTokenInUrl Then PageUrl = PageUrl & "t=" & AuditTrail.TableVar & "&" ' add page token
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
		If AuditTrail.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (AuditTrail.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (AuditTrail.TableVar = Request.QueryString("t"))
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
		If IsEmpty(AuditTrail) Then Set AuditTrail = New cAuditTrail
		Set Table = AuditTrail

		' Initialize urls
		Dim KeyUrl
		KeyUrl = ""
		If Request.QueryString("Id").Count > 0 Then
			ew_AddKey RecKey, "Id", Request.QueryString("Id")
			KeyUrl = KeyUrl & "&Id=" & Server.URLEncode(Request.QueryString("Id"))
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
		EW_TABLE_NAME = "AuditTrail"

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
		If Not Security.IsLoggedIn() Then
			Call Security.SaveLastUrl()
			Call Page_Terminate("login.asp")
		End If
		If Not Security.CanView Then
			Call Security.SaveLastUrl()
			Call Page_Terminate("audittraillist.asp")
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
		Set AuditTrail = Nothing
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
			If Request.QueryString("Id").Count > 0 Then
				AuditTrail.Id.QueryStringValue = Request.QueryString("Id")
			Else
				sReturnUrl = "audittraillist.asp" ' Return to list
			End If

			' Get action
			AuditTrail.CurrentAction = "I" ' Display form
			Select Case AuditTrail.CurrentAction
				Case "I" ' Get a record to display
					If Not LoadRow() Then ' Load record based on key
						If SuccessMessage = "" And FailureMessage = "" Then
							FailureMessage = Language.Phrase("NoRecord") ' Set no record message
						End If
						sReturnUrl = "audittraillist.asp" ' No matching record, return to list
					End If
			End Select
		Else
			sReturnUrl = "audittraillist.asp" ' Not page request, return to list
		End If
		If sReturnUrl <> "" Then Call Page_Terminate(sReturnUrl)

		' Render row
		AuditTrail.RowType = EW_ROWTYPE_VIEW
		Call AuditTrail.ResetAttrs()
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
				AuditTrail.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					AuditTrail.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = AuditTrail.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			AuditTrail.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			AuditTrail.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			AuditTrail.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = AuditTrail.KeyFilter

		' Call Row Selecting event
		Call AuditTrail.Row_Selecting(sFilter)

		' Load sql based on filter
		AuditTrail.CurrentFilter = sFilter
		sSql = AuditTrail.SQL
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
		Call AuditTrail.Row_Selected(RsRow)
		AuditTrail.Id.DbValue = RsRow("Id")
		AuditTrail.DateTime.DbValue = RsRow("DateTime")
		AuditTrail.Script.DbValue = RsRow("Script")
		AuditTrail.User.DbValue = RsRow("User")
		AuditTrail.Action.DbValue = RsRow("Action")
		AuditTrail.zTable.DbValue = RsRow("Table")
		AuditTrail.zField.DbValue = RsRow("Field")
		AuditTrail.KeyValue.DbValue = RsRow("KeyValue")
		AuditTrail.OldValue.DbValue = RsRow("OldValue")
		AuditTrail.NewValue.DbValue = RsRow("NewValue")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		AddUrl = AuditTrail.AddUrl
		EditUrl = AuditTrail.EditUrl("")
		CopyUrl = AuditTrail.CopyUrl("")
		DeleteUrl = AuditTrail.DeleteUrl
		ListUrl = AuditTrail.ListUrl

		' Call Row Rendering event
		Call AuditTrail.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' Id
		' DateTime
		' Script
		' User
		' Action
		' Table
		' Field
		' KeyValue
		' OldValue
		' NewValue
		' -----------
		'  View  Row
		' -----------

		If AuditTrail.RowType = EW_ROWTYPE_VIEW Then ' View row

			' Id
			AuditTrail.Id.ViewValue = AuditTrail.Id.CurrentValue
			AuditTrail.Id.ViewCustomAttributes = ""

			' DateTime
			AuditTrail.DateTime.ViewValue = AuditTrail.DateTime.CurrentValue
			AuditTrail.DateTime.ViewValue = ew_FormatDateTime(AuditTrail.DateTime.ViewValue, 7)
			AuditTrail.DateTime.ViewCustomAttributes = ""

			' Script
			AuditTrail.Script.ViewValue = AuditTrail.Script.CurrentValue
			AuditTrail.Script.ViewCustomAttributes = ""

			' User
			AuditTrail.User.ViewValue = AuditTrail.User.CurrentValue
			AuditTrail.User.ViewCustomAttributes = ""

			' Action
			AuditTrail.Action.ViewValue = AuditTrail.Action.CurrentValue
			AuditTrail.Action.ViewCustomAttributes = ""

			' Table
			AuditTrail.zTable.ViewValue = AuditTrail.zTable.CurrentValue
			AuditTrail.zTable.ViewCustomAttributes = ""

			' Field
			AuditTrail.zField.ViewValue = AuditTrail.zField.CurrentValue
			AuditTrail.zField.ViewCustomAttributes = ""

			' KeyValue
			AuditTrail.KeyValue.ViewValue = AuditTrail.KeyValue.CurrentValue
			AuditTrail.KeyValue.ViewCustomAttributes = ""

			' OldValue
			AuditTrail.OldValue.ViewValue = AuditTrail.OldValue.CurrentValue
			AuditTrail.OldValue.ViewCustomAttributes = ""

			' NewValue
			AuditTrail.NewValue.ViewValue = AuditTrail.NewValue.CurrentValue
			AuditTrail.NewValue.ViewCustomAttributes = ""

			' View refer script
			' DateTime

			AuditTrail.DateTime.LinkCustomAttributes = ""
			AuditTrail.DateTime.HrefValue = ""
			AuditTrail.DateTime.TooltipValue = ""

			' Script
			AuditTrail.Script.LinkCustomAttributes = ""
			AuditTrail.Script.HrefValue = ""
			AuditTrail.Script.TooltipValue = ""

			' User
			AuditTrail.User.LinkCustomAttributes = ""
			AuditTrail.User.HrefValue = ""
			AuditTrail.User.TooltipValue = ""

			' Action
			AuditTrail.Action.LinkCustomAttributes = ""
			AuditTrail.Action.HrefValue = ""
			AuditTrail.Action.TooltipValue = ""

			' Table
			AuditTrail.zTable.LinkCustomAttributes = ""
			AuditTrail.zTable.HrefValue = ""
			AuditTrail.zTable.TooltipValue = ""

			' Field
			AuditTrail.zField.LinkCustomAttributes = ""
			AuditTrail.zField.HrefValue = ""
			AuditTrail.zField.TooltipValue = ""

			' KeyValue
			AuditTrail.KeyValue.LinkCustomAttributes = ""
			AuditTrail.KeyValue.HrefValue = ""
			AuditTrail.KeyValue.TooltipValue = ""

			' OldValue
			AuditTrail.OldValue.LinkCustomAttributes = ""
			AuditTrail.OldValue.HrefValue = ""
			AuditTrail.OldValue.TooltipValue = ""

			' NewValue
			AuditTrail.NewValue.LinkCustomAttributes = ""
			AuditTrail.NewValue.HrefValue = ""
			AuditTrail.NewValue.TooltipValue = ""
		End If

		' Call Row Rendered event
		If AuditTrail.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call AuditTrail.Row_Rendered()
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
