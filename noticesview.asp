<%@ CodePage="1252" LCID="2057" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="ewcfg9.asp"-->
<!--#include file="noticesinfo.asp"-->
<!--#include file="usersinfo.asp"-->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim Notices_view
Set Notices_view = New cNotices_view
Set Page = Notices_view

' Page init processing
Call Notices_view.Page_Init()

' Page main processing
Call Notices_view.Page_Main()
%>
<!--#include file="header.asp"-->
<% If Notices.Export = "" Then %>
<script type="text/javascript">
<!--
// Create page object
var Notices_view = new ew_Page("Notices_view");
// page properties
Notices_view.PageID = "view"; // page ID
Notices_view.FormID = "fNoticesview"; // form ID
var EW_PAGE_ID = Notices_view.PageID; // for backward compatibility
// extend page with Form_CustomValidate function
Notices_view.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
<% If EW_CLIENT_VALIDATE Then %>
Notices_view.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Notices_view.ValidateRequired = false; // no JavaScript validation
<% End If %>
//-->
</script>
<script language="JavaScript" type="text/javascript">
<!--
// Write your client script here, no need to add script tags.
//-->
</script>
<% End If %>
<% Notices_view.ShowPageHeader() %>

</p>
<% If Notices.Export = "" Then %>
<p class="aspmaker">
<a class="btn btn-primary" href="<%= Notices_view.ListUrl %>"><i class="icon-arrow-left icon-white"></i>&nbsp;<%= Language.Phrase("BackToList") %></a>&nbsp;
<% If Security.CanAdd Then %>
<a class="btn btn-success" href="<%= Notices_view.AddUrl %>"><i class="icon-plus icon-white"></i>&nbsp;<%= Language.Phrase("ViewPageAddLink") %></a>&nbsp;
<% End If %>
<% If Security.CanEdit Then %>
<a class="btn btn-info" href="<%= Notices_view.EditUrl %>"><i class="icon-pencil icon-white"></i>&nbsp;<%= Language.Phrase("ViewPageEditLink") %></a>&nbsp;
<% End If %>
<% If Security.CanDelete Then %>
<a class="btn btn-danger" href="<%= Notices_view.DeleteUrl %>"><i class="icon-remove-circle icon-white"></i>&nbsp;<%= Language.Phrase("ViewPageDeleteLink") %></a>&nbsp;
<% End If %>
<% End If %>
</p>
<% Notices_view.ShowMessage %>
<p>
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel well">
<table cellspacing="0" class="ewTable">
<% If Notices.Title.Visible Then ' Title %>
	<tr id="r_Title"<%= Notices.RowAttributes %>>
		<td class="ewTableHeader"><B><%= Notices.Title.FldCaption %></B></td>
		<td<%= Notices.Title.CellAttributes %>>
<div<%= Notices.Title.ViewAttributes %>><%= Notices.Title.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Notices.Author.Visible Then ' Author %>
	<tr id="r_Author"<%= Notices.RowAttributes %>>
		<td class="ewTableHeader"><B><%= Notices.Author.FldCaption %></b></td>
		<td<%= Notices.Author.CellAttributes %>>
<div<%= Notices.Author.ViewAttributes %>><%= Notices.Author.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Notices.Sdate.Visible Then ' Sdate %>
	<tr id="r_Sdate"<%= Notices.RowAttributes %>>
		<td class="ewTableHeader"><B>Start Date</b></td>
		<td<%= Notices.Sdate.CellAttributes %>>
<div<%= Notices.Sdate.ViewAttributes %>><%= Notices.Sdate.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Notices.Edate.Visible Then ' Edate %>
	<tr id="r_Edate"<%= Notices.RowAttributes %>>
		<td class="ewTableHeader"><B>End Date</b></td>
		<td<%= Notices.Edate.CellAttributes %>>
<div<%= Notices.Edate.ViewAttributes %>><%= Notices.Edate.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Notices.Group.Visible Then ' Group %>
	<tr id="r_Group"<%= Notices.RowAttributes %>>
		<td class="ewTableHeader"><B><%= Notices.Group.FldCaption %></b></td>
		<td<%= Notices.Group.CellAttributes %>>
<div<%= Notices.Group.ViewAttributes %>><%= Notices.Group.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Notices.Notice.Visible Then ' Notice %>
	<tr id="r_Notice"<%= Notices.RowAttributes %>>
		<td class="ewTableHeader"><B><%= Notices.Notice.FldCaption %></b></td>
		<td<%= Notices.Notice.CellAttributes %>>
<div<%= Notices.Notice.ViewAttributes %>><%= Notices.Notice.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Notices.Approved.Visible Then ' Approved %>
	<tr id="r_Approved"<%= Notices.RowAttributes %>>
		<td class="ewTableHeader"><B><%= Notices.Approved.FldCaption %></b></td>
		<td<%= Notices.Approved.CellAttributes %>>
<% If ew_ConvertToBool(Notices.Approved.CurrentValue) Then %>
<input type="checkbox" value="<%= Notices.Approved.ViewValue %>" checked onclick="this.form.reset();" disabled="disabled">
<% Else %>
<input type="checkbox" value="<%= Notices.Approved.ViewValue %>" onclick="this.form.reset();" disabled="disabled">
<% End If %></td>
	</tr>
<% End If %>

</table>
</div>
</td></tr></table>
<p>
<%
Notices_view.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If Notices.Export = "" Then %>
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
Set Notices_view = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cNotices_view

	' Page ID
	Public Property Get PageID()
		PageID = "view"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Notices"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Notices_view"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If Notices.UseTokenInUrl Then PageUrl = PageUrl & "t=" & Notices.TableVar & "&" ' add page token
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
		If Notices.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (Notices.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (Notices.TableVar = Request.QueryString("t"))
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
		If IsEmpty(Notices) Then Set Notices = New cNotices
		Set Table = Notices

		' Initialize urls
		Dim KeyUrl
		KeyUrl = ""
		If Request.QueryString("Notice_ID").Count > 0 Then
			ew_AddKey RecKey, "Notice_ID", Request.QueryString("Notice_ID")
			KeyUrl = KeyUrl & "&Notice_ID=" & Server.URLEncode(Request.QueryString("Notice_ID"))
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
		EW_TABLE_NAME = "Notices"

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
			Call Page_Terminate("noticeslist.asp")
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
		Set Notices = Nothing
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
			If Request.QueryString("Notice_ID").Count > 0 Then
				Notices.Notice_ID.QueryStringValue = Request.QueryString("Notice_ID")
			Else
				sReturnUrl = "noticeslist.asp" ' Return to list
			End If

			' Get action
			Notices.CurrentAction = "I" ' Display form
			Select Case Notices.CurrentAction
				Case "I" ' Get a record to display
					If Not LoadRow() Then ' Load record based on key
						If SuccessMessage = "" And FailureMessage = "" Then
							FailureMessage = Language.Phrase("NoRecord") ' Set no record message
						End If
						sReturnUrl = "noticeslist.asp" ' No matching record, return to list
					End If
			End Select
		Else
			sReturnUrl = "noticeslist.asp" ' Not page request, return to list
		End If
		If sReturnUrl <> "" Then Call Page_Terminate(sReturnUrl)

		' Render row
		Notices.RowType = EW_ROWTYPE_VIEW
		Call Notices.ResetAttrs()
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
				Notices.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					Notices.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = Notices.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			Notices.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			Notices.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			Notices.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = Notices.KeyFilter

		' Call Row Selecting event
		Call Notices.Row_Selecting(sFilter)

		' Load sql based on filter
		Notices.CurrentFilter = sFilter
		sSql = Notices.SQL
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
		Call Notices.Row_Selected(RsRow)
		Notices.Notice_ID.DbValue = RsRow("Notice_ID")
		Notices.Title.DbValue = RsRow("Title")
		Notices.Author.DbValue = RsRow("Author")
		Notices.Sdate.DbValue = RsRow("Sdate")
		Notices.Edate.DbValue = RsRow("Edate")
		Notices.Group.DbValue = RsRow("Group")
		Notices.Notice.DbValue = RsRow("Notice")
		Notices.Approved.DbValue = ew_IIf(RsRow("Approved"), "1", "0")
		Notices.Archived.DbValue = ew_IIf(RsRow("Archived"), "1", "0")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		AddUrl = Notices.AddUrl
		EditUrl = Notices.EditUrl("")
		CopyUrl = Notices.CopyUrl("")
		DeleteUrl = Notices.DeleteUrl
		ListUrl = Notices.ListUrl

		' Call Row Rendering event
		Call Notices.Row_Rendering()

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

		If Notices.RowType = EW_ROWTYPE_VIEW Then ' View row

			' Title
			Notices.Title.ViewValue = Notices.Title.CurrentValue
			Notices.Title.ViewCustomAttributes = ""

			' Author
			Notices.Author.ViewValue = Notices.Author.CurrentValue
			Notices.Author.ViewCustomAttributes = ""

			' Sdate
			Notices.Sdate.ViewValue = Notices.Sdate.CurrentValue
			Notices.Sdate.ViewValue = ew_FormatDateTime(Notices.Sdate.ViewValue, 7)
			Notices.Sdate.ViewCustomAttributes = ""

			' Edate
			Notices.Edate.ViewValue = Notices.Edate.CurrentValue
			Notices.Edate.ViewValue = ew_FormatDateTime(Notices.Edate.ViewValue, 7)
			Notices.Edate.ViewCustomAttributes = ""

			' Group
			If Notices.Group.CurrentValue & "" <> "" Then
				sFilterWrk = "[Group] = '" & ew_AdjustSql(Notices.Group.CurrentValue) & "'"
			sSqlWrk = "SELECT DISTINCT [Group] FROM [Groups]"
			sWhereWrk = ""
			Call ew_AddFilter(sWhereWrk, sFilterWrk)
			If sWhereWrk <> "" Then sSqlWrk = sSqlWrk & " WHERE " & sWhereWrk
			sSqlWrk = sSqlWrk & " ORDER BY [Group] Asc"
				Set RsWrk = Conn.Execute(sSqlWrk)
				If Not RsWrk.Eof Then
					Notices.Group.ViewValue = RsWrk("Group")
				Else
					Notices.Group.ViewValue = Notices.Group.CurrentValue
				End If
				RsWrk.Close
				Set RsWrk = Nothing
			Else
				Notices.Group.ViewValue = Null
			End If
			Notices.Group.ViewCustomAttributes = ""

			' Notice
			Notices.Notice.ViewValue = Notices.Notice.CurrentValue
			Notices.Notice.ViewCustomAttributes = ""

			' Approved
			If ew_ConvertToBool(Notices.Approved.CurrentValue) Then
				Notices.Approved.ViewValue = ew_IIf(Notices.Approved.FldTagCaption(1) <> "", Notices.Approved.FldTagCaption(1), "Yes")
			Else
				Notices.Approved.ViewValue = ew_IIf(Notices.Approved.FldTagCaption(2) <> "", Notices.Approved.FldTagCaption(2), "No")
			End If
			Notices.Approved.ViewCustomAttributes = ""

			' Archived
			If ew_ConvertToBool(Notices.Archived.CurrentValue) Then
				Notices.Archived.ViewValue = ew_IIf(Notices.Archived.FldTagCaption(1) <> "", Notices.Archived.FldTagCaption(1), "Yes")
			Else
				Notices.Archived.ViewValue = ew_IIf(Notices.Archived.FldTagCaption(2) <> "", Notices.Archived.FldTagCaption(2), "No")
			End If
			Notices.Archived.ViewCustomAttributes = ""

			' View refer script
			' Title

			Notices.Title.LinkCustomAttributes = ""
			Notices.Title.HrefValue = ""
			Notices.Title.TooltipValue = ""

			' Author
			Notices.Author.LinkCustomAttributes = ""
			Notices.Author.HrefValue = ""
			Notices.Author.TooltipValue = ""

			' Sdate
			Notices.Sdate.LinkCustomAttributes = ""
			Notices.Sdate.HrefValue = ""
			Notices.Sdate.TooltipValue = ""

			' Edate
			Notices.Edate.LinkCustomAttributes = ""
			Notices.Edate.HrefValue = ""
			Notices.Edate.TooltipValue = ""

			' Group
			Notices.Group.LinkCustomAttributes = ""
			Notices.Group.HrefValue = ""
			Notices.Group.TooltipValue = ""

			' Notice
			Notices.Notice.LinkCustomAttributes = ""
			Notices.Notice.HrefValue = ""
			Notices.Notice.TooltipValue = ""

			' Approved
			Notices.Approved.LinkCustomAttributes = ""
			Notices.Approved.HrefValue = ""
			Notices.Approved.TooltipValue = ""

			' Archived
			Notices.Archived.LinkCustomAttributes = ""
			Notices.Archived.HrefValue = ""
			Notices.Archived.TooltipValue = ""
		End If

		' Call Row Rendered event
		If Notices.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Notices.Row_Rendered()
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
