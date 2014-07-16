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
Dim Users_view
Set Users_view = New cUsers_view
Set Page = Users_view

' Page init processing
Call Users_view.Page_Init()

' Page main processing
Call Users_view.Page_Main()
%>
<!--#include file="header.asp"-->
<% If Users.Export = "" Then %>
<script type="text/javascript">
<!--
// Create page object
var Users_view = new ew_Page("Users_view");
// page properties
Users_view.PageID = "view"; // page ID
Users_view.FormID = "fUsersview"; // form ID
var EW_PAGE_ID = Users_view.PageID; // for backward compatibility
// extend page with Form_CustomValidate function
Users_view.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
Users_view.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
Users_view.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Users_view.ValidateRequired = false; // no JavaScript validation
<% End If %>
//-->
</script>
<script language="JavaScript" type="text/javascript">
<!--
// Write your client script here, no need to add script tags.
//-->
</script>
<% End If %>
<% Users_view.ShowPageHeader() %>
<p class="aspmaker ewTitle"><%= Language.Phrase("View") %>&nbsp;<%= Language.Phrase("TblTypeTABLE") %><%= Users.TableCaption %>
&nbsp;&nbsp;<% Users_view.ExportOptions.Render "body", "" %>
</p>
<% If Users.Export = "" Then %>
<p class="aspmaker">
<a href="<%= Users_view.ListUrl %>"><%= Language.Phrase("BackToList") %></a>&nbsp;
<% If Security.CanAdd Then %>
<a href="<%= Users_view.AddUrl %>"><%= Language.Phrase("ViewPageAddLink") %></a>&nbsp;
<% End If %>
<% If Security.CanEdit Then %>
<a href="<%= Users_view.EditUrl %>"><%= Language.Phrase("ViewPageEditLink") %></a>&nbsp;
<% End If %>
<% If Security.CanDelete Then %>
<a href="<%= Users_view.DeleteUrl %>"><%= Language.Phrase("ViewPageDeleteLink") %></a>&nbsp;
<% End If %>
<% End If %>
</p>
<% Users_view.ShowMessage %>
<p>
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table cellspacing="0" class="ewTable">
<% If Users.Username.Visible Then ' Username %>
	<tr id="r_Username"<%= Users.RowAttributes %>>
		<td class="ewTableHeader"><%= Users.Username.FldCaption %></td>
		<td<%= Users.Username.CellAttributes %>>
<div<%= Users.Username.ViewAttributes %>><%= Users.Username.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Users.zEmail.Visible Then ' Email %>
	<tr id="r_zEmail"<%= Users.RowAttributes %>>
		<td class="ewTableHeader"><%= Users.zEmail.FldCaption %></td>
		<td<%= Users.zEmail.CellAttributes %>>
<div<%= Users.zEmail.ViewAttributes %>>
<% If Users.zEmail.LinkAttributes <> "" Then %>
<a<%= Users.zEmail.LinkAttributes %>><%= Users.zEmail.ViewValue %></a>
<% Else %>
<%= Users.zEmail.ViewValue %>
<% End If %>
</div>
</td>
	</tr>
<% End If %>
<% If Users.Permissions.Visible Then ' Permissions %>
	<tr id="r_Permissions"<%= Users.RowAttributes %>>
		<td class="ewTableHeader"><%= Users.Permissions.FldCaption %></td>
		<td<%= Users.Permissions.CellAttributes %>>
<div<%= Users.Permissions.ViewAttributes %>><%= Users.Permissions.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Users.Active.Visible Then ' Active %>
	<tr id="r_Active"<%= Users.RowAttributes %>>
		<td class="ewTableHeader"><%= Users.Active.FldCaption %></td>
		<td<%= Users.Active.CellAttributes %>>
<% If ew_ConvertToBool(Users.Active.CurrentValue) Then %>
<input type="checkbox" value="<%= Users.Active.ViewValue %>" checked onclick="this.form.reset();" disabled="disabled">
<% Else %>
<input type="checkbox" value="<%= Users.Active.ViewValue %>" onclick="this.form.reset();" disabled="disabled">
<% End If %></td>
	</tr>
<% End If %>
<% If Users.Profile.Visible Then ' Profile %>
	<tr id="r_Profile"<%= Users.RowAttributes %>>
		<td class="ewTableHeader"><%= Users.Profile.FldCaption %></td>
		<td<%= Users.Profile.CellAttributes %>>
<div<%= Users.Profile.ViewAttributes %>><%= Users.Profile.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Users.Theme.Visible Then ' Theme %>
	<tr id="r_Theme"<%= Users.RowAttributes %>>
		<td class="ewTableHeader"><%= Users.Theme.FldCaption %></td>
		<td<%= Users.Theme.CellAttributes %>>
<div<%= Users.Theme.ViewAttributes %>><%= Users.Theme.ViewValue %></div>
</td>
	</tr>
<% End If %>
</table>
</div>
</td></tr></table>
<p>
<%
Users_view.ShowPageFooter()
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
Set Users_view = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cUsers_view

	' Page ID
	Public Property Get PageID()
		PageID = "view"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Users"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Users_view"
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
		Dim KeyUrl
		KeyUrl = ""
		If Request.QueryString("Username").Count > 0 Then
			ew_AddKey RecKey, "Username", Request.QueryString("Username")
			KeyUrl = KeyUrl & "&Username=" & Server.URLEncode(Request.QueryString("Username"))
		End If
		ExportPrintUrl = PageUrl & "export=print" & KeyUrl
		ExportHtmlUrl = PageUrl & "export=html" & KeyUrl
		ExportExcelUrl = PageUrl & "export=excel" & KeyUrl
		ExportWordUrl = PageUrl & "export=word" & KeyUrl
		ExportXmlUrl = PageUrl & "export=xml" & KeyUrl
		ExportCsvUrl = PageUrl & "export=csv" & KeyUrl

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "view"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "Users"

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
			Call Page_Terminate("userslist.asp")
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
		Set Users = Nothing
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
			If Request.QueryString("Username").Count > 0 Then
				Users.Username.QueryStringValue = Request.QueryString("Username")
			Else
				sReturnUrl = "userslist.asp" ' Return to list
			End If

			' Get action
			Users.CurrentAction = "I" ' Display form
			Select Case Users.CurrentAction
				Case "I" ' Get a record to display
					If Not LoadRow() Then ' Load record based on key
						If SuccessMessage = "" And FailureMessage = "" Then
							FailureMessage = Language.Phrase("NoRecord") ' Set no record message
						End If
						sReturnUrl = "userslist.asp" ' No matching record, return to list
					End If
			End Select
		Else
			sReturnUrl = "userslist.asp" ' Not page request, return to list
		End If
		If sReturnUrl <> "" Then Call Page_Terminate(sReturnUrl)

		' Render row
		Users.RowType = EW_ROWTYPE_VIEW
		Call Users.ResetAttrs()
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

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		AddUrl = Users.AddUrl
		EditUrl = Users.EditUrl("")
		CopyUrl = Users.CopyUrl("")
		DeleteUrl = Users.DeleteUrl
		ListUrl = Users.ListUrl

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

			' Active
			If ew_ConvertToBool(Users.Active.CurrentValue) Then
				Users.Active.ViewValue = ew_IIf(Users.Active.FldTagCaption(1) <> "", Users.Active.FldTagCaption(1), "Yes")
			Else
				Users.Active.ViewValue = ew_IIf(Users.Active.FldTagCaption(2) <> "", Users.Active.FldTagCaption(2), "No")
			End If
			Users.Active.ViewCustomAttributes = ""

			' Profile
			Users.Profile.ViewValue = Users.Profile.CurrentValue
			Users.Profile.ViewCustomAttributes = ""

			' Theme
			If Users.Theme.CurrentValue & "" <> "" Then
				sFilterWrk = "[Theme_Name] = '" & ew_AdjustSql(Users.Theme.CurrentValue) & "'"
			sSqlWrk = "SELECT DISTINCT [Theme_Name] FROM [Themes]"
			sWhereWrk = ""
			Call ew_AddFilter(sWhereWrk, sFilterWrk)
			If sWhereWrk <> "" Then sSqlWrk = sSqlWrk & " WHERE " & sWhereWrk
			sSqlWrk = sSqlWrk & " ORDER BY [Theme_Name] Asc"
				Set RsWrk = Conn.Execute(sSqlWrk)
				If Not RsWrk.Eof Then
					Users.Theme.ViewValue = RsWrk("Theme_Name")
				Else
					Users.Theme.ViewValue = Users.Theme.CurrentValue
				End If
				RsWrk.Close
				Set RsWrk = Nothing
			Else
				Users.Theme.ViewValue = Null
			End If
			Users.Theme.ViewCustomAttributes = ""

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

			' Active
			Users.Active.LinkCustomAttributes = ""
			Users.Active.HrefValue = ""
			Users.Active.TooltipValue = ""

			' Profile
			Users.Profile.LinkCustomAttributes = ""
			Users.Profile.HrefValue = ""
			Users.Profile.TooltipValue = ""

			' Theme
			Users.Theme.LinkCustomAttributes = ""
			Users.Theme.HrefValue = ""
			Users.Theme.TooltipValue = ""
		End If

		' Call Row Rendered event
		If Users.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Users.Row_Rendered()
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
