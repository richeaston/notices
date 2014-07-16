<%@ CodePage="1252" LCID="2057" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="ewcfg9.asp"-->
<!--#include file="approved_noticesinfo.asp"-->
<!--#include file="usersinfo.asp"-->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim Approved_Notices_delete
Set Approved_Notices_delete = New cApproved_Notices_delete
Set Page = Approved_Notices_delete

' Page init processing
Call Approved_Notices_delete.Page_Init()

' Page main processing
Call Approved_Notices_delete.Page_Main()
%>
<!--#include file="header.asp"-->
<script type="text/javascript">
<!--
// Create page object
var Approved_Notices_delete = new ew_Page("Approved_Notices_delete");
// page properties
Approved_Notices_delete.PageID = "delete"; // page ID
Approved_Notices_delete.FormID = "fApproved_Noticesdelete"; // form ID
var EW_PAGE_ID = Approved_Notices_delete.PageID; // for backward compatibility
// extend page with Form_CustomValidate function
Approved_Notices_delete.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
<% If EW_CLIENT_VALIDATE Then %>
Approved_Notices_delete.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Approved_Notices_delete.ValidateRequired = false; // no JavaScript validation
<% End If %>
//-->
</script>
<script language="JavaScript" type="text/javascript">
<!--
// Write your client script here, no need to add script tags.
//-->
</script>
<% Approved_Notices_delete.ShowPageHeader() %>
<%

' Load records for display
Set Approved_Notices_delete.Recordset = Approved_Notices_delete.LoadRecordset()
Approved_Notices_delete.TotalRecs = Approved_Notices_delete.Recordset.RecordCount ' Get record count
If Approved_Notices_delete.TotalRecs <= 0 Then ' No record found, exit
	Approved_Notices_delete.Recordset.Close
	Set Approved_Notices_delete.Recordset = Nothing
	Call Approved_Notices_delete.Page_Terminate("approved_noticeslist.asp") ' Return to list
End If
%>
<p class="aspmaker ewTitle"><%= Language.Phrase("Delete") %>&nbsp;<%= Language.Phrase("TblTypeVIEW") %><%= Approved_Notices.TableCaption %></p>
<p class="aspmaker"><a href="<%= Approved_Notices.ReturnUrl %>"><%= Language.Phrase("GoBack") %></a></p>
<% Approved_Notices_delete.ShowMessage %>
<form action="<%= ew_CurrentPage %>" method="post">
<p>
<input type="hidden" name="t" id="t" value="Approved_Notices">
<input type="hidden" name="a_delete" id="a_delete" value="D">
<% For i = 0 to UBound(Approved_Notices_delete.RecKeys) %>
<input type="hidden" name="key_m" id="key_m" value="<%= ew_HtmlEncode(ew_GetKeyValue(Approved_Notices_delete.RecKeys(i))) %>">
<% Next %>
<table class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table cellspacing="0" class="ewTable ewTableSeparate">
<%= Approved_Notices.TableCustomInnerHTML %>
	<thead>
	<tr class="ewTableHeader">
		<td valign="top"><%= Approved_Notices.Title.FldCaption %></td>
		<td valign="top"><%= Approved_Notices.Author.FldCaption %></td>
		<td valign="top"><%= Approved_Notices.Sdate.FldCaption %></td>
		<td valign="top"><%= Approved_Notices.Edate.FldCaption %></td>
		<td valign="top"><%= Approved_Notices.Group.FldCaption %></td>
	</tr>
	</thead>
	<tbody>
<%
Approved_Notices_delete.RecCnt = 0
i = 0
Do While (Not Approved_Notices_delete.Recordset.Eof)
	Approved_Notices_delete.RecCnt = Approved_Notices_delete.RecCnt + 1

	' Set row properties
	Call Approved_Notices.ResetAttrs()
	Approved_Notices.RowType = EW_ROWTYPE_VIEW ' view

	' Get the field contents
	Call Approved_Notices_delete.LoadRowValues(Approved_Notices_delete.Recordset)

	' Render row
	Call Approved_Notices_delete.RenderRow()
%>
	<tr<%= Approved_Notices.RowAttributes %>>
		<td<%= Approved_Notices.Title.CellAttributes %>>
<div<%= Approved_Notices.Title.ViewAttributes %>><%= Approved_Notices.Title.ListViewValue %></div>
</td>
		<td<%= Approved_Notices.Author.CellAttributes %>>
<div<%= Approved_Notices.Author.ViewAttributes %>><%= Approved_Notices.Author.ListViewValue %></div>
</td>
		<td<%= Approved_Notices.Sdate.CellAttributes %>>
<div<%= Approved_Notices.Sdate.ViewAttributes %>><%= Approved_Notices.Sdate.ListViewValue %></div>
</td>
		<td<%= Approved_Notices.Edate.CellAttributes %>>
<div<%= Approved_Notices.Edate.ViewAttributes %>><%= Approved_Notices.Edate.ListViewValue %></div>
</td>
		<td<%= Approved_Notices.Group.CellAttributes %>>
<div<%= Approved_Notices.Group.ViewAttributes %>><%= Approved_Notices.Group.ListViewValue %></div>
</td>
	</tr>
<%
	Approved_Notices_delete.Recordset.MoveNext
Loop
Approved_Notices_delete.Recordset.Close
Set Approved_Notices_delete.Recordset = Nothing
%>
	</tbody>
</table>
</div>
</td></tr></table>
<p>
<input type="submit" name="Action" id="Action" value="<%= ew_BtnCaption(Language.Phrase("DeleteBtn")) %>">
</form>
<%
Approved_Notices_delete.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script language="JavaScript" type="text/javascript">
<!--
// Write your table-specific startup script here
// document.write("page loaded");
//-->
</script>
<!--#include file="footer.asp"-->
<%

' Drop page object
Set Approved_Notices_delete = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cApproved_Notices_delete

	' Page ID
	Public Property Get PageID()
		PageID = "delete"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Approved Notices"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Approved_Notices_delete"
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
		' Initialize other table object

		If IsEmpty(Users) Then Set Users = New cUsers

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "delete"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "Approved Notices"

		' Open connection to the database
		If IsEmpty(Conn) Then Call ew_Connect()
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
		If Not Security.CanDelete Then
			Call Security.SaveLastUrl()
			Call Page_Terminate("approved_noticeslist.asp")
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
		Set Approved_Notices = Nothing
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

	Dim TotalRecs
	Dim RecCnt
	Dim RecKeys
	Dim Recordset

	' Page main processing
	Sub Page_Main()
		Dim sFilter

		' Load Key Parameters
		RecKeys = Approved_Notices.GetRecordKeys() ' Load record keys
		sFilter = Approved_Notices.GetKeyFilter()
		If sFilter = "" Then
			Call Page_Terminate("approved_noticeslist.asp") ' Prevent SQL injection, return to list
		End If

		' Set up filter (Sql Where Clause) and get Return Sql
		' Sql constructor in Approved_Notices class, Approved_Noticesinfo.asp

		Approved_Notices.CurrentFilter = sFilter

		' Get action
		If Request.Form("a_delete").Count > 0 Then
			Approved_Notices.CurrentAction = Request.Form("a_delete")
		Else
			Approved_Notices.CurrentAction = "I"	' Display record
		End If
		Select Case Approved_Notices.CurrentAction
			Case "D" ' Delete
				Approved_Notices.SendEmail = True ' Send email on delete success
				If DeleteRows() Then ' delete rows
					SuccessMessage = Language.Phrase("DeleteSuccess") ' Set up success message
					Call Page_Terminate(Approved_Notices.ReturnUrl) ' Return to caller
				End If
		End Select
	End Sub

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

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
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
		End If

		' Call Row Rendered event
		If Approved_Notices.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Approved_Notices.Row_Rendered()
		End If
	End Sub

	'
	' Delete records based on current filter
	'
	Function DeleteRows()
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		Dim sKey, sThisKey, sKeyFld, arKeyFlds
		Dim sSql, RsDelete
		Dim RsOld
		DeleteRows = True
		sSql = Approved_Notices.SQL
		If Not Security.CanDelete Then
			FailureMessage = Language.Phrase("NoDeletePermission") ' No delete permission
			DeleteRows = False
			Exit Function
		End If
		Set RsDelete = Server.CreateObject("ADODB.Recordset")
		RsDelete.CursorLocation = EW_CURSORLOCATION
		RsDelete.Open sSql, Conn, 1, EW_RECORDSET_LOCKTYPE
		If Err.Number <> 0 Then
			FailureMessage = Err.Description
			RsDelete.Close
			Set RsDelete = Nothing
			DeleteRows = False
			Exit Function
		ElseIf RsDelete.Eof Then
			FailureMessage = Language.Phrase("NoRecord") ' No record found
			RsDelete.Close
			Set RsDelete = Nothing
			DeleteRows = False
			Exit Function
		End If
		Conn.BeginTrans

		' Clone old recordset object
		Set RsOld = ew_CloneRs(RsDelete)

		' Call row deleting event
		If DeleteRows Then
			RsDelete.MoveFirst
			Do While Not RsDelete.Eof
				DeleteRows = Approved_Notices.Row_Deleting(RsDelete)
				If Not DeleteRows Then Exit Do
				RsDelete.MoveNext
			Loop
			RsDelete.MoveFirst
		End If
		If DeleteRows Then
			sKey = ""
			RsDelete.MoveFirst
			Do While Not RsDelete.Eof
				sThisKey = ""
				If sThisKey <> "" Then sThisKey = sThisKey & EW_COMPOSITE_KEY_SEPARATOR
				sThisKey = sThisKey & RsDelete("Notice_ID")
				RsDelete.Delete
				If Err.Number <> 0 Then
					FailureMessage = Err.Description ' Set up error message
					DeleteRows = False
					Exit Do
				End If
				If sKey <> "" Then sKey = sKey & ", "
				sKey = sKey & sThisKey
				RsDelete.MoveNext
			Loop
		Else

			' Set up error message
			If Approved_Notices.CancelMessage <> "" Then
				FailureMessage = Approved_Notices.CancelMessage
				Approved_Notices.CancelMessage = ""
			Else
				FailureMessage = Language.Phrase("DeleteCancelled")
			End If
		End If
		If DeleteRows Then
			Conn.CommitTrans ' Commit the changes
			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				DeleteRows = False ' Delete failed
			End If
		Else
			Conn.RollbackTrans ' Rollback changes
		End If
		RsDelete.Close
		Set RsDelete = Nothing

		' Call row deleting event
		If DeleteRows Then
			If Not RsOld.Eof Then RsOld.MoveFirst
			Do While Not RsOld.Eof
				Call Approved_Notices.Row_Deleted(RsOld)
				RsOld.MoveNext
			Loop
		End If
		RsOld.Close
		Set RsOld = Nothing
	End Function

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
