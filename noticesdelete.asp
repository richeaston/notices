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
Dim Notices_delete
Set Notices_delete = New cNotices_delete
Set Page = Notices_delete

' Page init processing
Call Notices_delete.Page_Init()

' Page main processing
Call Notices_delete.Page_Main()
%>
<!--#include file="header.asp"-->
<script type="text/javascript">
<!--
// Create page object
var Notices_delete = new ew_Page("Notices_delete");
// page properties
Notices_delete.PageID = "delete"; // page ID
Notices_delete.FormID = "fNoticesdelete"; // form ID
var EW_PAGE_ID = Notices_delete.PageID; // for backward compatibility
// extend page with Form_CustomValidate function
Notices_delete.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
<% If EW_CLIENT_VALIDATE Then %>
Notices_delete.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Notices_delete.ValidateRequired = false; // no JavaScript validation
<% End If %>
//-->
</script>
<script language="JavaScript" type="text/javascript">
<!--
// Write your client script here, no need to add script tags.
//-->
</script>
<% Notices_delete.ShowPageHeader() %>
<%

' Load records for display
Set Notices_delete.Recordset = Notices_delete.LoadRecordset()
Notices_delete.TotalRecs = Notices_delete.Recordset.RecordCount ' Get record count
If Notices_delete.TotalRecs <= 0 Then ' No record found, exit
	Notices_delete.Recordset.Close
	Set Notices_delete.Recordset = Nothing
	Call Notices_delete.Page_Terminate("noticeslist.asp") ' Return to list
End If
%>
<p class="aspmaker ewTitle"><%= Language.Phrase("Delete") %>&nbsp;<%= Language.Phrase("TblTypeTABLE") %><%= Notices.TableCaption %></p>
<p class="aspmaker"><a href="<%= Notices.ReturnUrl %>"><%= Language.Phrase("GoBack") %></a></p>
<% Notices_delete.ShowMessage %>
<form action="<%= ew_CurrentPage %>" method="post">
<p>
<input type="hidden" name="t" id="t" value="Notices">
<input type="hidden" name="a_delete" id="a_delete" value="D">
<% For i = 0 to UBound(Notices_delete.RecKeys) %>
<input type="hidden" name="key_m" id="key_m" value="<%= ew_HtmlEncode(ew_GetKeyValue(Notices_delete.RecKeys(i))) %>">
<% Next %>
<table class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table cellspacing="0" class="ewTable ewTableSeparate">
<%= Notices.TableCustomInnerHTML %>
	<thead>
	<tr class="ewTableHeader">
		<td valign="top"><%= Notices.Title.FldCaption %></td>
		<td valign="top"><%= Notices.Author.FldCaption %></td>
		<td valign="top"><%= Notices.Sdate.FldCaption %></td>
		<td valign="top"><%= Notices.Edate.FldCaption %></td>
		<td valign="top"><%= Notices.Group.FldCaption %></td>
		<td valign="top"><%= Notices.Notice.FldCaption %></td>
	</tr>
	</thead>
	<tbody>
<%
Notices_delete.RecCnt = 0
i = 0
Do While (Not Notices_delete.Recordset.Eof)
	Notices_delete.RecCnt = Notices_delete.RecCnt + 1

	' Set row properties
	Call Notices.ResetAttrs()
	Notices.RowType = EW_ROWTYPE_VIEW ' view

	' Get the field contents
	Call Notices_delete.LoadRowValues(Notices_delete.Recordset)

	' Render row
	Call Notices_delete.RenderRow()
%>
	<tr<%= Notices.RowAttributes %>>
		<td<%= Notices.Title.CellAttributes %>>
<div<%= Notices.Title.ViewAttributes %>><%= Notices.Title.ListViewValue %></div>
</td>
		<td<%= Notices.Author.CellAttributes %>>
<div<%= Notices.Author.ViewAttributes %>><%= Notices.Author.ListViewValue %></div>
</td>
		<td<%= Notices.Sdate.CellAttributes %>>
<div<%= Notices.Sdate.ViewAttributes %>><%= Notices.Sdate.ListViewValue %></div>
</td>
		<td<%= Notices.Edate.CellAttributes %>>
<div<%= Notices.Edate.ViewAttributes %>><%= Notices.Edate.ListViewValue %></div>
</td>
		<td<%= Notices.Group.CellAttributes %>>
<div<%= Notices.Group.ViewAttributes %>><%= Notices.Group.ListViewValue %></div>
</td>
		<td<%= Notices.Notice.CellAttributes %>>
<div<%= Notices.Notice.ViewAttributes %>><%= Notices.Notice.ListViewValue %></div>
</td>
	</tr>
<%
	Notices_delete.Recordset.MoveNext
Loop
Notices_delete.Recordset.Close
Set Notices_delete.Recordset = Nothing
%>
	</tbody>
</table>
</div>
</td></tr></table>
<p>
<input type="submit" name="Action" id="Action" value="<%= ew_BtnCaption(Language.Phrase("DeleteBtn")) %>">
</form>
<%
Notices_delete.ShowPageFooter()
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
Set Notices_delete = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cNotices_delete

	' Page ID
	Public Property Get PageID()
		PageID = "delete"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Notices"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Notices_delete"
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
		' Initialize other table object

		If IsEmpty(Users) Then Set Users = New cUsers

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "delete"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "Notices"

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

	Dim TotalRecs
	Dim RecCnt
	Dim RecKeys
	Dim Recordset

	' Page main processing
	Sub Page_Main()
		Dim sFilter

		' Load Key Parameters
		RecKeys = Notices.GetRecordKeys() ' Load record keys
		sFilter = Notices.GetKeyFilter()
		If sFilter = "" Then
			Call Page_Terminate("noticeslist.asp") ' Prevent SQL injection, return to list
		End If

		' Set up filter (Sql Where Clause) and get Return Sql
		' Sql constructor in Notices class, Noticesinfo.asp

		Notices.CurrentFilter = sFilter

		' Get action
		If Request.Form("a_delete").Count > 0 Then
			Notices.CurrentAction = Request.Form("a_delete")
		Else
			Notices.CurrentAction = "I"	' Display record
		End If
		Select Case Notices.CurrentAction
			Case "D" ' Delete
				Notices.SendEmail = True ' Send email on delete success
				If DeleteRows() Then ' delete rows
					SuccessMessage = Language.Phrase("DeleteSuccess") ' Set up success message
					Call Page_Terminate(Notices.ReturnUrl) ' Return to caller
				End If
		End Select
	End Sub

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = Notices.CurrentFilter
		Call Notices.Recordset_Selecting(sFilter)
		Notices.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = Notices.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call Notices.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

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
		End If

		' Call Row Rendered event
		If Notices.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Notices.Row_Rendered()
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
		sSql = Notices.SQL
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
				DeleteRows = Notices.Row_Deleting(RsDelete)
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
			If Notices.CancelMessage <> "" Then
				FailureMessage = Notices.CancelMessage
				Notices.CancelMessage = ""
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
			If DeleteRows Then
				If Not RsOld.Eof Then RsOld.MoveFirst
				Do While Not RsOld.Eof
					Call WriteAuditTrailOnDelete(RsOld)
					RsOld.MoveNext
				Loop
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
				Call Notices.Row_Deleted(RsOld)
				RsOld.MoveNext
			Loop
		End If
		RsOld.Close
		Set RsOld = Nothing
	End Function

	' Write Audit Trail start/end for grid update
	Sub WriteAuditTrailDummy(typ)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		Dim table
		table = "Notices"

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

	' Write Audit Trail (delete page)
	Sub WriteAuditTrailOnDelete(RsSrc)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		Dim table
		table = "Notices"

		' Write Audit Trail
		Dim filePfx, curDateTime, id, user, action, field, keyvalue, oldvalue, newvalue
		Dim i
		filePfx = "log"
		curDateTime = ew_StdCurrentDateTime()
		id = Request.ServerVariables("SCRIPT_NAME")
    	user = CurrentUserName
		action = "D"
		newvalue = ""

		' Get key value
		Dim sKey
		sKey = ""
		If sKey <> "" Then sKey = sKey & EW_COMPOSITE_KEY_SEPARATOR
		sKey = sKey & RsSrc.Fields("Notice_ID")
		keyvalue = sKey

		' Notice_ID Field
		oldvalue = RsSrc("Notice_ID")
		Call ew_WriteAuditTrail(filePfx, curDateTime, id, user, action, table, "Notice_ID", keyvalue, oldvalue, newvalue)

		' Title Field
		oldvalue = RsSrc("Title")
		Call ew_WriteAuditTrail(filePfx, curDateTime, id, user, action, table, "Title", keyvalue, oldvalue, newvalue)

		' Author Field
		oldvalue = RsSrc("Author")
		Call ew_WriteAuditTrail(filePfx, curDateTime, id, user, action, table, "Author", keyvalue, oldvalue, newvalue)

		' Sdate Field
		oldvalue = RsSrc("Sdate")
		Call ew_WriteAuditTrail(filePfx, curDateTime, id, user, action, table, "Sdate", keyvalue, oldvalue, newvalue)

		' Edate Field
		oldvalue = RsSrc("Edate")
		Call ew_WriteAuditTrail(filePfx, curDateTime, id, user, action, table, "Edate", keyvalue, oldvalue, newvalue)

		' Group Field
		oldvalue = RsSrc("Group")
		Call ew_WriteAuditTrail(filePfx, curDateTime, id, user, action, table, "Group", keyvalue, oldvalue, newvalue)

		' Notice Field
		oldvalue = RsSrc("Notice")
		If Not EW_AUDIT_TRAIL_TO_DATABASE Then oldvalue = "[MEMO]"
		Call ew_WriteAuditTrail(filePfx, curDateTime, id, user, action, table, "Notice", keyvalue, oldvalue, newvalue)

		' Approved Field
		oldvalue = RsSrc("Approved")
		Call ew_WriteAuditTrail(filePfx, curDateTime, id, user, action, table, "Approved", keyvalue, oldvalue, newvalue)

		' Archived Field
		oldvalue = RsSrc("Archived")
		Call ew_WriteAuditTrail(filePfx, curDateTime, id, user, action, table, "Archived", keyvalue, oldvalue, newvalue)
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
