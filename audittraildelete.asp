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
Dim AuditTrail_delete
Set AuditTrail_delete = New cAuditTrail_delete
Set Page = AuditTrail_delete

' Page init processing
Call AuditTrail_delete.Page_Init()

' Page main processing
Call AuditTrail_delete.Page_Main()
%>
<!--#include file="header.asp"-->
<script type="text/javascript">
<!--
// Create page object
var AuditTrail_delete = new ew_Page("AuditTrail_delete");
// page properties
AuditTrail_delete.PageID = "delete"; // page ID
AuditTrail_delete.FormID = "fAuditTraildelete"; // form ID
var EW_PAGE_ID = AuditTrail_delete.PageID; // for backward compatibility
// extend page with Form_CustomValidate function
AuditTrail_delete.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
AuditTrail_delete.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
AuditTrail_delete.ValidateRequired = true; // uses JavaScript validation
<% Else %>
AuditTrail_delete.ValidateRequired = false; // no JavaScript validation
<% End If %>
//-->
</script>
<script language="JavaScript" type="text/javascript">
<!--
// Write your client script here, no need to add script tags.
//-->
</script>
<% AuditTrail_delete.ShowPageHeader() %>
<%

' Load records for display
Set AuditTrail_delete.Recordset = AuditTrail_delete.LoadRecordset()
AuditTrail_delete.TotalRecs = AuditTrail_delete.Recordset.RecordCount ' Get record count
If AuditTrail_delete.TotalRecs <= 0 Then ' No record found, exit
	AuditTrail_delete.Recordset.Close
	Set AuditTrail_delete.Recordset = Nothing
	Call AuditTrail_delete.Page_Terminate("audittraillist.asp") ' Return to list
End If
%>
<p class="aspmaker"><a class="btn btn-inverse" href="<%= AuditTrail.ReturnUrl %>"><i class="icon-arrow-left icon-white"></i>&nbsp;<%= Language.Phrase("GoBack") %></a></p>
<% AuditTrail_delete.ShowMessage %>
<form action="<%= ew_CurrentPage %>" method="post">
<p>
<input class="btn btn-danger" type="submit" name="Action" id="Action" value="<%= ew_BtnCaption(Language.Phrase("DeleteBtn")) %>">
<input type="hidden" name="t" id="t" value="AuditTrail">
<input type="hidden" name="a_delete" id="a_delete" value="D">
<% For i = 0 to UBound(AuditTrail_delete.RecKeys) %>
<input type="hidden" name="key_m" id="key_m" value="<%= ew_HtmlEncode(ew_GetKeyValue(AuditTrail_delete.RecKeys(i))) %>">
<% Next %>
<table class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel ">
<table cellspacing="0" class="ewTable ewTableSeparate">
<%= AuditTrail.TableCustomInnerHTML %>
	<thead>
	<tr class="ewTableHeader">
		<td valign="top"><%= AuditTrail.DateTime.FldCaption %></td>
		<td valign="top"><%= AuditTrail.User.FldCaption %></td>
		<td valign="top"><%= AuditTrail.Action.FldCaption %></td>
		<td valign="top"><%= AuditTrail.zTable.FldCaption %></td>
		<td valign="top"><%= AuditTrail.zField.FldCaption %></td>
	</tr>
	</thead>
	<tbody>
<%
AuditTrail_delete.RecCnt = 0
i = 0
Do While (Not AuditTrail_delete.Recordset.Eof)
	AuditTrail_delete.RecCnt = AuditTrail_delete.RecCnt + 1

	' Set row properties
	Call AuditTrail.ResetAttrs()
	AuditTrail.RowType = EW_ROWTYPE_VIEW ' view

	' Get the field contents
	Call AuditTrail_delete.LoadRowValues(AuditTrail_delete.Recordset)

	' Render row
	Call AuditTrail_delete.RenderRow()
%>
	<tr<%= AuditTrail.RowAttributes %>>
		<td<%= AuditTrail.DateTime.CellAttributes %>>
<div<%= AuditTrail.DateTime.ViewAttributes %>><%= AuditTrail.DateTime.ListViewValue %></div>
</td>
		<td<%= AuditTrail.User.CellAttributes %>>
<div<%= AuditTrail.User.ViewAttributes %>><%= AuditTrail.User.ListViewValue %></div>
</td>
		<td<%= AuditTrail.Action.CellAttributes %>>
<div<%= AuditTrail.Action.ViewAttributes %>><%= AuditTrail.Action.ListViewValue %></div>
</td>
		<td<%= AuditTrail.zTable.CellAttributes %>>
<div<%= AuditTrail.zTable.ViewAttributes %>><%= AuditTrail.zTable.ListViewValue %></div>
</td>
		<td<%= AuditTrail.zField.CellAttributes %>>
<div<%= AuditTrail.zField.ViewAttributes %>><%= AuditTrail.zField.ListViewValue %></div>
</td>
	</tr>
<%
	AuditTrail_delete.Recordset.MoveNext
Loop
AuditTrail_delete.Recordset.Close
Set AuditTrail_delete.Recordset = Nothing
%>
	</tbody>
</table>
</div>
</td></tr></table>
<p>
<input class="btn btn-danger" type="submit" name="Action" id="Action" value="<%= ew_BtnCaption(Language.Phrase("DeleteBtn")) %>">
</form>
<%
AuditTrail_delete.ShowPageFooter()
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
Set AuditTrail_delete = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cAuditTrail_delete

	' Page ID
	Public Property Get PageID()
		PageID = "delete"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "AuditTrail"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "AuditTrail_delete"
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
		' Initialize other table object

		If IsEmpty(Users) Then Set Users = New cUsers

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "delete"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "AuditTrail"

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

	Dim TotalRecs
	Dim RecCnt
	Dim RecKeys
	Dim Recordset

	' Page main processing
	Sub Page_Main()
		Dim sFilter

		' Load Key Parameters
		RecKeys = AuditTrail.GetRecordKeys() ' Load record keys
		sFilter = AuditTrail.GetKeyFilter()
		If sFilter = "" Then
			Call Page_Terminate("audittraillist.asp") ' Prevent SQL injection, return to list
		End If

		' Set up filter (Sql Where Clause) and get Return Sql
		' Sql constructor in AuditTrail class, AuditTrailinfo.asp

		AuditTrail.CurrentFilter = sFilter

		' Get action
		If Request.Form("a_delete").Count > 0 Then
			AuditTrail.CurrentAction = Request.Form("a_delete")
		Else
			AuditTrail.CurrentAction = "I"	' Display record
		End If
		Select Case AuditTrail.CurrentAction
			Case "D" ' Delete
				AuditTrail.SendEmail = True ' Send email on delete success
				If DeleteRows() Then ' delete rows
					SuccessMessage = Language.Phrase("DeleteSuccess") ' Set up success message
					Call Page_Terminate(AuditTrail.ReturnUrl) ' Return to caller
				End If
		End Select
	End Sub

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = AuditTrail.CurrentFilter
		Call AuditTrail.Recordset_Selecting(sFilter)
		AuditTrail.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = AuditTrail.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call AuditTrail.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

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

			' View refer script
			' DateTime

			AuditTrail.DateTime.LinkCustomAttributes = ""
			AuditTrail.DateTime.HrefValue = ""
			AuditTrail.DateTime.TooltipValue = ""

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
		End If

		' Call Row Rendered event
		If AuditTrail.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call AuditTrail.Row_Rendered()
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
		sSql = AuditTrail.SQL
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
				DeleteRows = AuditTrail.Row_Deleting(RsDelete)
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
				sThisKey = sThisKey & RsDelete("Id")
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
			If AuditTrail.CancelMessage <> "" Then
				FailureMessage = AuditTrail.CancelMessage
				AuditTrail.CancelMessage = ""
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
				Call AuditTrail.Row_Deleted(RsOld)
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
