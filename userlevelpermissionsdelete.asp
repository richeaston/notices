<%@ CodePage="1252" LCID="2057" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="ewcfg9.asp"-->
<!--#include file="userlevelpermissionsinfo.asp"-->
<!--#include file="usersinfo.asp"-->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim UserLevelPermissions_delete
Set UserLevelPermissions_delete = New cUserLevelPermissions_delete
Set Page = UserLevelPermissions_delete

' Page init processing
Call UserLevelPermissions_delete.Page_Init()

' Page main processing
Call UserLevelPermissions_delete.Page_Main()
%>
<!--#include file="header.asp"-->
<script type="text/javascript">
<!--
// Create page object
var UserLevelPermissions_delete = new ew_Page("UserLevelPermissions_delete");
// page properties
UserLevelPermissions_delete.PageID = "delete"; // page ID
UserLevelPermissions_delete.FormID = "fUserLevelPermissionsdelete"; // form ID
var EW_PAGE_ID = UserLevelPermissions_delete.PageID; // for backward compatibility
// extend page with Form_CustomValidate function
UserLevelPermissions_delete.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
UserLevelPermissions_delete.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
UserLevelPermissions_delete.ValidateRequired = true; // uses JavaScript validation
<% Else %>
UserLevelPermissions_delete.ValidateRequired = false; // no JavaScript validation
<% End If %>
//-->
</script>
<script language="JavaScript" type="text/javascript">
<!--
// Write your client script here, no need to add script tags.
//-->
</script>
<% UserLevelPermissions_delete.ShowPageHeader() %>
<%

' Load records for display
Set UserLevelPermissions_delete.Recordset = UserLevelPermissions_delete.LoadRecordset()
UserLevelPermissions_delete.TotalRecs = UserLevelPermissions_delete.Recordset.RecordCount ' Get record count
If UserLevelPermissions_delete.TotalRecs <= 0 Then ' No record found, exit
	UserLevelPermissions_delete.Recordset.Close
	Set UserLevelPermissions_delete.Recordset = Nothing
	Call UserLevelPermissions_delete.Page_Terminate("userlevelpermissionslist.asp") ' Return to list
End If
%>
<p class="aspmaker ewTitle"><%= Language.Phrase("Delete") %>&nbsp;<%= Language.Phrase("TblTypeTABLE") %><%= UserLevelPermissions.TableCaption %></p>
<p class="aspmaker"><a href="<%= UserLevelPermissions.ReturnUrl %>"><%= Language.Phrase("GoBack") %></a></p>
<% UserLevelPermissions_delete.ShowMessage %>
<form action="<%= ew_CurrentPage %>" method="post">
<p>
<input type="hidden" name="t" id="t" value="UserLevelPermissions">
<input type="hidden" name="a_delete" id="a_delete" value="D">
<% For i = 0 to UBound(UserLevelPermissions_delete.RecKeys) %>
<input type="hidden" name="key_m" id="key_m" value="<%= ew_HtmlEncode(ew_GetKeyValue(UserLevelPermissions_delete.RecKeys(i))) %>">
<% Next %>
<table class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table cellspacing="0" class="ewTable ewTableSeparate">
<%= UserLevelPermissions.TableCustomInnerHTML %>
	<thead>
	<tr class="ewTableHeader">
		<td valign="top"><%= UserLevelPermissions.UserLevelID.FldCaption %></td>
		<td valign="top"><%= UserLevelPermissions.zTableName.FldCaption %></td>
		<td valign="top"><%= UserLevelPermissions.Permission.FldCaption %></td>
	</tr>
	</thead>
	<tbody>
<%
UserLevelPermissions_delete.RecCnt = 0
i = 0
Do While (Not UserLevelPermissions_delete.Recordset.Eof)
	UserLevelPermissions_delete.RecCnt = UserLevelPermissions_delete.RecCnt + 1

	' Set row properties
	Call UserLevelPermissions.ResetAttrs()
	UserLevelPermissions.RowType = EW_ROWTYPE_VIEW ' view

	' Get the field contents
	Call UserLevelPermissions_delete.LoadRowValues(UserLevelPermissions_delete.Recordset)

	' Render row
	Call UserLevelPermissions_delete.RenderRow()
%>
	<tr<%= UserLevelPermissions.RowAttributes %>>
		<td<%= UserLevelPermissions.UserLevelID.CellAttributes %>>
<div<%= UserLevelPermissions.UserLevelID.ViewAttributes %>><%= UserLevelPermissions.UserLevelID.ListViewValue %></div>
</td>
		<td<%= UserLevelPermissions.zTableName.CellAttributes %>>
<div<%= UserLevelPermissions.zTableName.ViewAttributes %>><%= UserLevelPermissions.zTableName.ListViewValue %></div>
</td>
		<td<%= UserLevelPermissions.Permission.CellAttributes %>>
<div<%= UserLevelPermissions.Permission.ViewAttributes %>><%= UserLevelPermissions.Permission.ListViewValue %></div>
</td>
	</tr>
<%
	UserLevelPermissions_delete.Recordset.MoveNext
Loop
UserLevelPermissions_delete.Recordset.Close
Set UserLevelPermissions_delete.Recordset = Nothing
%>
	</tbody>
</table>
</div>
</td></tr></table>
<p>
<input type="submit" name="Action" id="Action" value="<%= ew_BtnCaption(Language.Phrase("DeleteBtn")) %>">
</form>
<%
UserLevelPermissions_delete.ShowPageFooter()
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
Set UserLevelPermissions_delete = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cUserLevelPermissions_delete

	' Page ID
	Public Property Get PageID()
		PageID = "delete"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "UserLevelPermissions"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "UserLevelPermissions_delete"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If UserLevelPermissions.UseTokenInUrl Then PageUrl = PageUrl & "t=" & UserLevelPermissions.TableVar & "&" ' add page token
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
		If UserLevelPermissions.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (UserLevelPermissions.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (UserLevelPermissions.TableVar = Request.QueryString("t"))
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
		If IsEmpty(UserLevelPermissions) Then Set UserLevelPermissions = New cUserLevelPermissions
		Set Table = UserLevelPermissions

		' Initialize urls
		' Initialize other table object

		If IsEmpty(Users) Then Set Users = New cUsers

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "delete"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "UserLevelPermissions"

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
		Set UserLevelPermissions = Nothing
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
		RecKeys = UserLevelPermissions.GetRecordKeys() ' Load record keys
		sFilter = UserLevelPermissions.GetKeyFilter()
		If sFilter = "" Then
			Call Page_Terminate("userlevelpermissionslist.asp") ' Prevent SQL injection, return to list
		End If

		' Set up filter (Sql Where Clause) and get Return Sql
		' Sql constructor in UserLevelPermissions class, UserLevelPermissionsinfo.asp

		UserLevelPermissions.CurrentFilter = sFilter

		' Get action
		If Request.Form("a_delete").Count > 0 Then
			UserLevelPermissions.CurrentAction = Request.Form("a_delete")
		Else
			UserLevelPermissions.CurrentAction = "I"	' Display record
		End If
		Select Case UserLevelPermissions.CurrentAction
			Case "D" ' Delete
				UserLevelPermissions.SendEmail = True ' Send email on delete success
				If DeleteRows() Then ' delete rows
					SuccessMessage = Language.Phrase("DeleteSuccess") ' Set up success message
					Call Page_Terminate(UserLevelPermissions.ReturnUrl) ' Return to caller
				End If
		End Select
	End Sub

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = UserLevelPermissions.CurrentFilter
		Call UserLevelPermissions.Recordset_Selecting(sFilter)
		UserLevelPermissions.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = UserLevelPermissions.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call UserLevelPermissions.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = UserLevelPermissions.KeyFilter

		' Call Row Selecting event
		Call UserLevelPermissions.Row_Selecting(sFilter)

		' Load sql based on filter
		UserLevelPermissions.CurrentFilter = sFilter
		sSql = UserLevelPermissions.SQL
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
		Call UserLevelPermissions.Row_Selected(RsRow)
		UserLevelPermissions.UserLevelID.DbValue = RsRow("UserLevelID")
		UserLevelPermissions.zTableName.DbValue = RsRow("TableName")
		UserLevelPermissions.Permission.DbValue = RsRow("Permission")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Call Row Rendering event

		Call UserLevelPermissions.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' UserLevelID
		' TableName
		' Permission
		' -----------
		'  View  Row
		' -----------

		If UserLevelPermissions.RowType = EW_ROWTYPE_VIEW Then ' View row

			' UserLevelID
			UserLevelPermissions.UserLevelID.ViewValue = UserLevelPermissions.UserLevelID.CurrentValue
			UserLevelPermissions.UserLevelID.ViewCustomAttributes = ""

			' TableName
			UserLevelPermissions.zTableName.ViewValue = UserLevelPermissions.zTableName.CurrentValue
			UserLevelPermissions.zTableName.ViewCustomAttributes = ""

			' Permission
			UserLevelPermissions.Permission.ViewValue = UserLevelPermissions.Permission.CurrentValue
			UserLevelPermissions.Permission.ViewCustomAttributes = ""

			' View refer script
			' UserLevelID

			UserLevelPermissions.UserLevelID.LinkCustomAttributes = ""
			UserLevelPermissions.UserLevelID.HrefValue = ""
			UserLevelPermissions.UserLevelID.TooltipValue = ""

			' TableName
			UserLevelPermissions.zTableName.LinkCustomAttributes = ""
			UserLevelPermissions.zTableName.HrefValue = ""
			UserLevelPermissions.zTableName.TooltipValue = ""

			' Permission
			UserLevelPermissions.Permission.LinkCustomAttributes = ""
			UserLevelPermissions.Permission.HrefValue = ""
			UserLevelPermissions.Permission.TooltipValue = ""
		End If

		' Call Row Rendered event
		If UserLevelPermissions.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call UserLevelPermissions.Row_Rendered()
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
		sSql = UserLevelPermissions.SQL
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
				DeleteRows = UserLevelPermissions.Row_Deleting(RsDelete)
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
				sThisKey = sThisKey & RsDelete("UserLevelID")
				If sThisKey <> "" Then sThisKey = sThisKey & EW_COMPOSITE_KEY_SEPARATOR
				sThisKey = sThisKey & RsDelete("TableName")
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
			If UserLevelPermissions.CancelMessage <> "" Then
				FailureMessage = UserLevelPermissions.CancelMessage
				UserLevelPermissions.CancelMessage = ""
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
				Call UserLevelPermissions.Row_Deleted(RsOld)
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
