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
Dim Users_delete
Set Users_delete = New cUsers_delete
Set Page = Users_delete

' Page init processing
Call Users_delete.Page_Init()

' Page main processing
Call Users_delete.Page_Main()
%>
<!--#include file="header.asp"-->
<script type="text/javascript">
<!--
// Create page object
var Users_delete = new ew_Page("Users_delete");
// page properties
Users_delete.PageID = "delete"; // page ID
Users_delete.FormID = "fUsersdelete"; // form ID
var EW_PAGE_ID = Users_delete.PageID; // for backward compatibility
// extend page with Form_CustomValidate function
Users_delete.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
Users_delete.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
Users_delete.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Users_delete.ValidateRequired = false; // no JavaScript validation
<% End If %>
//-->
</script>
<script language="JavaScript" type="text/javascript">
<!--
// Write your client script here, no need to add script tags.
//-->
</script>
<% Users_delete.ShowPageHeader() %>
<%

' Load records for display
Set Users_delete.Recordset = Users_delete.LoadRecordset()
Users_delete.TotalRecs = Users_delete.Recordset.RecordCount ' Get record count
If Users_delete.TotalRecs <= 0 Then ' No record found, exit
	Users_delete.Recordset.Close
	Set Users_delete.Recordset = Nothing
	Call Users_delete.Page_Terminate("userslist.asp") ' Return to list
End If
%>
<p class="aspmaker ewTitle"><%= Language.Phrase("Delete") %>&nbsp;<%= Language.Phrase("TblTypeTABLE") %><%= Users.TableCaption %></p>
<p class="aspmaker"><a href="<%= Users.ReturnUrl %>"><%= Language.Phrase("GoBack") %></a></p>
<% Users_delete.ShowMessage %>
<form action="<%= ew_CurrentPage %>" method="post">
<p>
<input type="hidden" name="t" id="t" value="Users">
<input type="hidden" name="a_delete" id="a_delete" value="D">
<% For i = 0 to UBound(Users_delete.RecKeys) %>
<input type="hidden" name="key_m" id="key_m" value="<%= ew_HtmlEncode(ew_GetKeyValue(Users_delete.RecKeys(i))) %>">
<% Next %>
<table class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table cellspacing="0" class="ewTable ewTableSeparate">
<%= Users.TableCustomInnerHTML %>
	<thead>
	<tr class="ewTableHeader">
		<td valign="top"><%= Users.Username.FldCaption %></td>
		<td valign="top"><%= Users.zEmail.FldCaption %></td>
		<td valign="top"><%= Users.Permissions.FldCaption %></td>
	</tr>
	</thead>
	<tbody>
<%
Users_delete.RecCnt = 0
i = 0
Do While (Not Users_delete.Recordset.Eof)
	Users_delete.RecCnt = Users_delete.RecCnt + 1

	' Set row properties
	Call Users.ResetAttrs()
	Users.RowType = EW_ROWTYPE_VIEW ' view

	' Get the field contents
	Call Users_delete.LoadRowValues(Users_delete.Recordset)

	' Render row
	Call Users_delete.RenderRow()
%>
	<tr<%= Users.RowAttributes %>>
		<td<%= Users.Username.CellAttributes %>>
<div<%= Users.Username.ViewAttributes %>><%= Users.Username.ListViewValue %></div>
</td>
		<td<%= Users.zEmail.CellAttributes %>>
<div<%= Users.zEmail.ViewAttributes %>>
<% If Users.zEmail.LinkAttributes <> "" Then %>
<a<%= Users.zEmail.LinkAttributes %>><%= Users.zEmail.ListViewValue %></a>
<% Else %>
<%= Users.zEmail.ListViewValue %>
<% End If %>
</div>
</td>
		<td<%= Users.Permissions.CellAttributes %>>
<div<%= Users.Permissions.ViewAttributes %>><%= Users.Permissions.ListViewValue %></div>
</td>
	</tr>
<%
	Users_delete.Recordset.MoveNext
Loop
Users_delete.Recordset.Close
Set Users_delete.Recordset = Nothing
%>
	</tbody>
</table>
</div>
</td></tr></table>
<p>
<input type="submit" name="Action" id="Action" value="<%= ew_BtnCaption(Language.Phrase("DeleteBtn")) %>">
</form>
<%
Users_delete.ShowPageFooter()
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
Set Users_delete = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cUsers_delete

	' Page ID
	Public Property Get PageID()
		PageID = "delete"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Users"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Users_delete"
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
		' Initialize form object

		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "delete"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "Users"

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

	Dim TotalRecs
	Dim RecCnt
	Dim RecKeys
	Dim Recordset

	' Page main processing
	Sub Page_Main()
		Dim sFilter

		' Load Key Parameters
		RecKeys = Users.GetRecordKeys() ' Load record keys
		sFilter = Users.GetKeyFilter()
		If sFilter = "" Then
			Call Page_Terminate("userslist.asp") ' Prevent SQL injection, return to list
		End If

		' Set up filter (Sql Where Clause) and get Return Sql
		' Sql constructor in Users class, Usersinfo.asp

		Users.CurrentFilter = sFilter

		' Get action
		If Request.Form("a_delete").Count > 0 Then
			Users.CurrentAction = Request.Form("a_delete")
		Else
			Users.CurrentAction = "I"	' Display record
		End If
		Select Case Users.CurrentAction
			Case "D" ' Delete
				Users.SendEmail = True ' Send email on delete success
				If DeleteRows() Then ' delete rows
					SuccessMessage = Language.Phrase("DeleteSuccess") ' Set up success message
					Call Page_Terminate(Users.ReturnUrl) ' Return to caller
				End If
		End Select
	End Sub

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = Users.CurrentFilter
		Call Users.Recordset_Selecting(sFilter)
		Users.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = Users.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call Users.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

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
		End If

		' Call Row Rendered event
		If Users.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Users.Row_Rendered()
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
		sSql = Users.SQL
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
		Call WriteAuditTrailDummy(Language.Phrase("BatchDeleteBegin")) ' Batch update begin

		' Clone old recordset object
		Set RsOld = ew_CloneRs(RsDelete)

		' Call row deleting event
		If DeleteRows Then
			RsDelete.MoveFirst
			Do While Not RsDelete.Eof
				DeleteRows = Users.Row_Deleting(RsDelete)
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
				sThisKey = sThisKey & RsDelete("Username")
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
			If Users.CancelMessage <> "" Then
				FailureMessage = Users.CancelMessage
				Users.CancelMessage = ""
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
				Call WriteAuditTrailDummy(Language.Phrase("BatchDeleteSuccess")) ' Batch delete success
			End If
		Else
			Conn.RollbackTrans ' Rollback changes
			Call WriteAuditTrailDummy(Language.Phrase("BatchDeleteRollback")) ' Batch delete rollback
		End If
		RsDelete.Close
		Set RsDelete = Nothing

		' Call row deleting event
		If DeleteRows Then
			If Not RsOld.Eof Then RsOld.MoveFirst
			Do While Not RsOld.Eof
				Call Users.Row_Deleted(RsOld)
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
		table = "Users"

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
		table = "Users"

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
		sKey = sKey & RsSrc.Fields("Username")
		keyvalue = sKey

		' Username Field
		oldvalue = RsSrc("Username")
		Call ew_WriteAuditTrail(filePfx, curDateTime, id, user, action, table, "Username", keyvalue, oldvalue, newvalue)

		' Password Field
		oldvalue = RsSrc("Password")
		Call ew_WriteAuditTrail(filePfx, curDateTime, id, user, action, table, "Password", keyvalue, oldvalue, newvalue)

		' Email Field
		oldvalue = RsSrc("Email")
		Call ew_WriteAuditTrail(filePfx, curDateTime, id, user, action, table, "Email", keyvalue, oldvalue, newvalue)

		' Permissions Field
		oldvalue = RsSrc("Permissions")
		Call ew_WriteAuditTrail(filePfx, curDateTime, id, user, action, table, "Permissions", keyvalue, oldvalue, newvalue)

		' Active Field
		oldvalue = RsSrc("Active")
		Call ew_WriteAuditTrail(filePfx, curDateTime, id, user, action, table, "Active", keyvalue, oldvalue, newvalue)

		' Profile Field
		oldvalue = RsSrc("Profile")
		If Not EW_AUDIT_TRAIL_TO_DATABASE Then oldvalue = "[MEMO]"
		Call ew_WriteAuditTrail(filePfx, curDateTime, id, user, action, table, "Profile", keyvalue, oldvalue, newvalue)

		' Theme Field
		oldvalue = RsSrc("Theme")
		Call ew_WriteAuditTrail(filePfx, curDateTime, id, user, action, table, "Theme", keyvalue, oldvalue, newvalue)
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
