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
Dim UserLevels_edit
Set UserLevels_edit = New cUserLevels_edit
Set Page = UserLevels_edit

' Page init processing
Call UserLevels_edit.Page_Init()

' Page main processing
Call UserLevels_edit.Page_Main()
%>
<!--#include file="header.asp"-->
<script type="text/javascript">
<!--
// Create page object
var UserLevels_edit = new ew_Page("UserLevels_edit");
// page properties
UserLevels_edit.PageID = "edit"; // page ID
UserLevels_edit.FormID = "fUserLevelsedit"; // form ID
var EW_PAGE_ID = UserLevels_edit.PageID; // for backward compatibility
// extend page with ValidateForm function
UserLevels_edit.ValidateForm = function(fobj) {
	ew_PostAutoSuggest(fobj);
	if (!this.ValidateRequired)
		return true; // ignore validation
	if (fobj.a_confirm && fobj.a_confirm.value == "F")
		return true;
	var i, elm, aelm, infix;
	var rowcnt = 1;
	for (i=0; i<rowcnt; i++) {
		infix = "";
		elm = fobj.elements["x" + infix + "_UserLevelID"];
		if (elm && !ew_HasValue(elm))
			return ew_OnError(this, elm, ewLanguage.Phrase("EnterRequiredField") + " - <%= ew_JsEncode2(UserLevels.UserLevelID.FldCaption) %>");
		elm = fobj.elements["x" + infix + "_UserLevelID"];
		if (elm && !ew_CheckInteger(elm.value))
			return ew_OnError(this, elm, "<%= ew_JsEncode2(UserLevels.UserLevelID.FldErrMsg) %>");
		elm = fobj.elements["x" + infix + "_UserLevelName"];
		if (elm && !ew_HasValue(elm))
			return ew_OnError(this, elm, ewLanguage.Phrase("EnterRequiredField") + " - <%= ew_JsEncode2(UserLevels.UserLevelName.FldCaption) %>");
		elmId = fobj.elements["x" + infix + "_UserLevelID"];
		elmName = fobj.elements["x" + infix + "_UserLevelName"];
		if (elmId && elmName) {
			elmId.value = elmId.value.replace(/^\s+|\s+$/, '');
			elmName.value = elmName.value.replace(/^\s+|\s+$/, '');
			if (elmId && !ew_CheckInteger(elmId.value))
				return ew_OnError(this, elmId, ewLanguage.Phrase("UserLevelIDInteger"));
			var level = parseInt(elmId.value);
			if (level == 0) {
				if (elmName.value.toLowerCase() != "default")
					return ew_OnError(this, elmName, ewLanguage.Phrase("UserLevelDefaultName"));
			} else if (level == -1) { 
				if (elmName.value.toLowerCase() != "administrator")
					return ew_OnError(this, elmName, ewLanguage.Phrase("UserLevelAdministratorName"));
			} else if (level < -1) {
				return ew_OnError(this, elmId, ewLanguage.Phrase("UserLevelIDIncorrect"));
			} else if (level > 0) { 
				if (elmName.value.toLowerCase() == "administrator" || elmName.value.toLowerCase() == "default")
					return ew_OnError(this, elmName, ewLanguage.Phrase("UserLevelNameIncorrect"));
			}
		}
		// Set up row object
		var row = {};
		row["index"] = infix;
		for (var j = 0; j < fobj.elements.length; j++) {
			var el = fobj.elements[j];
			var len = infix.length + 2;
			if (el.name.substr(0, len) == "x" + infix + "_") {
				var elname = "x_" + el.name.substr(len);
				if (ewLang.isObject(row[elname])) { // already exists
					if (ewLang.isArray(row[elname])) {
						row[elname][row[elname].length] = el; // add to array
					} else {
						row[elname] = [row[elname], el]; // convert to array
					}
				} else {
					row[elname] = el;
				}
			}
		}
		fobj.row = row;
		// Call Form Custom Validate event
		if (!this.Form_CustomValidate(fobj)) return false;
	}
	// Process detail page
	var detailpage = (fobj.detailpage) ? fobj.detailpage.value : "";
	if (detailpage != "") {
		return eval(detailpage+".ValidateForm(fobj)");
	}
	return true;
}
// extend page with Form_CustomValidate function
UserLevels_edit.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
UserLevels_edit.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
UserLevels_edit.ValidateRequired = true; // uses JavaScript validation
<% Else %>
UserLevels_edit.ValidateRequired = false; // no JavaScript validation
<% End If %>
//-->
</script>
<script type="text/javascript">
<!--
var ew_DHTMLEditors = [];
//-->
</script>
<script language="JavaScript" type="text/javascript">
<!--
// Write your client script here, no need to add script tags.
//-->
</script>
<% UserLevels_edit.ShowPageHeader() %>
<p class="aspmaker ewTitle"><%= Language.Phrase("Edit") %>&nbsp;<%= Language.Phrase("TblTypeTABLE") %><%= UserLevels.TableCaption %></p>
<p class="aspmaker"><a href="<%= UserLevels.ReturnUrl %>"><%= Language.Phrase("GoBack") %></a></p>
<% UserLevels_edit.ShowMessage %>
<form name="fUserLevelsedit" id="fUserLevelsedit" action="<%= ew_CurrentPage %>" method="post" onsubmit="return UserLevels_edit.ValidateForm(this);">
<p>
<input type="hidden" name="a_table" id="a_table" value="UserLevels">
<input type="hidden" name="a_edit" id="a_edit" value="U">
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table cellspacing="0" class="ewTable">
<% If UserLevels.UserLevelID.Visible Then ' UserLevelID %>
	<tr id="r_UserLevelID"<%= UserLevels.RowAttributes %>>
		<td class="ewTableHeader"><%= UserLevels.UserLevelID.FldCaption %><%= Language.Phrase("FieldRequiredIndicator") %></td>
		<td<%= UserLevels.UserLevelID.CellAttributes %>><span id="el_UserLevelID">
<div<%= UserLevels.UserLevelID.ViewAttributes %>><%= UserLevels.UserLevelID.EditValue %></div>
<input type="hidden" name="x_UserLevelID" id="x_UserLevelID" value="<%= Server.HTMLEncode(UserLevels.UserLevelID.CurrentValue&"") %>">
</span><%= UserLevels.UserLevelID.CustomMsg %></td>
	</tr>
<% End If %>
<% If UserLevels.UserLevelName.Visible Then ' UserLevelName %>
	<tr id="r_UserLevelName"<%= UserLevels.RowAttributes %>>
		<td class="ewTableHeader"><%= UserLevels.UserLevelName.FldCaption %><%= Language.Phrase("FieldRequiredIndicator") %></td>
		<td<%= UserLevels.UserLevelName.CellAttributes %>><span id="el_UserLevelName">
<input type="text" name="x_UserLevelName" id="x_UserLevelName" size="30" maxlength="255" value="<%= UserLevels.UserLevelName.EditValue %>"<%= UserLevels.UserLevelName.EditAttributes %>>
</span><%= UserLevels.UserLevelName.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</div>
</td></tr></table>
<p>
<input type="submit" name="btnAction" id="btnAction" value="<%= ew_BtnCaption(Language.Phrase("EditBtn")) %>">
</form>
<%
UserLevels_edit.ShowPageFooter()
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
Set UserLevels_edit = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cUserLevels_edit

	' Page ID
	Public Property Get PageID()
		PageID = "edit"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "UserLevels"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "UserLevels_edit"
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
		' Initialize other table object

		If IsEmpty(Users) Then Set Users = New cUsers

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "edit"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "UserLevels"

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

	' Create form object
	Set ObjForm = New cFormObj

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

	Dim DbMasterFilter, DbDetailFilter

	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()

		' Load key from QueryString
		If Request.QueryString("UserLevelID").Count > 0 Then
			UserLevels.UserLevelID.QueryStringValue = Request.QueryString("UserLevelID")
		End If
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			UserLevels.CurrentAction = ObjForm.GetValue("a_edit") ' Get action code
			Call LoadFormValues() ' Get form values

			' Validate Form
			If Not ValidateForm() Then
				UserLevels.CurrentAction = "" ' Form error, reset action
				FailureMessage = gsFormError
				UserLevels.EventCancelled = True ' Event cancelled
				Call LoadRow() ' Restore row
				Call RestoreFormValues() ' Restore form values if validate failed
			End If
		Else
			UserLevels.CurrentAction = "I" ' Default action is display
		End If

		' Check if valid key
		If UserLevels.UserLevelID.CurrentValue = "" Then Call Page_Terminate("userlevelslist.asp") ' Invalid key, return to list
		Select Case UserLevels.CurrentAction
			Case "I" ' Get a record to display
				If Not LoadRow() Then ' Load Record based on key
					FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("userlevelslist.asp") ' No matching record, return to list
				End If
			Case "U" ' Update
				UserLevels.SendEmail = True ' Send email on update success
				If EditRow() Then ' Update Record based on key
					SuccessMessage = Language.Phrase("UpdateSuccess") ' Update success
					Dim sReturnUrl
					sReturnUrl = UserLevels.ReturnUrl
					Call Page_Terminate(sReturnUrl) ' Return to caller
				Else
					UserLevels.EventCancelled = True ' Event cancelled
					Call LoadRow() ' Restore row
					Call RestoreFormValues() ' Restore form values if update failed
				End If
		End Select

		' Render the record
		UserLevels.RowType = EW_ROWTYPE_EDIT ' Render as edit

		' Render row
		Call UserLevels.ResetAttrs()
		Call RenderRow()
	End Sub

	' -----------------------------------------------------------------
	' Function Get upload files
	'
	Function GetUploadFiles()

		' Get upload data
		Dim index, confirmPage
		index = ObjForm.Index ' Save form index
		ObjForm.Index = 0
		confirmPage = (ObjForm.GetValue("a_confirm") & "" <> "")
		ObjForm.Index = index ' Restore form index
	End Function

	' -----------------------------------------------------------------
	' Load form values
	'
	Function LoadFormValues()

		' Load values from form
		If Not UserLevels.UserLevelID.FldIsDetailKey Then UserLevels.UserLevelID.FormValue = ObjForm.GetValue("x_UserLevelID")
		If Not UserLevels.UserLevelName.FldIsDetailKey Then UserLevels.UserLevelName.FormValue = ObjForm.GetValue("x_UserLevelName")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadRow()
		UserLevels.UserLevelID.CurrentValue = UserLevels.UserLevelID.FormValue
		UserLevels.UserLevelName.CurrentValue = UserLevels.UserLevelName.FormValue
	End Function

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

		' ----------
		'  Edit Row
		' ----------

		ElseIf UserLevels.RowType = EW_ROWTYPE_EDIT Then ' Edit row

			' UserLevelID
			UserLevels.UserLevelID.EditCustomAttributes = ""
			UserLevels.UserLevelID.EditValue = UserLevels.UserLevelID.CurrentValue
			UserLevels.UserLevelID.ViewCustomAttributes = ""

			' UserLevelName
			UserLevels.UserLevelName.EditCustomAttributes = ""
			UserLevels.UserLevelName.EditValue = ew_HtmlEncode(UserLevels.UserLevelName.CurrentValue)

			' Edit refer script
			' UserLevelID

			UserLevels.UserLevelID.HrefValue = ""

			' UserLevelName
			UserLevels.UserLevelName.HrefValue = ""
		End If
		If UserLevels.RowType = EW_ROWTYPE_ADD Or UserLevels.RowType = EW_ROWTYPE_EDIT Or UserLevels.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call UserLevels.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If UserLevels.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call UserLevels.Row_Rendered()
		End If
	End Sub

	' -----------------------------------------------------------------
	' Validate form
	'
	Function ValidateForm()

		' Initialize
		gsFormError = ""

		' Check if validation required
		If Not EW_SERVER_VALIDATE Then
			ValidateForm = (gsFormError = "")
			Exit Function
		End If
		If Not IsNull(UserLevels.UserLevelID.FormValue) And UserLevels.UserLevelID.FormValue&"" = "" Then
			Call ew_AddMessage(gsFormError, Language.Phrase("EnterRequiredField") & " - " & UserLevels.UserLevelID.FldCaption)
		End If
		If Not ew_CheckInteger(UserLevels.UserLevelID.FormValue) Then
			Call ew_AddMessage(gsFormError, UserLevels.UserLevelID.FldErrMsg)
		End If
		If Not IsNull(UserLevels.UserLevelName.FormValue) And UserLevels.UserLevelName.FormValue&"" = "" Then
			Call ew_AddMessage(gsFormError, Language.Phrase("EnterRequiredField") & " - " & UserLevels.UserLevelName.FldCaption)
		End If

		' Return validate result
		ValidateForm = (gsFormError = "")

		' Call Form Custom Validate event
		Dim sFormCustomError
		sFormCustomError = ""
		ValidateForm = ValidateForm And Form_CustomValidate(sFormCustomError)
		If sFormCustomError <> "" Then
			Call ew_AddMessage(gsFormError, sFormCustomError)
		End If
	End Function

	' -----------------------------------------------------------------
	' Update record based on key values
	'
	Function EditRow()
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		Dim Rs, sSql, sFilter
		Dim RsChk, sSqlChk, sFilterChk
		Dim bUpdateRow
		Dim RsOld, RsNew
		Dim sIdxErrMsg

		' Clear any previous errors
		Err.Clear
		sFilter = UserLevels.KeyFilter
		UserLevels.CurrentFilter  = sFilter
		sSql = UserLevels.SQL
		Set Rs = Server.CreateObject("ADODB.Recordset")
		Rs.CursorLocation = EW_CURSORLOCATION
		Rs.Open sSql, Conn, 1, EW_RECORDSET_LOCKTYPE
		If Err.Number <> 0 Then
			Message = Err.Description
			Rs.Close
			Set Rs = Nothing
			EditRow = False
			Exit Function
		End If

		' Clone old recordset object
		Set RsOld = ew_CloneRs(Rs)
		If Rs.Eof Then
			EditRow = False ' Update Failed
		Else

			' Field UserLevelID
			' Field UserLevelName

			Call UserLevels.UserLevelName.SetDbValue(Rs, UserLevels.UserLevelName.CurrentValue, "", UserLevels.UserLevelName.ReadOnly)

			' Check recordset update error
			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				Rs.Close
				Set Rs = Nothing
				EditRow = False
				Exit Function
			End If

			' Call Row Updating event
			bUpdateRow = UserLevels.Row_Updating(RsOld, Rs)
			If bUpdateRow Then

				' Clone new recordset object
				Set RsNew = ew_CloneRs(Rs)
				Rs.Update
				If Err.Number <> 0 Then
					FailureMessage = Err.Description
					EditRow = False
				Else
					EditRow = True
				End If
			Else
				Rs.CancelUpdate
				If UserLevels.CancelMessage <> "" Then
					FailureMessage = UserLevels.CancelMessage
					UserLevels.CancelMessage = ""
				Else
					FailureMessage = Language.Phrase("UpdateCancelled")
				End If
				EditRow = False
			End If
		End If

		' Call Row Updated event
		If EditRow Then
			Call UserLevels.Row_Updated(RsOld, RsNew)
		End If
		Rs.Close
		Set Rs = Nothing
		If IsObject(RsOld) Then
			RsOld.Close
			Set RsOld = Nothing
		End If
		If IsObject(RsNew) Then
			RsNew.Close
			Set RsNew = Nothing
		End If
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

	' Form Custom Validate event
	Function Form_CustomValidate(CustomError)

		'Return error message in CustomError
		Form_CustomValidate = True
	End Function
End Class
%>
