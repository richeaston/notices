<%@ CodePage="1252" LCID="2057" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="ewcfg9.asp"-->
<!--#include file="groupsinfo.asp"-->
<!--#include file="usersinfo.asp"-->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim Groups_edit
Set Groups_edit = New cGroups_edit
Set Page = Groups_edit

' Page init processing
Call Groups_edit.Page_Init()

' Page main processing
Call Groups_edit.Page_Main()
%>
<!--#include file="header.asp"-->
<script type="text/javascript">
<!--
// Create page object
var Groups_edit = new ew_Page("Groups_edit");
// page properties
Groups_edit.PageID = "edit"; // page ID
Groups_edit.FormID = "fGroupsedit"; // form ID
var EW_PAGE_ID = Groups_edit.PageID; // for backward compatibility
// extend page with ValidateForm function
Groups_edit.ValidateForm = function(fobj) {
	ew_PostAutoSuggest(fobj);
	if (!this.ValidateRequired)
		return true; // ignore validation
	if (fobj.a_confirm && fobj.a_confirm.value == "F")
		return true;
	var i, elm, aelm, infix;
	var rowcnt = 1;
	for (i=0; i<rowcnt; i++) {
		infix = "";
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
Groups_edit.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
<% If EW_CLIENT_VALIDATE Then %>
Groups_edit.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Groups_edit.ValidateRequired = false; // no JavaScript validation
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
<% Groups_edit.ShowPageHeader() %>
<p class="aspmaker ewTitle"><%= Language.Phrase("Edit") %>&nbsp;<%= Language.Phrase("TblTypeTABLE") %><%= Groups.TableCaption %></p>
<p class="aspmaker"><a href="<%= Groups.ReturnUrl %>"><%= Language.Phrase("GoBack") %></a></p>
<% Groups_edit.ShowMessage %>
<form name="fGroupsedit" id="fGroupsedit" action="<%= ew_CurrentPage %>" method="post" onsubmit="return Groups_edit.ValidateForm(this);">
<p>
<input type="hidden" name="a_table" id="a_table" value="Groups">
<input type="hidden" name="a_edit" id="a_edit" value="U">
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table cellspacing="0" class="ewTable">
<% If Groups.Group.Visible Then ' Group %>
	<tr id="r_Group"<%= Groups.RowAttributes %>>
		<td class="ewTableHeader"><%= Groups.Group.FldCaption %></td>
		<td<%= Groups.Group.CellAttributes %>><span id="el_Group">
<div<%= Groups.Group.ViewAttributes %>><%= Groups.Group.EditValue %></div>
<input type="hidden" name="x_Group" id="x_Group" value="<%= Server.HTMLEncode(Groups.Group.CurrentValue&"") %>">
</span><%= Groups.Group.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</div>
</td></tr></table>
<p>
<input type="submit" name="btnAction" id="btnAction" value="<%= ew_BtnCaption(Language.Phrase("EditBtn")) %>">
</form>
<%
Groups_edit.ShowPageFooter()
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
Set Groups_edit = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cGroups_edit

	' Page ID
	Public Property Get PageID()
		PageID = "edit"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Groups"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Groups_edit"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If Groups.UseTokenInUrl Then PageUrl = PageUrl & "t=" & Groups.TableVar & "&" ' add page token
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
		If Groups.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (Groups.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (Groups.TableVar = Request.QueryString("t"))
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
		If IsEmpty(Groups) Then Set Groups = New cGroups
		Set Table = Groups

		' Initialize urls
		' Initialize other table object

		If IsEmpty(Users) Then Set Users = New cUsers

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "edit"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "Groups"

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
		If Not Security.CanEdit Then
			Call Security.SaveLastUrl()
			Call Page_Terminate("groupslist.asp")
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
		Set Groups = Nothing
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
		If Request.QueryString("Group").Count > 0 Then
			Groups.Group.QueryStringValue = Request.QueryString("Group")
		End If
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			Groups.CurrentAction = ObjForm.GetValue("a_edit") ' Get action code
			Call LoadFormValues() ' Get form values

			' Validate Form
			If Not ValidateForm() Then
				Groups.CurrentAction = "" ' Form error, reset action
				FailureMessage = gsFormError
				Groups.EventCancelled = True ' Event cancelled
				Call LoadRow() ' Restore row
				Call RestoreFormValues() ' Restore form values if validate failed
			End If
		Else
			Groups.CurrentAction = "I" ' Default action is display
		End If

		' Check if valid key
		If Groups.Group.CurrentValue = "" Then Call Page_Terminate("groupslist.asp") ' Invalid key, return to list
		Select Case Groups.CurrentAction
			Case "I" ' Get a record to display
				If Not LoadRow() Then ' Load Record based on key
					FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("groupslist.asp") ' No matching record, return to list
				End If
			Case "U" ' Update
				Groups.SendEmail = True ' Send email on update success
				If EditRow() Then ' Update Record based on key
					SuccessMessage = Language.Phrase("UpdateSuccess") ' Update success
					Dim sReturnUrl
					sReturnUrl = Groups.ReturnUrl
					Call Page_Terminate(sReturnUrl) ' Return to caller
				Else
					Groups.EventCancelled = True ' Event cancelled
					Call LoadRow() ' Restore row
					Call RestoreFormValues() ' Restore form values if update failed
				End If
		End Select

		' Render the record
		Groups.RowType = EW_ROWTYPE_EDIT ' Render as edit

		' Render row
		Call Groups.ResetAttrs()
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
		If Not Groups.Group.FldIsDetailKey Then Groups.Group.FormValue = ObjForm.GetValue("x_Group")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadRow()
		Groups.Group.CurrentValue = Groups.Group.FormValue
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = Groups.KeyFilter

		' Call Row Selecting event
		Call Groups.Row_Selecting(sFilter)

		' Load sql based on filter
		Groups.CurrentFilter = sFilter
		sSql = Groups.SQL
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
		Call Groups.Row_Selected(RsRow)
		Groups.Group.DbValue = RsRow("Group")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Call Row Rendering event

		Call Groups.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' Group
		' -----------
		'  View  Row
		' -----------

		If Groups.RowType = EW_ROWTYPE_VIEW Then ' View row

			' Group
			Groups.Group.ViewValue = Groups.Group.CurrentValue
			Groups.Group.ViewCustomAttributes = ""

			' View refer script
			' Group

			Groups.Group.LinkCustomAttributes = ""
			Groups.Group.HrefValue = ""
			Groups.Group.TooltipValue = ""

		' ----------
		'  Edit Row
		' ----------

		ElseIf Groups.RowType = EW_ROWTYPE_EDIT Then ' Edit row

			' Group
			Groups.Group.EditCustomAttributes = ""
			Groups.Group.EditValue = Groups.Group.CurrentValue
			Groups.Group.ViewCustomAttributes = ""

			' Edit refer script
			' Group

			Groups.Group.HrefValue = ""
		End If
		If Groups.RowType = EW_ROWTYPE_ADD Or Groups.RowType = EW_ROWTYPE_EDIT Or Groups.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call Groups.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If Groups.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Groups.Row_Rendered()
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
		sFilter = Groups.KeyFilter
		Groups.CurrentFilter  = sFilter
		sSql = Groups.SQL
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

			' Field Group
			' Check recordset update error

			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				Rs.Close
				Set Rs = Nothing
				EditRow = False
				Exit Function
			End If

			' Call Row Updating event
			bUpdateRow = Groups.Row_Updating(RsOld, Rs)
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
				If Groups.CancelMessage <> "" Then
					FailureMessage = Groups.CancelMessage
					Groups.CancelMessage = ""
				Else
					FailureMessage = Language.Phrase("UpdateCancelled")
				End If
				EditRow = False
			End If
		End If

		' Call Row Updated event
		If EditRow Then
			Call Groups.Row_Updated(RsOld, RsNew)
		End If
		If EditRow Then
			Call WriteAuditTrailOnEdit(RsOld, RsNew)
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

	' Write Audit Trail start/end for grid update
	Sub WriteAuditTrailDummy(typ)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		Dim table
		table = "Groups"

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

	' Write Audit Trail (edit page)
	Sub WriteAuditTrailOnEdit(RsOld, RsNew)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		Dim table
		table = "Groups"

		' Get key value
		Dim sKey
		sKey = ""
		If sKey <> "" Then sKey = sKey & EW_COMPOSITE_KEY_SEPARATOR
		sKey = sKey & RsNew.Fields("Group")

		' Write Audit Trail
		Dim filePfx, curDateTime, id, user, action, field, keyvalue, oldvalue, newvalue
		Dim i
		filePfx = "log"
		curDateTime = ew_StdCurrentDateTime()
		id = Request.ServerVariables("SCRIPT_NAME")
    	user = CurrentUserName
		action = "U"
		For i = 0 to RsOld.Fields.Count - 1
			If RsOld.Fields(i).Type <> 205 Then ' Ignore Blob Field
				oldvalue = ew_Conv(RsOld.Fields(i).Value, RsOld.Fields(i).Type)
				newvalue = ew_Conv(RsNew.Fields(i).Value, RsNew.Fields(i).Type)
				If Not ew_CompareValue(oldvalue, newvalue) Then
					field = RsOld.Fields(i).Name
					keyvalue = sKey
					If (RsOld.Fields(i).Type = 201 Or RsOld.Fields(i).Type = 203) And Not EW_AUDIT_TRAIL_TO_DATABASE Then
						oldvalue = "[MEMO]"
						newvalue = "[MEMO]"
					End If
					Call ew_WriteAuditTrail(filePfx, curDateTime, id, user, action, table, field, keyvalue, oldvalue, newvalue)
				End If
			End If
		Next
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

	' Form Custom Validate event
	Function Form_CustomValidate(CustomError)

		'Return error message in CustomError
		Form_CustomValidate = True
	End Function
End Class
%>
