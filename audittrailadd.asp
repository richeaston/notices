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
Dim AuditTrail_add
Set AuditTrail_add = New cAuditTrail_add
Set Page = AuditTrail_add

' Page init processing
Call AuditTrail_add.Page_Init()

' Page main processing
Call AuditTrail_add.Page_Main()
%>
<!--#include file="header.asp"-->
<script type="text/javascript">
<!--
// Create page object
var AuditTrail_add = new ew_Page("AuditTrail_add");
// page properties
AuditTrail_add.PageID = "add"; // page ID
AuditTrail_add.FormID = "fAuditTrailadd"; // form ID
var EW_PAGE_ID = AuditTrail_add.PageID; // for backward compatibility
// extend page with ValidateForm function
AuditTrail_add.ValidateForm = function(fobj) {
	ew_PostAutoSuggest(fobj);
	if (!this.ValidateRequired)
		return true; // ignore validation
	if (fobj.a_confirm && fobj.a_confirm.value == "F")
		return true;
	var i, elm, aelm, infix;
	var rowcnt = 1;
	for (i=0; i<rowcnt; i++) {
		infix = "";
		elm = fobj.elements["x" + infix + "_DateTime"];
		if (elm && !ew_HasValue(elm))
			return ew_OnError(this, elm, ewLanguage.Phrase("EnterRequiredField") + " - <%= ew_JsEncode2(AuditTrail.DateTime.FldCaption) %>");
		elm = fobj.elements["x" + infix + "_DateTime"];
		if (elm && !ew_CheckEuroDate(elm.value))
			return ew_OnError(this, elm, "<%= ew_JsEncode2(AuditTrail.DateTime.FldErrMsg) %>");
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
AuditTrail_add.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
AuditTrail_add.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
AuditTrail_add.ValidateRequired = true; // uses JavaScript validation
<% Else %>
AuditTrail_add.ValidateRequired = false; // no JavaScript validation
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
<% AuditTrail_add.ShowPageHeader() %>
<p class="aspmaker ewTitle"><%= Language.Phrase("Add") %>&nbsp;<%= Language.Phrase("TblTypeTABLE") %><%= AuditTrail.TableCaption %></p>
<p class="aspmaker"><a href="<%= AuditTrail.ReturnUrl %>"><%= Language.Phrase("GoBack") %></a></p>
<% AuditTrail_add.ShowMessage %>
<form name="fAuditTrailadd" id="fAuditTrailadd" action="<%= ew_CurrentPage %>" method="post" onsubmit="return AuditTrail_add.ValidateForm(this);">
<p>
<input type="hidden" name="t" id="t" value="AuditTrail">
<input type="hidden" name="a_add" id="a_add" value="A">
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table cellspacing="0" class="ewTable">
<% If AuditTrail.DateTime.Visible Then ' DateTime %>
	<tr id="r_DateTime"<%= AuditTrail.RowAttributes %>>
		<td class="ewTableHeader"><%= AuditTrail.DateTime.FldCaption %><%= Language.Phrase("FieldRequiredIndicator") %></td>
		<td<%= AuditTrail.DateTime.CellAttributes %>><span id="el_DateTime">
<input type="text" name="x_DateTime" id="x_DateTime" value="<%= AuditTrail.DateTime.EditValue %>"<%= AuditTrail.DateTime.EditAttributes %>>
</span><%= AuditTrail.DateTime.CustomMsg %></td>
	</tr>
<% End If %>
<% If AuditTrail.Script.Visible Then ' Script %>
	<tr id="r_Script"<%= AuditTrail.RowAttributes %>>
		<td class="ewTableHeader"><%= AuditTrail.Script.FldCaption %></td>
		<td<%= AuditTrail.Script.CellAttributes %>><span id="el_Script">
<input type="text" name="x_Script" id="x_Script" size="30" maxlength="255" value="<%= AuditTrail.Script.EditValue %>"<%= AuditTrail.Script.EditAttributes %>>
</span><%= AuditTrail.Script.CustomMsg %></td>
	</tr>
<% End If %>
<% If AuditTrail.User.Visible Then ' User %>
	<tr id="r_User"<%= AuditTrail.RowAttributes %>>
		<td class="ewTableHeader"><%= AuditTrail.User.FldCaption %></td>
		<td<%= AuditTrail.User.CellAttributes %>><span id="el_User">
<input type="text" name="x_User" id="x_User" size="30" maxlength="255" value="<%= AuditTrail.User.EditValue %>"<%= AuditTrail.User.EditAttributes %>>
</span><%= AuditTrail.User.CustomMsg %></td>
	</tr>
<% End If %>
<% If AuditTrail.Action.Visible Then ' Action %>
	<tr id="r_Action"<%= AuditTrail.RowAttributes %>>
		<td class="ewTableHeader"><%= AuditTrail.Action.FldCaption %></td>
		<td<%= AuditTrail.Action.CellAttributes %>><span id="el_Action">
<input type="text" name="x_Action" id="x_Action" size="30" maxlength="255" value="<%= AuditTrail.Action.EditValue %>"<%= AuditTrail.Action.EditAttributes %>>
</span><%= AuditTrail.Action.CustomMsg %></td>
	</tr>
<% End If %>
<% If AuditTrail.zTable.Visible Then ' Table %>
	<tr id="r_zTable"<%= AuditTrail.RowAttributes %>>
		<td class="ewTableHeader"><%= AuditTrail.zTable.FldCaption %></td>
		<td<%= AuditTrail.zTable.CellAttributes %>><span id="el_zTable">
<input type="text" name="x_zTable" id="x_zTable" size="30" maxlength="255" value="<%= AuditTrail.zTable.EditValue %>"<%= AuditTrail.zTable.EditAttributes %>>
</span><%= AuditTrail.zTable.CustomMsg %></td>
	</tr>
<% End If %>
<% If AuditTrail.zField.Visible Then ' Field %>
	<tr id="r_zField"<%= AuditTrail.RowAttributes %>>
		<td class="ewTableHeader"><%= AuditTrail.zField.FldCaption %></td>
		<td<%= AuditTrail.zField.CellAttributes %>><span id="el_zField">
<input type="text" name="x_zField" id="x_zField" size="30" maxlength="255" value="<%= AuditTrail.zField.EditValue %>"<%= AuditTrail.zField.EditAttributes %>>
</span><%= AuditTrail.zField.CustomMsg %></td>
	</tr>
<% End If %>
<% If AuditTrail.KeyValue.Visible Then ' KeyValue %>
	<tr id="r_KeyValue"<%= AuditTrail.RowAttributes %>>
		<td class="ewTableHeader"><%= AuditTrail.KeyValue.FldCaption %></td>
		<td<%= AuditTrail.KeyValue.CellAttributes %>><span id="el_KeyValue">
<textarea name="x_KeyValue" id="x_KeyValue" cols="35" rows="4"<%= AuditTrail.KeyValue.EditAttributes %>><%= AuditTrail.KeyValue.EditValue %></textarea>
</span><%= AuditTrail.KeyValue.CustomMsg %></td>
	</tr>
<% End If %>
<% If AuditTrail.OldValue.Visible Then ' OldValue %>
	<tr id="r_OldValue"<%= AuditTrail.RowAttributes %>>
		<td class="ewTableHeader"><%= AuditTrail.OldValue.FldCaption %></td>
		<td<%= AuditTrail.OldValue.CellAttributes %>><span id="el_OldValue">
<textarea name="x_OldValue" id="x_OldValue" cols="35" rows="4"<%= AuditTrail.OldValue.EditAttributes %>><%= AuditTrail.OldValue.EditValue %></textarea>
</span><%= AuditTrail.OldValue.CustomMsg %></td>
	</tr>
<% End If %>
<% If AuditTrail.NewValue.Visible Then ' NewValue %>
	<tr id="r_NewValue"<%= AuditTrail.RowAttributes %>>
		<td class="ewTableHeader"><%= AuditTrail.NewValue.FldCaption %></td>
		<td<%= AuditTrail.NewValue.CellAttributes %>><span id="el_NewValue">
<textarea name="x_NewValue" id="x_NewValue" cols="35" rows="4"<%= AuditTrail.NewValue.EditAttributes %>><%= AuditTrail.NewValue.EditValue %></textarea>
</span><%= AuditTrail.NewValue.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</div>
</td></tr></table>
<p>
<input type="submit" name="btnAction" id="btnAction" value="<%= ew_BtnCaption(Language.Phrase("AddBtn")) %>">
</form>
<%
AuditTrail_add.ShowPageFooter()
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
Set AuditTrail_add = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cAuditTrail_add

	' Page ID
	Public Property Get PageID()
		PageID = "add"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "AuditTrail"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "AuditTrail_add"
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
		EW_PAGE_ID = "add"

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
		If Not Security.CanAdd Then
			Call Security.SaveLastUrl()
			Call Page_Terminate("audittraillist.asp")
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

	Dim DbMasterFilter, DbDetailFilter
	Dim Priv
	Dim OldRecordset
	Dim CopyRecord

	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()

		' Process form if post back
		If ObjForm.GetValue("a_add")&"" <> "" Then
			AuditTrail.CurrentAction = ObjForm.GetValue("a_add") ' Get form action
			CopyRecord = LoadOldRecord() ' Load old recordset
			Call LoadFormValues() ' Load form values

			' Validate Form
			If Not ValidateForm() Then
				AuditTrail.CurrentAction = "I" ' Form error, reset action
				AuditTrail.EventCancelled = True ' Event cancelled
				Call RestoreFormValues() ' Restore form values
				FailureMessage = gsFormError
			End If

		' Not post back
		Else

			' Load key values from QueryString
			CopyRecord = True
			If Request.QueryString("Id").Count > 0 Then
				AuditTrail.Id.QueryStringValue = Request.QueryString("Id")
				Call AuditTrail.SetKey("Id", AuditTrail.Id.CurrentValue) ' Set up key
			Else
				Call AuditTrail.SetKey("Id", "") ' Clear key
				CopyRecord = False
			End If
			If CopyRecord Then
				AuditTrail.CurrentAction = "C" ' Copy Record
			Else
				AuditTrail.CurrentAction = "I" ' Display Blank Record
				Call LoadDefaultValues() ' Load default values
			End If
		End If

		' Perform action based on action code
		Select Case AuditTrail.CurrentAction
			Case "I" ' Blank record, no action required
			Case "C" ' Copy an existing record
				If Not LoadRow() Then ' Load record based on key
					FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("audittraillist.asp") ' No matching record, return to list
				End If
			Case "A" ' Add new record
				AuditTrail.SendEmail = True ' Send email on add success
				If AddRow(OldRecordset) Then ' Add successful
					SuccessMessage = Language.Phrase("AddSuccess") ' Set up success message
					Dim sReturnUrl
					sReturnUrl = AuditTrail.ReturnUrl
					If ew_GetPageName(sReturnUrl) = "audittrailview.asp" Then sReturnUrl = AuditTrail.ViewUrl ' View paging, return to view page with keyurl directly
					Call Page_Terminate(sReturnUrl) ' Clean up and return
				Else
					AuditTrail.EventCancelled = True ' Event cancelled
					Call RestoreFormValues() ' Add failed, restore form values
				End If
		End Select

		' Render row based on row type
		AuditTrail.RowType = EW_ROWTYPE_ADD ' Render add type

		' Render row
		Call AuditTrail.ResetAttrs()
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
	' Load default values
	'
	Function LoadDefaultValues()
		AuditTrail.DateTime.CurrentValue = Null
		AuditTrail.DateTime.OldValue = AuditTrail.DateTime.CurrentValue
		AuditTrail.Script.CurrentValue = Null
		AuditTrail.Script.OldValue = AuditTrail.Script.CurrentValue
		AuditTrail.User.CurrentValue = Null
		AuditTrail.User.OldValue = AuditTrail.User.CurrentValue
		AuditTrail.Action.CurrentValue = Null
		AuditTrail.Action.OldValue = AuditTrail.Action.CurrentValue
		AuditTrail.zTable.CurrentValue = Null
		AuditTrail.zTable.OldValue = AuditTrail.zTable.CurrentValue
		AuditTrail.zField.CurrentValue = Null
		AuditTrail.zField.OldValue = AuditTrail.zField.CurrentValue
		AuditTrail.KeyValue.CurrentValue = Null
		AuditTrail.KeyValue.OldValue = AuditTrail.KeyValue.CurrentValue
		AuditTrail.OldValue.CurrentValue = Null
		AuditTrail.OldValue.OldValue = AuditTrail.OldValue.CurrentValue
		AuditTrail.NewValue.CurrentValue = Null
		AuditTrail.NewValue.OldValue = AuditTrail.NewValue.CurrentValue
	End Function

	' -----------------------------------------------------------------
	' Load form values
	'
	Function LoadFormValues()

		' Load values from form
		If Not AuditTrail.DateTime.FldIsDetailKey Then AuditTrail.DateTime.FormValue = ObjForm.GetValue("x_DateTime")
		If Not AuditTrail.DateTime.FldIsDetailKey Then AuditTrail.DateTime.CurrentValue = ew_UnFormatDateTime(AuditTrail.DateTime.CurrentValue, 7)
		If Not AuditTrail.Script.FldIsDetailKey Then AuditTrail.Script.FormValue = ObjForm.GetValue("x_Script")
		If Not AuditTrail.User.FldIsDetailKey Then AuditTrail.User.FormValue = ObjForm.GetValue("x_User")
		If Not AuditTrail.Action.FldIsDetailKey Then AuditTrail.Action.FormValue = ObjForm.GetValue("x_Action")
		If Not AuditTrail.zTable.FldIsDetailKey Then AuditTrail.zTable.FormValue = ObjForm.GetValue("x_zTable")
		If Not AuditTrail.zField.FldIsDetailKey Then AuditTrail.zField.FormValue = ObjForm.GetValue("x_zField")
		If Not AuditTrail.KeyValue.FldIsDetailKey Then AuditTrail.KeyValue.FormValue = ObjForm.GetValue("x_KeyValue")
		If Not AuditTrail.OldValue.FldIsDetailKey Then AuditTrail.OldValue.FormValue = ObjForm.GetValue("x_OldValue")
		If Not AuditTrail.NewValue.FldIsDetailKey Then AuditTrail.NewValue.FormValue = ObjForm.GetValue("x_NewValue")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadOldRecord()
		AuditTrail.DateTime.CurrentValue = AuditTrail.DateTime.FormValue
		AuditTrail.DateTime.CurrentValue = ew_UnFormatDateTime(AuditTrail.DateTime.CurrentValue, 7)
		AuditTrail.Script.CurrentValue = AuditTrail.Script.FormValue
		AuditTrail.User.CurrentValue = AuditTrail.User.FormValue
		AuditTrail.Action.CurrentValue = AuditTrail.Action.FormValue
		AuditTrail.zTable.CurrentValue = AuditTrail.zTable.FormValue
		AuditTrail.zField.CurrentValue = AuditTrail.zField.FormValue
		AuditTrail.KeyValue.CurrentValue = AuditTrail.KeyValue.FormValue
		AuditTrail.OldValue.CurrentValue = AuditTrail.OldValue.FormValue
		AuditTrail.NewValue.CurrentValue = AuditTrail.NewValue.FormValue
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

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If AuditTrail.GetKey("Id")&"" <> "" Then
			AuditTrail.Id.CurrentValue = AuditTrail.GetKey("Id") ' Id
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			AuditTrail.CurrentFilter = AuditTrail.KeyFilter
			Dim sSql
			sSql = AuditTrail.SQL
			Set OldRecordset = ew_LoadRecordset(sSql)
			Call LoadRowValues(OldRecordset) ' Load row values
		Else
			OldRecordset = Null
		End If
		LoadOldRecord = bValidKey
	End Function

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

		' ---------
		'  Add Row
		' ---------

		ElseIf AuditTrail.RowType = EW_ROWTYPE_ADD Then ' Add row

			' DateTime
			AuditTrail.DateTime.EditCustomAttributes = ""
			AuditTrail.DateTime.EditValue = ew_FormatDateTime(AuditTrail.DateTime.CurrentValue, 7)

			' Script
			AuditTrail.Script.EditCustomAttributes = ""
			AuditTrail.Script.EditValue = ew_HtmlEncode(AuditTrail.Script.CurrentValue)

			' User
			AuditTrail.User.EditCustomAttributes = ""
			AuditTrail.User.EditValue = ew_HtmlEncode(AuditTrail.User.CurrentValue)

			' Action
			AuditTrail.Action.EditCustomAttributes = ""
			AuditTrail.Action.EditValue = ew_HtmlEncode(AuditTrail.Action.CurrentValue)

			' Table
			AuditTrail.zTable.EditCustomAttributes = ""
			AuditTrail.zTable.EditValue = ew_HtmlEncode(AuditTrail.zTable.CurrentValue)

			' Field
			AuditTrail.zField.EditCustomAttributes = ""
			AuditTrail.zField.EditValue = ew_HtmlEncode(AuditTrail.zField.CurrentValue)

			' KeyValue
			AuditTrail.KeyValue.EditCustomAttributes = ""
			AuditTrail.KeyValue.EditValue = ew_HtmlEncode(AuditTrail.KeyValue.CurrentValue)

			' OldValue
			AuditTrail.OldValue.EditCustomAttributes = ""
			AuditTrail.OldValue.EditValue = ew_HtmlEncode(AuditTrail.OldValue.CurrentValue)

			' NewValue
			AuditTrail.NewValue.EditCustomAttributes = ""
			AuditTrail.NewValue.EditValue = ew_HtmlEncode(AuditTrail.NewValue.CurrentValue)

			' Edit refer script
			' DateTime

			AuditTrail.DateTime.HrefValue = ""

			' Script
			AuditTrail.Script.HrefValue = ""

			' User
			AuditTrail.User.HrefValue = ""

			' Action
			AuditTrail.Action.HrefValue = ""

			' Table
			AuditTrail.zTable.HrefValue = ""

			' Field
			AuditTrail.zField.HrefValue = ""

			' KeyValue
			AuditTrail.KeyValue.HrefValue = ""

			' OldValue
			AuditTrail.OldValue.HrefValue = ""

			' NewValue
			AuditTrail.NewValue.HrefValue = ""
		End If
		If AuditTrail.RowType = EW_ROWTYPE_ADD Or AuditTrail.RowType = EW_ROWTYPE_EDIT Or AuditTrail.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call AuditTrail.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If AuditTrail.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call AuditTrail.Row_Rendered()
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
		If Not IsNull(AuditTrail.DateTime.FormValue) And AuditTrail.DateTime.FormValue&"" = "" Then
			Call ew_AddMessage(gsFormError, Language.Phrase("EnterRequiredField") & " - " & AuditTrail.DateTime.FldCaption)
		End If
		If Not ew_CheckEuroDate(AuditTrail.DateTime.FormValue) Then
			Call ew_AddMessage(gsFormError, AuditTrail.DateTime.FldErrMsg)
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
	' Add record
	'
	Function AddRow(RsOld)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		Dim Rs, sSql, sFilter
		Dim RsNew
		Dim bInsertRow
		Dim RsChk
		Dim sIdxErrMsg

		' Clear any previous errors
		Err.Clear

		' Add new record
		sFilter = "(0 = 1)"
		AuditTrail.CurrentFilter = sFilter
		sSql = AuditTrail.SQL
		Set Rs = Server.CreateObject("ADODB.Recordset")
		Rs.CursorLocation = EW_CURSORLOCATION
		Rs.Open sSql, Conn, 1, EW_RECORDSET_LOCKTYPE
		Rs.AddNew
		If Err.Number <> 0 Then
			Message = Err.Description
			Rs.Close
			Set Rs = Nothing
			AddRow = False
			Exit Function
		End If

		' Field DateTime
		Call AuditTrail.DateTime.SetDbValue(Rs, ew_UnFormatDateTime(AuditTrail.DateTime.CurrentValue, 7), Now, False)

		' Field Script
		Call AuditTrail.Script.SetDbValue(Rs, AuditTrail.Script.CurrentValue, Null, False)

		' Field User
		Call AuditTrail.User.SetDbValue(Rs, AuditTrail.User.CurrentValue, Null, False)

		' Field Action
		Call AuditTrail.Action.SetDbValue(Rs, AuditTrail.Action.CurrentValue, Null, False)

		' Field Table
		Call AuditTrail.zTable.SetDbValue(Rs, AuditTrail.zTable.CurrentValue, Null, False)

		' Field Field
		Call AuditTrail.zField.SetDbValue(Rs, AuditTrail.zField.CurrentValue, Null, False)

		' Field KeyValue
		Call AuditTrail.KeyValue.SetDbValue(Rs, AuditTrail.KeyValue.CurrentValue, Null, False)

		' Field OldValue
		Call AuditTrail.OldValue.SetDbValue(Rs, AuditTrail.OldValue.CurrentValue, Null, False)

		' Field NewValue
		Call AuditTrail.NewValue.SetDbValue(Rs, AuditTrail.NewValue.CurrentValue, Null, False)

		' Check recordset update error
		If Err.Number <> 0 Then
			FailureMessage = Err.Description
			Rs.Close
			Set Rs = Nothing
			AddRow = False
			Exit Function
		End If

		' Call Row Inserting event
		bInsertRow = AuditTrail.Row_Inserting(RsOld, Rs)
		If bInsertRow Then

			' Clone new recordset object
			Set RsNew = ew_CloneRs(Rs)
			Rs.Update
			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				AddRow = False
			Else
				AddRow = True
			End If
		Else
			Rs.CancelUpdate
			If AuditTrail.CancelMessage <> "" Then
				FailureMessage = AuditTrail.CancelMessage
				AuditTrail.CancelMessage = ""
			Else
				FailureMessage = Language.Phrase("InsertCancelled")
			End If
			AddRow = False
		End If
		Rs.Close
		Set Rs = Nothing
		If AddRow Then
			AuditTrail.Id.DbValue = RsNew("Id")
		End If
		If AddRow Then

			' Call Row Inserted event
			Call AuditTrail.Row_Inserted(RsOld, RsNew)
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
