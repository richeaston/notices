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
Dim Notices_add
Set Notices_add = New cNotices_add
Set Page = Notices_add

' Page init processing
Call Notices_add.Page_Init()

' Page main processing
Call Notices_add.Page_Main()
%>
<!--#include file="header.asp"-->
<script type="text/javascript">
<!--
// Create page object
var Notices_add = new ew_Page("Notices_add");
// page properties
Notices_add.PageID = "add"; // page ID
Notices_add.FormID = "fNoticesadd"; // form ID
var EW_PAGE_ID = Notices_add.PageID; // for backward compatibility
// extend page with ValidateForm function
Notices_add.ValidateForm = function(fobj) {
	ew_PostAutoSuggest(fobj);
	if (!this.ValidateRequired)
		return true; // ignore validation
	if (fobj.a_confirm && fobj.a_confirm.value == "F")
		return true;
	var i, elm, aelm, infix;
	var rowcnt = 1;
	for (i=0; i<rowcnt; i++) {
		infix = "";
		elm = fobj.elements["x" + infix + "_Title"];
		if (elm && !ew_HasValue(elm))
			return ew_OnError(this, elm, ewLanguage.Phrase("EnterRequiredField") + " - <%= ew_JsEncode2(Notices.Title.FldCaption) %>");
		elm = fobj.elements["x" + infix + "_Author"];
		if (elm && !ew_HasValue(elm))
			return ew_OnError(this, elm, ewLanguage.Phrase("EnterRequiredField") + " - <%= ew_JsEncode2(Notices.Author.FldCaption) %>");
		elm = fobj.elements["x" + infix + "_Sdate"];
		if (elm && !ew_HasValue(elm))
			return ew_OnError(this, elm, ewLanguage.Phrase("EnterRequiredField") + " - <%= ew_JsEncode2(Notices.Sdate.FldCaption) %>");
		elm = fobj.elements["x" + infix + "_Sdate"];
		if (elm && !ew_CheckEuroDate(elm.value))
			return ew_OnError(this, elm, "<%= ew_JsEncode2(Notices.Sdate.FldErrMsg) %>");
		elm = fobj.elements["x" + infix + "_Edate"];
		if (elm && !ew_HasValue(elm))
			return ew_OnError(this, elm, ewLanguage.Phrase("EnterRequiredField") + " - <%= ew_JsEncode2(Notices.Edate.FldCaption) %>");
		elm = fobj.elements["x" + infix + "_Edate"];
		if (elm && !ew_CheckEuroDate(elm.value))
			return ew_OnError(this, elm, "<%= ew_JsEncode2(Notices.Edate.FldErrMsg) %>");
		elm = fobj.elements["x" + infix + "_Group"];
		if (elm && !ew_HasValue(elm))
			return ew_OnError(this, elm, ewLanguage.Phrase("EnterRequiredField") + " - <%= ew_JsEncode2(Notices.Group.FldCaption) %>");
		elm = fobj.elements["x" + infix + "_Notice"];
		if (elm && !ew_HasValue(elm))
			return ew_OnError(this, elm, ewLanguage.Phrase("EnterRequiredField") + " - <%= ew_JsEncode2(Notices.Notice.FldCaption) %>");
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
Notices_add.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
<% If EW_CLIENT_VALIDATE Then %>
Notices_add.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Notices_add.ValidateRequired = false; // no JavaScript validation
<% End If %>
//-->
</script>
<script type="text/javascript" src="ckeditor/ckeditor.js"></script>
<script type="text/javascript">
<!--
_width_multiplier = 20;
_height_multiplier = 60;
var ew_DHTMLEditors = [];
// update value from editor to textarea
function ew_UpdateTextArea() {
	if (typeof ew_DHTMLEditors != 'undefined' && typeof CKEDITOR != 'undefined') {			
		var inst;			
		for (inst in CKEDITOR.instances)
			CKEDITOR.instances[inst].updateElement();
	}
}
// update value from textarea to editor
function ew_UpdateDHTMLEditor(name) {
	if (typeof ew_DHTMLEditors != 'undefined' && typeof CKEDITOR != 'undefined') {
		var inst = CKEDITOR.instances[name];		
		if (inst)
			inst.setData(inst.element.value);
	}
}
// focus editor
function ew_FocusDHTMLEditor(name) {
	if (typeof ew_DHTMLEditors != 'undefined' && typeof CKEDITOR != 'undefined') {
		var inst = CKEDITOR.instances[name];	
		if (inst)
			inst.focus();
	}
}
//-->
</script>
<link rel="stylesheet" type="text/css" media="all" href="calendar/calendar-white-min.css" title="">
<script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/lang/calendar-en.js"></script>
<script type="text/javascript" src="calendar/calendar-setup.js"></script>
<script language="JavaScript" type="text/javascript">
<!--
// Write your client script here, no need to add script tags.
//-->
</script>
<% Notices_add.ShowPageHeader() %>
<p><a class="btn btn-primary" href="<%= Notices.ReturnUrl %>"><i class="icon-arrow-left icon-white"></i>&nbsp;<%= Language.Phrase("GoBack") %></a></p>
<% Notices_add.ShowMessage %>
<form class="form-horizontal" name="fNoticesadd" id="fNoticesadd" action="<%= ew_CurrentPage %>" method="post" onsubmit="return Notices_add.ValidateForm(this);">
<p>
<input type="hidden" name="t" id="t" value="Notices">
<input type="hidden" name="a_add" id="a_add" value="A">
<div class="well Span8">
<Div class="control-group">
	<% If Notices.Title.Visible Then ' Title %>
	<div  class="input-prepend">
	<span class="add-on"><%= Notices.Title.FldCaption %></span>
	<input type="text" name="x_Title" id="x_Title" size="30" maxlength="50" placeholder="Notice Title" value="<%= Notices.Title.EditValue %>"<%= Notices.Title.EditAttributes %>>
	</div>
<% End If %>
&nbsp;
<% If Notices.Author.Visible Then ' Author %>
	<div  class="input-prepend">
	<span class="add-on"><%= Notices.Author.FldCaption %></span>
    <input type="text" name="x_Author" id="x_Author" size="30" maxlength="50" placeholder="Notice Author" value="<%= Notices.Author.EditValue %>"<%= Notices.Author.EditAttributes %>>
	</div>
	
<% End If %>
</div>
<Div class="control-group">
<% If Notices.Sdate.Visible Then ' Sdate %>
	<div class="input-prepend input-append">
	<span class="add-on">Start Date</span>
	<input type="text" name="x_Sdate" id="x_Sdate" value="<%= Notices.Sdate.EditValue %>">
	<a href="#" class="add-on btn" id="cal_x_Sdate" name="cal_x_Sdate" alt="<%= Language.Phrase("PickDate") %>" title="<%= Language.Phrase("PickDate") %>" style="cursor:pointer;cursor:hand;"><i class="icon-calendar"></i></a>
	<script type="text/javascript">
	Calendar.setup({
		inputField: "x_Sdate", // input field id
		ifFormat: "%d/%m/%Y", // date format
		button: "cal_x_Sdate" // button id
		});
	</script>
	</span>
	</div>
<% End If %>
&nbsp;
<% If Notices.Edate.Visible Then ' Edate %>
	<div class="input-prepend input-append">
	<span class="add-on">End Date</span>
	<input type="text" name="x_Edate" id="x_Edate" value="<%= Notices.Edate.EditValue %>">
	<a href="#" class="add-on btn" id="cal_x_Edate" name="cal_x_Edate" alt="<%= Language.Phrase("PickDate") %>" title="<%= Language.Phrase("PickDate") %>" style="cursor:pointer;cursor:hand;"><i class="icon-calendar"></i></a>
	<script type="text/javascript">
		Calendar.setup({
			inputField: "x_Edate", // input field id
			ifFormat: "%d/%m/%Y", // date format
			button: "cal_x_Edate" // button id
			});
	</script>
	</span>
	</div>
<% End If %>
</div>
<% If Notices.Group.Visible Then ' Group %>
	<div class="control-group input-prepend input-append">
	<span class="add-on"><%= Notices.Group.FldCaption %></span>
	<select id="x_Group" name="x_Group"<%= Notices.Group.EditAttributes %>>
	<%
	emptywrk = True
	If IsArray(Notices.Group.EditValue) Then
		arwrk = Notices.Group.EditValue
		For rowcntwrk = 0 To UBound(arwrk, 2)
		If arwrk(0, rowcntwrk)&"" = Notices.Group.CurrentValue&"" Then
			selwrk = " selected=""selected"""
			emptywrk = False
		Else
			selwrk = ""
		End If
	%>
	<option value="<%= Server.HtmlEncode(arwrk(0, rowcntwrk)&"") %>"<%= selwrk %>>
	<%= arwrk(1, rowcntwrk) %>
	</option>
	<%
		Next
		End If
	%>
	</select>
	</span>
	</div>
<% End If %>
<% If Notices.Notice.Visible Then ' Notice %>
	<div class="control-group">
	<textarea class="input-block-level" name="x_Notice" id="x_Notice" rows="8" placeholder="Your notice content"><%= Notices.Notice.EditValue %></textarea>
	<script type="text/javascript">
<!--
<% If Notices.Notice.ReadOnly Then %>
new ew_ReadOnlyTextArea('x_Notice', 40*_width_multiplier, 4*_height_multiplier);
<% Else %>ew_DHTMLEditors.push(new ew_DHTMLEditor("x_Notice", function() {
	var oCKeditor = CKEDITOR.replace('x_Notice', { width: 40*_width_multiplier, height: 4*_height_multiplier, autoUpdateElement: false, baseHref: 'ckeditor/'});
	this.active = true;
}));
<% End If %>
-->
</script>
	</div>
<% End If %>
</div>
<p>
<Button Class="btn btn-danger" type="submit" name="btnAction" id="btnAction"><i class="icon-ok icon-white"></i>&nbsp;Submit New Notice</button>
</form>
<%
Notices_add.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
<!--
ew_CreateEditor();  // Create DHTML editor(s)
//-->
</script>
<script language="JavaScript" type="text/javascript">
<!--
// Write your table-specific startup script here
// document.write("page loaded");
//-->
</script>
<!--#include file="footer.asp"-->
<%

' Drop page object
Set Notices_add = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cNotices_add

	' Page ID
	Public Property Get PageID()
		PageID = "add"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Notices"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Notices_add"
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
		EW_PAGE_ID = "add"

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

		' Table Permission loading event
		Call Security.TablePermission_Loading()
		Call Security.LoadCurrentUserLevel(TableName)

		' Table Permission loaded event
		Call Security.TablePermission_Loaded()

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
			Notices.CurrentAction = ObjForm.GetValue("a_add") ' Get form action
			CopyRecord = LoadOldRecord() ' Load old recordset
			Call LoadFormValues() ' Load form values

			' Validate Form
			If Not ValidateForm() Then
				Notices.CurrentAction = "I" ' Form error, reset action
				Notices.EventCancelled = True ' Event cancelled
				Call RestoreFormValues() ' Restore form values
				FailureMessage = gsFormError
			End If

		' Not post back
		Else

			' Load key values from QueryString
			CopyRecord = True
			If Request.QueryString("Notice_ID").Count > 0 Then
				Notices.Notice_ID.QueryStringValue = Request.QueryString("Notice_ID")
				Call Notices.SetKey("Notice_ID", Notices.Notice_ID.CurrentValue) ' Set up key
			Else
				Call Notices.SetKey("Notice_ID", "") ' Clear key
				CopyRecord = False
			End If
			If CopyRecord Then
				Notices.CurrentAction = "C" ' Copy Record
			Else
				Notices.CurrentAction = "I" ' Display Blank Record
				Call LoadDefaultValues() ' Load default values
			End If
		End If

		' Perform action based on action code
		Select Case Notices.CurrentAction
			Case "I" ' Blank record, no action required
			Case "C" ' Copy an existing record
				If Not LoadRow() Then ' Load record based on key
					FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("noticeslist.asp") ' No matching record, return to list
				End If
			Case "A" ' Add new record
				Notices.SendEmail = True ' Send email on add success
				If AddRow(OldRecordset) Then ' Add successful
					SuccessMessage = Language.Phrase("AddSuccess") ' Set up success message
					Dim sReturnUrl
					sReturnUrl = Notices.ReturnUrl
					If ew_GetPageName(sReturnUrl) = "noticesview.asp" Then sReturnUrl = Notices.ViewUrl ' View paging, return to view page with keyurl directly
					Call Page_Terminate(sReturnUrl) ' Clean up and return
				Else
					Notices.EventCancelled = True ' Event cancelled
					Call RestoreFormValues() ' Add failed, restore form values
				End If
		End Select

		' Render row based on row type
		Notices.RowType = EW_ROWTYPE_ADD ' Render add type

		' Render row
		Call Notices.ResetAttrs()
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
		Notices.Title.CurrentValue = Null
		Notices.Title.OldValue = Notices.Title.CurrentValue
		Notices.Author.CurrentValue = Null
		Notices.Author.OldValue = Notices.Author.CurrentValue
		Notices.Sdate.CurrentValue = Now()
		Notices.Edate.CurrentValue = Now()
		Notices.Group.CurrentValue = Null
		Notices.Group.OldValue = Notices.Group.CurrentValue
		Notices.Notice.CurrentValue = Null
		Notices.Notice.OldValue = Notices.Notice.CurrentValue
	End Function

	' -----------------------------------------------------------------
	' Load form values
	'
	Function LoadFormValues()

		' Load values from form
		If Not Notices.Title.FldIsDetailKey Then Notices.Title.FormValue = ObjForm.GetValue("x_Title")
		If Not Notices.Author.FldIsDetailKey Then Notices.Author.FormValue = ObjForm.GetValue("x_Author")
		If Not Notices.Sdate.FldIsDetailKey Then Notices.Sdate.FormValue = ObjForm.GetValue("x_Sdate")
		If Not Notices.Sdate.FldIsDetailKey Then Notices.Sdate.CurrentValue = ew_UnFormatDateTime(Notices.Sdate.CurrentValue, 7)
		If Not Notices.Edate.FldIsDetailKey Then Notices.Edate.FormValue = ObjForm.GetValue("x_Edate")
		If Not Notices.Edate.FldIsDetailKey Then Notices.Edate.CurrentValue = ew_UnFormatDateTime(Notices.Edate.CurrentValue, 7)
		If Not Notices.Group.FldIsDetailKey Then Notices.Group.FormValue = ObjForm.GetValue("x_Group")
		If Not Notices.Notice.FldIsDetailKey Then Notices.Notice.FormValue = ObjForm.GetValue("x_Notice")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadOldRecord()
		Notices.Title.CurrentValue = Notices.Title.FormValue
		Notices.Author.CurrentValue = Notices.Author.FormValue
		Notices.Sdate.CurrentValue = Notices.Sdate.FormValue
		Notices.Sdate.CurrentValue = ew_UnFormatDateTime(Notices.Sdate.CurrentValue, 7)
		Notices.Edate.CurrentValue = Notices.Edate.FormValue
		Notices.Edate.CurrentValue = ew_UnFormatDateTime(Notices.Edate.CurrentValue, 7)
		Notices.Group.CurrentValue = Notices.Group.FormValue
		Notices.Notice.CurrentValue = Notices.Notice.FormValue
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

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If Notices.GetKey("Notice_ID")&"" <> "" Then
			Notices.Notice_ID.CurrentValue = Notices.GetKey("Notice_ID") ' Notice_ID
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			Notices.CurrentFilter = Notices.KeyFilter
			Dim sSql
			sSql = Notices.SQL
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

		' ---------
		'  Add Row
		' ---------

		ElseIf Notices.RowType = EW_ROWTYPE_ADD Then ' Add row

			' Title
			Notices.Title.EditCustomAttributes = ""
			Notices.Title.EditValue = ew_HtmlEncode(Notices.Title.CurrentValue)

			' Author
			Notices.Author.EditCustomAttributes = ""
			Notices.Author.EditValue = ew_HtmlEncode(Notices.Author.CurrentValue)

			' Sdate
			Notices.Sdate.EditCustomAttributes = ""
			Notices.Sdate.EditValue = ew_FormatDateTime(Notices.Sdate.CurrentValue, 7)

			' Edate
			Notices.Edate.EditCustomAttributes = ""
			Notices.Edate.EditValue = ew_FormatDateTime(Notices.Edate.CurrentValue, 7)

			' Group
			Notices.Group.EditCustomAttributes = ""
				sFilterWrk = ""
			sSqlWrk = "SELECT DISTINCT [Group], [Group] AS [DispFld], '' AS [Disp2Fld], '' AS [Disp3Fld], '' AS [Disp4Fld], '' AS [SelectFilterFld] FROM [Groups]"
			sWhereWrk = ""
			Call ew_AddFilter(sWhereWrk, sFilterWrk)
			If sWhereWrk <> "" Then sSqlWrk = sSqlWrk & " WHERE " & sWhereWrk
			sSqlWrk = sSqlWrk & " ORDER BY [Group] Asc"
			Set RsWrk = Server.CreateObject("ADODB.Recordset")
			RsWrk.Open sSqlWrk, Conn
			If Not RsWrk.Eof Then
				arwrk = RsWrk.GetRows
			Else
				arwrk = ""
			End If
			RsWrk.Close
			Set RsWrk = Nothing
			arwrk = ew_AddItemToArray(arwrk, 0, Array("", Language.Phrase("PleaseSelect")))
			Notices.Group.EditValue = arwrk

			' Notice
			Notices.Notice.EditCustomAttributes = ""
			Notices.Notice.EditValue = ew_HtmlEncode(Notices.Notice.CurrentValue)

			' Edit refer script
			' Title

			Notices.Title.HrefValue = ""

			' Author
			Notices.Author.HrefValue = ""

			' Sdate
			Notices.Sdate.HrefValue = ""

			' Edate
			Notices.Edate.HrefValue = ""

			' Group
			Notices.Group.HrefValue = ""

			' Notice
			Notices.Notice.HrefValue = ""
		End If
		If Notices.RowType = EW_ROWTYPE_ADD Or Notices.RowType = EW_ROWTYPE_EDIT Or Notices.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call Notices.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If Notices.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Notices.Row_Rendered()
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
		If Not IsNull(Notices.Title.FormValue) And Notices.Title.FormValue&"" = "" Then
			Call ew_AddMessage(gsFormError, Language.Phrase("EnterRequiredField") & " - " & Notices.Title.FldCaption)
		End If
		If Not IsNull(Notices.Author.FormValue) And Notices.Author.FormValue&"" = "" Then
			Call ew_AddMessage(gsFormError, Language.Phrase("EnterRequiredField") & " - " & Notices.Author.FldCaption)
		End If
		If Not IsNull(Notices.Sdate.FormValue) And Notices.Sdate.FormValue&"" = "" Then
			Call ew_AddMessage(gsFormError, Language.Phrase("EnterRequiredField") & " - " & Notices.Sdate.FldCaption)
		End If
		If Not ew_CheckEuroDate(Notices.Sdate.FormValue) Then
			Call ew_AddMessage(gsFormError, Notices.Sdate.FldErrMsg)
		End If
		If Not IsNull(Notices.Edate.FormValue) And Notices.Edate.FormValue&"" = "" Then
			Call ew_AddMessage(gsFormError, Language.Phrase("EnterRequiredField") & " - " & Notices.Edate.FldCaption)
		End If
		If Not ew_CheckEuroDate(Notices.Edate.FormValue) Then
			Call ew_AddMessage(gsFormError, Notices.Edate.FldErrMsg)
		End If
		If Not IsNull(Notices.Group.FormValue) And Notices.Group.FormValue&"" = "" Then
			Call ew_AddMessage(gsFormError, Language.Phrase("EnterRequiredField") & " - " & Notices.Group.FldCaption)
		End If
		If Not IsNull(Notices.Notice.FormValue) And Notices.Notice.FormValue&"" = "" Then
			Call ew_AddMessage(gsFormError, Language.Phrase("EnterRequiredField") & " - " & Notices.Notice.FldCaption)
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
		Notices.CurrentFilter = sFilter
		sSql = Notices.SQL
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

		' Field Title
		Call Notices.Title.SetDbValue(Rs, Notices.Title.CurrentValue, Null, False)

		' Field Author
		Call Notices.Author.SetDbValue(Rs, Notices.Author.CurrentValue, Null, False)

		' Field Sdate
		Call Notices.Sdate.SetDbValue(Rs, ew_UnFormatDateTime(Notices.Sdate.CurrentValue, 7), Null, False)

		' Field Edate
		Call Notices.Edate.SetDbValue(Rs, ew_UnFormatDateTime(Notices.Edate.CurrentValue, 7), Null, False)

		' Field Group
		Call Notices.Group.SetDbValue(Rs, Notices.Group.CurrentValue, Null, False)

		' Field Notice
		Call Notices.Notice.SetDbValue(Rs, Notices.Notice.CurrentValue, Null, False)

		' Check recordset update error
		If Err.Number <> 0 Then
			FailureMessage = Err.Description
			Rs.Close
			Set Rs = Nothing
			AddRow = False
			Exit Function
		End If

		' Call Row Inserting event
		bInsertRow = Notices.Row_Inserting(RsOld, Rs)
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
			If Notices.CancelMessage <> "" Then
				FailureMessage = Notices.CancelMessage
				Notices.CancelMessage = ""
			Else
				FailureMessage = Language.Phrase("InsertCancelled")
			End If
			AddRow = False
		End If
		Rs.Close
		Set Rs = Nothing
		If AddRow Then
			Notices.Notice_ID.DbValue = RsNew("Notice_ID")
		End If
		If AddRow Then

			' Call Row Inserted event
			Call Notices.Row_Inserted(RsOld, RsNew)
			Call WriteAuditTrailOnAdd(RsNew)
			If Notices.SendEmail Then Call SendEmailOnAdd(RsNew)
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

	' Write Audit Trail (add page)
	Sub WriteAuditTrailOnAdd(RsSrc)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		Dim table
		table = "Notices"

		' Get key value
		Dim sKey
		sKey = ""
		If sKey <> "" Then sKey = sKey & EW_COMPOSITE_KEY_SEPARATOR
		sKey = sKey & RsSrc.Fields("Notice_ID")

		' Write Audit Trail
		Dim filePfx, curDateTime, id, user, action, field, keyvalue, oldvalue, newvalue
		Dim i
		filePfx = "log"
		curDateTime = ew_StdCurrentDateTime()
		id = Request.ServerVariables("SCRIPT_NAME")
    	user = CurrentUserName
		action = "A"
		keyvalue = sKey
		oldvalue = ""

		' Title Field
		newvalue = RsSrc("Title")
		Call ew_WriteAuditTrail(filePfx, curDateTime, id, user, action, table, "Title", keyvalue, oldvalue, newvalue)

		' Author Field
		newvalue = RsSrc("Author")
		Call ew_WriteAuditTrail(filePfx, curDateTime, id, user, action, table, "Author", keyvalue, oldvalue, newvalue)

		' Sdate Field
		newvalue = RsSrc("Sdate")
		Call ew_WriteAuditTrail(filePfx, curDateTime, id, user, action, table, "Sdate", keyvalue, oldvalue, newvalue)

		' Edate Field
		newvalue = RsSrc("Edate")
		Call ew_WriteAuditTrail(filePfx, curDateTime, id, user, action, table, "Edate", keyvalue, oldvalue, newvalue)

		' Group Field
		newvalue = RsSrc("Group")
		Call ew_WriteAuditTrail(filePfx, curDateTime, id, user, action, table, "Group", keyvalue, oldvalue, newvalue)

		' Notice Field
		newvalue = RsSrc("Notice")
		If Not EW_AUDIT_TRAIL_TO_DATABASE Then newvalue = "[MEMO]"
		Call ew_WriteAuditTrail(filePfx, curDateTime, id, user, action, table, "Notice", keyvalue, oldvalue, newvalue)
	End Sub

	' Send email after add success
	Sub SendEmailOnAdd(RsSrc)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		Dim sFn, sSubject, sTable, sKey, sAction
		sFn = "txt/newnotice.txt"
		sTable = "Notices"
		sAuthor = RsSrc.Fields("Author")
		sTitle = RsSrc.Fields("Title")
		sNotice = RsSrc.Fields("Notice")
		sSubject = "A new notice (" & sTitle & ") has been submitted by " & sAuthor
		sAction = "The following Notice <B>(" & sTitle & ")</b> have been submitted by <B>" & sAuthor & "</B><BR><BR>Please check the notices system to approve this <a href='http://kbl-web/notices/' target='_blank'>notice</a>."
		
		
		
		' Get key value
		sKey = ""
		If sKey <> "" Then sKey = sKey & EW_COMPOSITE_KEY_SEPARATOR
		sKey = sKey & RsSrc.Fields("Notice_ID")
		Dim Email, bEmailSent
		Set Email = New cEmail
		Email.Load(sFn)
		Email.ReplaceSender(EW_SENDER_EMAIL) ' Replace Sender
		Email.ReplaceRecipient(EW_RECIPIENT_EMAIL) ' Replace Recipient
		Email.ReplaceSubject(sSubject) ' Replace Subject
		Email.ReplaceContent "<!--action-->", sAction
		If EW_EMAIL_CHARSET <> "" Then Email.Charset = EW_EMAIL_CHARSET
		Set EventArgs = Server.CreateObject("Scripting.Dictionary")
		EventArgs.Add "RsNew", RsSrc
		If Notices.Email_Sending(Email, EventArgs) Then
			bEmailSent = Email.Send()
		Else
			bEmailSent = False
		End If
		Set EventArgs = Nothing

		' Send email failed
		If Not bEmailSent Then
			FailureMessage = Replace(Replace(Language.Phrase("FailedToSendMail"),"%n",Email.SendErrNumber),"%e",Email.SendErrDescription)
		End If
		Set Email = Nothing
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
