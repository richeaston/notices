<%@ CodePage="1252" LCID="2057" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="ewcfg9.asp"-->
<!--#include file="unapproved_noticesinfo.asp"-->
<!--#include file="usersinfo.asp"-->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim Unapproved_Notices_edit
Set Unapproved_Notices_edit = New cUnapproved_Notices_edit
Set Page = Unapproved_Notices_edit

' Page init processing
Call Unapproved_Notices_edit.Page_Init()

' Page main processing
Call Unapproved_Notices_edit.Page_Main()
%>
<!--#include file="header.asp"-->
<script type="text/javascript">
<!--
// Create page object
var Unapproved_Notices_edit = new ew_Page("Unapproved_Notices_edit");
// page properties
Unapproved_Notices_edit.PageID = "edit"; // page ID
Unapproved_Notices_edit.FormID = "fUnapproved_Noticesedit"; // form ID
var EW_PAGE_ID = Unapproved_Notices_edit.PageID; // for backward compatibility
// extend page with ValidateForm function
Unapproved_Notices_edit.ValidateForm = function(fobj) {
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
			return ew_OnError(this, elm, ewLanguage.Phrase("EnterRequiredField") + " - <%= ew_JsEncode2(Unapproved_Notices.Title.FldCaption) %>");
		elm = fobj.elements["x" + infix + "_Sdate"];
		if (elm && !ew_HasValue(elm))
			return ew_OnError(this, elm, ewLanguage.Phrase("EnterRequiredField") + " - <%= ew_JsEncode2(Unapproved_Notices.Sdate.FldCaption) %>");
		elm = fobj.elements["x" + infix + "_Sdate"];
		if (elm && !ew_CheckEuroDate(elm.value))
			return ew_OnError(this, elm, "<%= ew_JsEncode2(Unapproved_Notices.Sdate.FldErrMsg) %>");
		elm = fobj.elements["x" + infix + "_Edate"];
		if (elm && !ew_HasValue(elm))
			return ew_OnError(this, elm, ewLanguage.Phrase("EnterRequiredField") + " - <%= ew_JsEncode2(Unapproved_Notices.Edate.FldCaption) %>");
		elm = fobj.elements["x" + infix + "_Edate"];
		if (elm && !ew_CheckEuroDate(elm.value))
			return ew_OnError(this, elm, "<%= ew_JsEncode2(Unapproved_Notices.Edate.FldErrMsg) %>");
		elm = fobj.elements["x" + infix + "_Group"];
		if (elm && !ew_HasValue(elm))
			return ew_OnError(this, elm, ewLanguage.Phrase("EnterRequiredField") + " - <%= ew_JsEncode2(Unapproved_Notices.Group.FldCaption) %>");
		elm = fobj.elements["x" + infix + "_Notice"];
		if (elm && !ew_HasValue(elm))
			return ew_OnError(this, elm, ewLanguage.Phrase("EnterRequiredField") + " - <%= ew_JsEncode2(Unapproved_Notices.Notice.FldCaption) %>");
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
Unapproved_Notices_edit.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
Unapproved_Notices_edit.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
Unapproved_Notices_edit.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Unapproved_Notices_edit.ValidateRequired = false; // no JavaScript validation
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
<link rel="stylesheet" type="text/css" media="all" href="calendar/calendar-win2k-cold-1.css" title="win2k-1">
<script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/lang/calendar-en.js"></script>
<script type="text/javascript" src="calendar/calendar-setup.js"></script>
<script language="JavaScript" type="text/javascript">
<!--
// Write your client script here, no need to add script tags.
//-->
</script>
<% Unapproved_Notices_edit.ShowPageHeader() %>
<p class="aspmaker"><a class="btn btn-inverse" href="<%= Unapproved_Notices.ReturnUrl %>"><i class="icon-arrow-left icon-white"></i>&nbsp;<%= Language.Phrase("GoBack") %></a></p>
<% Unapproved_Notices_edit.ShowMessage %>
<div class="well">
<form name="fUnapproved_Noticesedit" id="fUnapproved_Noticesedit" action="<%= ew_CurrentPage %>" method="post" onsubmit="return Unapproved_Notices_edit.ValidateForm(this);">
<p>
<input type="hidden" name="a_table" id="a_table" value="Unapproved_Notices">
<input type="hidden" name="a_edit" id="a_edit" value="U">
<% If Unapproved_Notices.CurrentAction = "F" Then ' Confirm page %>
<input type="hidden" name="a_confirm" id="a_confirm" value="F">
<% End If %>
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table cellspacing="0" class="ewTable">
<% If Unapproved_Notices.Title.Visible Then ' Title %>
	<tr id="r_Title"<%= Unapproved_Notices.RowAttributes %>>
		<td class="ewTableHeader"><%= Unapproved_Notices.Title.FldCaption %><%= Language.Phrase("FieldRequiredIndicator") %></td>
		<td<%= Unapproved_Notices.Title.CellAttributes %>><span id="el_Title">
<% If Unapproved_Notices.CurrentAction <> "F" Then %>
<input type="text" name="x_Title" id="x_Title" size="30" maxlength="50" value="<%= Unapproved_Notices.Title.EditValue %>"<%= Unapproved_Notices.Title.EditAttributes %>>
<% Else %>
<div<%= Unapproved_Notices.Title.ViewAttributes %>><%= Unapproved_Notices.Title.ViewValue %></div>
<input type="hidden" name="x_Title" id="x_Title" value="<%= Server.HTMLEncode(Unapproved_Notices.Title.FormValue&"") %>">
<% End If %>
</span><%= Unapproved_Notices.Title.CustomMsg %></td>
	</tr>
<% End If %>
<% If Unapproved_Notices.Sdate.Visible Then ' Sdate %>
	<tr id="r_Sdate"<%= Unapproved_Notices.RowAttributes %>>
		<td class="ewTableHeader"><%= Unapproved_Notices.Sdate.FldCaption %><%= Language.Phrase("FieldRequiredIndicator") %></td>
		<td<%= Unapproved_Notices.Sdate.CellAttributes %>><span id="el_Sdate">
<% If Unapproved_Notices.CurrentAction <> "F" Then %>
<input type="text" name="x_Sdate" id="x_Sdate" value="<%= Unapproved_Notices.Sdate.EditValue %>"<%= Unapproved_Notices.Sdate.EditAttributes %>>
&nbsp;<img src="images/calendar.png" id="cal_x_Sdate" name="cal_x_Sdate" alt="<%= Language.Phrase("PickDate") %>" title="<%= Language.Phrase("PickDate") %>" style="cursor:pointer;cursor:hand;">
<script type="text/javascript">
Calendar.setup({
	inputField: "x_Sdate", // input field id
	ifFormat: "%d/%m/%Y", // date format
	button: "cal_x_Sdate" // button id
});
</script>
<% Else %>
<div<%= Unapproved_Notices.Sdate.ViewAttributes %>><%= Unapproved_Notices.Sdate.ViewValue %></div>
<input type="hidden" name="x_Sdate" id="x_Sdate" value="<%= Server.HTMLEncode(Unapproved_Notices.Sdate.FormValue&"") %>">
<% End If %>
</span><%= Unapproved_Notices.Sdate.CustomMsg %></td>
	</tr>
<% End If %>
<% If Unapproved_Notices.Edate.Visible Then ' Edate %>
	<tr id="r_Edate"<%= Unapproved_Notices.RowAttributes %>>
		<td class="input-prepend"><span class="addon"><%= Unapproved_Notices.Edate.FldCaption %><%= Language.Phrase("FieldRequiredIndicator") %></span>>
		<td<%= Unapproved_Notices.Edate.CellAttributes %>><span id="el_Edate">
<% If Unapproved_Notices.CurrentAction <> "F" Then %>
<input type="text" name="x_Edate" id="x_Edate" value="<%= Unapproved_Notices.Edate.EditValue %>"<%= Unapproved_Notices.Edate.EditAttributes %>>
&nbsp;<img src="images/calendar.png" id="cal_x_Edate" name="cal_x_Edate" alt="<%= Language.Phrase("PickDate") %>" title="<%= Language.Phrase("PickDate") %>" style="cursor:pointer;cursor:hand;">
<script type="text/javascript">
Calendar.setup({
	inputField: "x_Edate", // input field id
	ifFormat: "%d/%m/%Y", // date format
	button: "cal_x_Edate" // button id
});
</script>
<% Else %>
<div<%= Unapproved_Notices.Edate.ViewAttributes %>><%= Unapproved_Notices.Edate.ViewValue %></div>
<input type="hidden" name="x_Edate" id="x_Edate" value="<%= Server.HTMLEncode(Unapproved_Notices.Edate.FormValue&"") %>">
<% End If %>
</span><%= Unapproved_Notices.Edate.CustomMsg %></td>
	</tr>
<% End If %>
<% If Unapproved_Notices.Group.Visible Then ' Group %>
	<tr id="r_Group"<%= Unapproved_Notices.RowAttributes %>>
		<td><%= Unapproved_Notices.Group.FldCaption %><%= Language.Phrase("FieldRequiredIndicator") %></td>
		<td<%= Unapproved_Notices.Group.CellAttributes %>><span id="el_Group">
<% If Unapproved_Notices.CurrentAction <> "F" Then %>
<select id="x_Group" name="x_Group"<%= Unapproved_Notices.Group.EditAttributes %>>
<%
emptywrk = True
If IsArray(Unapproved_Notices.Group.EditValue) Then
	arwrk = Unapproved_Notices.Group.EditValue
	For rowcntwrk = 0 To UBound(arwrk, 2)
		If arwrk(0, rowcntwrk)&"" = Unapproved_Notices.Group.CurrentValue&"" Then
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
<% Else %>
<div<%= Unapproved_Notices.Group.ViewAttributes %>><%= Unapproved_Notices.Group.ViewValue %></div>
<input type="hidden" name="x_Group" id="x_Group" value="<%= Server.HTMLEncode(Unapproved_Notices.Group.FormValue&"") %>">
<% End If %>
</span><%= Unapproved_Notices.Group.CustomMsg %></td>
	</tr>
<% End If %>
<% If Unapproved_Notices.Notice.Visible Then ' Notice %>
	<tr id="r_Notice"<%= Unapproved_Notices.RowAttributes %>>
		<td class="ewTableHeader"><%= Unapproved_Notices.Notice.FldCaption %><%= Language.Phrase("FieldRequiredIndicator") %></td>
		<td<%= Unapproved_Notices.Notice.CellAttributes %>><span id="el_Notice">
<% If Unapproved_Notices.CurrentAction <> "F" Then %>
<textarea class="input-block-level" name="x_Notice" id="x_Notice" rows="8"><%= Unapproved_Notices.Notice.EditValue %></textarea>

<% Else %>
<div<%= Unapproved_Notices.Notice.ViewAttributes %>><%= Unapproved_Notices.Notice.ViewValue %></div>
<input type="hidden" name="x_Notice" id="x_Notice" value="<%= Server.HTMLEncode(Unapproved_Notices.Notice.FormValue&"") %>">
<% End If %>
</span><%= Unapproved_Notices.Notice.CustomMsg %></td>
	</tr>
<% End If %>
<% If Unapproved_Notices.Approved.Visible Then ' Approved %>
	<tr id="r_Approved"<%= Unapproved_Notices.RowAttributes %>>
		<td class="ewTableHeader"><%= Unapproved_Notices.Approved.FldCaption %></td>
		<td<%= Unapproved_Notices.Approved.CellAttributes %>><span id="el_Approved">
<% If Unapproved_Notices.CurrentAction <> "F" Then %>
<% selwrk = ew_IIf(ew_ConvertToBool(Unapproved_Notices.Approved.CurrentValue), " checked=""checked""", "") %>
<input type="checkbox" name="x_Approved" id="x_Approved" value="1"<%= selwrk %><%= Unapproved_Notices.Approved.EditAttributes %>>
<% Else %>
<% If ew_ConvertToBool(Unapproved_Notices.Approved.CurrentValue) Then %>
<input type="checkbox" value="<%= Unapproved_Notices.Approved.ViewValue %>" checked onclick="this.form.reset();" disabled="disabled">
<% Else %>
<input type="checkbox" value="<%= Unapproved_Notices.Approved.ViewValue %>" onclick="this.form.reset();" disabled="disabled">
<% End If %>
<input type="hidden" name="x_Approved" id="x_Approved" value="<%= Server.HTMLEncode(Unapproved_Notices.Approved.FormValue&"") %>">
<% End If %>
</span><%= Unapproved_Notices.Approved.CustomMsg %></td>
	</tr>
<% End If %>
<!--
<% If Unapproved_Notices.Archived.Visible Then ' Archived %>
	<tr id="r_Archived"<%= Unapproved_Notices.RowAttributes %>>
		<td class="ewTableHeader"><%= Unapproved_Notices.Archived.FldCaption %></td>
		<td<%= Unapproved_Notices.Archived.CellAttributes %>><span id="el_Archived">
<% If Unapproved_Notices.CurrentAction <> "F" Then %>
<% selwrk = ew_IIf(ew_ConvertToBool(Unapproved_Notices.Archived.CurrentValue), " checked=""checked""", "") %>
<input type="checkbox" name="x_Archived" id="x_Archived" value="1"<%= selwrk %><%= Unapproved_Notices.Archived.EditAttributes %>>
<% Else %>
<% If ew_ConvertToBool(Unapproved_Notices.Archived.CurrentValue) Then %>
<input type="checkbox" value="<%= Unapproved_Notices.Archived.ViewValue %>" checked onclick="this.form.reset();" disabled="disabled">
<% Else %>
<input type="checkbox" value="<%= Unapproved_Notices.Archived.ViewValue %>" onclick="this.form.reset();" disabled="disabled">
<% End If %>
<input type="hidden" name="x_Archived" id="x_Archived" value="<%= Server.HTMLEncode(Unapproved_Notices.Archived.FormValue&"") %>">
<% End If %>
</span><%= Unapproved_Notices.Archived.CustomMsg %></td>
	</tr>
<% End If %>
-->
</table>
</div>
</td></tr></table>
</div>
<input type="hidden" name="x_Notice_ID" id="x_Notice_ID" value="<%= Server.HTMLEncode(Unapproved_Notices.Notice_ID.CurrentValue&"") %>">
<p>
<% If Unapproved_Notices.CurrentAction <> "F" Then ' Confirm page %>
<button class="btn btn-primary" type="submit" name="btnAction" id="btnAction" onclick="this.form.a_edit.value='F';"><i class="icon-ok icon-white"></i>&nbsp;Edit Notice</button>
<% Else %>
<input type="submit" name="btnCancel" id="btnCancel" value="<%= ew_BtnCaption(Language.Phrase("CancelBtn")) %>" onclick="this.form.a_edit.value='X';">
<input type="submit" name="btnAction" id="btnAction" value="<%= ew_BtnCaption(Language.Phrase("ConfirmBtn")) %>">
<% End If %>
</form>
<% If Unapproved_Notices.CurrentAction <> "F" Then %>
<% End If %>
<%
Unapproved_Notices_edit.ShowPageFooter()
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
Set Unapproved_Notices_edit = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cUnapproved_Notices_edit

	' Page ID
	Public Property Get PageID()
		PageID = "edit"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Unapproved Notices"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Unapproved_Notices_edit"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If Unapproved_Notices.UseTokenInUrl Then PageUrl = PageUrl & "t=" & Unapproved_Notices.TableVar & "&" ' add page token
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
		If Unapproved_Notices.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (Unapproved_Notices.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (Unapproved_Notices.TableVar = Request.QueryString("t"))
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
		If IsEmpty(Unapproved_Notices) Then Set Unapproved_Notices = New cUnapproved_Notices
		Set Table = Unapproved_Notices

		' Initialize urls
		' Initialize other table object

		If IsEmpty(Users) Then Set Users = New cUsers

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "edit"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "Unapproved Notices"

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
			Call Page_Terminate("unapproved_noticeslist.asp")
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
		Set Unapproved_Notices = Nothing
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
		If Request.QueryString("Notice_ID").Count > 0 Then
			Unapproved_Notices.Notice_ID.QueryStringValue = Request.QueryString("Notice_ID")
		End If
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			Unapproved_Notices.CurrentAction = ObjForm.GetValue("a_edit") ' Get action code
			Call LoadFormValues() ' Get form values

			' Validate Form
			If Not ValidateForm() Then
				Unapproved_Notices.CurrentAction = "" ' Form error, reset action
				FailureMessage = gsFormError
				Unapproved_Notices.EventCancelled = True ' Event cancelled
				Call LoadRow() ' Restore row
				Call RestoreFormValues() ' Restore form values if validate failed
			End If
		Else
			Unapproved_Notices.CurrentAction = "I" ' Default action is display
		End If

		' Check if valid key
		If Unapproved_Notices.Notice_ID.CurrentValue = "" Then Call Page_Terminate("unapproved_noticeslist.asp") ' Invalid key, return to list
		Select Case Unapproved_Notices.CurrentAction
			Case "I" ' Get a record to display
				If Not LoadRow() Then ' Load Record based on key
					FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("unapproved_noticeslist.asp") ' No matching record, return to list
				End If
			Case "U" ' Update
				Unapproved_Notices.SendEmail = True ' Send email on update success
				If EditRow() Then ' Update Record based on key
					SuccessMessage = Language.Phrase("UpdateSuccess") ' Update success
					Dim sReturnUrl
					sReturnUrl = Unapproved_Notices.ReturnUrl
					Call Page_Terminate(sReturnUrl) ' Return to caller
				Else
					Unapproved_Notices.EventCancelled = True ' Event cancelled
					Call LoadRow() ' Restore row
					Call RestoreFormValues() ' Restore form values if update failed
				End If
		End Select

		' Render the record
		If Unapproved_Notices.CurrentAction = "F" Then ' Confirm page
			Unapproved_Notices.RowType = EW_ROWTYPE_VIEW ' Render as view
		Else
			Unapproved_Notices.RowType = EW_ROWTYPE_EDIT ' Render as edit
		End If

		' Render row
		Call Unapproved_Notices.ResetAttrs()
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
		If Not Unapproved_Notices.Title.FldIsDetailKey Then Unapproved_Notices.Title.FormValue = ObjForm.GetValue("x_Title")
		If Not Unapproved_Notices.Sdate.FldIsDetailKey Then Unapproved_Notices.Sdate.FormValue = ObjForm.GetValue("x_Sdate")
		If Not Unapproved_Notices.Sdate.FldIsDetailKey Then Unapproved_Notices.Sdate.CurrentValue = ew_UnFormatDateTime(Unapproved_Notices.Sdate.CurrentValue, 7)
		If Not Unapproved_Notices.Edate.FldIsDetailKey Then Unapproved_Notices.Edate.FormValue = ObjForm.GetValue("x_Edate")
		If Not Unapproved_Notices.Edate.FldIsDetailKey Then Unapproved_Notices.Edate.CurrentValue = ew_UnFormatDateTime(Unapproved_Notices.Edate.CurrentValue, 7)
		If Not Unapproved_Notices.Group.FldIsDetailKey Then Unapproved_Notices.Group.FormValue = ObjForm.GetValue("x_Group")
		If Not Unapproved_Notices.Notice.FldIsDetailKey Then Unapproved_Notices.Notice.FormValue = ObjForm.GetValue("x_Notice")
		If Not Unapproved_Notices.Approved.FldIsDetailKey Then Unapproved_Notices.Approved.FormValue = ObjForm.GetValue("x_Approved")
		If Not Unapproved_Notices.Archived.FldIsDetailKey Then Unapproved_Notices.Archived.FormValue = ObjForm.GetValue("x_Archived")
		If Not Unapproved_Notices.Notice_ID.FldIsDetailKey Then Unapproved_Notices.Notice_ID.FormValue = ObjForm.GetValue("x_Notice_ID")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadRow()
		Unapproved_Notices.Title.CurrentValue = Unapproved_Notices.Title.FormValue
		Unapproved_Notices.Sdate.CurrentValue = Unapproved_Notices.Sdate.FormValue
		Unapproved_Notices.Sdate.CurrentValue = ew_UnFormatDateTime(Unapproved_Notices.Sdate.CurrentValue, 7)
		Unapproved_Notices.Edate.CurrentValue = Unapproved_Notices.Edate.FormValue
		Unapproved_Notices.Edate.CurrentValue = ew_UnFormatDateTime(Unapproved_Notices.Edate.CurrentValue, 7)
		Unapproved_Notices.Group.CurrentValue = Unapproved_Notices.Group.FormValue
		Unapproved_Notices.Notice.CurrentValue = Unapproved_Notices.Notice.FormValue
		Unapproved_Notices.Approved.CurrentValue = Unapproved_Notices.Approved.FormValue
		Unapproved_Notices.Archived.CurrentValue = Unapproved_Notices.Archived.FormValue
		Unapproved_Notices.Notice_ID.CurrentValue = Unapproved_Notices.Notice_ID.FormValue
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = Unapproved_Notices.KeyFilter

		' Call Row Selecting event
		Call Unapproved_Notices.Row_Selecting(sFilter)

		' Load sql based on filter
		Unapproved_Notices.CurrentFilter = sFilter
		sSql = Unapproved_Notices.SQL
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
		Call Unapproved_Notices.Row_Selected(RsRow)
		Unapproved_Notices.Notice_ID.DbValue = RsRow("Notice_ID")
		Unapproved_Notices.Title.DbValue = RsRow("Title")
		Unapproved_Notices.Author.DbValue = RsRow("Author")
		Unapproved_Notices.Sdate.DbValue = RsRow("Sdate")
		Unapproved_Notices.Edate.DbValue = RsRow("Edate")
		Unapproved_Notices.Group.DbValue = RsRow("Group")
		Unapproved_Notices.Notice.DbValue = RsRow("Notice")
		Unapproved_Notices.Approved.DbValue = ew_IIf(RsRow("Approved"), "1", "0")
		Unapproved_Notices.Archived.DbValue = ew_IIf(RsRow("Archived"), "1", "0")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Call Row Rendering event

		Call Unapproved_Notices.Row_Rendering()

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

		If Unapproved_Notices.RowType = EW_ROWTYPE_VIEW Then ' View row

			' Title
			Unapproved_Notices.Title.ViewValue = Unapproved_Notices.Title.CurrentValue
			Unapproved_Notices.Title.ViewCustomAttributes = ""

			' Author
			Unapproved_Notices.Author.ViewValue = Unapproved_Notices.Author.CurrentValue
			Unapproved_Notices.Author.ViewCustomAttributes = ""

			' Sdate
			Unapproved_Notices.Sdate.ViewValue = Unapproved_Notices.Sdate.CurrentValue
			Unapproved_Notices.Sdate.ViewValue = ew_FormatDateTime(Unapproved_Notices.Sdate.ViewValue, 7)
			Unapproved_Notices.Sdate.ViewCustomAttributes = ""

			' Edate
			Unapproved_Notices.Edate.ViewValue = Unapproved_Notices.Edate.CurrentValue
			Unapproved_Notices.Edate.ViewValue = ew_FormatDateTime(Unapproved_Notices.Edate.ViewValue, 7)
			Unapproved_Notices.Edate.ViewCustomAttributes = ""

			' Group
			If Unapproved_Notices.Group.CurrentValue & "" <> "" Then
				sFilterWrk = "[Group] = '" & ew_AdjustSql(Unapproved_Notices.Group.CurrentValue) & "'"
			sSqlWrk = "SELECT DISTINCT [Group] FROM [Groups]"
			sWhereWrk = ""
			Call ew_AddFilter(sWhereWrk, sFilterWrk)
			If sWhereWrk <> "" Then sSqlWrk = sSqlWrk & " WHERE " & sWhereWrk
			sSqlWrk = sSqlWrk & " ORDER BY [Group] Asc"
				Set RsWrk = Conn.Execute(sSqlWrk)
				If Not RsWrk.Eof Then
					Unapproved_Notices.Group.ViewValue = RsWrk("Group")
				Else
					Unapproved_Notices.Group.ViewValue = Unapproved_Notices.Group.CurrentValue
				End If
				RsWrk.Close
				Set RsWrk = Nothing
			Else
				Unapproved_Notices.Group.ViewValue = Null
			End If
			Unapproved_Notices.Group.ViewCustomAttributes = ""

			' Notice
			Unapproved_Notices.Notice.ViewValue = Unapproved_Notices.Notice.CurrentValue
			Unapproved_Notices.Notice.ViewCustomAttributes = ""

			' Approved
			If ew_ConvertToBool(Unapproved_Notices.Approved.CurrentValue) Then
				Unapproved_Notices.Approved.ViewValue = ew_IIf(Unapproved_Notices.Approved.FldTagCaption(1) <> "", Unapproved_Notices.Approved.FldTagCaption(1), "Yes")
			Else
				Unapproved_Notices.Approved.ViewValue = ew_IIf(Unapproved_Notices.Approved.FldTagCaption(2) <> "", Unapproved_Notices.Approved.FldTagCaption(2), "No")
			End If
			Unapproved_Notices.Approved.ViewCustomAttributes = ""

			' Archived
			If ew_ConvertToBool(Unapproved_Notices.Archived.CurrentValue) Then
				Unapproved_Notices.Archived.ViewValue = ew_IIf(Unapproved_Notices.Archived.FldTagCaption(1) <> "", Unapproved_Notices.Archived.FldTagCaption(1), "Yes")
			Else
				Unapproved_Notices.Archived.ViewValue = ew_IIf(Unapproved_Notices.Archived.FldTagCaption(2) <> "", Unapproved_Notices.Archived.FldTagCaption(2), "No")
			End If
			Unapproved_Notices.Archived.ViewCustomAttributes = ""

			' View refer script
			' Title

			Unapproved_Notices.Title.LinkCustomAttributes = ""
			Unapproved_Notices.Title.HrefValue = ""
			Unapproved_Notices.Title.TooltipValue = ""

			' Sdate
			Unapproved_Notices.Sdate.LinkCustomAttributes = ""
			Unapproved_Notices.Sdate.HrefValue = ""
			Unapproved_Notices.Sdate.TooltipValue = ""

			' Edate
			Unapproved_Notices.Edate.LinkCustomAttributes = ""
			Unapproved_Notices.Edate.HrefValue = ""
			Unapproved_Notices.Edate.TooltipValue = ""

			' Group
			Unapproved_Notices.Group.LinkCustomAttributes = ""
			Unapproved_Notices.Group.HrefValue = ""
			Unapproved_Notices.Group.TooltipValue = ""

			' Notice
			Unapproved_Notices.Notice.LinkCustomAttributes = ""
			Unapproved_Notices.Notice.HrefValue = ""
			Unapproved_Notices.Notice.TooltipValue = ""

			' Approved
			Unapproved_Notices.Approved.LinkCustomAttributes = ""
			Unapproved_Notices.Approved.HrefValue = ""
			Unapproved_Notices.Approved.TooltipValue = ""

			' Archived
			Unapproved_Notices.Archived.LinkCustomAttributes = ""
			Unapproved_Notices.Archived.HrefValue = ""
			Unapproved_Notices.Archived.TooltipValue = ""

		' ----------
		'  Edit Row
		' ----------

		ElseIf Unapproved_Notices.RowType = EW_ROWTYPE_EDIT Then ' Edit row

			' Title
			Unapproved_Notices.Title.EditCustomAttributes = ""
			Unapproved_Notices.Title.EditValue = ew_HtmlEncode(Unapproved_Notices.Title.CurrentValue)

			' Sdate
			Unapproved_Notices.Sdate.EditCustomAttributes = ""
			Unapproved_Notices.Sdate.EditValue = ew_FormatDateTime(Unapproved_Notices.Sdate.CurrentValue, 7)

			' Edate
			Unapproved_Notices.Edate.EditCustomAttributes = ""
			Unapproved_Notices.Edate.EditValue = ew_FormatDateTime(Unapproved_Notices.Edate.CurrentValue, 7)

			' Group
			Unapproved_Notices.Group.EditCustomAttributes = ""
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
			Unapproved_Notices.Group.EditValue = arwrk

			' Notice
			Unapproved_Notices.Notice.EditCustomAttributes = ""
			Unapproved_Notices.Notice.EditValue = ew_HtmlEncode(Unapproved_Notices.Notice.CurrentValue)

			' Approved
			Unapproved_Notices.Approved.EditCustomAttributes = ""
			Redim arwrk(1, 1)
			arwrk(0, 0) = "1"
			arwrk(1, 0) = ew_IIf(Unapproved_Notices.Approved.FldTagCaption(1) <> "", Unapproved_Notices.Approved.FldTagCaption(1), "Yes")
			arwrk(0, 1) = "0"
			arwrk(1, 1) = ew_IIf(Unapproved_Notices.Approved.FldTagCaption(2) <> "", Unapproved_Notices.Approved.FldTagCaption(2), "No")
			Unapproved_Notices.Approved.EditValue = arwrk

			' Archived
			Unapproved_Notices.Archived.EditCustomAttributes = ""
			Redim arwrk(1, 1)
			arwrk(0, 0) = "1"
			arwrk(1, 0) = ew_IIf(Unapproved_Notices.Archived.FldTagCaption(1) <> "", Unapproved_Notices.Archived.FldTagCaption(1), "Yes")
			arwrk(0, 1) = "0"
			arwrk(1, 1) = ew_IIf(Unapproved_Notices.Archived.FldTagCaption(2) <> "", Unapproved_Notices.Archived.FldTagCaption(2), "No")
			Unapproved_Notices.Archived.EditValue = arwrk

			' Edit refer script
			' Title

			Unapproved_Notices.Title.HrefValue = ""

			' Sdate
			Unapproved_Notices.Sdate.HrefValue = ""

			' Edate
			Unapproved_Notices.Edate.HrefValue = ""

			' Group
			Unapproved_Notices.Group.HrefValue = ""

			' Notice
			Unapproved_Notices.Notice.HrefValue = ""

			' Approved
			Unapproved_Notices.Approved.HrefValue = ""

			' Archived
			Unapproved_Notices.Archived.HrefValue = ""
		End If
		If Unapproved_Notices.RowType = EW_ROWTYPE_ADD Or Unapproved_Notices.RowType = EW_ROWTYPE_EDIT Or Unapproved_Notices.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call Unapproved_Notices.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If Unapproved_Notices.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Unapproved_Notices.Row_Rendered()
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
		If Not IsNull(Unapproved_Notices.Title.FormValue) And Unapproved_Notices.Title.FormValue&"" = "" Then
			Call ew_AddMessage(gsFormError, Language.Phrase("EnterRequiredField") & " - " & Unapproved_Notices.Title.FldCaption)
		End If
		If Not IsNull(Unapproved_Notices.Sdate.FormValue) And Unapproved_Notices.Sdate.FormValue&"" = "" Then
			Call ew_AddMessage(gsFormError, Language.Phrase("EnterRequiredField") & " - " & Unapproved_Notices.Sdate.FldCaption)
		End If
		If Not ew_CheckEuroDate(Unapproved_Notices.Sdate.FormValue) Then
			Call ew_AddMessage(gsFormError, Unapproved_Notices.Sdate.FldErrMsg)
		End If
		If Not IsNull(Unapproved_Notices.Edate.FormValue) And Unapproved_Notices.Edate.FormValue&"" = "" Then
			Call ew_AddMessage(gsFormError, Language.Phrase("EnterRequiredField") & " - " & Unapproved_Notices.Edate.FldCaption)
		End If
		If Not ew_CheckEuroDate(Unapproved_Notices.Edate.FormValue) Then
			Call ew_AddMessage(gsFormError, Unapproved_Notices.Edate.FldErrMsg)
		End If
		If Not IsNull(Unapproved_Notices.Group.FormValue) And Unapproved_Notices.Group.FormValue&"" = "" Then
			Call ew_AddMessage(gsFormError, Language.Phrase("EnterRequiredField") & " - " & Unapproved_Notices.Group.FldCaption)
		End If
		If Not IsNull(Unapproved_Notices.Notice.FormValue) And Unapproved_Notices.Notice.FormValue&"" = "" Then
			Call ew_AddMessage(gsFormError, Language.Phrase("EnterRequiredField") & " - " & Unapproved_Notices.Notice.FldCaption)
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
		sFilter = Unapproved_Notices.KeyFilter
		Unapproved_Notices.CurrentFilter  = sFilter
		sSql = Unapproved_Notices.SQL
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

			' Field Title
			Call Unapproved_Notices.Title.SetDbValue(Rs, Unapproved_Notices.Title.CurrentValue, Null, Unapproved_Notices.Title.ReadOnly)

			' Field Sdate
			Call Unapproved_Notices.Sdate.SetDbValue(Rs, ew_UnFormatDateTime(Unapproved_Notices.Sdate.CurrentValue, 7), Null, Unapproved_Notices.Sdate.ReadOnly)

			' Field Edate
			Call Unapproved_Notices.Edate.SetDbValue(Rs, ew_UnFormatDateTime(Unapproved_Notices.Edate.CurrentValue, 7), Null, Unapproved_Notices.Edate.ReadOnly)

			' Field Group
			Call Unapproved_Notices.Group.SetDbValue(Rs, Unapproved_Notices.Group.CurrentValue, Null, Unapproved_Notices.Group.ReadOnly)

			' Field Notice
			Call Unapproved_Notices.Notice.SetDbValue(Rs, Unapproved_Notices.Notice.CurrentValue, Null, Unapproved_Notices.Notice.ReadOnly)

			' Field Approved
			boolwrk = Unapproved_Notices.Approved.CurrentValue
			If boolwrk&"" <> "1" And boolwrk&"" <> "0" Then boolwrk = ew_IIf(boolwrk&"" <> "", "1", "0")
			Call Unapproved_Notices.Approved.SetDbValue(Rs, boolwrk, Null, Unapproved_Notices.Approved.ReadOnly)

			' Field Archived
			boolwrk = Unapproved_Notices.Archived.CurrentValue
			If boolwrk&"" <> "1" And boolwrk&"" <> "0" Then boolwrk = ew_IIf(boolwrk&"" <> "", "1", "0")
			Call Unapproved_Notices.Archived.SetDbValue(Rs, boolwrk, Null, Unapproved_Notices.Archived.ReadOnly)

			' Check recordset update error
			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				Rs.Close
				Set Rs = Nothing
				EditRow = False
				Exit Function
			End If

			' Call Row Updating event
			bUpdateRow = Unapproved_Notices.Row_Updating(RsOld, Rs)
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
				If Unapproved_Notices.CancelMessage <> "" Then
					FailureMessage = Unapproved_Notices.CancelMessage
					Unapproved_Notices.CancelMessage = ""
				Else
					FailureMessage = Language.Phrase("UpdateCancelled")
				End If
				EditRow = False
			End If
		End If

		' Call Row Updated event
		If EditRow Then
			Call Unapproved_Notices.Row_Updated(RsOld, RsNew)
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
