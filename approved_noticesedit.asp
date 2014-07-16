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
Dim Approved_Notices_edit
Set Approved_Notices_edit = New cApproved_Notices_edit
Set Page = Approved_Notices_edit

' Page init processing
Call Approved_Notices_edit.Page_Init()

' Page main processing
Call Approved_Notices_edit.Page_Main()
%>
<!--#include file="header.asp"-->
<script type="text/javascript">
<!--
// Create page object
var Approved_Notices_edit = new ew_Page("Approved_Notices_edit");
// page properties
Approved_Notices_edit.PageID = "edit"; // page ID
Approved_Notices_edit.FormID = "fApproved_Noticesedit"; // form ID
var EW_PAGE_ID = Approved_Notices_edit.PageID; // for backward compatibility
// extend page with ValidateForm function
Approved_Notices_edit.ValidateForm = function(fobj) {
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
			return ew_OnError(this, elm, ewLanguage.Phrase("EnterRequiredField") + " - <%= ew_JsEncode2(Approved_Notices.Title.FldCaption) %>");
		elm = fobj.elements["x" + infix + "_Author"];
		if (elm && !ew_HasValue(elm))
			return ew_OnError(this, elm, ewLanguage.Phrase("EnterRequiredField") + " - <%= ew_JsEncode2(Approved_Notices.Author.FldCaption) %>");
		elm = fobj.elements["x" + infix + "_Sdate"];
		if (elm && !ew_HasValue(elm))
			return ew_OnError(this, elm, ewLanguage.Phrase("EnterRequiredField") + " - <%= ew_JsEncode2(Approved_Notices.Sdate.FldCaption) %>");
		elm = fobj.elements["x" + infix + "_Sdate"];
		if (elm && !ew_CheckEuroDate(elm.value))
			return ew_OnError(this, elm, "<%= ew_JsEncode2(Approved_Notices.Sdate.FldErrMsg) %>");
		elm = fobj.elements["x" + infix + "_Edate"];
		if (elm && !ew_HasValue(elm))
			return ew_OnError(this, elm, ewLanguage.Phrase("EnterRequiredField") + " - <%= ew_JsEncode2(Approved_Notices.Edate.FldCaption) %>");
		elm = fobj.elements["x" + infix + "_Edate"];
		if (elm && !ew_CheckEuroDate(elm.value))
			return ew_OnError(this, elm, "<%= ew_JsEncode2(Approved_Notices.Edate.FldErrMsg) %>");
		elm = fobj.elements["x" + infix + "_Group"];
		if (elm && !ew_HasValue(elm))
			return ew_OnError(this, elm, ewLanguage.Phrase("EnterRequiredField") + " - <%= ew_JsEncode2(Approved_Notices.Group.FldCaption) %>");
		elm = fobj.elements["x" + infix + "_Notice"];
		if (elm && !ew_HasValue(elm))
			return ew_OnError(this, elm, ewLanguage.Phrase("EnterRequiredField") + " - <%= ew_JsEncode2(Approved_Notices.Notice.FldCaption) %>");
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
Approved_Notices_edit.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
<% If EW_CLIENT_VALIDATE Then %>
Approved_Notices_edit.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Approved_Notices_edit.ValidateRequired = false; // no JavaScript validation
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
<% Approved_Notices_edit.ShowPageHeader() %>
<p class="aspmaker ewTitle"><%= Language.Phrase("Edit") %>&nbsp;<%= Language.Phrase("TblTypeVIEW") %><%= Approved_Notices.TableCaption %></p>
<p class="aspmaker"><a href="<%= Approved_Notices.ReturnUrl %>"><%= Language.Phrase("GoBack") %></a></p>
<% Approved_Notices_edit.ShowMessage %>
<form name="fApproved_Noticesedit" id="fApproved_Noticesedit" action="<%= ew_CurrentPage %>" method="post" onsubmit="return Approved_Notices_edit.ValidateForm(this);">
<p>
<input type="hidden" name="a_table" id="a_table" value="Approved_Notices">
<input type="hidden" name="a_edit" id="a_edit" value="U">
<% If Approved_Notices.CurrentAction = "F" Then ' Confirm page %>
<input type="hidden" name="a_confirm" id="a_confirm" value="F">
<% End If %>
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table cellspacing="0" class="ewTable">
<% If Approved_Notices.Title.Visible Then ' Title %>
	<tr id="r_Title"<%= Approved_Notices.RowAttributes %>>
		<td class="ewTableHeader"><%= Approved_Notices.Title.FldCaption %><%= Language.Phrase("FieldRequiredIndicator") %></td>
		<td<%= Approved_Notices.Title.CellAttributes %>><span id="el_Title">
<% If Approved_Notices.CurrentAction <> "F" Then %>
<input type="text" name="x_Title" id="x_Title" size="30" maxlength="50" value="<%= Approved_Notices.Title.EditValue %>"<%= Approved_Notices.Title.EditAttributes %>>
<% Else %>
<div<%= Approved_Notices.Title.ViewAttributes %>><%= Approved_Notices.Title.ViewValue %></div>
<input type="hidden" name="x_Title" id="x_Title" value="<%= Server.HTMLEncode(Approved_Notices.Title.FormValue&"") %>">
<% End If %>
</span><%= Approved_Notices.Title.CustomMsg %></td>
	</tr>
<% End If %>
<% If Approved_Notices.Author.Visible Then ' Author %>
	<tr id="r_Author"<%= Approved_Notices.RowAttributes %>>
		<td class="ewTableHeader"><%= Approved_Notices.Author.FldCaption %><%= Language.Phrase("FieldRequiredIndicator") %></td>
		<td<%= Approved_Notices.Author.CellAttributes %>><span id="el_Author">
<% If Approved_Notices.CurrentAction <> "F" Then %>
<input type="text" name="x_Author" id="x_Author" size="30" maxlength="50" value="<%= Approved_Notices.Author.EditValue %>"<%= Approved_Notices.Author.EditAttributes %>>
<% Else %>
<div<%= Approved_Notices.Author.ViewAttributes %>><%= Approved_Notices.Author.ViewValue %></div>
<input type="hidden" name="x_Author" id="x_Author" value="<%= Server.HTMLEncode(Approved_Notices.Author.FormValue&"") %>">
<% End If %>
</span><%= Approved_Notices.Author.CustomMsg %></td>
	</tr>
<% End If %>
<% If Approved_Notices.Sdate.Visible Then ' Sdate %>
	<tr id="r_Sdate"<%= Approved_Notices.RowAttributes %>>
		<td class="ewTableHeader"><%= Approved_Notices.Sdate.FldCaption %><%= Language.Phrase("FieldRequiredIndicator") %></td>
		<td<%= Approved_Notices.Sdate.CellAttributes %>><span id="el_Sdate">
<% If Approved_Notices.CurrentAction <> "F" Then %>
<input type="text" name="x_Sdate" id="x_Sdate" value="<%= Approved_Notices.Sdate.EditValue %>"<%= Approved_Notices.Sdate.EditAttributes %>>
&nbsp;<img src="images/calendar.png" id="cal_x_Sdate" name="cal_x_Sdate" alt="<%= Language.Phrase("PickDate") %>" title="<%= Language.Phrase("PickDate") %>" style="cursor:pointer;cursor:hand;">
<script type="text/javascript">
Calendar.setup({
	inputField: "x_Sdate", // input field id
	ifFormat: "%d/%m/%Y", // date format
	button: "cal_x_Sdate" // button id
});
</script>
<% Else %>
<div<%= Approved_Notices.Sdate.ViewAttributes %>><%= Approved_Notices.Sdate.ViewValue %></div>
<input type="hidden" name="x_Sdate" id="x_Sdate" value="<%= Server.HTMLEncode(Approved_Notices.Sdate.FormValue&"") %>">
<% End If %>
</span><%= Approved_Notices.Sdate.CustomMsg %></td>
	</tr>
<% End If %>
<% If Approved_Notices.Edate.Visible Then ' Edate %>
	<tr id="r_Edate"<%= Approved_Notices.RowAttributes %>>
		<td class="ewTableHeader"><%= Approved_Notices.Edate.FldCaption %><%= Language.Phrase("FieldRequiredIndicator") %></td>
		<td<%= Approved_Notices.Edate.CellAttributes %>><span id="el_Edate">
<% If Approved_Notices.CurrentAction <> "F" Then %>
<input type="text" name="x_Edate" id="x_Edate" value="<%= Approved_Notices.Edate.EditValue %>"<%= Approved_Notices.Edate.EditAttributes %>>
&nbsp;<img src="images/calendar.png" id="cal_x_Edate" name="cal_x_Edate" alt="<%= Language.Phrase("PickDate") %>" title="<%= Language.Phrase("PickDate") %>" style="cursor:pointer;cursor:hand;">
<script type="text/javascript">
Calendar.setup({
	inputField: "x_Edate", // input field id
	ifFormat: "%d/%m/%Y", // date format
	button: "cal_x_Edate" // button id
});
</script>
<% Else %>
<div<%= Approved_Notices.Edate.ViewAttributes %>><%= Approved_Notices.Edate.ViewValue %></div>
<input type="hidden" name="x_Edate" id="x_Edate" value="<%= Server.HTMLEncode(Approved_Notices.Edate.FormValue&"") %>">
<% End If %>
</span><%= Approved_Notices.Edate.CustomMsg %></td>
	</tr>
<% End If %>
<% If Approved_Notices.Group.Visible Then ' Group %>
	<tr id="r_Group"<%= Approved_Notices.RowAttributes %>>
		<td class="ewTableHeader"><%= Approved_Notices.Group.FldCaption %><%= Language.Phrase("FieldRequiredIndicator") %></td>
		<td<%= Approved_Notices.Group.CellAttributes %>><span id="el_Group">
<% If Approved_Notices.CurrentAction <> "F" Then %>
<select id="x_Group" name="x_Group"<%= Approved_Notices.Group.EditAttributes %>>
<%
emptywrk = True
If IsArray(Approved_Notices.Group.EditValue) Then
	arwrk = Approved_Notices.Group.EditValue
	For rowcntwrk = 0 To UBound(arwrk, 2)
		If arwrk(0, rowcntwrk)&"" = Approved_Notices.Group.CurrentValue&"" Then
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
<div<%= Approved_Notices.Group.ViewAttributes %>><%= Approved_Notices.Group.ViewValue %></div>
<input type="hidden" name="x_Group" id="x_Group" value="<%= Server.HTMLEncode(Approved_Notices.Group.FormValue&"") %>">
<% End If %>
</span><%= Approved_Notices.Group.CustomMsg %></td>
	</tr>
<% End If %>
<% If Approved_Notices.Notice.Visible Then ' Notice %>
	<tr id="r_Notice"<%= Approved_Notices.RowAttributes %>>
		<td class="ewTableHeader"><%= Approved_Notices.Notice.FldCaption %><%= Language.Phrase("FieldRequiredIndicator") %></td>
		<td<%= Approved_Notices.Notice.CellAttributes %>><span id="el_Notice">
<% If Approved_Notices.CurrentAction <> "F" Then %>
<textarea name="x_Notice" id="x_Notice" cols="40" rows="4"<%= Approved_Notices.Notice.EditAttributes %>><%= Approved_Notices.Notice.EditValue %></textarea>
<script type="text/javascript">
<!--
<% If Approved_Notices.Notice.ReadOnly Then %>
new ew_ReadOnlyTextArea('x_Notice', 40*_width_multiplier, 4*_height_multiplier);
<% Else %>ew_DHTMLEditors.push(new ew_DHTMLEditor("x_Notice", function() {
	var oCKeditor = CKEDITOR.replace('x_Notice', { width: 40*_width_multiplier, height: 4*_height_multiplier, autoUpdateElement: false, baseHref: 'ckeditor/'});
	this.active = true;
}));
<% End If %>
-->
</script>
<% Else %>
<div<%= Approved_Notices.Notice.ViewAttributes %>><%= Approved_Notices.Notice.ViewValue %></div>
<input type="hidden" name="x_Notice" id="x_Notice" value="<%= Server.HTMLEncode(Approved_Notices.Notice.FormValue&"") %>">
<% End If %>
</span><%= Approved_Notices.Notice.CustomMsg %></td>
	</tr>
<% End If %>
<% If Approved_Notices.Approved.Visible Then ' Approved %>
	<tr id="r_Approved"<%= Approved_Notices.RowAttributes %>>
		<td class="ewTableHeader"><%= Approved_Notices.Approved.FldCaption %></td>
		<td<%= Approved_Notices.Approved.CellAttributes %>><span id="el_Approved">
<% If Approved_Notices.CurrentAction <> "F" Then %>
<% selwrk = ew_IIf(ew_ConvertToBool(Approved_Notices.Approved.CurrentValue), " checked=""checked""", "") %>
<input type="checkbox" name="x_Approved" id="x_Approved" value="1"<%= selwrk %><%= Approved_Notices.Approved.EditAttributes %>>
<% Else %>
<% If ew_ConvertToBool(Approved_Notices.Approved.CurrentValue) Then %>
<input type="checkbox" value="<%= Approved_Notices.Approved.ViewValue %>" checked onclick="this.form.reset();" disabled="disabled">
<% Else %>
<input type="checkbox" value="<%= Approved_Notices.Approved.ViewValue %>" onclick="this.form.reset();" disabled="disabled">
<% End If %>
<input type="hidden" name="x_Approved" id="x_Approved" value="<%= Server.HTMLEncode(Approved_Notices.Approved.FormValue&"") %>">
<% End If %>
</span><%= Approved_Notices.Approved.CustomMsg %></td>
	</tr>
<% End If %>
<% If Approved_Notices.Archived.Visible Then ' Archived %>
	<tr id="r_Archived"<%= Approved_Notices.RowAttributes %>>
		<td class="ewTableHeader"><%= Approved_Notices.Archived.FldCaption %></td>
		<td<%= Approved_Notices.Archived.CellAttributes %>><span id="el_Archived">
<% If Approved_Notices.CurrentAction <> "F" Then %>
<% selwrk = ew_IIf(ew_ConvertToBool(Approved_Notices.Archived.CurrentValue), " checked=""checked""", "") %>
<input type="checkbox" name="x_Archived" id="x_Archived" value="1"<%= selwrk %><%= Approved_Notices.Archived.EditAttributes %>>
<% Else %>
<% If ew_ConvertToBool(Approved_Notices.Archived.CurrentValue) Then %>
<input type="checkbox" value="<%= Approved_Notices.Archived.ViewValue %>" checked onclick="this.form.reset();" disabled="disabled">
<% Else %>
<input type="checkbox" value="<%= Approved_Notices.Archived.ViewValue %>" onclick="this.form.reset();" disabled="disabled">
<% End If %>
<input type="hidden" name="x_Archived" id="x_Archived" value="<%= Server.HTMLEncode(Approved_Notices.Archived.FormValue&"") %>">
<% End If %>
</span><%= Approved_Notices.Archived.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</div>
</td></tr></table>
<input type="hidden" name="x_Notice_ID" id="x_Notice_ID" value="<%= Server.HTMLEncode(Approved_Notices.Notice_ID.CurrentValue&"") %>">
<p>
<% If Approved_Notices.CurrentAction <> "F" Then ' Confirm page %>
<input type="submit" name="btnAction" id="btnAction" value="<%= ew_BtnCaption(Language.Phrase("EditBtn")) %>" onclick="this.form.a_edit.value='F';">
<% Else %>
<input type="submit" name="btnCancel" id="btnCancel" value="<%= ew_BtnCaption(Language.Phrase("CancelBtn")) %>" onclick="this.form.a_edit.value='X';">
<input type="submit" name="btnAction" id="btnAction" value="<%= ew_BtnCaption(Language.Phrase("ConfirmBtn")) %>">
<% End If %>
</form>
<% If Approved_Notices.CurrentAction <> "F" Then %>
<% End If %>
<%
Approved_Notices_edit.ShowPageFooter()
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
Set Approved_Notices_edit = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cApproved_Notices_edit

	' Page ID
	Public Property Get PageID()
		PageID = "edit"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Approved Notices"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Approved_Notices_edit"
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
		EW_PAGE_ID = "edit"

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
		If Not Security.CanEdit Then
			Call Security.SaveLastUrl()
			Call Page_Terminate("approved_noticeslist.asp")
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

	Dim DbMasterFilter, DbDetailFilter

	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()

		' Load key from QueryString
		If Request.QueryString("Notice_ID").Count > 0 Then
			Approved_Notices.Notice_ID.QueryStringValue = Request.QueryString("Notice_ID")
		End If
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			Approved_Notices.CurrentAction = ObjForm.GetValue("a_edit") ' Get action code
			Call LoadFormValues() ' Get form values

			' Validate Form
			If Not ValidateForm() Then
				Approved_Notices.CurrentAction = "" ' Form error, reset action
				FailureMessage = gsFormError
				Approved_Notices.EventCancelled = True ' Event cancelled
				Call LoadRow() ' Restore row
				Call RestoreFormValues() ' Restore form values if validate failed
			End If
		Else
			Approved_Notices.CurrentAction = "I" ' Default action is display
		End If

		' Check if valid key
		If Approved_Notices.Notice_ID.CurrentValue = "" Then Call Page_Terminate("approved_noticeslist.asp") ' Invalid key, return to list
		Select Case Approved_Notices.CurrentAction
			Case "I" ' Get a record to display
				If Not LoadRow() Then ' Load Record based on key
					FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("approved_noticeslist.asp") ' No matching record, return to list
				End If
			Case "U" ' Update
				Approved_Notices.SendEmail = True ' Send email on update success
				If EditRow() Then ' Update Record based on key
					SuccessMessage = Language.Phrase("UpdateSuccess") ' Update success
					Dim sReturnUrl
					sReturnUrl = Approved_Notices.ReturnUrl
					Call Page_Terminate(sReturnUrl) ' Return to caller
				Else
					Approved_Notices.EventCancelled = True ' Event cancelled
					Call LoadRow() ' Restore row
					Call RestoreFormValues() ' Restore form values if update failed
				End If
		End Select

		' Render the record
		If Approved_Notices.CurrentAction = "F" Then ' Confirm page
			Approved_Notices.RowType = EW_ROWTYPE_VIEW ' Render as view
		Else
			Approved_Notices.RowType = EW_ROWTYPE_EDIT ' Render as edit
		End If

		' Render row
		Call Approved_Notices.ResetAttrs()
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
		If Not Approved_Notices.Title.FldIsDetailKey Then Approved_Notices.Title.FormValue = ObjForm.GetValue("x_Title")
		If Not Approved_Notices.Author.FldIsDetailKey Then Approved_Notices.Author.FormValue = ObjForm.GetValue("x_Author")
		If Not Approved_Notices.Sdate.FldIsDetailKey Then Approved_Notices.Sdate.FormValue = ObjForm.GetValue("x_Sdate")
		If Not Approved_Notices.Sdate.FldIsDetailKey Then Approved_Notices.Sdate.CurrentValue = ew_UnFormatDateTime(Approved_Notices.Sdate.CurrentValue, 7)
		If Not Approved_Notices.Edate.FldIsDetailKey Then Approved_Notices.Edate.FormValue = ObjForm.GetValue("x_Edate")
		If Not Approved_Notices.Edate.FldIsDetailKey Then Approved_Notices.Edate.CurrentValue = ew_UnFormatDateTime(Approved_Notices.Edate.CurrentValue, 7)
		If Not Approved_Notices.Group.FldIsDetailKey Then Approved_Notices.Group.FormValue = ObjForm.GetValue("x_Group")
		If Not Approved_Notices.Notice.FldIsDetailKey Then Approved_Notices.Notice.FormValue = ObjForm.GetValue("x_Notice")
		If Not Approved_Notices.Approved.FldIsDetailKey Then Approved_Notices.Approved.FormValue = ObjForm.GetValue("x_Approved")
		If Not Approved_Notices.Archived.FldIsDetailKey Then Approved_Notices.Archived.FormValue = ObjForm.GetValue("x_Archived")
		If Not Approved_Notices.Notice_ID.FldIsDetailKey Then Approved_Notices.Notice_ID.FormValue = ObjForm.GetValue("x_Notice_ID")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadRow()
		Approved_Notices.Title.CurrentValue = Approved_Notices.Title.FormValue
		Approved_Notices.Author.CurrentValue = Approved_Notices.Author.FormValue
		Approved_Notices.Sdate.CurrentValue = Approved_Notices.Sdate.FormValue
		Approved_Notices.Sdate.CurrentValue = ew_UnFormatDateTime(Approved_Notices.Sdate.CurrentValue, 7)
		Approved_Notices.Edate.CurrentValue = Approved_Notices.Edate.FormValue
		Approved_Notices.Edate.CurrentValue = ew_UnFormatDateTime(Approved_Notices.Edate.CurrentValue, 7)
		Approved_Notices.Group.CurrentValue = Approved_Notices.Group.FormValue
		Approved_Notices.Notice.CurrentValue = Approved_Notices.Notice.FormValue
		Approved_Notices.Approved.CurrentValue = Approved_Notices.Approved.FormValue
		Approved_Notices.Archived.CurrentValue = Approved_Notices.Archived.FormValue
		Approved_Notices.Notice_ID.CurrentValue = Approved_Notices.Notice_ID.FormValue
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

			' Notice
			Approved_Notices.Notice.ViewValue = Approved_Notices.Notice.CurrentValue
			Approved_Notices.Notice.ViewCustomAttributes = ""

			' Approved
			If ew_ConvertToBool(Approved_Notices.Approved.CurrentValue) Then
				Approved_Notices.Approved.ViewValue = ew_IIf(Approved_Notices.Approved.FldTagCaption(1) <> "", Approved_Notices.Approved.FldTagCaption(1), "Yes")
			Else
				Approved_Notices.Approved.ViewValue = ew_IIf(Approved_Notices.Approved.FldTagCaption(2) <> "", Approved_Notices.Approved.FldTagCaption(2), "No")
			End If
			Approved_Notices.Approved.ViewCustomAttributes = ""

			' Archived
			If ew_ConvertToBool(Approved_Notices.Archived.CurrentValue) Then
				Approved_Notices.Archived.ViewValue = ew_IIf(Approved_Notices.Archived.FldTagCaption(1) <> "", Approved_Notices.Archived.FldTagCaption(1), "Yes")
			Else
				Approved_Notices.Archived.ViewValue = ew_IIf(Approved_Notices.Archived.FldTagCaption(2) <> "", Approved_Notices.Archived.FldTagCaption(2), "No")
			End If
			Approved_Notices.Archived.ViewCustomAttributes = ""

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

			' Notice
			Approved_Notices.Notice.LinkCustomAttributes = ""
			Approved_Notices.Notice.HrefValue = ""
			Approved_Notices.Notice.TooltipValue = ""

			' Approved
			Approved_Notices.Approved.LinkCustomAttributes = ""
			Approved_Notices.Approved.HrefValue = ""
			Approved_Notices.Approved.TooltipValue = ""

			' Archived
			Approved_Notices.Archived.LinkCustomAttributes = ""
			Approved_Notices.Archived.HrefValue = ""
			Approved_Notices.Archived.TooltipValue = ""

		' ----------
		'  Edit Row
		' ----------

		ElseIf Approved_Notices.RowType = EW_ROWTYPE_EDIT Then ' Edit row

			' Title
			Approved_Notices.Title.EditCustomAttributes = ""
			Approved_Notices.Title.EditValue = ew_HtmlEncode(Approved_Notices.Title.CurrentValue)

			' Author
			Approved_Notices.Author.EditCustomAttributes = ""
			Approved_Notices.Author.EditValue = ew_HtmlEncode(Approved_Notices.Author.CurrentValue)

			' Sdate
			Approved_Notices.Sdate.EditCustomAttributes = ""
			Approved_Notices.Sdate.EditValue = ew_FormatDateTime(Approved_Notices.Sdate.CurrentValue, 7)

			' Edate
			Approved_Notices.Edate.EditCustomAttributes = ""
			Approved_Notices.Edate.EditValue = ew_FormatDateTime(Approved_Notices.Edate.CurrentValue, 7)

			' Group
			Approved_Notices.Group.EditCustomAttributes = ""
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
			Approved_Notices.Group.EditValue = arwrk

			' Notice
			Approved_Notices.Notice.EditCustomAttributes = ""
			Approved_Notices.Notice.EditValue = ew_HtmlEncode(Approved_Notices.Notice.CurrentValue)

			' Approved
			Approved_Notices.Approved.EditCustomAttributes = ""
			Redim arwrk(1, 1)
			arwrk(0, 0) = "1"
			arwrk(1, 0) = ew_IIf(Approved_Notices.Approved.FldTagCaption(1) <> "", Approved_Notices.Approved.FldTagCaption(1), "Yes")
			arwrk(0, 1) = "0"
			arwrk(1, 1) = ew_IIf(Approved_Notices.Approved.FldTagCaption(2) <> "", Approved_Notices.Approved.FldTagCaption(2), "No")
			Approved_Notices.Approved.EditValue = arwrk

			' Archived
			Approved_Notices.Archived.EditCustomAttributes = ""
			Redim arwrk(1, 1)
			arwrk(0, 0) = "1"
			arwrk(1, 0) = ew_IIf(Approved_Notices.Archived.FldTagCaption(1) <> "", Approved_Notices.Archived.FldTagCaption(1), "Yes")
			arwrk(0, 1) = "0"
			arwrk(1, 1) = ew_IIf(Approved_Notices.Archived.FldTagCaption(2) <> "", Approved_Notices.Archived.FldTagCaption(2), "No")
			Approved_Notices.Archived.EditValue = arwrk

			' Edit refer script
			' Title

			Approved_Notices.Title.HrefValue = ""

			' Author
			Approved_Notices.Author.HrefValue = ""

			' Sdate
			Approved_Notices.Sdate.HrefValue = ""

			' Edate
			Approved_Notices.Edate.HrefValue = ""

			' Group
			Approved_Notices.Group.HrefValue = ""

			' Notice
			Approved_Notices.Notice.HrefValue = ""

			' Approved
			Approved_Notices.Approved.HrefValue = ""

			' Archived
			Approved_Notices.Archived.HrefValue = ""
		End If
		If Approved_Notices.RowType = EW_ROWTYPE_ADD Or Approved_Notices.RowType = EW_ROWTYPE_EDIT Or Approved_Notices.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call Approved_Notices.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If Approved_Notices.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Approved_Notices.Row_Rendered()
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
		If Not IsNull(Approved_Notices.Title.FormValue) And Approved_Notices.Title.FormValue&"" = "" Then
			Call ew_AddMessage(gsFormError, Language.Phrase("EnterRequiredField") & " - " & Approved_Notices.Title.FldCaption)
		End If
		If Not IsNull(Approved_Notices.Author.FormValue) And Approved_Notices.Author.FormValue&"" = "" Then
			Call ew_AddMessage(gsFormError, Language.Phrase("EnterRequiredField") & " - " & Approved_Notices.Author.FldCaption)
		End If
		If Not IsNull(Approved_Notices.Sdate.FormValue) And Approved_Notices.Sdate.FormValue&"" = "" Then
			Call ew_AddMessage(gsFormError, Language.Phrase("EnterRequiredField") & " - " & Approved_Notices.Sdate.FldCaption)
		End If
		If Not ew_CheckEuroDate(Approved_Notices.Sdate.FormValue) Then
			Call ew_AddMessage(gsFormError, Approved_Notices.Sdate.FldErrMsg)
		End If
		If Not IsNull(Approved_Notices.Edate.FormValue) And Approved_Notices.Edate.FormValue&"" = "" Then
			Call ew_AddMessage(gsFormError, Language.Phrase("EnterRequiredField") & " - " & Approved_Notices.Edate.FldCaption)
		End If
		If Not ew_CheckEuroDate(Approved_Notices.Edate.FormValue) Then
			Call ew_AddMessage(gsFormError, Approved_Notices.Edate.FldErrMsg)
		End If
		If Not IsNull(Approved_Notices.Group.FormValue) And Approved_Notices.Group.FormValue&"" = "" Then
			Call ew_AddMessage(gsFormError, Language.Phrase("EnterRequiredField") & " - " & Approved_Notices.Group.FldCaption)
		End If
		If Not IsNull(Approved_Notices.Notice.FormValue) And Approved_Notices.Notice.FormValue&"" = "" Then
			Call ew_AddMessage(gsFormError, Language.Phrase("EnterRequiredField") & " - " & Approved_Notices.Notice.FldCaption)
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
		sFilter = Approved_Notices.KeyFilter
		Approved_Notices.CurrentFilter  = sFilter
		sSql = Approved_Notices.SQL
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
			Call Approved_Notices.Title.SetDbValue(Rs, Approved_Notices.Title.CurrentValue, Null, Approved_Notices.Title.ReadOnly)

			' Field Author
			Call Approved_Notices.Author.SetDbValue(Rs, Approved_Notices.Author.CurrentValue, Null, Approved_Notices.Author.ReadOnly)

			' Field Sdate
			Call Approved_Notices.Sdate.SetDbValue(Rs, ew_UnFormatDateTime(Approved_Notices.Sdate.CurrentValue, 7), Null, Approved_Notices.Sdate.ReadOnly)

			' Field Edate
			Call Approved_Notices.Edate.SetDbValue(Rs, ew_UnFormatDateTime(Approved_Notices.Edate.CurrentValue, 7), Null, Approved_Notices.Edate.ReadOnly)

			' Field Group
			Call Approved_Notices.Group.SetDbValue(Rs, Approved_Notices.Group.CurrentValue, Null, Approved_Notices.Group.ReadOnly)

			' Field Notice
			Call Approved_Notices.Notice.SetDbValue(Rs, Approved_Notices.Notice.CurrentValue, Null, Approved_Notices.Notice.ReadOnly)

			' Field Approved
			boolwrk = Approved_Notices.Approved.CurrentValue
			If boolwrk&"" <> "1" And boolwrk&"" <> "0" Then boolwrk = ew_IIf(boolwrk&"" <> "", "1", "0")
			Call Approved_Notices.Approved.SetDbValue(Rs, boolwrk, Null, Approved_Notices.Approved.ReadOnly)

			' Field Archived
			boolwrk = Approved_Notices.Archived.CurrentValue
			If boolwrk&"" <> "1" And boolwrk&"" <> "0" Then boolwrk = ew_IIf(boolwrk&"" <> "", "1", "0")
			Call Approved_Notices.Archived.SetDbValue(Rs, boolwrk, Null, Approved_Notices.Archived.ReadOnly)

			' Check recordset update error
			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				Rs.Close
				Set Rs = Nothing
				EditRow = False
				Exit Function
			End If

			' Call Row Updating event
			bUpdateRow = Approved_Notices.Row_Updating(RsOld, Rs)
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
				If Approved_Notices.CancelMessage <> "" Then
					FailureMessage = Approved_Notices.CancelMessage
					Approved_Notices.CancelMessage = ""
				Else
					FailureMessage = Language.Phrase("UpdateCancelled")
				End If
				EditRow = False
			End If
		End If

		' Call Row Updated event
		If EditRow Then
			Call Approved_Notices.Row_Updated(RsOld, RsNew)
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
