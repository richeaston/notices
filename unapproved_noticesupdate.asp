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
Dim Unapproved_Notices_update
Set Unapproved_Notices_update = New cUnapproved_Notices_update
Set Page = Unapproved_Notices_update

' Page init processing
Call Unapproved_Notices_update.Page_Init()

' Page main processing
Call Unapproved_Notices_update.Page_Main()
%>
<!--#include file="header.asp"-->
<script type="text/javascript">
<!--
// Create page object
var Unapproved_Notices_update = new ew_Page("Unapproved_Notices_update");
// page properties
Unapproved_Notices_update.PageID = "update"; // page ID
Unapproved_Notices_update.FormID = "fUnapproved_Noticesupdate"; // form ID
var EW_PAGE_ID = Unapproved_Notices_update.PageID; // for backward compatibility
// extend page with ValidateForm function
Unapproved_Notices_update.ValidateForm = function(fobj) {
	ew_PostAutoSuggest(fobj);
	if (!this.ValidateRequired)
		return true; // ignore validation
	if (fobj.a_confirm && fobj.a_confirm.value == "F")
		return true;
	if (!ew_UpdateSelected(fobj)) {
		alert(ewLanguage.Phrase("NoFieldSelected"));
		return false;
	}
	var uelm;
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
	return true;
}
// extend page with Form_CustomValidate function
Unapproved_Notices_update.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
Unapproved_Notices_update.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
Unapproved_Notices_update.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Unapproved_Notices_update.ValidateRequired = false; // no JavaScript validation
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
<% Unapproved_Notices_update.ShowPageHeader() %>
<p><a class="btn btn-inverse" href="<%= Unapproved_Notices.ReturnUrl %>"><i class="icon-arrow-left icon-white"></i>&nbsp;<%= Language.Phrase("BackToList") %></a></p>
<% Unapproved_Notices_update.ShowMessage %>
<form name="fUnapproved_Noticesupdate" id="fUnapproved_Noticesupdate" action="<%= ew_CurrentPage %>" method="post" onsubmit="return Unapproved_Notices_update.ValidateForm(this);">
<p>
<input type="hidden" checked name="t" id="t" value="Unapproved_Notices">
<input type="hidden" checked name="a_update" id="a_update" value="U">
<% If Unapproved_Notices.CurrentAction = "F" Then ' Confirm page %>
<input type="hidden" checked name="a_confirm" id="a_confirm" value="F">
<% End If %>
<% For i = 0 to UBound(Unapproved_Notices_update.RecKeys) %>
<input type="hidden" checked name="key_m" id="key_m" value="<%= ew_HtmlEncode(ew_GetKeyValue(Unapproved_Notices_update.RecKeys(i))) %>">
<% Next %>
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<div class="well">
<table cellspacing="0" class="ewTable ewTableSeparate ">
	<tr class="ewTableHeader">
		<td><input type="hidden" checked name="u" id="u"  onclick="ew_SelectAll(this);"></td>
		<td></td>
		<td></td>
	</tr>
<% If Unapproved_Notices.Approved.Visible Then ' Approved %>
	<tr id="r_Approved"<%= Unapproved_Notices.RowAttributes %>>
		<td<%= Unapproved_Notices.Approved.CellAttributes %>>
<% If Unapproved_Notices.CurrentAction <> "F" Then %>
<input type="checkbox" readonly checked name="u_Approved" id="u_Approved" value="1"<% If Unapproved_Notices.Approved.MultiUpdate = "1" Then Response.Write " checked=""checked""" %>>
<% Else %>
<input type="hidden" checked onclick="this.form.reset();" disabled="disabled"<% If Unapproved_Notices.Approved.MultiUpdate = "1" Then Response.Write " checked=""checked""" %>>
<input type="hidden" checked name="u_Approved" id="u_Approved" value="<%= Unapproved_Notices.Approved.MultiUpdate %>">
<% End If %>
</td>
		<td<%= Unapproved_Notices.Approved.CellAttributes %>><%= Unapproved_Notices.Approved.FldCaption %></td>
		<td<%= Unapproved_Notices.Approved.CellAttributes %>><span id="el_Approved">
<% If Unapproved_Notices.CurrentAction <> "F" Then %>
<% selwrk = ew_IIf(ew_ConvertToBool(Unapproved_Notices.Approved.CurrentValue), " checked=""checked""", "") %>
<input type="checkbox" readonly Checked name="x_Approved" id="x_Approved" value="1"<%= selwrk %><%= Unapproved_Notices.Approved.EditAttributes %>>
<% Else %>
<% If ew_ConvertToBool(Unapproved_Notices.Approved.CurrentValue) Then %>
<input type="checkbox" checked value="<%= Unapproved_Notices.Approved.ViewValue %>" onclick="this.form.reset();" disabled="disabled">
<% Else %>
<input type="checkbox"  value="<%= Unapproved_Notices.Approved.ViewValue %>" checked onclick="this.form.reset();" disabled="disabled">
<% End If %>
<input type="hidden" checked name="x_Approved" id="x_Approved" value="<%= Server.HTMLEncode(Unapproved_Notices.Approved.FormValue&"") %>">
<% End If %>
</span><%= Unapproved_Notices.Approved.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</td></tr></table>
<p>
<% If Unapproved_Notices.CurrentAction <> "F" Then ' Confirm page %>
<input type="submit" name="btnAction" id="btnAction" value="<%= ew_BtnCaption(Language.Phrase("UpdateBtn")) %>" onclick="this.form.a_update.value='F';">
<% Else %>
<input type="submit" name="btnCancel" id="btnCancel" value="<%= ew_BtnCaption(Language.Phrase("CancelBtn")) %>" onclick="this.form.a_update.value='X';">
<input type="submit" name="btnAction" id="btnAction" value="<%= ew_BtnCaption(Language.Phrase("ConfirmBtn")) %>">
<% End If %>
</form>
<% If Unapproved_Notices.CurrentAction <> "F" Then %>
<% End If %>
<%
Unapproved_Notices_update.ShowPageFooter()
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
Set Unapproved_Notices_update = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cUnapproved_Notices_update

	' Page ID
	Public Property Get PageID()
		PageID = "update"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Unapproved Notices"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Unapproved_Notices_update"
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
		EW_PAGE_ID = "update"

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

	Dim RecKeys
	Dim Disabled
	Dim Recordset

	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()
		Dim sKeyName
		Dim sKey
		Dim nKeySelected
		Dim bUpdateSelected
		RecKeys = Unapproved_Notices.GetRecordKeys() ' Load record keys
		If ObjForm.GetValue("a_update")&"" <> "" Then

			' Get action
			Unapproved_Notices.CurrentAction = ObjForm.GetValue("a_update")
			Call LoadFormValues() ' Get form values

			' Validate form
			If Not ValidateForm() Then
				Unapproved_Notices.CurrentAction = "I" ' Form error, reset action
				FailureMessage = gsFormError
			End If
		Else
			Call LoadMultiUpdateValues() ' Load initial values to form
		End If
		If Not IsArray(RecKeys) Then
			Call Page_Terminate("unapproved_noticeslist.asp") ' No records selected, return to list
		End If
		Select Case Unapproved_Notices.CurrentAction
			Case "U" ' Update
				If UpdateRows() Then ' Update Records based on key
					SuccessMessage = Language.Phrase("UpdateSuccess") ' Set update success message
					Call Page_Terminate(Unapproved_Notices.ReturnUrl) ' Return to caller
				Else
					Call RestoreFormValues() ' Restore form values
				End If
		End Select

		' Render row
		If Unapproved_Notices.CurrentAction = "F" Then ' Confirm page
			Unapproved_Notices.RowType = EW_ROWTYPE_VIEW ' Render view
			Disabled = " disabled=""disabled"""
		Else
			Unapproved_Notices.RowType = EW_ROWTYPE_EDIT ' Render edit
			Disabled = ""
		End If

		' Render row
		Call Unapproved_Notices.ResetAttrs()
		Call RenderRow()
	End Sub

	' -----------------------------------------------------------------
	' Load initial values to form if field values are identical in all selected records
	'
	Sub LoadMultiUpdateValues()
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		Dim Rs, i
		Unapproved_Notices.CurrentFilter = Unapproved_Notices.GetKeyFilter()

		' Load recordset
		Set Rs = LoadRecordset()
		i = 1
		Do While Not Rs.Eof
			If i = 1 Then
				Unapproved_Notices.Approved.DbValue = Rs("Approved")
			Else
				If Not ew_CompareValue(Unapproved_Notices.Approved.DbValue, Rs("Approved")) Then
					Unapproved_Notices.Approved.CurrentValue = Null
				End If
			End If
			i = i + 1
			Rs.MoveNext
		Loop
		Rs.Close
		Set Rs = Nothing
	End Sub

	' -----------------------------------------------------------------
	'  Set up key value
	'
	Function SetupKeyValues(key)
		Dim sKeyFld
		Dim sWrkFilter, sFilter
		sKeyFld = key
		If Not IsNumeric(sKeyFld) Then
			SetupKeyValues = False
			Exit Function
		End If
		Unapproved_Notices.Notice_ID.CurrentValue = sKeyFld ' Set up key value
		SetupKeyValues = True
	End Function

	' -----------------------------------------------------------------
	' Update all selected rows
	'
	Function UpdateRows()
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		Dim sKey, sThisKey
		Dim Rs, RsOld, RsNew, sSql, i
		Conn.BeginTrans

		' Get old recordset
		Unapproved_Notices.CurrentFilter = Unapproved_Notices.GetKeyFilter()
		sSql = Unapproved_Notices.SQL
		Set Rs = Server.CreateObject("ADODB.Recordset")
		Rs.CursorLocation = EW_CURSORLOCATION
		Rs.Open sSql, Conn, 1, EW_RECORDSET_LOCKTYPE
		Set RsOld = ew_CloneRs(Rs)

		' Update all rows
		sKey = ""
		For i = 0 to UBound(RecKeys)
			If SetupKeyValues(RecKeys(i)) Then
				sThisKey = RecKeys(i)
				Unapproved_Notices.SendEmail = False ' Do not send email on update success
				UpdateRows = EditRow() ' Update this row
			Else
				UpdateRows = False
			End If
			If Not UpdateRows Then Exit For ' Update failed
			If sKey <> "" Then sKey = sKey & ", "
			sKey = sKey & sThisKey
		Next
		If UpdateRows Then
			Conn.CommitTrans ' Commit transaction

			' Get new recordset
			Set Rs = Conn.Execute(sSql)
			Set RsNew = ew_CloneRs(Rs)
		Else
			Conn.RollbackTrans ' Rollback transaction
		End If
		Set Rs = Nothing
		Set RsOld = Nothing
		Set RsNew = Nothing
	End Function

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
		If Not Unapproved_Notices.Approved.FldIsDetailKey Then Unapproved_Notices.Approved.FormValue = ObjForm.GetValue("x_Approved")
		Unapproved_Notices.Approved.MultiUpdate = ObjForm.GetValue("u_Approved")
		If Not Unapproved_Notices.Notice_ID.FldIsDetailKey Then Unapproved_Notices.Notice_ID.FormValue = ObjForm.GetValue("x_Notice_ID")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Unapproved_Notices.Approved.CurrentValue = Unapproved_Notices.Approved.FormValue
		Unapproved_Notices.Notice_ID.CurrentValue = Unapproved_Notices.Notice_ID.FormValue
	End Function

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = Unapproved_Notices.CurrentFilter
		Call Unapproved_Notices.Recordset_Selecting(sFilter)
		Unapproved_Notices.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = Unapproved_Notices.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call Unapproved_Notices.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

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

			' View refer script
			' Approved

			Unapproved_Notices.Approved.LinkCustomAttributes = ""
			Unapproved_Notices.Approved.HrefValue = ""
			Unapproved_Notices.Approved.TooltipValue = ""

		' ----------
		'  Edit Row
		' ----------

		ElseIf Unapproved_Notices.RowType = EW_ROWTYPE_EDIT Then ' Edit row

			' Approved
			Unapproved_Notices.Approved.EditCustomAttributes = ""
			Redim arwrk(1, 1)
			arwrk(0, 0) = "1"
			arwrk(1, 0) = ew_IIf(Unapproved_Notices.Approved.FldTagCaption(1) <> "", Unapproved_Notices.Approved.FldTagCaption(1), "Yes")
			arwrk(0, 1) = "0"
			arwrk(1, 1) = ew_IIf(Unapproved_Notices.Approved.FldTagCaption(2) <> "", Unapproved_Notices.Approved.FldTagCaption(2), "No")
			Unapproved_Notices.Approved.EditValue = arwrk

			' Edit refer script
			' Approved

			Unapproved_Notices.Approved.HrefValue = ""
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
		Dim lUpdateCnt
		lUpdateCnt = 0
		If Unapproved_Notices.Approved.MultiUpdate = "1" Then lUpdateCnt = lUpdateCnt + 1
		If lUpdateCnt = 0 Then
			gsFormError = Language.Phrase("NoFieldSelected")
			ValidateForm = False
			Exit Function
		End If

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

			' Field Approved
			boolwrk = Unapproved_Notices.Approved.CurrentValue
			If boolwrk&"" <> "1" And boolwrk&"" <> "0" Then boolwrk = ew_IIf(boolwrk&"" <> "", "1", "0")
			Call Unapproved_Notices.Approved.SetDbValue(Rs, boolwrk, Null, Unapproved_Notices.Approved.ReadOnly Or Unapproved_Notices.Approved.MultiUpdate&"" <> "1")

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
