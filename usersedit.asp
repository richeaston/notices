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
Dim Users_edit
Set Users_edit = New cUsers_edit
Set Page = Users_edit

' Page init processing
Call Users_edit.Page_Init()

' Page main processing
Call Users_edit.Page_Main()
%>
<!--#include file="header.asp"-->
<script type="text/javascript">
<!--
// Create page object
var Users_edit = new ew_Page("Users_edit");
// page properties
Users_edit.PageID = "edit"; // page ID
Users_edit.FormID = "fUsersedit"; // form ID
var EW_PAGE_ID = Users_edit.PageID; // for backward compatibility
// extend page with ValidateForm function
Users_edit.ValidateForm = function(fobj) {
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
Users_edit.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
Users_edit.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
Users_edit.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Users_edit.ValidateRequired = false; // no JavaScript validation
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
<% Users_edit.ShowPageHeader() %>
<p class="aspmaker"><a class="btn btn-inverse" href="<%= Users.ReturnUrl %>"><i class="icon-backward icon-white"></i>&nbsp;<%= Language.Phrase("GoBack") %></a></p>
<% Users_edit.ShowMessage %>
<form name="fUsersedit" id="fUsersedit" action="<%= ew_CurrentPage %>" method="post" onsubmit="return Users_edit.ValidateForm(this);">
<p>
<input type="hidden" name="a_table" id="a_table" value="Users">
<input type="hidden" name="a_edit" id="a_edit" value="U">
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel well">
<table cellspacing="0" class="ewTable">
<% If Users.Password.Visible Then ' Password %>
	<tr id="r_Password"<%= Users.RowAttributes %>>
		<td class="ewTableHeader"><%= Users.Password.FldCaption %></td>
		<td<%= Users.Password.CellAttributes %>><span id="el_Password">
<input type="password" name="x_Password" id="x_Password" value="<%= Users.Password.EditValue %>" size="30" maxlength="50"<%= Users.Password.EditAttributes %>>
</span><%= Users.Password.CustomMsg %></td>
	</tr>
<% End If %>
<% If Users.zEmail.Visible Then ' Email %>
	<tr id="r_zEmail"<%= Users.RowAttributes %>>
		<td class="ewTableHeader"><%= Users.zEmail.FldCaption %></td>
		<td<%= Users.zEmail.CellAttributes %>><span id="el_zEmail">
<input type="text" name="x_zEmail" id="x_zEmail" size="30" maxlength="255" value="<%= Users.zEmail.EditValue %>"<%= Users.zEmail.EditAttributes %>>
</span><%= Users.zEmail.CustomMsg %></td>
	</tr>
<% End If %>
<% If Users.Permissions.Visible Then ' Permissions %>
	<tr id="r_Permissions"<%= Users.RowAttributes %>>
		<td class="ewTableHeader"><%= Users.Permissions.FldCaption %></td>
		<td<%= Users.Permissions.CellAttributes %>><span id="el_Permissions">
<% If Not Security.IsAdmin And Security.IsLoggedIn() Then ' Non system admin %>
<div<%= Users.Permissions.ViewAttributes %>><%= Users.Permissions.EditValue %></div>
<% Else %>
<select id="x_Permissions" name="x_Permissions"<%= Users.Permissions.EditAttributes %>>
<%
emptywrk = True
If IsArray(Users.Permissions.EditValue) Then
	arwrk = Users.Permissions.EditValue
	For rowcntwrk = 0 To UBound(arwrk, 2)
		If arwrk(0, rowcntwrk)&"" = Users.Permissions.CurrentValue&"" Then
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
<% End If %>
</span><%= Users.Permissions.CustomMsg %></td>
	</tr>
<% End If %>
<% If Users.Active.Visible Then ' Active %>
	<tr id="r_Active"<%= Users.RowAttributes %>>
		<td class="ewTableHeader"><%= Users.Active.FldCaption %></td>
		<td<%= Users.Active.CellAttributes %>><span id="el_Active">
<% selwrk = ew_IIf(ew_ConvertToBool(Users.Active.CurrentValue), " checked=""checked""", "") %>
<input type="checkbox" name="x_Active" id="x_Active" value="1"<%= selwrk %><%= Users.Active.EditAttributes %>>
</span><%= Users.Active.CustomMsg %></td>
	</tr>
<% End If %>
<% If Users.Profile.Visible Then ' Profile %>
	<tr id="r_Profile"<%= Users.RowAttributes %>>
		<td class="ewTableHeader"><%= Users.Profile.FldCaption %></td>
		<td<%= Users.Profile.CellAttributes %>><span id="el_Profile">
<textarea name="x_Profile" id="x_Profile" cols="35" rows="4"<%= Users.Profile.EditAttributes %>><%= Users.Profile.EditValue %></textarea>
</span><%= Users.Profile.CustomMsg %></td>
	</tr>
<% End If %>
<% If Users.Theme.Visible Then ' Theme %>
	<tr id="r_Theme"<%= Users.RowAttributes %>>
		<td class="ewTableHeader"><%= Users.Theme.FldCaption %></td>
		<td<%= Users.Theme.CellAttributes %>><span id="el_Theme">
<select id="x_Theme" name="x_Theme"<%= Users.Theme.EditAttributes %>>
<%
emptywrk = True
If IsArray(Users.Theme.EditValue) Then
	arwrk = Users.Theme.EditValue
	For rowcntwrk = 0 To UBound(arwrk, 2)
		If arwrk(0, rowcntwrk)&"" = Users.Theme.CurrentValue&"" Then
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
</span><%= Users.Theme.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</div>
</td></tr></table>
<input type="hidden" name="x_Username" id="x_Username" value="<%= Server.HTMLEncode(Users.Username.CurrentValue&"") %>">
<p>
<input class="btn btn-primary" type="submit" name="btnAction" id="btnAction" value="<%= ew_BtnCaption(Language.Phrase("EditBtn")) %>">

</form>
<%
Users_edit.ShowPageFooter()
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
Set Users_edit = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cUsers_edit

	' Page ID
	Public Property Get PageID()
		PageID = "edit"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Users"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Users_edit"
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
		EW_PAGE_ID = "edit"

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
		If Not Security.CanEdit Then
			Call Security.SaveLastUrl()
			Call Page_Terminate("userslist.asp")
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

	Dim DbMasterFilter, DbDetailFilter

	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()

		' Load key from QueryString
		If Request.QueryString("Username").Count > 0 Then
			Users.Username.QueryStringValue = Request.QueryString("Username")
		End If
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			Users.CurrentAction = ObjForm.GetValue("a_edit") ' Get action code
			Call LoadFormValues() ' Get form values

			' Validate Form
			If Not ValidateForm() Then
				Users.CurrentAction = "" ' Form error, reset action
				FailureMessage = gsFormError
				Users.EventCancelled = True ' Event cancelled
				Call LoadRow() ' Restore row
				Call RestoreFormValues() ' Restore form values if validate failed
			End If
		Else
			Users.CurrentAction = "I" ' Default action is display
		End If

		' Check if valid key
		If Users.Username.CurrentValue = "" Then Call Page_Terminate("userslist.asp") ' Invalid key, return to list
		Select Case Users.CurrentAction
			Case "I" ' Get a record to display
				If Not LoadRow() Then ' Load Record based on key
					FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("userslist.asp") ' No matching record, return to list
				End If
			Case "U" ' Update
				Users.SendEmail = True ' Send email on update success
				If EditRow() Then ' Update Record based on key
					SuccessMessage = Language.Phrase("UpdateSuccess") ' Update success
					Dim sReturnUrl
					sReturnUrl = Users.ReturnUrl
					Call Page_Terminate(sReturnUrl) ' Return to caller
				Else
					Users.EventCancelled = True ' Event cancelled
					Call LoadRow() ' Restore row
					Call RestoreFormValues() ' Restore form values if update failed
				End If
		End Select

		' Render the record
		Users.RowType = EW_ROWTYPE_EDIT ' Render as edit

		' Render row
		Call Users.ResetAttrs()
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
		If Not Users.Password.FldIsDetailKey Then Users.Password.FormValue = ObjForm.GetValue("x_Password")
		If Not Users.zEmail.FldIsDetailKey Then Users.zEmail.FormValue = ObjForm.GetValue("x_zEmail")
		If Not Users.Permissions.FldIsDetailKey Then Users.Permissions.FormValue = ObjForm.GetValue("x_Permissions")
		If Not Users.Active.FldIsDetailKey Then Users.Active.FormValue = ObjForm.GetValue("x_Active")
		If Not Users.Profile.FldIsDetailKey Then Users.Profile.FormValue = ObjForm.GetValue("x_Profile")
		If Not Users.Theme.FldIsDetailKey Then Users.Theme.FormValue = ObjForm.GetValue("x_Theme")
		If Not Users.Username.FldIsDetailKey Then Users.Username.FormValue = ObjForm.GetValue("x_Username")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadRow()
		Users.Password.CurrentValue = Users.Password.FormValue
		Users.zEmail.CurrentValue = Users.zEmail.FormValue
		Users.Permissions.CurrentValue = Users.Permissions.FormValue
		Users.Active.CurrentValue = Users.Active.FormValue
		Users.Profile.CurrentValue = Users.Profile.FormValue
		Users.Theme.CurrentValue = Users.Theme.FormValue
		Users.Username.CurrentValue = Users.Username.FormValue
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

			' Password
			Users.Password.ViewValue = "********"
			Users.Password.ViewCustomAttributes = ""

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

			' Active
			If ew_ConvertToBool(Users.Active.CurrentValue) Then
				Users.Active.ViewValue = ew_IIf(Users.Active.FldTagCaption(1) <> "", Users.Active.FldTagCaption(1), "Yes")
			Else
				Users.Active.ViewValue = ew_IIf(Users.Active.FldTagCaption(2) <> "", Users.Active.FldTagCaption(2), "No")
			End If
			Users.Active.ViewCustomAttributes = ""

			' Profile
			Users.Profile.ViewValue = Users.Profile.CurrentValue
			Users.Profile.ViewCustomAttributes = ""

			' Theme
			If Users.Theme.CurrentValue & "" <> "" Then
				sFilterWrk = "[Theme_Name] = '" & ew_AdjustSql(Users.Theme.CurrentValue) & "'"
			sSqlWrk = "SELECT DISTINCT [Theme_Name] FROM [Themes]"
			sWhereWrk = ""
			Call ew_AddFilter(sWhereWrk, sFilterWrk)
			If sWhereWrk <> "" Then sSqlWrk = sSqlWrk & " WHERE " & sWhereWrk
			sSqlWrk = sSqlWrk & " ORDER BY [Theme_Name] Asc"
				Set RsWrk = Conn.Execute(sSqlWrk)
				If Not RsWrk.Eof Then
					Users.Theme.ViewValue = RsWrk("Theme_Name")
				Else
					Users.Theme.ViewValue = Users.Theme.CurrentValue
				End If
				RsWrk.Close
				Set RsWrk = Nothing
			Else
				Users.Theme.ViewValue = Null
			End If
			Users.Theme.ViewCustomAttributes = ""

			' View refer script
			' Password

			Users.Password.LinkCustomAttributes = ""
			Users.Password.HrefValue = ""
			Users.Password.TooltipValue = ""

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

			' Active
			Users.Active.LinkCustomAttributes = ""
			Users.Active.HrefValue = ""
			Users.Active.TooltipValue = ""

			' Profile
			Users.Profile.LinkCustomAttributes = ""
			Users.Profile.HrefValue = ""
			Users.Profile.TooltipValue = ""

			' Theme
			Users.Theme.LinkCustomAttributes = ""
			Users.Theme.HrefValue = ""
			Users.Theme.TooltipValue = ""

		' ----------
		'  Edit Row
		' ----------

		ElseIf Users.RowType = EW_ROWTYPE_EDIT Then ' Edit row

			' Password
			Users.Password.EditCustomAttributes = ""
			Users.Password.EditValue = ew_HtmlEncode(Users.Password.CurrentValue)

			' Email
			Users.zEmail.EditCustomAttributes = ""
			Users.zEmail.EditValue = ew_HtmlEncode(Users.zEmail.CurrentValue)

			' Permissions
			Users.Permissions.EditCustomAttributes = ""
			If Not Security.CanAdmin Then ' System admin
				Users.Permissions.EditValue = "********"
			Else
				sFilterWrk = ""
			sSqlWrk = "SELECT [UserLevelID], [UserLevelName] AS [DispFld], '' AS [Disp2Fld], '' AS [Disp3Fld], '' AS [Disp4Fld], '' AS [SelectFilterFld] FROM [UserLevels]"
			sWhereWrk = ""
			Call ew_AddFilter(sWhereWrk, sFilterWrk)
			If sWhereWrk <> "" Then sSqlWrk = sSqlWrk & " WHERE " & sWhereWrk
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
			Users.Permissions.EditValue = arwrk
			End If

			' Active
			Users.Active.EditCustomAttributes = ""
			Redim arwrk(1, 1)
			arwrk(0, 0) = "1"
			arwrk(1, 0) = ew_IIf(Users.Active.FldTagCaption(1) <> "", Users.Active.FldTagCaption(1), "Yes")
			arwrk(0, 1) = "0"
			arwrk(1, 1) = ew_IIf(Users.Active.FldTagCaption(2) <> "", Users.Active.FldTagCaption(2), "No")
			Users.Active.EditValue = arwrk

			' Profile
			Users.Profile.EditCustomAttributes = ""
			Users.Profile.EditValue = ew_HtmlEncode(Users.Profile.CurrentValue)

			' Theme
			Users.Theme.EditCustomAttributes = ""
				sFilterWrk = ""
			sSqlWrk = "SELECT DISTINCT [Theme_Name], [Theme_Name] AS [DispFld], '' AS [Disp2Fld], '' AS [Disp3Fld], '' AS [Disp4Fld], '' AS [SelectFilterFld] FROM [Themes]"
			sWhereWrk = ""
			Call ew_AddFilter(sWhereWrk, sFilterWrk)
			If sWhereWrk <> "" Then sSqlWrk = sSqlWrk & " WHERE " & sWhereWrk
			sSqlWrk = sSqlWrk & " ORDER BY [Theme_Name] Asc"
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
			Users.Theme.EditValue = arwrk

			' Edit refer script
			' Password

			Users.Password.HrefValue = ""

			' Email
			If Not ew_Empty(Users.zEmail.CurrentValue) Then
				Users.zEmail.HrefValue = "mailto:" & ew_IIf(Users.zEmail.EditValue<>"", Users.zEmail.EditValue, Users.zEmail.CurrentValue)
				Users.zEmail.LinkAttrs.AddAttribute "target", "", True ' Add target
				If Users.Export <> "" Then Users.zEmail.HrefValue = ew_ConvertFullUrl(Users.zEmail.HrefValue)
			Else
				Users.zEmail.HrefValue = ""
			End If

			' Permissions
			Users.Permissions.HrefValue = ""

			' Active
			Users.Active.HrefValue = ""

			' Profile
			Users.Profile.HrefValue = ""

			' Theme
			Users.Theme.HrefValue = ""
		End If
		If Users.RowType = EW_ROWTYPE_ADD Or Users.RowType = EW_ROWTYPE_EDIT Or Users.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call Users.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If Users.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Users.Row_Rendered()
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
		sFilter = Users.KeyFilter
		Users.CurrentFilter  = sFilter
		sSql = Users.SQL
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

			' Field Password
			If Not EW_CASE_SENSITIVE_PASSWORD And Not IsNull(Users.Password.CurrentValue) Then Users.Password.CurrentValue = LCase(Users.Password.CurrentValue)
			If EW_ENCRYPTED_PASSWORD And Not IsNull(Users.Password.CurrentValue) And (Users.Password.CurrentValue <> Rs("Password")) Then Users.Password.CurrentValue = MD5(Users.Password.CurrentValue)
			Call Users.Password.SetDbValue(Rs, Users.Password.CurrentValue, Null, Users.Password.ReadOnly)

			' Field Email
			Call Users.zEmail.SetDbValue(Rs, Users.zEmail.CurrentValue, Null, Users.zEmail.ReadOnly)

			' Field Permissions
						If Security.CanAdmin Then ' System admin
			Call Users.Permissions.SetDbValue(Rs, Users.Permissions.CurrentValue, Null, Users.Permissions.ReadOnly)
			End If

			' Field Active
			boolwrk = Users.Active.CurrentValue
			If boolwrk&"" <> "1" And boolwrk&"" <> "0" Then boolwrk = ew_IIf(boolwrk&"" <> "", "1", "0")
			Call Users.Active.SetDbValue(Rs, boolwrk, Null, Users.Active.ReadOnly)

			' Field Profile
			Call Users.Profile.SetDbValue(Rs, Users.Profile.CurrentValue, Null, Users.Profile.ReadOnly)

			' Field Theme
			Call Users.Theme.SetDbValue(Rs, Users.Theme.CurrentValue, Null, Users.Theme.ReadOnly)

			' Check recordset update error
			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				Rs.Close
				Set Rs = Nothing
				EditRow = False
				Exit Function
			End If

			' Call Row Updating event
			bUpdateRow = Users.Row_Updating(RsOld, Rs)
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
				If Users.CancelMessage <> "" Then
					FailureMessage = Users.CancelMessage
					Users.CancelMessage = ""
				Else
					FailureMessage = Language.Phrase("UpdateCancelled")
				End If
				EditRow = False
			End If
		End If

		' Call Row Updated event
		If EditRow Then
			Call Users.Row_Updated(RsOld, RsNew)
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

	' Write Audit Trail (edit page)
	Sub WriteAuditTrailOnEdit(RsOld, RsNew)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		Dim table
		table = "Users"

		' Get key value
		Dim sKey
		sKey = ""
		If sKey <> "" Then sKey = sKey & EW_COMPOSITE_KEY_SEPARATOR
		sKey = sKey & RsNew.Fields("Username")

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
