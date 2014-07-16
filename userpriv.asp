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
Dim userpriv
Set userpriv = New cuserpriv
Set Page = userpriv

' Page init processing
Call userpriv.Page_Init()

' Page main processing
Call userpriv.Page_Main()
%>
<!--#include file="header.asp"-->
<script type="text/javascript">
<!--
// Create page object
var userpriv = new ew_Page("userpriv");
// page properties
userpriv.PageID = "userpriv"; // page ID
userpriv.FormID = "fUserLevelsuserpriv"; // form ID
var EW_PAGE_ID = userpriv.PageID; // for backward compatibility
// extend page with Form_CustomValidate function
userpriv.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
userpriv.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
userpriv.ValidateRequired = true; // uses JavaScript validation
<% Else %>
userpriv.ValidateRequired = false; // no JavaScript validation
<% End If %>
//-->
</script>
<script language="JavaScript" type="text/javascript">
<!--
// Write your client script here, no need to add script tags.
//-->
</script>
<p class="aspmaker ewTitle"><%= Language.Phrase("UserLevelPermission") %></p>
<p class="aspmaker"><a href="userlevelslist.asp"><%= Language.Phrase("BackToList") %></a></p>
<p class="aspmaker"><%= Language.Phrase("UserLevel") %>:&nbsp;<%= Security.GetUserLevelName(UserLevels.UserLevelID.CurrentValue) %>(<%= UserLevels.UserLevelID.CurrentValue %>)</p>
<% userpriv.ShowMessage %>
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<form name="userpriv" id="userpriv" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="t" id="t" value="UserLevels">
<input type="hidden" name="a_edit" id="a_edit" value="U">
<!-- hidden tag for User Level ID -->
<input type="hidden" name="x_UserLevelID" id="x_UserLevelID" value="<%= UserLevels.UserLevelID.CurrentValue %>">
<table cellspacing="0" data-rowhighlightclass="ewTableHighlightRow" data-rowselectclass="ewTableSelectRow" data-roweditclass="ewTableEditRow" class="ewTable ewTableSeparate">
	<thead>
	<tr class="ewTableHeader">
		<td><%= Language.Phrase("TableOrView") %></td>
		<td><%= Language.Phrase("PermissionAddCopy") %><input type="checkbox" name="Add" id="Add" onclick="ew_SelectAll(this);"<%= userpriv.Disabled %>></td>
		<td><%= Language.Phrase("PermissionDelete") %><input type="checkbox" name="Delete" id="Delete" onclick="ew_SelectAll(this);"<%= userpriv.Disabled %>></td>
		<td><%= Language.Phrase("PermissionEdit") %><input type="checkbox" name="Edit" id="Edit" onclick="ew_SelectAll(this);"<%= userpriv.Disabled %>></td>
<% If EW_USER_LEVEL_COMPAT Then %>
		<td><%= Language.Phrase("PermissionListSearchView") %><input type="checkbox" name="List" id="List" onclick="ew_SelectAll(this);"<%= userpriv.Disabled %>></td>
<% Else %>
		<td><%= Language.Phrase("PermissionList") %><input type="checkbox" name="List" id="List" onclick="ew_SelectAll(this);"<%= userpriv.Disabled %>></td>
		<td><%= Language.Phrase("PermissionView") %><input type="checkbox" name="View" id="View" onclick="ew_SelectAll(this);"<%= userpriv.Disabled %>></td>
		<td><%= Language.Phrase("PermissionSearch") %><input type="checkbox" name="Search" id="Search" onclick="ew_SelectAll(this);"<%= userpriv.Disabled %>></td>
<% End If %>
	</tr>
	</thead>
	<tbody>
<%
For i = LBound(EW_USER_LEVEL_TABLE_NAME) to UBound(EW_USER_LEVEL_TABLE_NAME)
	userpriv.TempPriv = Security.GetUserLevelPrivEx(EW_USER_LEVEL_TABLE_NAME(i), UserLevels.UserLevelID.CurrentValue)

	' Set row properties
	Call UserLevels.ResetAttrs()
	UserLevels.RowAttrs.AddAttributes Array(Array("onmouseover", "ew_MouseOver(event, this);"), Array("onmouseout", "ew_MouseOut(event, this);"), Array("onclick", "ew_Click(event, this);"))
%>
	<tr<%= UserLevels.RowAttributes %>>
		<td><span class="aspmaker"><%= userpriv.GetTableCaption(i) %></span></td>
		<td align="center"><input type="checkbox" name="Add_<%= i %>" id="Add_<%= i %>" value="1"<% If (userpriv.TempPriv And EW_ALLOW_ADD) = EW_ALLOW_ADD Then %> checked="checked"<% End If %><%= userpriv.Disabled %>></td>
		<td align="center"><input type="checkbox" name="Delete_<%= i %>" id="Delete_<%= i %>" value="2"<% If (userpriv.TempPriv And EW_ALLOW_DELETE) = EW_ALLOW_DELETE Then %> checked="checked"<% End If %><%= userpriv.Disabled %>></td>
		<td align="center"><input type="checkbox" name="Edit_<%= i %>" id="Edit_<%= i %>" value="4"<% If (userpriv.TempPriv And EW_ALLOW_EDIT) = EW_ALLOW_EDIT Then %> checked="checked"<% End If %><%= userpriv.Disabled %>></td>
<% If EW_USER_LEVEL_COMPAT Then %>
		<td align="center"><input type="checkbox" name="List_<%= i %>" id="List_<%= i %>" value="8"<% If (userpriv.TempPriv And EW_ALLOW_LIST) = EW_ALLOW_LIST Then %> checked="checked"<% End If %><%= userpriv.Disabled %>></td>
<% Else %>
		<td align="center"><input type="checkbox" name="List_<%= i %>" id="List_<%= i %>" value="8"<% If (userpriv.TempPriv And EW_ALLOW_LIST) = EW_ALLOW_LIST Then %> checked="checked"<% End If %><%= userpriv.Disabled %>></td>
		<td align="center"><input type="checkbox" name="View_<%= i %>" id="View_<%= i %>" value="32"<% If (userpriv.TempPriv And EW_ALLOW_VIEW) = EW_ALLOW_VIEW Then %> checked="checked"<% End If %><%= userpriv.Disabled %>></td>
		<td align="center"><input type="checkbox" name="Search_<%= i %>" id="Search_<%= i %>" value="64"<% If (userpriv.TempPriv And EW_ALLOW_SEARCH) = EW_ALLOW_SEARCH Then %> checked="checked"<% End If %><%= userpriv.Disabled %>></td>
<% End If %>
	</tr>
<% Next %>
	</tbody>				
</table>
</div>
</td></tr></table>
<p>
<input type="submit" name="btnSubmit" id="btnSubmit" value="<%= ew_BtnCaption(Language.Phrase("Update")) %>"<%= userpriv.Disabled %>>
</form>
<script language="JavaScript" type="text/javascript">
<!--
// Write your startup script here
// document.write("page loaded");
//-->
</script>
<!--#include file="footer.asp"-->
<%

' Drop page object
Set userpriv = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cuserpriv

	' Page ID
	Public Property Get PageID()
		PageID = "userpriv"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "userpriv"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
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
		IsPageRequest = True
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
		EW_PAGE_ID = "userpriv"

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
		Call Security.LoadCurrentUserLevel("UserLevels")

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

	Dim TempPriv, Disabled
	Dim Privileges
	Dim TableNameCount
	Dim ReportLanguage

	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()

		' Try to load ASP Report Maker language file
		' Note: The langauge IDs must be the same in both projects

		If ew_ReportLanguageFolder <> "" And IsEmpty(ReportLanguage) Then
			Set ReportLanguage = New cLanguage
			ReportLanguage.LanguageFolder = ew_ReportLanguageFolder
			Call ReportLanguage.LoadPhrases()
		End If
		If Not IsArray(EW_USER_LEVEL_TABLE_NAME) Then
			FailureMessage = Language.Phrase("NoTableGenerated")
			Call Page_Terminate("userlevelslist.asp") ' Return to list
		End If
		ReDim Privileges(UBound(EW_USER_LEVEL_TABLE_NAME))

		' Get action
		If Request.Form("a_edit").Count <= 0 Then
			UserLevels.CurrentAction = "I"	' Display with input box

			' Load key from QueryString
			If Request.QueryString("UserLevelID").Count > 0 Then
				UserLevels.UserLevelID.QueryStringValue = Request.QueryString("UserLevelID")
			Else
				Call Page_Terminate("userlevelslist.asp") ' Return to list
			End If
			If UserLevels.UserLevelID.QueryStringValue = "-1" Then
				Disabled = " disabled=""disabled"""
			Else
				Disabled = ""
			End If
		Else
			UserLevels.CurrentAction = Request.Form("a_edit")

			' Get fields from form
			UserLevels.UserLevelID.FormValue = Request.Form("x_UserLevelID")
			For i = LBound(EW_USER_LEVEL_TABLE_NAME) to UBound(EW_USER_LEVEL_TABLE_NAME)
				If EW_USER_LEVEL_COMPAT Then
					Privileges(i) = CInt(Request.Form("Add_" & i)) + _
						CInt(Request.Form("Delete_" & i)) + CInt(Request.Form("Edit_" & i)) + _
						CInt(Request.Form("List_" & i))
				Else
					Privileges(i) = CInt(Request.Form("Add_" & i)) + _
						CInt(Request.Form("Delete_" & i)) + CInt(Request.Form("Edit_" & i)) + _
						CInt(Request.Form("List_" & i)) + CInt(Request.Form("View_" & i)) + _
						CInt(Request.Form("Search_" & i))
				End If
			Next
		End If
		Select Case UserLevels.CurrentAction
			Case "I" ' Display
				Security.SetUpUserLevelEx() ' Get all user level info
			Case "U" ' Update
				If EditRow() Then ' Update Record based on key
					SuccessMessage = Language.Phrase("UpdateSuccess") ' Set update success message

					' Alternatively, comment out the following line to go back to this page
					Call Page_Terminate("userlevelslist.asp") ' Return to list
				End If
		End Select
	End Sub

	' -----------------------------------------------------------------
	' Update privileges
	'
	Function EditRow()
		Dim Sql
		Dim RowCnt, i
		For i = LBound(EW_USER_LEVEL_TABLE_NAME) to UBound(EW_USER_LEVEL_TABLE_NAME)
			Sql = "UPDATE " & EW_USER_LEVEL_PRIV_TABLE & " SET " & EW_USER_LEVEL_PRIV_PRIV_FIELD & " = " & Privileges(i) & " WHERE " & _
				EW_USER_LEVEL_PRIV_TABLE_NAME_FIELD & " = '" & ew_AdjustSql(EW_USER_LEVEL_TABLE_NAME(i)) & "' AND " & _
				EW_USER_LEVEL_PRIV_USER_LEVEL_ID_FIELD & " = " & UserLevels.UserLevelID.CurrentValue
			Conn.Execute Sql, RowCnt
			If RowCnt = 0 Then
				Sql = "INSERT INTO " & EW_USER_LEVEL_PRIV_TABLE & " (" & EW_USER_LEVEL_PRIV_TABLE_NAME_FIELD & ", " & EW_USER_LEVEL_PRIV_USER_LEVEL_ID_FIELD & ", " & EW_USER_LEVEL_PRIV_PRIV_FIELD & ") VALUES ('" & ew_AdjustSql(EW_USER_LEVEL_TABLE_NAME(i)) & "', " & UserLevels.UserLevelID.CurrentValue & ", " & Privileges(i) & ")"
				Conn.Execute Sql, RowCnt
			End If
		Next
		EditRow = True
	End Function

	' -------------------
	' Get table caption
	Function GetTableCaption(i)
		If i <= UBound(EW_USER_LEVEL_TABLE_NAME) Then
			If i <= UBound(EW_USER_LEVEL_TABLE_VAR) Then
				GetTableCaption = Language.TablePhrase(EW_USER_LEVEL_TABLE_VAR(i), "TblCaption")
				Dim report
				report = (Mid(EW_USER_LEVEL_TABLE_NAME(i), 1, Len(EW_TABLE_PREFIX)) = EW_TABLE_PREFIX)
				If report And IsObject(ReportLanguage) Then
					GetTableCaption = ReportLanguage.TablePhrase(EW_USER_LEVEL_TABLE_VAR(i), "TblCaption")
				End If
			End If
			If GetTableCaption = "" Then
				If i <= UBound(EW_USER_LEVEL_TABLE_CAPTION) Then
					GetTableCaption = EW_USER_LEVEL_TABLE_CAPTION(i)
				End If
			End If
			If GetTableCaption = "" Then
				GetTableCaption = EW_USER_LEVEL_TABLE_NAME(i)
				If Left(GetTableCaption, Len(EW_REPORT_TABLE_PREFIX)) = EW_REPORT_TABLE_PREFIX Then GetTableCaption = Mid(GetTableCaption, Len(EW_REPORT_TABLE_PREFIX)+1)
			End If
			If Left(EW_USER_LEVEL_TABLE_NAME(i), Len(EW_REPORT_TABLE_PREFIX)) = EW_REPORT_TABLE_PREFIX Then
				GetTableCaption = GetTableCaption & "&nbsp;(" & Language.Phrase("Report") & ")"
			End If
		Else
			GetTableCaption = ""
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
End Class
%>
