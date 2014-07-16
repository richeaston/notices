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
Dim changepwd
Set changepwd = New cchangepwd
Set Page = changepwd

' Page init processing
Call changepwd.Page_Init()

' Page main processing
Call changepwd.Page_Main()
%>
<!--#include file="header.asp"-->
<script language="JavaScript" type="text/javascript">
<!--
// Write your client script here, no need to add script tags.
//-->
</script>
<% changepwd.ShowPageHeader() %>
<script type="text/javascript">
<!--
var changepwd = new ew_Page("changepwd");
// extend page with ValidateForm function
changepwd.ValidateForm = function(fobj)
{
	if (!this.ValidateRequired)
		return true; // ignore validation
	if  (!ew_HasValue(fobj.opwd))
		return ew_OnError(this, fobj.opwd, ewLanguage.Phrase("EnterOldPassword"));
	if  (!ew_HasValue(fobj.npwd))
		return ew_OnError(this, fobj.npwd, ewLanguage.Phrase("EnterNewPassword"));
	if  (fobj.npwd.value != fobj.cpwd.value)
		return ew_OnError(this, fobj.cpwd, ewLanguage.Phrase("MismatchPassword"));
	// Call Form Custom Validate event
	if (!this.Form_CustomValidate(fobj)) return false;
	return true;
}
// extend page with Form_CustomValidate function
changepwd.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// requires js validation
<% If EW_CLIENT_VALIDATE Then %>
changepwd.ValidateRequired = true;
<% Else %>
changepwd.ValidateRequired = false;
<% End If %>
//-->
</script>
<p class="aspmaker ewTitle"><%= Language.Phrase("ChangePwdPage") %></p>
<% changepwd.ShowMessage %>
<form action="<%= ew_CurrentPage %>" method="post" onsubmit="return changepwd.ValidateForm(this);">
<table border="0" cellspacing="0" cellpadding="4">
	<tr>
		<td><span class="aspmaker"><%= Language.Phrase("OldPassword") %></span></td>
		<td><span class="aspmaker"><input type="password" name="opwd" id="opwd" size="20"></span></td>
	</tr>
	<tr>
		<td><span class="aspmaker"><%= Language.Phrase("NewPassword") %></span></td>
		<td><span class="aspmaker"><input type="password" name="npwd" id="npwd" size="20"></span></td>
	</tr>
	<tr>
		<td><span class="aspmaker"><%= Language.Phrase("ConfirmPassword") %></span></td>
		<td><span class="aspmaker"><input type="password" name="cpwd" id="cpwd" size="20"></span></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td><span class="aspmaker"><input type="submit" name="submit" id="submit" value="<%= ew_BtnCaption(Language.Phrase("ChangePwdBtn")) %>"></span></td>
	</tr>
</table>
</form>
<br>
<%
changepwd.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script language="JavaScript" type="text/javascript">
<!--
// Write your startup script here
// document.write("page loaded");
//-->
</script>
<!--#include file="footer.asp"-->
<%

' Drop page object
Set changepwd = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cchangepwd

	' Page ID
	Public Property Get PageID()
		PageID = "changepwd"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "changepwd"
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
		If IsEmpty(Users) Then Set Users = New cUsers

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "changepwd"

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
		If Not Security.IsLoggedIn() Or Security.IsSysAdmin() Then Call Page_Terminate("login.asp")
		Call Security.LoadCurrentUserLevel("Users")

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
	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()
		Dim bValidPwd
		Dim bPwdUpdated, sUsername, sOPwd, sNPwd, sCPwd, sEmail, sFilter, sSql, RsUser
		If Request.Form <> "" Then
			bPwdUpdated = False

			' Setup variables
			sUsername = Security.CurrentUserName
			sOPwd = Request.Form("opwd")
			sNPwd = Request.Form("npwd")
			sCPwd = Request.Form("cpwd")
			If ValidateForm(sOPwd, sNPwd, sCPwd) Then
				sFilter = Replace(EW_USER_NAME_FILTER, "%u", ew_AdjustSql(sUsername))

				' Set up filter (Sql Where Clause) and get Return Sql
				' Sql constructor in Users class, Usersinfo.asp

				Users.CurrentFilter = sFilter
				sSql = Users.SQL
				Set RsUser = Server.CreateObject("ADODB.Recordset")
				RsUser.CursorLocation = EW_CURSORLOCATION
				RsUser.Open sSql, Conn, 1, EW_RECORDSET_LOCKTYPE
				If Not RsUser.Eof Then
					If ew_ComparePassword(RsUser("Password"), sOPwd) Then
						bValidPwd = True
						bValidPwd = User_ChangePassword(RsUser, sUsername, sOPwd, sNPwd)
						If bValidPwd Then
							If Not EW_CASE_SENSITIVE_PASSWORD Then sNPwd = LCase(sNPwd)
							If EW_ENCRYPTED_PASSWORD Then
								RsUser("Password") = ew_EncryptPassword(sNPwd) ' Change Password
							Else
								RsUser("Password") = sNPwd ' Change Password
							End If
							sEmail = RsUser("Email")
							RsUser.Update
							bPwdUpdated = True
						Else
							FailureMessage = Language.Phrase("InvalidNewPassword")
						End If
					Else
						FailureMessage = Language.Phrase("InvalidPassword")
					End If
				End If
				If bPwdUpdated Then

					' Send Email
					Dim Email, bEmailSent
					Set Email = New cEmail
					Email.Load("txt/changepwd.txt")
					Email.ReplaceSender(EW_SENDER_EMAIL) ' Replace Sender
					Email.ReplaceRecipient(sEmail) ' Replace Recipient
					Email.ReplaceContent "<!--$Password-->", sNPwd
					If EW_EMAIL_CHARSET <> "" Then Email.Charset = EW_EMAIL_CHARSET
					Set EventArgs = Server.CreateObject("Scripting.Dictionary")
					EventArgs.Add "Rs", RsUser
					If Email_Sending(Email, EventArgs) Then
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
					SuccessMessage = Language.Phrase("PasswordChanged") ' set up message
					Call Page_Terminate("default.asp") ' exit page and clean up
				End If
				RsUser.Close
				Set RsUser = Nothing
			Else
				FailureMessage = gsFormError
			End If
		End If
	End Sub

	' -----------------------------------------------------------------
	' Validate form
	'
	Function ValidateForm(opwd, npwd, cpwd)

		' Initialize
		gsFormError = ""

		' Check if validation required
		If Not EW_SERVER_VALIDATE Then
			ValidateForm = True
			Exit Function
		End If
		If opwd = "" Then
			Call ew_AddMessage(gsFormError, Language.Phrase("EnterOldPassword"))
		End If
		If npwd = "" Then
			Call ew_AddMessage(gsFormError, Language.Phrase("EnterNewPassword"))
		End If
		If npwd <> cpwd Then
			Call ew_AddMessage(gsFormError, Language.Phrase("MismatchPassword"))
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

	' Email Sending event
	Function Email_Sending(Email, Args)

		'Response.Write Email.AsString
		'Response.Write "Keys of Args: " & Join(Args.Keys, ", ")
		'Response.End

		Email_Sending = True
	End Function

	' Form Custom Validate event
	Function Form_CustomValidate(CustomError)

		'Return error message in CustomError
		Form_CustomValidate = True
	End Function

	' User ChangePassword event
	Function User_ChangePassword(Rs, usr, oldpwd, newpwd)

		' Return FALSE to abort
		User_ChangePassword = True
	End Function
End Class
%>
