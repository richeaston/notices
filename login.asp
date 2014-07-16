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
Dim login
Set login = New clogin
Set Page = login

' Page init processing
Call login.Page_Init()

' Page main processing
Call login.Page_Main()
%>
<!--#include file="header.asp"-->
<script language="JavaScript" type="text/javascript">
<!--
// Write your client script here, no need to add script tags.
//-->
</script>
<% login.ShowPageHeader() %>
<script type="text/javascript">
<!--
var login = new ew_Page("login");
// extend page with ValidateForm function
login.ValidateForm = function(fobj)
{
	if (!this.ValidateRequired)
		return true; // ignore validation
	if (!ew_HasValue(fobj.username))
		return ew_OnError(this, fobj.username, ewLanguage.Phrase("EnterUid"));
	if (!ew_HasValue(fobj.password))
		return ew_OnError(this, fobj.password, ewLanguage.Phrase("EnterPwd"));
	// Call Form Custom Validate event
	if (!this.Form_CustomValidate(fobj)) return false;
	return true;
}
// extend page with Form_CustomValidate function
login.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// requires js validation
<% If EW_CLIENT_VALIDATE Then %>
login.ValidateRequired = true;
<% Else %>
login.ValidateRequired = false;
<% End If %>
//-->
</script>
<% login.ShowMessage %>
<form action="<%= ew_CurrentPage %>"  method="post" onsubmit="return login.ValidateForm(this);">
<div class="account-container">
	<div class="content clearfix">
		<form class="" action="documents.asp" method="post">
			<h1 class="mylogo">Sign In</h1>		
			<div class="login-fields">
				<p>Sign in using your registered account:</p>
				<div class="field  input-prepend">
					<span class="add-on"><i class="icon-user"></i></span>
					<input type="text" id="username" name="username" value="" placeholder="Username" class="login username-field">
				</div> <!-- /field -->
				
				<div class="field input-prepend" title="Password">
					<span class="add-on"><i class="icon-lock"></i></span>
					<input type="password" id="password" name="password" value="<%= login.sUsername %>" placeholder="Password" class="login password-field">
				</div> <!-- /password -->
				
			</div> <!-- /login-fields -->
			<div class="form-inline"> 
			<label><input type="radio" name="rememberme" id="rememberme" value="a"<% If login.sLoginType = "a" Then %> checked="checked"<% End If %>>&nbsp;Remember Me</label><br>
			<label><input type="radio" name="rememberme" id="rememberme" value="u"<% If login.sLoginType = "u" Then %>  checked="checked"<% End If %>>&nbsp;Save my username</label><br>
			<label><input type="radio" name="rememberme" id="rememberme" value=""<% If login.sLoginType = "" Then %> checked="checked"<% End If %>>&nbsp;Always Ask</label><br>
			
			</div>
			<div class="login-actions inline">
				<div class="btn-group button-login pull-right">	
							<button type="submit" class="button btn btn-warning"><i class="icon-off icon-white"></i>&nbsp;Sign In</button>
				</div>
			</div> <!-- .actions -->
			<div class="description"><a href="forgotpwd.asp"><%= Language.Phrase("ForgotPwd") %></a></div>
		</form>
	</div> <!-- /content -->
</div>				
<div class="clearfix"></div>



</form>

<%
login.ShowPageFooter()
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
Set login = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class clogin

	' Page ID
	Public Property Get PageID()
		PageID = "login"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "login"
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
		EW_PAGE_ID = "login"

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
		Set Security = New cAdvancedSecurity

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

	Dim sUsername
	Dim sLoginType

	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()
		Dim bValidate, bValidPwd
		Dim sPassword
		Dim sLastUrl
		sLastUrl = Security.LastUrl ' Get Last Url
		If sLastUrl = "" Then sLastUrl = "default.asp"
		If IsLoggingIn() Then
			sUsername = Session(EW_SESSION_USER_PROFILE_USER_NAME)
			sPassword = Session(EW_SESSION_USER_PROFILE_PASSWORD)
			sLoginType = Session(EW_SESSION_USER_PROFILE_LOGIN_TYPE)
			bValidPwd = Security.ValidateUser(sUsername, sPassword, False)
			If bValidPwd Then
				Session(EW_SESSION_USER_PROFILE_USER_NAME) = ""
				Session(EW_SESSION_USER_PROFILE_PASSWORD) = ""
				Session(EW_SESSION_USER_PROFILE_LOGIN_TYPE) = ""
			End If
		Else
			If Not Security.IsLoggedIn() Then Call Security.AutoLogin()
			Call Security.LoadUserLevel() ' Load user level
			If Request.Form <> "" Then

				' Setup variables
				sUsername = ew_RemoveXSS(Request.Form("Username"))
				sPassword = Request.Form("Password")
				sLoginType = LCase(Request.Form("rememberme"))
				bValidate = ValidateForm(sUsername, sPassword)
				If Not bValidate Then
					FailureMessage = gsFormError
				End If
				Session(EW_SESSION_USER_PROFILE_USER_NAME) = sUsername ' Save login user name
				Session(EW_SESSION_USER_PROFILE_LOGIN_TYPE) = sLoginType ' Save login type

				' Max login attempt checking
				If UserProfile.LoadProfileFromDatabase(sUsername) Then
					If UserProfile.ExceedLoginRetry() Then
						bValidate = False
						FailureMessage = Replace(Language.Phrase("ExceedMaxRetry"), "%t", EW_USER_PROFILE_RETRY_LOCKOUT)
					End If
					UserProfile.SaveProfileToDatabase sUsername
				End If
			Else
				If Security.IsLoggedIn() Then
					If FailureMessage = "" Then Page_Terminate(sLastUrl) ' Return to last accessed page
				End If
				bValidate = False

				' Restore settings
				sUsername = Request.Cookies(EW_PROJECT_NAME)("username")
				If Request.Cookies(EW_PROJECT_NAME)("autologin") = "autologin" Then
					sLoginType = "a"
				ElseIf Request.Cookies(EW_PROJECT_NAME)("autologin") = "rememberUsername" Then
					sLoginType = "u"
				Else
					sLoginType = ""
				End If
			End If
			bValidPwd = False
			If bValidate Then

				' Call logging in event
				bValidate = User_LoggingIn(sUsername, sPassword)
				If bValidate Then
					bValidPwd = Security.ValidateUser(sUsername, sPassword, False) ' Manual login
					If Not bValidPwd Then
						If FailureMessage = "" Then FailureMessage = Language.Phrase("InvalidUidPwd") ' Invalid user id/password
					End If
				Else
					If FailureMessage = "" Then FailureMessage = Language.Phrase("LoginCancelled") ' Login cancelled
				End If
			End If
		End If
		If bValidPwd Then

			' Write cookies
			If sLoginType = "a" Then ' Auto login
				Response.Cookies(EW_PROJECT_NAME)("autologin") = "autologin" ' Set up autologin cookies
				Response.Cookies(EW_PROJECT_NAME)("username") = sUsername ' Set up user name cookies
				Response.Cookies(EW_PROJECT_NAME)("password") = ew_Encode(TEAencrypt(sPassword, EW_RANDOM_KEY)) ' Set up password cookies
			ElseIf sLoginType = "u" Then ' Remember user name
				Response.Cookies(EW_PROJECT_NAME)("autologin") = "rememberUsername" ' Set up remember user name cookies
				Response.Cookies(EW_PROJECT_NAME)("username") = sUsername ' Set up user name cookies
			Else
				Response.Cookies(EW_PROJECT_NAME)("autologin") = "" ' Clear autologin cookies
			End If
			Response.Cookies(EW_PROJECT_NAME).Expires = DateAdd("d", EW_COOKIE_EXPIRY_TIME, Date)

			' Call loggedin event
			Call User_LoggedIn(sUsername)
			Call Page_Terminate(sLastUrl) ' Return to last accessed url
		ElseIf sUsername <> "" And sPassword <> "" Then

			' Call user login error event
			Call User_LoginError(sUsername, sPassword)
		End If
	End Sub

	' -----------------------------------------------------------------
	' Validate form
	'
	Function ValidateForm(usr, pwd)

		' Initialize
		gsFormError = ""

		' Check if validation required
		If Not EW_SERVER_VALIDATE Then
			ValidateForm = True
			Exit Function
		End If
		If usr = "" Then
			Call ew_AddMessage(gsFormError, Language.Phrase("EnterUid"))
		End If
		If pwd = "" Then
			Call ew_AddMessage(gsFormError, Language.Phrase("EnterPwd"))
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

	' User Logging In event
	Function User_LoggingIn(usr, pwd)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next

		' Enter your code here
		' To cancel, set return value to False

		User_LoggingIn = True
	End Function

	' User Logged In event
	Sub User_LoggedIn(usr)

		' Response.Write "User Logged In"
	End Sub

	' User Login Error event
	Sub User_LoginError(usr, pwd)

		' Response.Write "User Login Error"
	End Sub

	' Form Custom Validate event
	Function Form_CustomValidate(CustomError)

		'Return error message in CustomError
		Form_CustomValidate = True
	End Function
End Class
%>
