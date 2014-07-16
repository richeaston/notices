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
Dim forgotpwd
Set forgotpwd = New cforgotpwd
Set Page = forgotpwd

' Page init processing
Call forgotpwd.Page_Init()

' Page main processing
Call forgotpwd.Page_Main()
%>
<!--#include file="header.asp"-->
<script language="JavaScript" type="text/javascript">
<!--
// Write your client script here, no need to add script tags.
//-->
</script>
<% forgotpwd.ShowPageHeader() %>
<script type="text/javascript">
<!--
var forgotpwd = new ew_Page("forgotpwd");
// extend page with ValidateForm function
forgotpwd.ValidateForm = function(fobj)
{
	if (!this.ValidateRequired)
		return true; // ignore validation
	if  (!ew_HasValue(fobj.email))
		return ew_OnError(this, fobj.email, ewLanguage.Phrase("EnterValidEmail"));
	if  (!ew_CheckEmail(fobj.email.value))
		return ew_OnError(this, fobj.email, ewLanguage.Phrase("EnterValidEmail"));
	// Call Form Custom Validate event
	if (!this.Form_CustomValidate(fobj)) return false;
	return true;
}
// extend page with Form_CustomValidate function
forgotpwd.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// requires js validation
<% If EW_CLIENT_VALIDATE Then %>
forgotpwd.ValidateRequired = true;
<% Else %>
forgotpwd.ValidateRequired = false;
<% End If %>
//-->
</script>
<p class="aspmaker ewTitle"><%= Language.Phrase("RequestPwdPage") %></p>
<p class="aspmaker"><a href="login.asp"><%= Language.Phrase("BackToLogin") %></a></p>
<% forgotpwd.ShowMessage %>
<form action="<%= ew_CurrentPage %>" method="post" onsubmit="return forgotpwd.ValidateForm(this);">
<table border="0" cellspacing="0" cellpadding="4">
	<tr>
		<td><span class="aspmaker"><%= Language.Phrase("UserEmail") %></span></td>
		<td><span class="aspmaker"><input type="text" name="email" id="email" value="<%= ew_HtmlEncode(forgotpwd.EmailAddress) %>" size="30" maxlength="255"></span></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td><span class="aspmaker"><input type="submit" name="submit" id="submit" value="<%= ew_BtnCaption(Language.Phrase("SendPwd")) %>"></span></td>
	</tr>
</table>
</form>
<br>
<%
forgotpwd.ShowPageFooter()
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
Set forgotpwd = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cforgotpwd

	' Page ID
	Public Property Get PageID()
		PageID = "forgotpwd"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "forgotpwd"
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
		EW_PAGE_ID = "forgotpwd"

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

	Dim EmailAddress

	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()
		Dim bValidEmail
		Dim sFilter, sSql, RsUser
		If Request.Form <> "" Then
			bValidEmail = False

			' Setup variables
			EmailAddress = Request.Form("email")
			If ValidateForm(EmailAddress) Then

				' Set up filter (Sql Where Clause) and get Return Sql
				' Sql constructor in Users class, Usersinfo.asp

				sFilter = Replace(EW_USER_EMAIL_FILTER, "%e", ew_AdjustSql(EmailAddress))
				Users.CurrentFilter = sFilter
				sSql = Users.SQL
				Set RsUser = Server.CreateObject("ADODB.Recordset")
				RsUser.CursorLocation = EW_CURSORLOCATION
				RsUser.Open sSql, Conn, 1, EW_RECORDSET_LOCKTYPE
				If Not RsUser.Eof Then
					bValidEmail = True

					' Call User Recover Password event
					bValidEmail = User_RecoverPassword(RsUser)
					If bValidEmail Then
						Dim sUserName, sPassword
						sUserName = RsUser("Username")
						sPassword = RsUser("Password")
						If EW_ENCRYPTED_PASSWORD Then
							sPassword = Mid(sPassword, 1, 16) ' Use first 16 characters only
							RsUser("Password") = ew_EncryptPassword(sPassword) ' Reset password
							RsUser.Update
						End If
					End If
				Else
					FailureMessage = Language.Phrase("InvalidEmail")
				End If
				If bValidEmail Then
					Dim Email, bEmailSent
					Set Email = New cEmail
					Email.Load("txt/forgotpwd.txt")
					Email.ReplaceSender(EW_SENDER_EMAIL) ' Replace Sender
					Email.ReplaceRecipient(EmailAddress) ' Replace Recipient
					Email.ReplaceContent "<!--$UserName-->", sUserName
					Email.ReplaceContent "<!--$Password-->", sPassword
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
				Else
					bEmailSent = False
				End If
				RsUser.Close
				Set RsUser = Nothing
				If bEmailSent Then
					SuccessMessage = Language.Phrase("PwdEmailSent") ' Set success message
					Call Page_Terminate("login.asp") ' Return to login page
				End If
			Else
				FailureMessage = gsFormError
			End If
		End If
	End Sub

	' -----------------------------------------------------------------
	' Validate form
	'
	Function ValidateForm(email)

		' Initialize
		gsFormError = ""

		' Check if validation required
		If Not EW_SERVER_VALIDATE Then
			ValidateForm = True
			Exit Function
		End If
		If email = "" Then
			Call ew_AddMessage(gsFormError, Language.Phrase("EnterValidEmail"))
		End If
		If Not ew_CheckEmail(email) Then
			Call ew_AddMessage(gsFormError, Language.Phrase("EnterValidEmail"))
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

	' User RecoverPassword event
	Function User_RecoverPassword(Rs)

	  'Response.Write "User_RecoverPassword"
	  User_RecoverPassword = True
	End Function
End Class
%>
