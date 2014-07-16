<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
	<title>Notices</title>
<% If gsExport = "" Or gsExport = "print" Then %>
<link rel="stylesheet" type="text/css" href="<%= ew_YuiHost %>build/container/assets/skins/sam/container.css">
<link rel="stylesheet" type="text/css" href="<%= ew_YuiHost %>build/resize/assets/skins/sam/resize.css">
<!--<link rel="stylesheet" type="text/css" href="<%= EW_PROJECT_STYLESHEET_FILENAME %>">-->



<% End If %>
<% If gsExport = "" Or gsExport = "print" Then %>
<script type="text/javascript" src="<%= ew_YuiHost %>build/utilities/utilities.js"></script>
<script type="text/javascript" src="<%= ew_YuiHost %>build/container/container-min.js"></script>
<script type="text/javascript" src="<%= ew_YuiHost %>build/resize/resize-min.js"></script>
<script type="text/javascript">
<!--
var EW_LANGUAGE_ID = "<%= gsLanguage %>";
var EW_DATE_SEPARATOR = "/"; 
if (EW_DATE_SEPARATOR == "") EW_DATE_SEPARATOR = "/"; // Default date separator
var EW_UPLOAD_ALLOWED_FILE_EXT = "gif,jpg,jpeg,bmp,png,doc,xls,pdf,zip"; // Allowed upload file extension
var EW_FIELD_SEP = ", "; // Default field separator
// Ajax settings
var EW_RECORD_DELIMITER = "\r";
var EW_FIELD_DELIMITER = "|";
var EW_LOOKUP_FILE_NAME = "ewlookup9.asp"; // Lookup file name
var EW_AUTO_SUGGEST_MAX_ENTRIES = <%= EW_AUTO_SUGGEST_MAX_ENTRIES %>; // Auto-Suggest max entries
// Common JavaScript messages
var EW_ADDOPT_BUTTON_SUBMIT_TEXT = "<%= ew_JsEncode2(ew_BtnCaption(Language.Phrase("AddBtn"))) %>";
var EW_EMAIL_EXPORT_BUTTON_SUBMIT_TEXT = "<%= ew_JsEncode2(ew_BtnCaption(Language.Phrase("SendEmailBtn"))) %>";
var EW_BUTTON_CANCEL_TEXT = "<%= ew_JsEncode2(ew_BtnCaption(Language.Phrase("CancelBtn"))) %>";
var ewTooltipDiv;
var ew_TooltipTimer = null;
//-->
</script>
<script type="text/javascript" src="js/ew9.js"></script>
<script type="text/javascript" src="js/ewvalidator.js"></script>
<script type="text/javascript" src="js/userfn8.js"></script>
<script type="text/javascript">
<!--
<%= Language.ToJSON() %>
//-->
</script>
<script language="JavaScript" type="text/javascript">
<!--
// Write your client script here, no need to add script tags.
//-->
</script>
<% If gsExport = "" Or gsExport = "print" Then %>
<% if IsLoggedIn() then 
   dim myname, mytheme, themeSQL, themeRS
   myname = Session(EW_SESSION_USER_NAME) 
   themeSQL = "SELECT * FROM [users] WHERE [username] = '"& myname &"';" 
   Set themeRS = Server.CreateObject("ADODB.Recordset")
   themeRS.Open themeSQL, Conn

   ' Display result
   mytheme = "css/" &  themeRS("Theme") & ".css"
   %>
   	<link rel="stylesheet" type="text/css" href="<%= mytheme %>">
	<link rel="stylesheet" type="text/css" href="css/style.css">
	<%
	themeRS.Close
	Set themeRS = Nothing 
	else %>
	<link rel="stylesheet" type="text/css" href="css/bootstrap.css">
	<link rel="stylesheet" type="text/css" href="css/style.css">
<% End If %>

<% End If %>
<% End If %>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="shortcut icon" type="image/vnd.microsoft.icon" href="<%= ew_ConvertFullUrl("Favicon.ico") %>"><link rel="icon" type="image/vnd.microsoft.icon" href="<%= ew_ConvertFullUrl("Favicon.ico") %>">
<link href='http://fonts.googleapis.com/css?family=Oleo+Script' rel='stylesheet' type='text/css'>
<meta name="" content="">
</head>
<body>
<%
Dim rsunav,strSQLun, result
strSQLun = "Select count(*) AS unnoticecount from [Notices] where [approved] = False;"
Set rsunav = Server.CreateObject("ADODB.Recordset")	
rsunav.Open strSQLun, Conn 
%>
<% result = rsunav("unnoticecount")%>
<% if IsLoggedIn() and result > 0 and result < 2 then %>
<div class="gmsg">
  <button type="button" class="close" data-dismiss="alert"><i class="icon-remove-circle icon-white"></i></button>
  <p>Welcome <span class="muted myuser"><%= Session(EW_SESSION_USER_NAME) %></span><br/>
  <hr width="98%">
  <img src="images/exclamation.png" border="0">&nbsp;There is&nbsp;&nbsp;<span class="gmsg-count"><%= result %></span>&nbsp;&nbsp;Notice that is awaiting your approval.</p>
</div>
<% elseif IsLoggedIn() and result => 2 then %>
<div class="gmsg">
  <button type="button" class="close" data-dismiss="alert"><i class="icon-remove-circle icon-white"></i></button>
  <p>Welcome <span class="muted myuser"><%= Session(EW_SESSION_USER_NAME) %></span>
  <hr width="98%">
  <img src="images/exclamation.png" border="0">&nbsp;There are&nbsp;&nbsp;<span class="gmsg-count"><%= result %></span>&nbsp;&nbsp;Notices that are awaiting your approval.</p>
</div>
<% end if %>
<% If gsExport = "" Then %>
<div class="container">
  <div>
 <div>
	 	  <% if IsLoggedIn() then %>
		  <a href="logout.asp" class="btn btn-warning pull-right"><i class="icon-off icon-white"></i>&nbsp;Logout</a>
		  <!--<Span class="pull-right"><strong>Welcome</strong> <span class="muted"><%= Session(EW_SESSION_USER_NAME) %></span>&nbsp;&nbsp;&nbsp;</span>-->
		  <% else %>
		  <%
		  Thispage = Request.ServerVariables("URL")
		  %>
		  <% if Thispage <> "/notices/login.asp" then %>
		  <a href="login.asp" class="btn btn-warning pull-right"><i class="icon-off icon-white"></i>&nbsp;Login</a>
	 	  <% end if %>
		  <% end if %>
		  <h1><img src="favicon.ico" border="0" width="40">&nbsp;&nbsp;<span Class="mylogo">Notices</span></h1>
		  <p class="description">Keeping you informed</p>

	 </div>	<!-- header (end) -->
	<!-- content (begin) -->
  <table class="container">
		<tr>	
			<td class="navbar-inverse">
				<% Server.Execute("ewmenu.asp") %>
			</td>
		</TR>
		<TR>
		<td>
			<!-- right column (begin) -->
<% End If %>
