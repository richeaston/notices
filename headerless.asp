<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
	<title><%= Language.ProjectPhrase("BodyTitle") %></title>
<% If gsExport = "" Or gsExport = "print" Then %>
<link rel="stylesheet" type="text/css" href="<%= ew_YuiHost %>build/container/assets/skins/sam/container.css">
<link rel="stylesheet" type="text/css" href="<%= ew_YuiHost %>build/resize/assets/skins/sam/resize.css">
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
	<link rel="stylesheet" type="text/css" href="css/style.css">
	<link rel="stylesheet" type="text/css" href="css/Cerulean.css">
<% End If %>


<% End If %>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="shortcut icon" type="image/vnd.microsoft.icon" href="<%= ew_ConvertFullUrl("Favicon.ico") %>"><link rel="icon" type="image/vnd.microsoft.icon" href="<%= ew_ConvertFullUrl("Favicon.ico") %>">
<link href='http://fonts.googleapis.com/css?family=Oleo+Script' rel='stylesheet' type='text/css'>
<meta name="generator" content="ASPMaker v9.0.2">
</head>
<body class="scrollerbg">
<% If gsExport = "" Then %>
<div >
  <table >
		<TR>
		<td>
			<!-- right column (begin) -->
<% End If %>
