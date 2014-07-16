<%

' ASPMaker 9 configuration file
' - contains all web site configuration settings

Const EW_PROJECT_NAME = "Billboard" ' Project Name
Const EW_MAX_EMAIL_RECIPIENT = 3

' Auto suggest max entries
Const EW_AUTO_SUGGEST_MAX_ENTRIES = 10

' Language settings
Const EW_LANGUAGE_FOLDER = "lang/"
Dim EW_LANGUAGE_FILE(0)
EW_LANGUAGE_FILE(0) = Array("en", "", "english.xml")
Const EW_LANGUAGE_DEFAULT_ID = "en"
Dim EW_SESSION_LANGUAGE_FILE_CACHE
EW_SESSION_LANGUAGE_FILE_CACHE = EW_PROJECT_NAME & "_LanguageFile_B97ie5vAbl2Ky6ac" ' Language File Cache
Dim EW_SESSION_LANGUAGE_CACHE
EW_SESSION_LANGUAGE_CACHE = EW_PROJECT_NAME & "_Language_B97ie5vAbl2Ky6ac" ' Language Cache
Dim EW_SESSION_LANGUAGE_ID
EW_SESSION_LANGUAGE_ID = EW_PROJECT_NAME & "_LanguageId" ' Language ID

' Response.Buffer setting
Const EW_RESPONSE_BUFFER = True

' Menu
Const EW_ITEM_TEMPLATE_CLASSNAME = "ewTemplate"
Const EW_ITEM_TABLE_CLASSNAME = "ewItemTable"

' Debug flag
Const EW_DEBUG_ENABLED = False ' True to debug / False to skip

' Remove XSS
Const EW_REMOVE_XSS = True ' True to Remove XSS / False to skip

' XSS Array
Dim EW_XSS_ARRAY
EW_XSS_ARRAY = Array("javascript", "vbscript", "expression", "<applet", "<meta", "<xml", "<blink", "<link", "<style", "<script", "<embed", "<object", "<iframe", "<frame", "<frameset", "<ilayer", "<layer", "<bgsound", "<title", "<base", _
	"onabort", "onactivate", "onafterprint", "onafterupdate", "onbeforeactivate", "onbeforecopy", "onbeforecut", "onbeforedeactivate", "onbeforeeditfocus", "onbeforepaste", "onbeforeprint", "onbeforeunload", "onbeforeupdate", "onblur", "onbounce", "oncellchange", "onchange", "onclick", "oncontextmenu", "oncontrolselect", "oncopy", "oncut", "ondataavailable", "ondatasetchanged", "ondatasetcomplete", "ondblclick", "ondeactivate", "ondrag", "ondragend", "ondragenter", "ondragleave", "ondragover", "ondragstart", "ondrop", "onerror", "onerrorupdate", "onfilterchange", "onfinish", "onfocus", "onfocusin", "onfocusout", "onhelp", "onkeydown", "onkeypress", "onkeyup", "onlayoutcomplete", "onload", "onlosecapture", "onmousedown", "onmouseenter", "onmouseleave", "onmousemove", "onmouseout", "onmouseover", "onmouseup", "onmousewheel", "onmove", "onmoveend", "onmovestart", "onpaste", "onpropertychange", "onreadystatechange", "onreset", "onresize", "onresizeend", "onresizestart", "onrowenter", "onrowexit", "onrowsdelete", "onrowsinserted", "onscroll", "onselect", "onselectionchange", "onselectstart", "onstart", "onstop", "onsubmit", "onunload")

' Session names
Dim EW_SESSION_STATUS
EW_SESSION_STATUS = EW_PROJECT_NAME & "_Status" ' Login Status
Dim EW_SESSION_USER_NAME
EW_SESSION_USER_NAME = EW_SESSION_STATUS & "_UserName" ' User Name
Dim EW_SESSION_USER_ID
EW_SESSION_USER_ID = EW_SESSION_STATUS & "_UserID" ' User ID
Dim EW_SESSION_USER_PROFILE, EW_SESSION_USER_PROFILE_USER_NAME, EW_SESSION_USER_PROFILE_PASSWORD, EW_SESSION_USER_PROFILE_LOGIN_TYPE
EW_SESSION_USER_PROFILE = EW_SESSION_STATUS & "_UserProfile" ' User Profile
EW_SESSION_USER_PROFILE_USER_NAME = EW_SESSION_USER_PROFILE & "_UserName"
EW_SESSION_USER_PROFILE_PASSWORD = EW_SESSION_USER_PROFILE & "_Password"
EW_SESSION_USER_PROFILE_LOGIN_TYPE = EW_SESSION_USER_PROFILE & "_LoginType"
Dim EW_SESSION_USER_LEVEL_ID
EW_SESSION_USER_LEVEL_ID = EW_SESSION_STATUS & "_UserLevel" ' User Level ID
Dim EW_SESSION_USER_LEVEL
EW_SESSION_USER_LEVEL = EW_SESSION_STATUS & "_UserLevelValue" ' User Level
Dim EW_SESSION_PARENT_USER_ID
EW_SESSION_PARENT_USER_ID = EW_SESSION_STATUS & "_ParentUserID" ' Parent User ID
Dim EW_SESSION_SYS_ADMIN
EW_SESSION_SYS_ADMIN = EW_PROJECT_NAME & "_SysAdmin" ' System Admin
Dim EW_SESSION_AR_USER_LEVEL
EW_SESSION_AR_USER_LEVEL = EW_PROJECT_NAME & "_arUserLevel" ' User Level Array
Dim EW_SESSION_AR_USER_LEVEL_PRIV
EW_SESSION_AR_USER_LEVEL_PRIV = EW_PROJECT_NAME & "_arUserLevelPriv" ' User Level Privilege Array
Dim EW_SESSION_SECURITY
EW_SESSION_SECURITY = EW_PROJECT_NAME & "_Security" ' Security Array
Dim EW_SESSION_MESSAGE
EW_SESSION_MESSAGE = EW_PROJECT_NAME & "_Message" ' System Message
Dim EW_SESSION_FAILURE_MESSAGE, EW_SESSION_SUCCESS_MESSAGE
EW_SESSION_FAILURE_MESSAGE = EW_PROJECT_NAME & "_Failure_Message" ' System error message
EW_SESSION_SUCCESS_MESSAGE = EW_PROJECT_NAME & "_Success_Message" ' System message
Dim EW_SESSION_INLINE_MODE
EW_SESSION_INLINE_MODE = EW_PROJECT_NAME & "_InlineMode" ' Inline Mode

' Charset
Const EW_CHARSET = "windows-1252" ' Project charset

' Css file name
Const EW_PROJECT_STYLESHEET_FILENAME = "css/billboard.css"

' Database settings
Dim EW_DB_CONNECTION_STRING ' DB Connection String
EW_DB_CONNECTION_STRING = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("notices.mdb") & ";"
Const EW_DB_QUOTE_START = "["
Const EW_DB_QUOTE_END = "]"
Const EW_IS_MSACCESS = True ' Access
Const EW_IS_MSSQL = False ' MS SQL
Const EW_IS_MYSQL = False ' MySQL
Const EW_IS_ORACLE = False ' Oracle
Const EW_IS_POSTGRESQL = False ' PostgreSQL
Const EW_CURSORLOCATION = 2 ' Cursor location
Const EW_RECORDSET_LOCKTYPE = 2 ' Recordset lock type
Const EW_DATATYPE_NUMBER = 1
Const EW_DATATYPE_DATE = 2
Const EW_DATATYPE_STRING = 3
Const EW_DATATYPE_BOOLEAN = 4
Const EW_DATATYPE_MEMO = 5
Const EW_DATATYPE_BLOB = 6
Const EW_DATATYPE_TIME = 7
Const EW_DATATYPE_GUID = 8
Const EW_DATATYPE_XML = 9
Const EW_DATATYPE_OTHER = 10
Const EW_COMPOSITE_KEY_SEPARATOR = "," ' Composite key separator
Const EW_EMAIL_KEYWORD_SEPARATOR = "" ' Email keyword separator
Const EW_EMAIL_CHARSET = "windows-1252" ' Email charset
Const EW_HIGHLIGHT_COMPARE = 1 ' Highlight compare mode
Const EW_ROWTYPE_VIEW = 1 ' Row type view
Const EW_ROWTYPE_ADD = 2 ' Row type add
Const EW_ROWTYPE_EDIT = 3 ' Row type edit
Const EW_ROWTYPE_SEARCH = 4 ' Row type search
Const EW_ROWTYPE_MASTER = 5 ' Row type master record
Const EW_ROWTYPE_AGGREGATEINIT = 6 ' Row type aggregate init
Const EW_ROWTYPE_AGGREGATE = 7 ' Row type aggregate

' Table specific names
Const EW_TABLE_PREFIX = "||ASPReportMaker||"
Const EW_TABLE_REC_PER_PAGE = "recperpage" ' Records per page
Const EW_TABLE_START_REC = "start" ' Start record
Const EW_TABLE_PAGE_NO = "pageno" ' Page number
Const EW_TABLE_BASIC_SEARCH = "psearch" ' Basic search keyword
Const EW_TABLE_BASIC_SEARCH_TYPE = "psearchtype" ' Basic search type
Const EW_TABLE_ADVANCED_SEARCH = "advsrch" ' Advanced search
Const EW_TABLE_SEARCH_WHERE = "searchwhere" ' Search where clause
Const EW_TABLE_WHERE = "where" ' Table where
Const EW_TABLE_WHERE_LIST = "where_list" ' Table where (list page)
Const EW_TABLE_ORDER_BY = "orderby" ' Table order by
Const EW_TABLE_ORDER_BY_LIST = "orderby_list" ' Table order by (list page)
Const EW_TABLE_SORT = "sort" ' Table sort
Const EW_TABLE_KEY = "key" ' Table key
Const EW_TABLE_SHOW_MASTER = "showmaster" ' Table show master
Const EW_TABLE_SHOW_DETAIL = "showdetail" ' Table show detail
Const EW_TABLE_MASTER_TABLE = "mastertable" ' Master table
Const EW_TABLE_DETAIL_TABLE = "detailtable" ' Detail table
Const EW_TABLE_MASTER_FILTER = "masterfilter" ' Master filter
Const EW_TABLE_DETAIL_FILTER = "detailfilter" ' Detail filter
Const EW_TABLE_RETURN_URL = "return" ' Return url
Const EW_TABLE_EXPORT_RETURN_URL = "exportreturn" ' Export return url
Const EW_TABLE_GRID_ADD_ROW_COUNT = "gridaddcnt" ' Grid add row count

' Audit Trail
Const EW_AUDIT_TRAIL_TO_DATABASE = True ' Write audit trail to DB
Const EW_AUDIT_TRAIL_TABLE_NAME = "AuditTrail" ' Audit trail table name
Const EW_AUDIT_TRAIL_FIELD_NAME_DATETIME = "DateTime" ' Audit trail DateTime field name
Const EW_AUDIT_TRAIL_FIELD_NAME_SCRIPT = "Script" ' Audit trail Script field name
Const EW_AUDIT_TRAIL_FIELD_NAME_USER = "User" ' Audit trail User field name
Const EW_AUDIT_TRAIL_FIELD_NAME_ACTION = "Action" ' Audit trail Action field name
Const EW_AUDIT_TRAIL_FIELD_NAME_TABLE = "Table" ' Audit trail Table field name
Const EW_AUDIT_TRAIL_FIELD_NAME_FIELD = "Field" ' Audit trail Field field name
Const EW_AUDIT_TRAIL_FIELD_NAME_KEYVALUE = "KeyValue" ' Audit trail Key Value field name
Const EW_AUDIT_TRAIL_FIELD_NAME_OLDVALUE = "OldValue" ' Audit trail Old Value field name
Const EW_AUDIT_TRAIL_FIELD_NAME_NEWVALUE = "NewValue" ' Audit trail New Value field name

' Security specific
Const EW_AUDIT_TRAIL_PATH = "" ' Audit trail path
Const EW_ADMIN_USER_NAME = "" ' Administrator user name
Const EW_ADMIN_PASSWORD = "" ' Administrator password
Const EW_USE_CUSTOM_LOGIN = True ' Use custom login
Const EW_ENCRYPTED_PASSWORD = False ' Use encrypted password
Const EW_CASE_SENSITIVE_PASSWORD = True ' Case Sensitive password

' Dynamic user level table
Dim EW_USER_LEVEL_TABLE_NAME, EW_USER_LEVEL_TABLE_CAPTION, EW_USER_LEVEL_TABLE_VAR
ReDim EW_USER_LEVEL_TABLE_NAME(6)
ReDim EW_USER_LEVEL_TABLE_CAPTION(6)
ReDim EW_USER_LEVEL_TABLE_VAR(6)
EW_USER_LEVEL_TABLE_NAME(0) = "Groups"
EW_USER_LEVEL_TABLE_CAPTION(0) = "Groups"
EW_USER_LEVEL_TABLE_VAR(0) = "Groups"
EW_USER_LEVEL_TABLE_NAME(1) = "Notices"
EW_USER_LEVEL_TABLE_CAPTION(1) = "Notices"
EW_USER_LEVEL_TABLE_VAR(1) = "Notices"
EW_USER_LEVEL_TABLE_NAME(2) = "Users"
EW_USER_LEVEL_TABLE_CAPTION(2) = "Users"
EW_USER_LEVEL_TABLE_VAR(2) = "Users"
EW_USER_LEVEL_TABLE_NAME(3) = "Approved Notices"
EW_USER_LEVEL_TABLE_CAPTION(3) = "Approved Notices"
EW_USER_LEVEL_TABLE_VAR(3) = "Approved_Notices"
EW_USER_LEVEL_TABLE_NAME(4) = "Unapproved Notices"
EW_USER_LEVEL_TABLE_CAPTION(4) = "Unapproved Notices"
EW_USER_LEVEL_TABLE_VAR(4) = "Unapproved_Notices"
EW_USER_LEVEL_TABLE_NAME(5) = "AuditTrail"
EW_USER_LEVEL_TABLE_CAPTION(5) = "Audit Trail"
EW_USER_LEVEL_TABLE_VAR(5) = "AuditTrail"
EW_USER_LEVEL_TABLE_NAME(6) = "Themes"
EW_USER_LEVEL_TABLE_CAPTION(6) = "Themes"
EW_USER_LEVEL_TABLE_VAR(6) = "Themes"

' Dynamic user level settings
' User level definition table/field names

Const EW_USER_LEVEL_TABLE = "[UserLevels]"
Const EW_USER_LEVEL_ID_FIELD = "[UserLevelID]"
Const EW_USER_LEVEL_NAME_FIELD = "[UserLevelName]"

' User Level privileges table/field names
Const EW_USER_LEVEL_PRIV_TABLE = "[UserLevelPermissions]"
Const EW_USER_LEVEL_PRIV_TABLE_NAME_FIELD = "[TableName]"
Const EW_USER_LEVEL_PRIV_USER_LEVEL_ID_FIELD = "[UserLevelID]"
Const EW_USER_LEVEL_PRIV_PRIV_FIELD = "[Permission]"

' User level constants
Const EW_USER_LEVEL_COMPAT = True ' Use old user level values

'Const EW_USER_LEVEL_COMPAT = False ' Use new user level values (separate values for View/Search)
Const EW_ALLOW_ADD = 1 ' Add
Const EW_ALLOW_DELETE = 2 ' Delete
Const EW_ALLOW_EDIT = 4 ' Edit
Const EW_ALLOW_LIST = 8 ' List
Dim EW_ALLOW_VIEW, EW_ALLOW_SEARCH ' View / Search
If EW_USER_LEVEL_COMPAT Then
	EW_ALLOW_VIEW = 8 ' View
	EW_ALLOW_SEARCH = 8 ' Search
Else
	EW_ALLOW_VIEW = 32 ' View
	EW_ALLOW_SEARCH = 64 ' Search
End If
Const EW_ALLOW_REPORT = 8 ' Report
Const EW_ALLOW_ADMIN = 16 ' Admin

' Hierarchical User ID
Const EW_USER_ID_IS_HIERARCHICAL = True ' True to show all level / False to show 1 level

' Use subquery for master/detail user id checking
Const EW_USE_SUBQUERY_FOR_MASTER_USER_ID = False ' True to use subquery / False to skip

' User table filters
Const EW_USER_TABLE = "[Users]"
Const EW_USER_NAME_FILTER = "([Username] = '%u')"
Const EW_USER_ID_FILTER = ""
Const EW_USER_EMAIL_FILTER = "([Email] = '%e')"
Const EW_USER_ACTIVATE_FILTER = ""
Const EW_USER_PROFILE_FIELD_NAME = "Profile"

' User Profile Constants
Const EW_USER_PROFILE_KEY_SEPARATOR = "="
Const EW_USER_PROFILE_FIELD_SEPARATOR = ","
Const EW_USER_PROFILE_SESSION_ID = "SessionID"
Const EW_USER_PROFILE_LAST_ACCESSED_DATE_TIME = "LastAccessedDateTime"
Const EW_USER_PROFILE_SESSION_TIMEOUT = 20
Const EW_USER_PROFILE_LOGIN_RETRY_COUNT = "LoginRetryCount"
Const EW_USER_PROFILE_LAST_BAD_LOGIN_DATE_TIME = "LastBadLoginDateTime"
Const EW_USER_PROFILE_MAX_RETRY = 3
Const EW_USER_PROFILE_RETRY_LOCKOUT = 20
Const EW_USER_PROFILE_LAST_PASSWORD_CHANGED_DATE = "LastPasswordChangedDate"
Const EW_USER_PROFILE_PASSWORD_EXPIRE = 90

' Date separator / format
Const EW_DATE_SEPARATOR = "/"
Const EW_DEFAULT_DATE_FORMAT = 7
Const EW_UNFORMAT_YEAR = 50 ' Unformat year
%>
<!--#include file="email_cfg.asp"-->
<%
'Const EW_MAX_EMAIL_RECIPIENT = 3 ' already defined in shared
Const EW_MAX_EMAIL_SENT_COUNT = 3
Dim EW_EXPORT_EMAIL_COUNTER
EW_EXPORT_EMAIL_COUNTER = EW_SESSION_STATUS & "_EmailCounter"

' File upload constants
Const EW_UPLOAD_DEST_PATH = "" ' Upload destination path
Const EW_UPLOAD_ALLOWED_FILE_EXT = "gif,jpg,jpeg,bmp,png,doc,xls,pdf,zip" ' Allowed file extensions
Const EW_UPLOAD_CHARSET = "windows-1252" ' Upload charset
Const EW_MAX_FILE_SIZE = 2000000 ' Max file size
Const EW_THUMBNAIL_DEFAULT_WIDTH = 0 ' Thumbnail default width
Const EW_THUMBNAIL_DEFAULT_HEIGHT = 0 ' Thumbnail default height
Const EW_THUMBNAIL_DEFAULT_INTERPOLATION = 1 ' Thumbnail default interpolation

' Export original value
Const EW_EXPORT_ORIGINAL_VALUE = False ' True to export original value
Const EW_EXPORT_FIELD_CAPTION = False ' True to export field caption
Const EW_EXPORT_CSS_STYLES = True ' True to export css styles
Const EW_EXPORT_MASTER_RECORD = True ' True to export master record
Const EW_EXPORT_MASTER_RECORD_FOR_CSV = False ' True to export master record for CSV

' Use token in Url
Const EW_USE_TOKEN_IN_URL = False ' do not use token in url

' Const EW_USE_TOKEN_IN_URL = True ' use token in url
' Use ILIKE for PostgreSql

Const EW_USE_ILIKE_FOR_POSTGRESQL = True

' Use collation for MySQL
Const EW_LIKE_COLLATION_FOR_MYSQL = ""

' Null / Not Null values
Const EW_NULL_VALUE = "##null##"
Const EW_NOT_NULL_VALUE = "##notnull##"

' Search multi value option
' 1 - no multi value
' 2 - AND all multi values
' 3 - OR all multi values

Const EW_SEARCH_MULTI_VALUE_OPTION = 3

' Validate option
Const EW_CLIENT_VALIDATE = True
Const EW_SERVER_VALIDATE = False

' Blob field byte count for Hash value calculation
Const EW_BLOB_FIELD_BYTE_COUNT = 200

' Use DOM XML
Const EW_USE_DOM_XML = False

' Cookie expiry time
Const EW_COOKIE_EXPIRY_TIME = 365

' -------------------------------------
'  ASPMaker 9 global variables (begin)
'
' Global variables
' Common

Dim Page ' Common page object
Dim Table ' Main table
Dim MasterPage ' Master page

'Dim MasterTable ' Master table
Dim Language ' Language object
Dim gsLanguage ' Current language
Dim EW_PAGE_ID ' Page ID
Dim EW_TABLE_NAME ' Table name
Dim Conn, Rs, RsDtl, i
Dim Security
Dim UserProfile
Dim ObjForm
Dim PagerItem
Dim EventArgs
Dim StartTimer ' Start timer

' Used for debug
Dim gsDebugMsg

' Used by ValidateForm/ValidateSearch
Dim gsFormError, gsSearchError

' Used by *master.asp
Dim gsMasterReturnUrl

' Used by header.asp, export checking
Dim gsExport, gsExportFile
Dim gsEmailSender, gsEmailRecipient, gsEmailCc, gsEmailBcc, gsEmailSubject, gsEmailContent, gsEmailContentType
Dim gsEmailErrNo, gsEmailErrDesc

' Used by system generated functions
Dim sSqlWrk, sWhereWrk, sLookupTblFilter, RsWrk, ari, arwrk, armultiwrk, rowswrk, rowcntwrk, selwrk, jswrk, emptywrk, boolwrk
Dim sFilterWrk
Dim sNewFileName

' Lookup
Dim EW_RECORD_DELIMITER, EW_FIELD_DELIMITER
EW_RECORD_DELIMITER = vbCr
EW_FIELD_DELIMITER = "|"

'
'  ASPMaker 9 global variables (end)
' -----------------------------------

%>
<!--#include file="ewruserlevel.asp"-->
<%

' Menu
Const EW_MENUBAR_CLASSNAME = "ewMenuBarVertical"
Const EW_MENUBAR_ITEM_CLASSNAME = ""
Const EW_MENUBAR_ITEM_LABEL_CLASSNAME = ""
Const EW_MENU_CLASSNAME = "ewMenuBarVertical"
Const EW_MENU_ITEM_CLASSNAME = ""
Const EW_MENU_ITEM_LABEL_CLASSNAME = ""
%>
