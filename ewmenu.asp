<%@ CodePage="1252" LCID="2057" %>
<!--#include file="ewcfg9.asp"-->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<%

' Variables for checking menu items for user level
Const EW_SESSION_MENU_AR_USER_LEVEL_PRIV = "Billboard_arUserLevelPriv" ' User Level Privilege Array
Const EW_SESSION_MENU_USER_LEVEL = "Billboard_Status_UserLevelValue" ' User level value
Const EW_MENU_ALLOW_ADMIN = 16
Dim arMenuUserLevelPriv

' Restore user level privilege
If IsArray(Session(EW_SESSION_MENU_AR_USER_LEVEL_PRIV)) Then
	arMenuUserLevelPriv = Session(EW_SESSION_MENU_AR_USER_LEVEL_PRIV)
End If

' Check if menu item is allowed for current user level
Function AllowListMenu(TableName)
		Dim i, priv, thispriv, userlevellist
		userlevellist = CurrentUserLevelList ' Get user level id list
		If userlevellist & "" = "" Then ' Not defined, just get user level
			userlevellist = CurrentUserLevel
		End If
		If IsLoggedIn() Then
			If IsListItem(userlevellist, "-1") Then
				AllowListMenu = True
			Else
				AllowListMenu = False
				priv = 0
				If IsArray(arMenuUserLevelPriv) Then
					For i = 0 to UBound(arMenuUserLevelPriv, 2)
						If CStr(arMenuUserLevelPriv(0, i)) = CStr(TableName) Then
							If IsListItem(userlevellist, arMenuUserLevelPriv(1, i)) Then
								thispriv = arMenuUserLevelPriv(2, i)
								If IsNull(thispriv) Then thispriv = 0
								If Not IsNumeric(thispriv) Then thispriv = 0
								thispriv = CLng(thispriv)
								priv = priv Or thispriv
							End If
						End If
					Next
				End If
				AllowListMenu = CBool(priv And 8)
			End If
		Else
			AllowListMenu = False
		End If
End Function

' Is list item
Function IsListItem(list, item)
	Dim ar, i
	If list = "" Then
		IsListItem = False
	Else
		ar = Split(list, ",")
		For i = 0 to UBound(ar)
			If CStr(item&"") = CStr(Trim(ar(i)&"")) Then
				IsListItem = True
				Exit Function
			End If
		Next
		IsListItem = False
	End If
End Function
%>
<%

' Get Menu Text
Function GetMenuText(Id, Text)
	GetMenuText = Language.MenuPhrase(Id, "MenuText")
	If GetMenuText = "" Then GetMenuText = Text
End Function
%>
<!-- Begin Main Menu -->
<div Class="navbar">
<div class="navbar-inner">
	 <ul class="nav">
<%
' Open connection to the database
If IsEmpty(Conn) Then Call ew_Connect()
%>

<%
Dim rsall,strSQLall, resultall
strSQLall = "Select count(*) AS allnoticecount from [Notices] where Date() >= [sdate] AND Date() <= [edate] AND [approved] = True and [group] <> 'Digital Signage';"
Set rsall = Server.CreateObject("ADODB.Recordset")	
rsall.Open strSQLall, Conn 
%>
<% resultall = rsall("allnoticecount") %>
<% if resultall > 0 then %>
	 <li><a href="Approved_noticeslist.asp?cmd=reset"><i class="icon-bullhorn icon-white"></i>&nbsp;All Notices&nbsp;<span class="label label-success"><%= resultall %></span></a></li>
<% else %>
	 <li><a href="Approved_noticeslist.asp?cmd=reset"><i class="icon-bullhorn icon-white"></i>&nbsp;All Notices</a></li>
<% end if %>
<%
rsall.close
Set rsall = Nothing
Set strSQLall = Nothing
	%>
	 <li class="divider-vertical"></li>
<% if isloggedin() then %>
	 <li class='dropdown'><a class='dropdown-toggle' data-toggle='dropdown' href='#'><i class="icon-filter icon-white"></i>&nbsp;Filter Notices<b class='caret'></b></a>
	 	 <ul class='dropdown-menu'>
<%
Dim rsa,strSQLa, resulta
strSQLa = "Select count(*) AS anoticecount from [Notices] where Date() >= [sdate] AND Date() <= [edate] AND [approved] = True AND [Group] = 'Year9';"
Set rsa = Server.CreateObject("ADODB.Recordset")	
rsa.Open strSQLa, Conn 
%>
<% resulta = rsa("anoticecount") %>
<% if resulta > 0 then %>
		 	 <li><a href="year9_noticeslist.asp?cmd=reset">Year 9&nbsp;<span class="label label-success"><%= resulta %></span></a></li>
<% end if %>			 
<%
rsa.close
Set rsa = Nothing
Set strSQLa = Nothing
	%><%
Dim rsb,strSQLb, resultb
strSQLb = "Select count(*) AS bnoticecount from [Notices] where Date() >= [sdate] AND Date() <= [edate] AND [approved] = True AND [Group] = 'Year10';"
Set rsb = Server.CreateObject("ADODB.Recordset")	
rsb.Open strSQLb, Conn 
%>
<% resultb = rsb("bnoticecount") %>
<% if resultb > 0 then %>
		 	 <li><a href="year10_noticeslist.asp?cmd=reset">Year 10&nbsp;<span class="label label-success"><%= resultb %></span></a></li>
<% end if %>
<%
rsb.close
Set rsb = Nothing
Set strSQLb = Nothing
	%><%
Dim rsc,strSQLc, resultc
strSQLc = "Select count(*) AS cnoticecount from [Notices] where Date() >= [sdate] AND Date() <= [edate] AND [approved] = True AND [Group] = 'Year11';"
Set rsc = Server.CreateObject("ADODB.Recordset")	
rsc.Open strSQLc, Conn 
%>
<% resultc = rsc("cnoticecount") %>
<% if resultc > 0 then %>
		 	 <li><a href="year11_noticeslist.asp?cmd=reset">Year 11&nbsp;<span class="label label-success"><%= resultc %></span></a></li>
<% end if %>
<%
rsc.close
Set rsc = Nothing
Set strSQLc = Nothing
	%><%
Dim rsd,strSQLd, resultd
strSQLd = "Select count(*) AS dnoticecount from [Notices] where Date() >= [sdate] AND Date() <= [edate] AND [approved] = True AND [Group] = 'Year12';"
Set rsd = Server.CreateObject("ADODB.Recordset")	
rsd.Open strSQLd, Conn 
%>
<% resultd = rsd("dnoticecount") %>
<% if resultd > 0 then %>
		 	 <li><a href="year12_noticeslist.asp?cmd=reset">Year 12&nbsp;<span class="label label-success"><%= resultd %></span></a></li>
<% end if %>
<%
rsd.close
Set rsd = Nothing
Set strSQLd = Nothing
	%>
	<%
Dim rse,strSQLe, resulte
strSQLe = "Select count(*) AS enoticecount from [Notices] where Date() >= [sdate] AND Date() <= [edate] AND [approved] = True AND [Group] = 'Year13';"
Set rse = Server.CreateObject("ADODB.Recordset")	
rse.Open strSQLe, Conn 
%>
<% resulte = rse("enoticecount") %>
<% if resulte > 0 then %>
		 	 <li><a href="year13_noticeslist.asp?cmd=reset">Year 13&nbsp;<span class="label label-success"><%= resulte %></span></a></li>
<% end if %>
<%
rse.close
Set rse = Nothing
Set strSQLe = Nothing
	%>
		<%
Dim rsf,strSQLf, resultf
strSQLf = "Select count(*) AS enoticecount from [Notices] where Date() >= [sdate] AND Date() <= [edate] AND [approved] = True AND [Group] = 'Post16';"
Set rsf = Server.CreateObject("ADODB.Recordset")	
rsf.Open strSQLf, Conn 
%>
<% resultf = rsf("enoticecount") %>
<% if resultf > 0 then %>
		 	 <li><a href="post16_noticeslist.asp?cmd=reset">Post 16&nbsp;<span class="label label-success"><%= resultf %></span></a></li>
<% end if %>
<%
rsf.close
Set rsf = Nothing
Set strSQLf = Nothing
	%>

	
	<%
Dim rsws,strSQLws, resultws
strSQLws = "Select count(*) AS wsnoticecount from [Notices] where Date() >= [sdate] AND Date() <= [edate] AND [approved] = True AND [Group] = 'WholeSchool';"
Set rsws = Server.CreateObject("ADODB.Recordset")	
rsws.Open strSQLws, Conn 
%>
<% resultws = rsws("wsnoticecount") %>
<% if resultws > 0 then %>
		 	 <li class="divider"></li>
			 <li><a href="approved_noticeslist.asp?psearch=WholeSchool">Whole School&nbsp;<span class="label label-success"><%= resultws %></span></a></li>
<% end if %>
<%
rsws.close
Set rsws = Nothing
Set strSQLws = Nothing
	%>	
	<li class="divider"></li>
						 <li><a href="Headerless_notices.asp?cmd=reset">Notice Slider</a></li>					 
						 <li><a href="Headerless_scroller.asp?cmd=reset">Digital Signage</a></li>
			<% if IsLoggedIn() then %>
			<li><a href="frog_notices.asp?cmd=reset" target="_blank">Frog Notices</a></li>
			<% end if %>
		 			
			</ul>
	<li class="divider-vertical"></li>
<% end if %>
	 <li class='dropdown'><a class='dropdown-toggle' data-toggle='dropdown' href='#'><i class="icon-question-sign icon-white"></i>&nbsp;Help<b class='caret'></b></a>
	 	 <ul class='dropdown-menu'>
		 <li><a href="docs/billboard userguide.pdf" target="_blank"><i class="icon-book"></i>&nbsp;User Guide</a></li>
		 <% if IsLoggedIn() then %>
			<li><a href="docs/Notices admin guide.pdf" target="_blank"><i class="icon-book"></i>&nbsp;Admin Guide</a></li>
		 <% end if %>
		 </ul>
	 <li class="divider-vertical"></li>
			
	 <% if IsLoggedIn() then %>
	 <li class='dropdown'><a class='dropdown-toggle' data-toggle='dropdown' href='#'><i class="icon-cog icon-white"></i>&nbsp;Administration<b class='caret'></b></a>
	 	 <ul class='dropdown-menu'>
<%
Dim rsunav,strSQLun, result
strSQLun = "Select count(*) AS unnoticecount from [Notices] where [approved] = False;"
Set rsunav = Server.CreateObject("ADODB.Recordset")	
rsunav.Open strSQLun, Conn 
%>
<% result = rsunav("unnoticecount")%>
<% if result > 0 then %>
		  	  <li><a href="unapproved_noticeslist.asp"><i class="icon-exclamation-sign"></i>&nbsp;Unapproved Notices&nbsp;<span class="label label-important"><%= result %></span></a></li>
<% end if %>
<%
rsunav.close
Set rsunav = Nothing
Set strSQLun = Nothing
	%>
			 <li><a href="groupslist.asp"><i class="icon-th"></i>&nbsp;Groups</a></li>
<%
Dim rsactive, SQLactive, Active, rsinactive, SQLinactive, inactive, rsexpired, SQLexpired, expired, archived, SQLdigi, srdigi, digi
SQLactive = "Select count(*) AS activecount from [Notices] Where Date() >= [sdate] AND Date() <= [edate] AND [approved] = True and [group] <> 'digital signage';"
Set rsactive = Server.CreateObject("ADODB.Recordset")	
rsactive.Open SQLactive, Conn 
SQLinactive = "Select count(*) AS inactivecount from [Notices] Where Date() < [sdate] AND [approved] = True;"
Set rsinactive = Server.CreateObject("ADODB.Recordset")	
rsinactive.Open SQLinactive, Conn 
SQLexpired = "Select count(*) AS expiredcount from [Notices] Where Date() > [Edate] AND [approved] = True;"
Set rsexpired = Server.CreateObject("ADODB.Recordset")	
rsexpired.Open SQLexpired, Conn 
SQLdigi = "Select count(*) AS digicount from [Notices] Where Date() >= [sdate] AND Date() <= [edate] AND [approved] = True and [group] = 'digital signage';"
Set rsdigi = Server.CreateObject("ADODB.Recordset")	
rsdigi.Open SQLdigi, Conn 
%>
<% Active = rsactive("activecount")%>
<% archived = (rsinactive("inactivecount")+rsexpired("expiredcount"))%>
<% inactive = rsinactive("inactivecount")%>
<% expired = rsexpired("expiredcount")%>
<% digi = rsdigi("digicount")%>


			 <li><a href="noticeslist.asp"><i class="icon-bullhorn"></i>&nbsp;Notices&nbsp;
			 <% if Active > 0 then %>
				<span rel="tooltip" title="Active notices" class="label label-success"><%= Active %></span>&nbsp;
			 <% end if %>
			 <% if inactive > 0 then %>
				<span rel="tooltip" title="Pending notices" class="label label-warning"><%= inactive %></span>&nbsp;
			 <% end if %>
			 <% if expired > 0 then %>
				<span rel="tooltip" title="Expired notices" class="label"><%= expired %></span>
			 <% end if %>
			 <% if active > 0 or Inactive > 0 or expired > 0 then %>
			 &nbsp;|&nbsp;
			 <% end if %>
			 <% if digi > 0 then %>
				<i class="icon-eye-open"></i>&nbsp;<span rel="tooltip" title="Digital Signage notices" class="label label-info"><%= digi %></span>
			 <% end if %>
			 </a></li>
<%
rsactive.close
rsinactive.close
rsexpired.close
rsdigi.close
Set rsactive = Nothing
Set rsinactive = Nothing
Set rsexpired = Nothing
Set rsdigi = Nothing
Set SQLactive = Nothing
Set SQLdigi = Nothing
Set SQLinactive = Nothing	
Set SQLexpired = Nothing	
%>
<% if session(EW_SESSION_USER_LEVEL_ID) = -1 then %>
			 <li class="divider"></li>
			 <li><a href="themeslist.asp"><i class="icon-picture"></i>&nbsp;Themes</a></li>
			 <li class="divider"></li>
			 <li><a href="userslist.asp"><i class="icon-user"></i>&nbsp;Users</a></li>
			 <li><a href="userlevelslist.asp"><i class="icon-warning-sign"></i>&nbsp;User Levels</a></li>
			 <li class="divider"></li>
			 <li><a href="audittraillist.asp"><i class="icon-fire"></i>&nbsp;Audit Trail</a></li>
<% End If %>
		</ul>
	<li class="divider-vertical"></li>
	<% end if %>
	<% if IsLoggedIn() And session(EW_SESSION_USER_LEVEL_ID) <> -1 then %>
	<li class='dropdown'><a class='dropdown-toggle' data-toggle='dropdown' href='#'><i class="icon-certificate icon-white"></i>&nbsp;Profile<b class='caret'></b></a>
	 	 <ul class='dropdown-menu'>
		 <li><a href="changepwd.asp"><i class="icon-lock"></i>&nbsp;Change password</a></li>
		 </ul>
	</li>
	<% end if %>
</UL>
<form class="navbar-form pull-right" action="noticesadd.asp">
<button type="submit" class="btn btn-success"><i class="icon-pencil icon-white"></i>&nbsp;New Notice</button>
</form>

</div>
</div>
<!-- End Main Menu -->
