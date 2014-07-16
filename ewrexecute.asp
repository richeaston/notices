' Compatibility codes for ASP Report Maker
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
EW_SESSION_LANGUAGE_FILE_CACHE = EW_PROJECT_NAME & "_LanguageFile_WkQV5eL40cJoon2B" ' Language File Cache
Dim EW_SESSION_LANGUAGE_CACHE
EW_SESSION_LANGUAGE_CACHE = EW_PROJECT_NAME & "_Language_WkQV5eL40cJoon2B" ' Language Cache
Dim EW_SESSION_LANGUAGE_ID
EW_SESSION_LANGUAGE_ID = EW_PROJECT_NAME & "_LanguageId" ' Language ID
'
' *** DO NOT CHANGE BELOW
'
' Init language object
Set Language = new cLanguage
Call Language.LoadPhrases()
' ------------------------
'  Language class (begin)
'
Class cLanguage
	Dim LanguageId
	Dim objDOM
	Dim Col
	Dim LanguageFolder
	Dim Key
	' Class initialize
	Private Sub Class_Initialize
		LanguageFolder = EW_LANGUAGE_FOLDER
	End Sub
	' Load phrases
	Public Sub LoadPhrases()
		' Set up file list
		LoadFileList()
		' Set up language id
		If Request.QueryString("language") <> "" Then
			LanguageId = Request.QueryString("language")
			Session(EW_SESSION_LANGUAGE_ID) = LanguageId
		ElseIf Session(EW_SESSION_LANGUAGE_ID) <> "" Then
			LanguageId = Session(EW_SESSION_LANGUAGE_ID)
		Else
			LanguageId = EW_LANGUAGE_DEFAULT_ID
		End If
		gsLanguage = LanguageId
		If EW_USE_DOM_XML Then
			Set objDOM = ew_CreateXmlDom()
			objDOM.async = False
		Else
			Set Col = Server.CreateObject("Scripting.Dictionary")
		End If
		' Load current language
		Load(LanguageId)
	End Sub
	' Terminate
	Private Sub Class_Terminate()
		If EW_USE_DOM_XML Then
			Set objDOM = Nothing
		Else
			Set Col = Nothing
		End If
	End Sub
	' Load language file list
	Private Sub LoadFileList()
		If IsArray(EW_LANGUAGE_FILE) Then
			For i = 0 to UBound(EW_LANGUAGE_FILE)
				EW_LANGUAGE_FILE(i)(1) = LoadFileDesc(Server.MapPath(LanguageFolder & EW_LANGUAGE_FILE(i)(2)))
			Next
		End If
	End Sub
	' Load language file description
	Private Function LoadFileDesc(File)
		LoadFileDesc = ""
		Set objDOM = ew_CreateXmlDom()
		objDOM.async = False
		objDOM.Load(File)
		If objDOM.ParseError.ErrorCode = 0 Then
			LoadFileDesc = GetNodeAtt(objDOM.documentElement, "desc")
		End If
	End Function
	' Load language file
	Private Sub Load(id)
		Dim sFileName
		sFileName = GetFileName(id)
		If sFileName = "" Then
			sFileName = GetFileName(EW_LANGUAGE_DEFAULT_ID)
		End If
		If sFileName = "" Then Exit Sub
		If EW_USE_DOM_XML Then
			objDOM.Load(sFileName)
			If objDOM.ParseError.ErrorCode = 0 Then
				objDOM.setProperty "SelectionLanguage", "XPath"
			End If
		Else
			XmlToCollection(sFileName)
		End If
	End Sub
	Private Sub IterateNodes(Node)
		If Node.baseName = vbNullString Then Exit Sub
		Dim Index, Id
		If Node.nodeType = 1 And Node.baseName <> "ew-language" Then ' NODE_ELEMENT
			Id = ""
			If Node.attributes.length > 0 Then
				Id = Node.getAttribute("id")
			End If
			If Node.hasChildNodes Then
				Key = Key & Node.baseName & "/"
				If Id <> "" Then Key = Key & Id & "/"
			End If
			If Id <> "" And Not Node.hasChildNodes Then ' phrase
				Id = Node.baseName & "/" & Id
				If Node.getAttribute("client") = "1" Then Id = Id & "/1"
				If Id <> "" Then 
					Col(Key & Id) = Node.getAttribute("value")
					'Response.Write Key & Id & "=" & Node.getAttribute("value") & "<br>"
				End If
			End If
		End If
		If Node.hasChildNodes Then
			For Index = 0 To Node.childNodes.length - 1
				IterateNodes Node.childNodes(Index)
			Next
			Index	=	InStrRev(Key, "/"	&	Node.baseName & "/")
			If Index > 0	Then Key = Left(Key, Index)
		End If
	End Sub
	' Convert XML to Collection
	Private Sub XmlToCollection(File)
		Dim I, xmlr
		Key = "/"
		Set xmlr = ew_CreateXmlDom()
		xmlr.async = False
		xmlr.Load(File)
		For I = 0 To xmlr.childNodes.length - 1
			IterateNodes xmlr.childNodes(I)
		Next
		Set xmlr = Nothing
	End Sub
	' Get language file name
	Private Function GetFileName(Id)
		GetFileName = ""
		If IsArray(EW_LANGUAGE_FILE) Then
			For i = 0 to UBound(EW_LANGUAGE_FILE)
				If EW_LANGUAGE_FILE(i)(0) = Id Then
					GetFileName = Server.MapPath(LanguageFolder & EW_LANGUAGE_FILE(i)(2))
					Exit For
				End If
			Next
		End If
	End Function
	' Get node attribute
	Private Function GetNodeAtt(Node, Att)
		If Not (Node Is Nothing) Then
			GetNodeAtt = Node.getAttribute(Att)
		Else
			GetNodeAtt = ""
		End If
	End Function
	' Get phrase
	Public Function Phrase(Id)
		If EW_USE_DOM_XML Then
			Phrase = GetNodeAtt(objDOM.SelectSingleNode("//global/phrase[@id='" & LCase(Id) & "']"), "value")
		Else
			If Col.Exists("/global/phrase/" & LCase(Id)) Then
				Phrase = Col("/global/phrase/" & LCase(Id))
			ElseIf Col.Exists("/global/phrase/" & LCase(Id) & "/1") Then
				Phrase = Col("/global/phrase/" & LCase(Id) & "/1")
			Else
				Phrase = ""
			End If
		End If
	End Function
	' Set phrase
	Public Sub SetPhrase(Id, Value)
		If Not EW_USE_DOM_XML Then
			If Col.Exists("/global/phrase/" & LCase(Id)) Then
				Col("/global/phrase/" & LCase(Id)) = Value
			ElseIf Col.Exists("/global/phrase/" & LCase(Id) & "/1") Then
				Col("/global/phrase/" & LCase(Id) & "/1") = Value
			End If
		End If
	End Sub
	' Get project phrase
	Public Function ProjectPhrase(Id)
		If EW_USE_DOM_XML Then
			ProjectPhrase = GetNodeAtt(objDOM.SelectSingleNode("//project/phrase[@id='" & LCase(Id) & "']"), "value")
		Else
			ProjectPhrase = Col("/project/phrase/" & LCase(Id))
		End If  
	End Function
	' Set project phrase
	Public Sub SetProjectPhrase(Id, Value)
		If Not EW_USE_DOM_XML Then
			Col("/project/phrase/" & LCase(Id)) = Value
		End If  
	End Sub
	' Get menu phrase
	Public Function MenuPhrase(MenuId, Id)
		If EW_USE_DOM_XML Then
			MenuPhrase = GetNodeAtt(objDOM.SelectSingleNode("//project/menu[@id='" & MenuId & "']/phrase[@id='" & LCase(Id) & "']"), "value")
		Else
			MenuPhrase = Col("/project/menu/" & MenuId & "/phrase/" & LCase(Id))
		End If  
	End Function
	' Set menu phrase
	Public Sub SetMenuPhrase(MenuId, Id, Value)
		If Not EW_USE_DOM_XML Then
			Col("/project/menu/" & MenuId & "/phrase/" & LCase(Id)) = Value
		End If  
	End Sub
	' Get table phrase
	Public Function TablePhrase(TblVar, Id)
		If EW_USE_DOM_XML Then
			TablePhrase = GetNodeAtt(objDOM.SelectSingleNode("//project/table[@id='" & LCase(TblVar) & "']/phrase[@id='" & LCase(Id) & "']"), "value")
		Else
			TablePhrase = Col("/project/table/" & LCase(TblVar) & "/phrase/" & LCase(Id))
		End If  
	End Function
	' Set table phrase
	Public Sub SetTablePhrase(TblVar, Id, Value)
		If Not EW_USE_DOM_XML Then
			Col("/project/table/" & LCase(TblVar) & "/phrase/" & LCase(Id)) = Value
		End If  
	End Sub
	' Get field phrase
	Public Function FieldPhrase(TblVar, FldVar, Id)
		If EW_USE_DOM_XML Then
			FieldPhrase = GetNodeAtt(objDOM.SelectSingleNode("//project/table[@id='" & LCase(TblVar) & "']/field[@id='" & LCase(FldVar) & "']/phrase[@id='" & LCase(Id) & "']"), "value")
		Else
			FieldPhrase = Col("/project/table/" & LCase(TblVar) & "/field/" & LCase(FldVar) & "/phrase/" & LCase(Id))
		End If  
	End Function
	' Set field phrase
	Public Sub SetFieldPhrase(TblVar, FldVar, Id, Value)
		If Not EW_USE_DOM_XML Then
			Col("/project/table/" & LCase(TblVar) & "/field/" & LCase(FldVar) & "/phrase/" & LCase(Id)) = Value
		End If  
	End Sub
	' Output XML as JSON
	Public Function XmlToJSON(XPath)
		Dim Node, NodeList, Id, Value, Str
		Set NodeList = objDOM.selectNodes(XPath)
		Str = "{"
		For Each Node In NodeList
			Id = GetNodeAtt(Node, "id")
			Value = GetNodeAtt(Node, "value")
			Str = Str & """" & ew_JsEncode2(Id) & """:""" & ew_JsEncode2(Value) & ""","
		Next  
		If Right(Str, 1) = "," Then Str = Left(Str, Len(Str)-1)
		Str = Str & "}"
		XmlToJSON = Str
	End Function
	' Output collection as JSON
	Public Function CollectionToJSON(Prefix, Suffix)
		Dim Name, Id, Str, Pos, Keys, I
		Str = "{"
		Keys = Col.Keys
		For I = 0 To Ubound(Keys)
			Name = Keys(I)
			If Left(Name, Len(Prefix)) = Prefix Then
				If Suffix <> "" And Right(Name, Len(Suffix)) = Suffix Then
					Pos = InStrRev(Name, Suffix)
					Id = Mid(Name, Len(Prefix) + 1, Pos - Len(Prefix) - 1)
				Else
					Id = Mid(Name, Len(Prefix) + 1)
				End If
				Str = Str & """" & ew_JsEncode2(Id) & """:""" & ew_JsEncode2(Col(Name)) & ""","
			End If
		Next  
		If Right(Str, 1) = "," Then Str = Left(Str, Len(Str)-1)
		Str = Str & "}"
		CollectionToJSON = Str
	End Function
	' Output all phrases as JSON
	Public Function AllToJSON()
		If EW_USE_DOM_XML Then
			AllToJSON ="var ewLanguage = new ew_Language(" & XmlToJSON("//global/phrase") & ");"
		Else
			AllToJSON = "var ewLanguage = new ew_Language(" & CollectionToJSON("/global/phrase/", "") & ");"
		End If
	End Function
	' Output client phrases as JSON
	Public Function ToJSON()
		If EW_USE_DOM_XML Then
			ToJSON = "var ewLanguage = new ew_Language(" & XmlToJSON("//global/phrase[@client='1']") & ");"
		Else
			ToJSON = "var ewLanguage = new ew_Language(" & CollectionToJSON("/global/phrase/", "/1") & ");"
		End If
	End Function
End Class
'
'  Language class (end)
' ----------------------
' Report language folder
Function ew_ReportLanguageFolder
	On Error Resume Next
	ew_ReportLanguageFolder = EW_REPORT_LANGUAGE_FOLDER
End Function
' Adjust text for caption
Function ew_BtnCaption(Caption)
	Dim Min, Pad
	Min = 10
	ew_BtnCaption = Caption
	If Len(Caption) < Min Then
		Pad = Abs(Int((Min - Len(Caption))/2*-1))
		ew_BtnCaption = String(Pad, " ") & Caption & String(Pad, " ")
	End If
End Function
' Encode value for single-quoted JavaScript string
Function ew_JsEncode(val)
	val = Replace(val & "", "\", "\\")
	val = Replace(val, "'", "\'")
'	val = Replace(val, vbCrLf, "\r\n")
'	val = Replace(val, vbCr, "\r")
'	val = Replace(val, vbLf, "\n")
	val = Replace(val, vbCrLf, "<br>")
	val = Replace(val, vbCr, "<br>")
	val = Replace(val, vbLf, "<br>")
	ew_JsEncode = val
End Function
' Encode value for double-quoted Javascript string
Function ew_JsEncode2(val)
	val = Replace(val & "", "\", "\\")
	val = Replace(val, """", "\""")
'	val = Replace(val, vbCrLf, "\r\n")
'	val = Replace(val, vbCr, "\r")
'	val = Replace(val, vbLf, "\n")
	val = Replace(val, vbCrLf, "<br>")
	val = Replace(val, vbCr, "<br>")
	val = Replace(val, vbLf, "<br>")
	ew_JsEncode2 = val
End Function
' Encode value to single-quoted Javascript string for HTML attributes
Function ew_JsEncode3(val)
	val = Replace(val & "", "\", "\\")
	val = Replace(val, "'", "\'")
	val = Replace(val, """", "&quot;")
	ew_JsEncode3 = val
End Function
' Get full url
Function ew_FullUrl()
	ew_FullUrl = ew_DomainUrl() & Request.ServerVariables("SCRIPT_NAME")
End Function 
' Get domain url
Function ew_DomainUrl()
	Dim sUrl, bSSL, sPort, defPort
	sUrl = "http"
	bSSL = ew_IsHttps()
	sPort = Request.ServerVariables("SERVER_PORT")
	If bSSL Then defPort = "443" Else defPort = "80"
	If sPort = defPort Then sPort = "" Else sPort = ":" & sPort
	If bSSL Then sUrl = sUrl & "s"
	sUrl = sUrl & "://"
	sUrl = sUrl & Request.ServerVariables("SERVER_NAME") & sPort
	ew_DomainUrl = sUrl
End Function 
' YUI files host
Function ew_YuiHost()
	ew_YuiHost = "yui290/" ' Use local files
End Function
' Check if HTTPS
Function ew_IsHttps()
	ew_IsHttps = (Request.ServerVariables("HTTPS") <> "" And Request.ServerVariables("HTTPS") <> "off")
End Function
' Get current url
Function ew_CurrentUrl()
	Dim s, q
	s = Request.ServerVariables("SCRIPT_NAME")
	q = Request.ServerVariables("QUERY_STRING")
	If q <> "" Then s = s & "?" & q
	ew_CurrentUrl = s
End Function
' Convert to full url
Function ew_ConvertFullUrl(url)
	Dim sUrl
	If url = "" Then
		ew_ConvertFullUrl = ""
	ElseIf Instr(url, "://") > 0 Then
		ew_ConvertFullUrl = url
	Else
		sUrl = ew_FullUrl
		ew_ConvertFullUrl = Mid(sUrl, 1, InStrRev(sUrl, "/")) & url
	End If
End Function
' Create XML Dom object
Function ew_CreateXmlDom()
	On Error Resume Next
	Dim ProgId
	ProgId = Array("MSXML2.DOMDocument", "Microsoft.XMLDOM") ' Add other ProgID here
	Dim i
	For i = 0 To UBound(ProgId)
		Set ew_CreateXmlDom = Server.CreateObject(ProgId(i))
		If Err.Number = 0 Then Exit For
	Next
End Function
' *** DO NOT CHANGE
