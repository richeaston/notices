<%

' ASPMaker configuration for Table AuditTrail
Dim AuditTrail

' Define table class
Class cAuditTrail

	' Class Initialize
	Private Sub Class_Initialize()
		UseTokenInUrl = EW_USE_TOKEN_IN_URL
		ExportOriginalValue = EW_EXPORT_ORIGINAL_VALUE
		ExportAll = True
		Set RowAttrs = New cAttributes ' Row attributes
		AllowAddDeleteRow = ew_AllowAddDeleteRow() ' Allow add/delete row
		DetailAdd = False ' Allow detail add
		DetailEdit = False ' Allow detail edit
		GridAddRowCount = 5 ' Grid add row count
		Call ew_SetArObj(Fields, "Id", Id)
		Call ew_SetArObj(Fields, "DateTime", DateTime)
		Call ew_SetArObj(Fields, "Script", Script)
		Call ew_SetArObj(Fields, "User", User)
		Call ew_SetArObj(Fields, "Action", Action)
		Call ew_SetArObj(Fields, "zTable", zTable)
		Call ew_SetArObj(Fields, "zField", zField)
		Call ew_SetArObj(Fields, "KeyValue", KeyValue)
		Call ew_SetArObj(Fields, "OldValue", OldValue)
		Call ew_SetArObj(Fields, "NewValue", NewValue)
	End Sub

	' Reset attributes for table object
	Public Sub ResetAttrs()
		CssClass = ""
		CssStyle = ""
		RowAttrs.Clear()
		Dim i, fld
		If IsArray(Fields) Then
			For i = 0 to UBound(Fields,2)
				Set fld = Fields(1,i)
				Call fld.ResetAttrs()
			Next
		End If
	End Sub

	' Setup field titles
	Public Sub SetupFieldTitles()
		Dim i, fld
		If IsArray(Fields) Then
			For i = 0 to UBound(Fields,2)
				Set fld = Fields(1,i)
				If fld.FldTitle <> "" Then
					fld.EditAttrs.AddAttribute "onmouseover", "ew_ShowTitle(this, '" & ew_JsEncode3(fld.FldTitle) & "');", True
					fld.EditAttrs.AddAttribute "onmouseout", "ew_HideTooltip();", True
				End If
			Next
		End If
	End Sub

	' Define table level constants
	' Use table token in Url

	Dim UseTokenInUrl

	' Table variable
	Public Property Get TableVar()
		TableVar = "AuditTrail"
	End Property

	' Table name
	Public Property Get TableName()
		TableName = "AuditTrail"
	End Property

	' Table type
	Public Property Get TableType()
		TableType = "TABLE"
	End Property

	' Table caption
	Public Property Get TableCaption()
		TableCaption = Language.TablePhrase(TableVar, "TblCaption")
	End Property

	' Page caption
	Public Property Get PageCaption(Page)
		PageCaption = Language.TablePhrase(TableVar, "TblPageCaption" & Page)
		If PageCaption = "" Then PageCaption = "Page " & Page
	End Property

	' Export Return Page
	Public Property Get ExportReturnUrl()
		If Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_EXPORT_RETURN_URL) <> "" Then
			ExportReturnUrl = Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_EXPORT_RETURN_URL)
		Else
			ExportReturnUrl = ew_CurrentPage
		End If
	End Property

	Public Property Let ExportReturnUrl(v)
		Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_EXPORT_RETURN_URL) = v
	End Property

	' Records per page
	Public Property Get RecordsPerPage()
		RecordsPerPage = Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_REC_PER_PAGE)
	End Property

	Public Property Let RecordsPerPage(v)
		Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_REC_PER_PAGE) = v
	End Property

	' Start record number
	Public Property Get StartRecordNumber()
		StartRecordNumber = Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_START_REC)
	End Property

	Public Property Let StartRecordNumber(v)
		Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_START_REC) = v
	End Property

	' Search Highlight Name
	Public Property Get HighlightName()
		HighlightName = "AuditTrail_Highlight"
	End Property

	' Advanced search
	Public Function GetAdvancedSearch(fld)
		GetAdvancedSearch = Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_ADVANCED_SEARCH & "_" & fld)
	End Function

	Public Function SetAdvancedSearch(fld, v)
		If Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_ADVANCED_SEARCH & "_" & fld) <> v Then
			Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_ADVANCED_SEARCH & "_" & fld) = v
		End If
	End Function
	Dim BasicSearchKeyword
	Dim BasicSearchType

	' Basic search Keyword
	Public Property Get SessionBasicSearchKeyword()
		SessionBasicSearchKeyword = Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_BASIC_SEARCH)
	End Property

	Public Property Let SessionBasicSearchKeyword(v)
		Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_BASIC_SEARCH) = v
	End Property

	' Basic Search Type
	Public Property Get SessionBasicSearchType()
		SessionBasicSearchType = Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_BASIC_SEARCH_TYPE)
	End Property

	Public Property Let SessionBasicSearchType(v)
		Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_BASIC_SEARCH_TYPE) = v
	End Property

	' Search where clause
	Public Property Get SearchWhere()
		SearchWhere = Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_SEARCH_WHERE)
	End Property

	Public Property Let SearchWhere(v)
		Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_SEARCH_WHERE) = v
	End Property

	' Single column sort
	Public Sub UpdateSort(ofld)
		Dim sSortField, sLastSort, sThisSort
		If CurrentOrder = ofld.FldName Then
			sSortField = ofld.FldExpression
			sLastSort = ofld.Sort
			If CurrentOrderType = "ASC" Or CurrentOrderType = "DESC" Then
				sThisSort = CurrentOrderType
			Else
				If sLastSort = "ASC" Then sThisSort = "DESC" Else sThisSort = "ASC"
			End If
			ofld.Sort = sThisSort
			SessionOrderBy = sSortField & " " & sThisSort ' Save to Session
		Else
			ofld.Sort = ""
		End If
	End Sub

	' Session WHERE Clause
	Public Property Get SessionWhere()
		SessionWhere = Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_WHERE)
	End Property

	Public Property Let SessionWhere(v)
		Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_WHERE) = v
	End Property

	' Session ORDER BY
	Public Property Get SessionOrderBy()
		SessionOrderBy = Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_ORDER_BY)
	End Property

	Public Property Let SessionOrderBy(v)
		Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_ORDER_BY) = v
	End Property

	' Session Key
	Public Function GetKey(fld)
		GetKey = Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_KEY & "_" & fld)
	End Function

	Public Function SetKey(fld, v)
		Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_KEY & "_" & fld) = v
	End Function

	' Table level SQL
	Public Property Get SqlSelect() ' Select
		SqlSelect = "SELECT * FROM [AuditTrail]"
	End Property

	Private Property Get TableFilter()
		TableFilter = ""
	End Property

	Public Property Get SqlWhere() ' Where
		Dim sWhere
		sWhere = ""
		Call ew_AddFilter(sWhere, TableFilter)
		SqlWhere = sWhere
	End Property

	Public Property Get SqlGroupBy() ' Group By
		SqlGroupBy = ""
	End Property

	Public Property Get SqlHaving() ' Having
		SqlHaving = ""
	End Property

	Public Property Get SqlOrderBy() ' Order By
		SqlOrderBy = ""
	End Property

	' SQL variables
	Dim CurrentFilter ' Current filter
	Dim CurrentOrder ' Current order
	Dim CurrentOrderType ' Current order type

	' Get sql
	Public Function GetSQL(where, orderby)
		GetSQL = ew_BuildSelectSql(SqlSelect, SqlWhere, SqlGroupBy, SqlHaving, SqlOrderBy, where, orderby)
	End Function

	' Table sql
	Public Property Get SQL()
		Dim sFilter, sSort
		sFilter = CurrentFilter
		sSort = SessionOrderBy
		SQL = ew_BuildSelectSql(SqlSelect, SqlWhere, SqlGroupBy, SqlHaving, SqlOrderBy, sFilter, sSort)
	End Property

	' Return table sql with list page filter
	Public Property Get ListSQL()
		Dim sFilter, sSort
		sFilter = SessionWhere
		Call ew_AddFilter(sFilter, CurrentFilter)
		sSort = SessionOrderBy
		ListSQL = ew_BuildSelectSql(SqlSelect, SqlWhere, SqlGroupBy, SqlHaving, SqlOrderBy, sFilter, sSort)
	End Property

	' Key filter for table
	Private Property Get SqlKeyFilter()
		SqlKeyFilter = "[Id] = @Id@"
	End Property

	' Return Key filter for table
	Public Property Get KeyFilter()
		Dim sKeyFilter
		sKeyFilter = SqlKeyFilter
		If Not IsNumeric(Id.CurrentValue) Then
			sKeyFilter = "0=1" ' Invalid key
		End If
		sKeyFilter = Replace(sKeyFilter, "@Id@", ew_AdjustSql(Id.CurrentValue)) ' Replace key value
		KeyFilter = sKeyFilter
	End Property

	' Return url
	Public Property Get ReturnUrl()

		' Get referer url automatically
		If Request.ServerVariables("HTTP_REFERER") <> "" Then
			If ew_ReferPage <> ew_CurrentPage And ew_ReferPage <> "login.asp" Then ' Referer not same page or login page
				Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_RETURN_URL) = Request.ServerVariables("HTTP_REFERER") ' Save to Session
			End If
		End If
		If Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_RETURN_URL) <> "" Then
			ReturnUrl = Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_RETURN_URL)
		Else
			ReturnUrl = "audittraillist.asp"
		End If
	End Property

	' List url
	Public Function ListUrl()
		ListUrl = "audittraillist.asp"
	End Function

	' View url
	Public Function ViewUrl()
		ViewUrl = KeyUrl("audittrailview.asp", UrlParm(""))
	End Function

	' Add url
	Public Function AddUrl()
		AddUrl = "audittrailadd.asp"

'		Dim sUrlParm
'		sUrlParm = UrlParm("")
'		If sUrlParm <> "" Then AddUrl = AddUrl & "?" & sUrlParm

	End Function

	' Edit url
	Public Function EditUrl(parm)
		EditUrl = KeyUrl("audittrailedit.asp", UrlParm(parm))
	End Function

	' Inline edit url
	Public Function InlineEditUrl()
		InlineEditUrl = KeyUrl(ew_CurrentPage, UrlParm("a=edit"))
	End Function

	' Copy url
	Public Function CopyUrl(parm)
		CopyUrl = KeyUrl("audittrailadd.asp", UrlParm(parm))
	End Function

	' Inline copy url
	Public Function InlineCopyUrl()
		InlineCopyUrl = KeyUrl(ew_CurrentPage, UrlParm("a=copy"))
	End Function

	' Delete url
	Public Function DeleteUrl()
		DeleteUrl = KeyUrl("audittraildelete.asp", UrlParm(""))
	End Function

	' Key url
	Public Function KeyUrl(url, parm)
		Dim sUrl: sUrl = url & "?"
		If parm <> "" Then sUrl = sUrl & parm & "&"
		If Not IsNull(Id.CurrentValue) Then
			sUrl = sUrl & "Id=" & Id.CurrentValue
		Else
			KeyUrl = "javascript:alert(ewLanguage.Phrase('InvalidRecord'));"
			Exit Function
		End If
		KeyUrl = sUrl
	End Function

	' Sort Url
	Public Property Get SortUrl(fld)
		If CurrentAction <> "" Or Export <> "" Or (fld.FldType = 201 Or fld.FldType = 203 Or fld.FldType = 205 Or fld.FldType = 141) Then
			SortUrl = ""
		ElseIf fld.Sortable Then
			SortUrl = ew_CurrentPage
			Dim sUrlParm
			sUrlParm = UrlParm("order=" & Server.URLEncode(fld.FldName) & "&amp;ordertype=" & fld.ReverseSort)
			SortUrl = SortUrl & "?" & sUrlParm
		Else
			SortUrl = ""
		End If
	End Property

	' Url parm
	Function UrlParm(parm)
		If UseTokenInUrl Then
			UrlParm = "t=AuditTrail"
		Else
			UrlParm = ""
		End If
		If parm <> "" Then
			If UrlParm <> "" Then UrlParm = UrlParm & "&"
			UrlParm = UrlParm & parm
		End If
	End Function

	' Get record keys from Form/QueryString/Session
	Public Function GetRecordKeys()
		Dim arKeys, arKey, cnt, i, bHasKey
		bHasKey = False

		' Check ObjForm first
		If IsObject(ObjForm) And Not (ObjForm Is Nothing) Then
			ObjForm.Index = 0
			If ObjForm.HasValue("key_m") Then
				arKeys = ObjForm.GetValue("key_m")
				If Not IsArray(arKeys) Then
					arKeys = Array(arKeys)
				End If
				bHasKey = True
			End If
		End If

		' Check Form/QueryString
		If Not bHasKey Then
			If Request.Form("key_m").Count > 0 Then
				cnt = Request.Form("key_m").Count
				ReDim arKeys(cnt-1)
				For i = 1 to cnt ' Set up keys
					arKeys(i-1) = Request.Form("key_m")(i)
				Next
			ElseIf Request.QueryString("key_m").Count > 0 Then
				cnt = Request.QueryString("key_m").Count
				ReDim arKeys(cnt-1)
				For i = 1 to cnt ' Set up keys
					arKeys(i-1) = Request.QueryString("key_m")(i)
				Next
			ElseIf Request.QueryString <> "" Then
				ReDim arKeys(0)
				arKeys(0) = Request.QueryString("Id") ' Id

				'GetRecordKeys = arKeys ' do not return yet, so the values will also be checked by the following code
			End If
		End If

		' Check keys
		Dim ar, key
		If IsArray(arKeys) Then
			For i = 0 to UBound(arKeys)
				key = arKeys(i)
						Dim skip
						skip = False
						If Not IsNumeric(key) Then skip = True
						If Not skip Then
							If IsArray(ar) Then
								ReDim Preserve ar(UBound(ar)+1)
							Else
								ReDim ar(0)
							End If
							ar(UBound(ar)) = key
						End If
			Next
		End If
		GetRecordKeys = ar
	End Function

	' Get key filter
	Public Function GetKeyFilter()
		Dim arKeys, sKeyFilter, i, key
		arKeys = GetRecordKeys()
		sKeyFilter = ""
		If IsArray(arKeys) Then
			For i = 0 to UBound(arKeys)
				key = arKeys(i)
				If sKeyFilter <> "" Then sKeyFilter = sKeyFilter & " OR "
				Id.CurrentValue = key
				sKeyFilter = sKeyFilter & "(" & KeyFilter & ")"
			Next
		End If
		GetKeyFilter = sKeyFilter
	End Function

	' Function LoadRecordCount
	' - Load record count based on filter
	Public Function LoadRecordCount(sFilter)
		Dim wrkrs
		Set wrkrs = LoadRs(sFilter)
		If Not wrkrs Is Nothing Then
			LoadRecordCount = wrkrs.RecordCount
		Else
			LoadRecordCount = 0
		End If
		Set wrkrs = Nothing
	End Function

	' Function LoadRs
	' - Load Rows based on filter
	Public Function LoadRs(sFilter)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		Dim RsRows, sSql

		' Set up filter (Sql Where Clause) and get Return Sql
		CurrentFilter = sFilter
		sSql = SQL
		Err.Clear
		Set RsRows = Server.CreateObject("ADODB.Recordset")
		RsRows.CursorLocation = EW_CURSORLOCATION
		RsRows.Open sSql, Conn, 3, 1, 1 ' adOpenStatic, adLockReadOnly, adCmdText
		If Err.Number <> 0 Then
			Err.Clear
			Set LoadRs = Nothing
			RsRows.Close
			Set RsRows = Nothing
		ElseIf RsRows.Eof Then
			Set LoadRs = Nothing
			RsRows.Close
			Set RsRows = Nothing
		Else
			Set LoadRs = RsRows
		End If
	End Function

	' Load row values from recordset
	Public Sub LoadListRowValues(RsRow)
		Id.DbValue = RsRow("Id")
		DateTime.DbValue = RsRow("DateTime")
		Script.DbValue = RsRow("Script")
		User.DbValue = RsRow("User")
		Action.DbValue = RsRow("Action")
		zTable.DbValue = RsRow("Table")
		zField.DbValue = RsRow("Field")
		KeyValue.DbValue = RsRow("KeyValue")
		OldValue.DbValue = RsRow("OldValue")
		NewValue.DbValue = RsRow("NewValue")
	End Sub

	' Render list row values
	Sub RenderListRow()

		'
		'  Common render codes
		'
		' Id
		' DateTime
		' Script
		' User
		' Action
		' Table
		' Field
		' KeyValue
		' OldValue
		' NewValue
		' Call Row Rendering event

		Call Row_Rendering()

		'
		'  Render for View
		'
		' Id

		Id.ViewValue = Id.CurrentValue
		Id.ViewCustomAttributes = ""

		' DateTime
		DateTime.ViewValue = DateTime.CurrentValue
		DateTime.ViewValue = ew_FormatDateTime(DateTime.ViewValue, 7)
		DateTime.ViewCustomAttributes = ""

		' Script
		Script.ViewValue = Script.CurrentValue
		Script.ViewCustomAttributes = ""

		' User
		User.ViewValue = User.CurrentValue
		User.ViewCustomAttributes = ""

		' Action
		Action.ViewValue = Action.CurrentValue
		Action.ViewCustomAttributes = ""

		' Table
		zTable.ViewValue = zTable.CurrentValue
		zTable.ViewCustomAttributes = ""

		' Field
		zField.ViewValue = zField.CurrentValue
		zField.ViewCustomAttributes = ""

		' KeyValue
		KeyValue.ViewValue = KeyValue.CurrentValue
		KeyValue.ViewCustomAttributes = ""

		' OldValue
		OldValue.ViewValue = OldValue.CurrentValue
		OldValue.ViewCustomAttributes = ""

		' NewValue
		NewValue.ViewValue = NewValue.CurrentValue
		NewValue.ViewCustomAttributes = ""

		' Id
		Id.LinkCustomAttributes = ""
		Id.HrefValue = ""
		Id.TooltipValue = ""

		' DateTime
		DateTime.LinkCustomAttributes = ""
		DateTime.HrefValue = ""
		DateTime.TooltipValue = ""

		' Script
		Script.LinkCustomAttributes = ""
		Script.HrefValue = ""
		Script.TooltipValue = ""

		' User
		User.LinkCustomAttributes = ""
		User.HrefValue = ""
		User.TooltipValue = ""

		' Action
		Action.LinkCustomAttributes = ""
		Action.HrefValue = ""
		Action.TooltipValue = ""

		' Table
		zTable.LinkCustomAttributes = ""
		zTable.HrefValue = ""
		zTable.TooltipValue = ""

		' Field
		zField.LinkCustomAttributes = ""
		zField.HrefValue = ""
		zField.TooltipValue = ""

		' KeyValue
		KeyValue.LinkCustomAttributes = ""
		KeyValue.HrefValue = ""
		KeyValue.TooltipValue = ""

		' OldValue
		OldValue.LinkCustomAttributes = ""
		OldValue.HrefValue = ""
		OldValue.TooltipValue = ""

		' NewValue
		NewValue.LinkCustomAttributes = ""
		NewValue.HrefValue = ""
		NewValue.TooltipValue = ""

		' Call Row Rendered event
		If RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Row_Rendered()
		End If
	End Sub

	' Aggregate list row values
	Public Sub AggregateListRowValues()
	End Sub

	' Aggregate list row (for rendering)
	Sub AggregateListRow()
	End Sub

	' Export data in Xml Format
	Public Sub ExportXmlDocument(XmlDoc, HasParent, Recordset, StartRec, StopRec, ExportPageType)
		If Not IsObject(Recordset) Or Not IsObject(XmlDoc) Then
			Exit Sub
		End If
		If Not HasParent Then
			Call XmlDoc.AddRoot(TableVar)
		End If

		' Move to first record
		Dim RecCnt, RowCnt
		RecCnt = StartRec - 1
		If Not Recordset.Eof Then
			Recordset.MoveFirst()
			If StartRec > 1 Then Recordset.Move(StartRec - 1)
		End If
		Do While Not Recordset.Eof And RecCnt < StopRec
			RecCnt = RecCnt + 1
			If CLng(RecCnt) >= CLng(StartRec) Then
				RowCnt = CLng(RecCnt) - CLng(StartRec) + 1
				Call LoadListRowValues(Recordset)

				' Render row
				RowType = EW_ROWTYPE_VIEW ' Render view
				Call ResetAttrs()
				Call RenderListRow()
				If HasParent Then
					Call XmlDoc.AddRow(TableVar, "")
				Else
					Call XmlDoc.AddRow("", "")
				End If
				If ExportPageType = "view" Then
					Call XmlDoc.AddField("DateTime", DateTime.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("Script", Script.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("User", User.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("Action", Action.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("zTable", zTable.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("zField", zField.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("KeyValue", KeyValue.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("OldValue", OldValue.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("NewValue", NewValue.ExportValue(Export, ExportOriginalValue))
				Else
					Call XmlDoc.AddField("Id", Id.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("DateTime", DateTime.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("Script", Script.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("User", User.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("Action", Action.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("zTable", zTable.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("zField", zField.ExportValue(Export, ExportOriginalValue))
				End If
			End If
			Recordset.MoveNext()
		Loop
	End Sub

	' Export data in HTML/CSV/Word/Excel/Email format
	Public Sub ExportDocument(Doc, Recordset, StartRec, StopRec, ExportPageType)
		If Not IsObject(Recordset) Or Not IsObject(Doc) Then
			Exit Sub
		End If

		' Write header
		Call Doc.ExportTableHeader()
		If Doc.Horizontal Then ' Horizontal format, write header
			Call Doc.BeginExportRow(0)
			If ExportPageType = "view" Then
				Call Doc.ExportCaption(DateTime)
				Call Doc.ExportCaption(Script)
				Call Doc.ExportCaption(User)
				Call Doc.ExportCaption(Action)
				Call Doc.ExportCaption(zTable)
				Call Doc.ExportCaption(zField)
				Call Doc.ExportCaption(KeyValue)
				Call Doc.ExportCaption(OldValue)
				Call Doc.ExportCaption(NewValue)
			Else
				Call Doc.ExportCaption(Id)
				Call Doc.ExportCaption(DateTime)
				Call Doc.ExportCaption(Script)
				Call Doc.ExportCaption(User)
				Call Doc.ExportCaption(Action)
				Call Doc.ExportCaption(zTable)
				Call Doc.ExportCaption(zField)
			End If
			Call Doc.EndExportRow()
		End If

		' Move to first record
		Dim RecCnt, RowCnt
		RecCnt = StartRec - 1
		If Not Recordset.Eof Then
			Recordset.MoveFirst()
			If StartRec > 1 Then Recordset.Move(StartRec - 1)
		End If
		Do While Not Recordset.Eof And CLng(RecCnt) < CLng(StopRec)
			RecCnt = RecCnt + 1
			If CLng(RecCnt) >= CLng(StartRec) Then
				RowCnt = CLng(RecCnt) - CLng(StartRec) + 1
				Call LoadListRowValues(Recordset)

				' Render row
				RowType = EW_ROWTYPE_VIEW ' Render view
				Call ResetAttrs()
				Call RenderListRow()
				Call Doc.BeginExportRow(RowCnt)
				If ExportPageType = "view" Then
					Call Doc.ExportField(DateTime)
					Call Doc.ExportField(Script)
					Call Doc.ExportField(User)
					Call Doc.ExportField(Action)
					Call Doc.ExportField(zTable)
					Call Doc.ExportField(zField)
					Call Doc.ExportField(KeyValue)
					Call Doc.ExportField(OldValue)
					Call Doc.ExportField(NewValue)
				Else
					Call Doc.ExportField(Id)
					Call Doc.ExportField(DateTime)
					Call Doc.ExportField(Script)
					Call Doc.ExportField(User)
					Call Doc.ExportField(Action)
					Call Doc.ExportField(zTable)
					Call Doc.ExportField(zField)
				End If
				Call Doc.EndExportRow()
			End If
			Recordset.MoveNext()
		Loop
		Call Doc.ExportTableFooter()
	End Sub
	Dim CurrentAction ' Current action
	Dim LastAction ' Last action
	Dim CurrentMode ' Current mode
	Dim UpdateConflict ' Update conflict
	Dim EventName ' Event name
	Dim EventCancelled ' Event cancelled
	Dim CancelMessage ' Cancel message
	Dim AllowAddDeleteRow ' Allow add/delete row
	Dim DetailAdd ' Allow detail add
	Dim DetailEdit ' Allow detail edit
	Dim GridAddRowCount ' Grid add row count

	' Check current action
	' - Add
	Public Function IsAdd()
		IsAdd = (CurrentAction = "add")
	End Function

	' - Copy
	Public Function IsCopy()
		IsCopy = (CurrentAction = "copy" Or CurrentAction = "C")
	End Function

	' - Edit
	Public Function IsEdit()
		IsEdit = (CurrentAction = "edit")
	End Function

	' - Delete
	Public Function IsDelete()
		IsDelete = (CurrentAction = "D")
	End Function

	' - Confirm
	Public Function IsConfirm()
		IsConfirm = (CurrentAction = "F")
	End Function

	' - Overwrite
	Public Function IsOverwrite()
		IsOverwrite = (CurrentAction = "overwrite")
	End Function

	' - Cancel
	Public Function IsCancel()
		IsCancel = (CurrentAction = "cancel")
	End Function

	' - Grid add
	Public Function IsGridAdd()
		IsGridAdd = (CurrentAction = "gridadd")
	End Function

	' - Grid edit
	Public Function IsGridEdit()
		IsGridEdit = (CurrentAction = "gridedit")
	End Function

	' - Insert
	Public Function IsInsert()
		IsInsert = (CurrentAction = "insert" Or CurrentAction = "A")
	End Function

	' - Update
	Public Function IsUpdate()
		IsUpdate = (CurrentAction = "update" Or CurrentAction = "U")
	End Function

	' - Grid update
	Public Function IsGridUpdate()
		IsGridUpdate = (CurrentAction = "gridupdate")
	End Function

	' - Grid insert
	Public Function IsGridInsert()
		IsGridInsert = (CurrentAction = "gridinsert")
	End Function

	' - Grid overwrite
	Public Function IsGridOverwrite()
		IsGridOverwrite = (CurrentAction = "gridoverwrite")
	End Function

	' Check last action
	' - Cancelled
	Public Function IsCancelled()
		IsCancelled = (LastAction = "cancel" And CurrentAction = "")
	End Function

	' - Inline inserted
	Public Function IsInlineInserted()
		IsInlineInserted = (LastAction = "insert" And CurrentAction = "")
	End Function

	' - Inline updated
	Public Function IsInlineUpdated()
		IsInlineUpdated = (LastAction = "update" And CurrentAction = "")
	End Function

	' - Grid updated
	Public Function IsGridUpdated()
		IsGridUpdated = (LastAction = "gridupdate" And CurrentAction = "")
	End Function

	' - Grid inserted
	Public Function IsGridInserted()
		IsGridInserted = (LastAction = "gridinsert" And CurrentAction = "")
	End Function

	' Row Type
	Private m_RowType

	Public Property Get RowType()
		RowType = m_RowType
	End Property

	Public Property Let RowType(v)
		m_RowType = v
	End Property
	Dim CssClass ' Css class
	Dim CssStyle' Css style

'	Dim RowClientEvents ' Row client events
	Dim RowAttrs ' Row attributes

	' Row Styles
	Public Property Get RowStyles()
		Dim sAtt, Value
		Dim sStyle, sClass
		sAtt = ""
		sStyle = CssStyle
		If RowAttrs.Exists("style") Then
			Value = RowAttrs.Item("style")
			If Trim(Value) <> "" Then
				sStyle = sStyle & " " & Value
			End If
		End If
		sClass = CssClass
		If RowAttrs.Exists("class") Then
			Value = RowAttrs.Item("class")
			If Trim(Value) <> "" Then
				sClass = sClass & " " & Value
			End If
		End If
		If Trim(sStyle) <> "" Then
			sAtt = sAtt & " style=""" & Trim(sStyle) & """" 
		End If
		If Trim(sClass) <> "" Then
			sAtt = sAtt & " class=""" & Trim(sClass) & """" 
		End If
		RowStyles = sAtt
	End Property

	' Row Attribute
	Public Property Get RowAttributes()
		Dim sAtt, Attr, Value, i
		sAtt = RowStyles
		If m_Export = "" Then

'			If Trim(RowClientEvents) <> "" Then
'				sAtt = sAtt & " " & Trim(RowClientEvents)
'			End If

			For i = 0 to UBound(RowAttrs.Attributes)
				Attr = RowAttrs.Attributes(i)(0)
				Value = RowAttrs.Attributes(i)(1)
				If Attr <> "style" And Attr <> "class" And Attr <> "" And Value <> "" Then
					sAtt = sAtt & " " & Attr & "=""" & Value & """"
				End If
			Next
		End If
		RowAttributes = sAtt
	End Property

	' Export
	Private m_Export

	Public Property Get Export()
		Export = m_Export
	End Property

	Public Property Let Export(v)
		m_Export = v
	End Property

	' Export Original Value
	Dim ExportOriginalValue

	' Export All
	Dim ExportAll

	' Send Email
	Dim SendEmail

	' Custom Inner Html
	Dim TableCustomInnerHtml

	' ----------------
	'  Field objects
	' ----------------
	' Field Id
	Private m_Id

	Public Property Get Id()
		If Not IsObject(m_Id) Then
			Set m_Id = NewFldObj("AuditTrail", "AuditTrail", "x_Id", "Id", "[Id]", 3, 7, "", False, False, "FORMATTED TEXT")
			m_Id.FldDefaultErrMsg = Language.Phrase("IncorrectInteger")
		End If
		Set Id = m_Id
	End Property

	' Field DateTime
	Private m_DateTime

	Public Property Get DateTime()
		If Not IsObject(m_DateTime) Then
			Set m_DateTime = NewFldObj("AuditTrail", "AuditTrail", "x_DateTime", "DateTime", "[DateTime]", 135, 7, "", False, False, "FORMATTED TEXT")
			m_DateTime.FldDefaultErrMsg = Replace(Language.Phrase("IncorrectDateDMY"), "%s", "/")
		End If
		Set DateTime = m_DateTime
	End Property

	' Field Script
	Private m_Script

	Public Property Get Script()
		If Not IsObject(m_Script) Then
			Set m_Script = NewFldObj("AuditTrail", "AuditTrail", "x_Script", "Script", "[Script]", 202, 7, "", False, False, "FORMATTED TEXT")
		End If
		Set Script = m_Script
	End Property

	' Field User
	Private m_User

	Public Property Get User()
		If Not IsObject(m_User) Then
			Set m_User = NewFldObj("AuditTrail", "AuditTrail", "x_User", "User", "[User]", 202, 7, "", False, False, "FORMATTED TEXT")
		End If
		Set User = m_User
	End Property

	' Field Action
	Private m_Action

	Public Property Get Action()
		If Not IsObject(m_Action) Then
			Set m_Action = NewFldObj("AuditTrail", "AuditTrail", "x_Action", "Action", "[Action]", 202, 7, "", False, False, "FORMATTED TEXT")
		End If
		Set Action = m_Action
	End Property

	' Field Table
	Private m_zTable

	Public Property Get zTable()
		If Not IsObject(m_zTable) Then
			Set m_zTable = NewFldObj("AuditTrail", "AuditTrail", "x_zTable", "Table", "[Table]", 202, 7, "", False, False, "FORMATTED TEXT")
		End If
		Set zTable = m_zTable
	End Property

	' Field Field
	Private m_zField

	Public Property Get zField()
		If Not IsObject(m_zField) Then
			Set m_zField = NewFldObj("AuditTrail", "AuditTrail", "x_zField", "Field", "[Field]", 202, 7, "", False, False, "FORMATTED TEXT")
		End If
		Set zField = m_zField
	End Property

	' Field KeyValue
	Private m_KeyValue

	Public Property Get KeyValue()
		If Not IsObject(m_KeyValue) Then
			Set m_KeyValue = NewFldObj("AuditTrail", "AuditTrail", "x_KeyValue", "KeyValue", "[KeyValue]", 203, 7, "", False, False, "FORMATTED TEXT")
		End If
		Set KeyValue = m_KeyValue
	End Property

	' Field OldValue
	Private m_OldValue

	Public Property Get OldValue()
		If Not IsObject(m_OldValue) Then
			Set m_OldValue = NewFldObj("AuditTrail", "AuditTrail", "x_OldValue", "OldValue", "[OldValue]", 203, 7, "", False, False, "FORMATTED TEXT")
		End If
		Set OldValue = m_OldValue
	End Property

	' Field NewValue
	Private m_NewValue

	Public Property Get NewValue()
		If Not IsObject(m_NewValue) Then
			Set m_NewValue = NewFldObj("AuditTrail", "AuditTrail", "x_NewValue", "NewValue", "[NewValue]", 203, 7, "", False, False, "FORMATTED TEXT")
		End If
		Set NewValue = m_NewValue
	End Property
	Dim Fields ' Fields

	' Create new field object
	Private Function NewFldObj(TblVar, TblName, FldVar, FldName, FldExpression, FldType, FldDtFormat, FldVirtualExp, FldVirtual, FldForceSelect, FldViewTag)
		Dim fld
		Set fld = New cField
		fld.TblVar = TblVar
		fld.TblName = TblName
		fld.FldVar = FldVar
		fld.FldName = FldName
		fld.FldExpression = FldExpression
		fld.FldType = FldType
		fld.FldDataType = ew_FieldDataType(FldType)
		fld.FldDateTimeFormat = FldDtFormat
		fld.FldVirtualExpression = FldVirtualExp
		fld.FldIsVirtual = FldVirtual
		fld.FldForceSelection = FldForceSelect
		fld.FldViewTag = FldViewTag
		Set NewFldObj = fld
	End Function

	' Table level events
	' Recordset Selecting event
	Sub Recordset_Selecting(filter)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next

		' Enter your code here	
	End Sub

	' Recordset Selected event
	Sub Recordset_Selected(rs)

		'Response.Write "Recordset Selected"
	End Sub

	' Recordset Search Validated event
	Sub Recordset_SearchValidated()
		If Not EW_DEBUG_ENABLED Then On Error Resume Next

		' Enter your code here
	End Sub

	' Recordset Searching event
	Sub Recordset_Searching(filter)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next

		' Enter your code here	
	End Sub

	' Row_Selecting event
	Sub Row_Selecting(filter)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next

		' Enter your code here	
	End Sub

	' Row Selected event
	Sub Row_Selected(rs)

		'Response.Write "Row Selected"
	End Sub

	' Row Inserting event
	Function Row_Inserting(rsold, rsnew)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next

		' Enter your code here
		' To cancel, set return value to False

		Row_Inserting = True
	End Function

	' Row Inserted event
	Sub Row_Inserted(rsold, rsnew)

		' Response.Write "Row Inserted"
	End Sub

	' Row Updating event
	Function Row_Updating(rsold, rsnew)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next

		' Enter your code here
		' To cancel, set return value to False

		Row_Updating = True
	End Function

	' Row Updated event
	Sub Row_Updated(rsold, rsnew)

		' Response.Write "Row Updated"
	End Sub

	' Row Update Conflict event
	Function Row_UpdateConflict(rsold, rsnew)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next

		' Enter your code here
		' To ignore conflict, set return value to False

		Row_UpdateConflict = True
	End Function

	' Row Deleting event
	Function Row_Deleting(rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next

		' Enter your code here
		' To cancel, set return value to False

		Row_Deleting = True
	End Function

	' Row Deleted event
	Sub Row_Deleted(rs)

		' Response.Write "Row Deleted"
	End Sub

	' Email Sending event
	Function Email_Sending(Email, Args)

		'Response.Write Email.AsString
		'Response.Write "Keys of Args: " & Join(Args.Keys, ", ")
		'Response.End

		Email_Sending = True
	End Function

	' Row Rendering event
	Sub Row_Rendering()
		If Not EW_DEBUG_ENABLED Then On Error Resume Next

		' Enter your code here	
	End Sub

	' Row Rendered event
	Sub Row_Rendered()

		' To view properties of field class, use:
		' Response.Write <FieldName>.AsString() 

	End Sub

	' Class terminate
	Private Sub Class_Terminate
		If IsObject(m_Id) Then Set m_Id = Nothing
		If IsObject(m_DateTime) Then Set m_DateTime = Nothing
		If IsObject(m_Script) Then Set m_Script = Nothing
		If IsObject(m_User) Then Set m_User = Nothing
		If IsObject(m_Action) Then Set m_Action = Nothing
		If IsObject(m_zTable) Then Set m_zTable = Nothing
		If IsObject(m_zField) Then Set m_zField = Nothing
		If IsObject(m_KeyValue) Then Set m_KeyValue = Nothing
		If IsObject(m_OldValue) Then Set m_OldValue = Nothing
		If IsObject(m_NewValue) Then Set m_NewValue = Nothing
		Set RowAttrs = Nothing
	End Sub
End Class
%>
