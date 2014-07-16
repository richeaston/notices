<%

' -----------------------------------------
'  ASPMaker 9 Common classes and functions
'
' -----------------------
Class cAttributes
	Dim Attributes

	' Class Initialize
	Private Sub Class_Initialize()
		Clear()
	End Sub

	Public Sub Clear()
		Attributes = Array() ' Clear attributes
	End Sub

	Public Function Exists(Attr)
		Dim i, att
		att = Trim(Attr)
		If att <> "" Then
			For i = 0 to UBound(Attributes)
				If LCase(Attributes(i)(0)) = LCase(att) Then
					Exists = True
					Exit Function
				End If
			Next
		End If
		Exists = False
	End Function

	Public Function Item(Attr)
		Dim i, att
		att = Trim(Attr)
		If att <> "" Then
			For i = 0 to UBound(Attributes)
				If LCase(Attributes(i)(0)) = LCase(att) Then
					Item = Attributes(i)(1)
					Exit Function
				End If
			Next
		End If
		Item = ""
	End Function

	' Add attributes
	Public Sub AddAttributes(Attrs)
		Dim i
		If IsArray(Attrs) Then
			For i = 0 to UBound(Attrs)
				If IsArray(Attrs(i)) Then
					If UBound(Attrs(i)) >= 1 Then
						AddAttribute Attrs(i)(0), Attrs(i)(1), True
					End If
				End If
			Next
		End If
	End Sub

	' Add attribute
	Public Sub AddAttribute(Attr, Value, Append)
		Dim opt
		opt = ew_IIf(Append, "a", "p") ' Append / Prepend
		If Not UpdateAttr(Attr, Value, opt) Then
			AddAttr Attr, Value
		End If
	End Sub

	' Update attribute
	Public Sub UpdateAttribute(Attr, Value)
		If Not UpdateAttr(Attr, Value, "u")  Then ' Update the attribute
			AddAttr Attr, Value
		End If
	End Sub

	' Append attribute
	Public Sub Append(Attr, Value)
		AddAttribute Attr, Value, True
	End Sub

	' Prepend attribute
	Public Sub Prepend(Attr, Value)
		AddAttribute Attr, Value, False
	End Sub

	' Update attribute based on option
	Private Function UpdateAttr(Attr, Value, Opt)
		Dim i, att, val
		att = Trim(Attr)
		val = Trim(Value)
		If att <> "" And val <> "" Then
			For i = 0 to UBound(Attributes)
				If LCase(Attributes(i)(0)) = LCase(att) Then
					If Opt = "a" Then ' Append
						Attributes(i)(1) = Attributes(i)(1) & " " & val
					ElseIf Opt = "p" Then ' Prepend
						Attributes(i)(1) = val & " " & Attributes(i)(1)
					Else ' Assume update
						Attributes(i)(1) = val
					End If
					UpdateAttr = True
					Exit Function
				End If
			Next
		End If
		UpdateAttr = False
	End Function

	' Add attribute to array
	Private Function AddAttr(Attr, Value)
		Dim att, val
		att = Trim(Attr)
		val = Trim(Value)
		If att <> "" And val <> "" Then
			If Ubound(Attributes) < 0 Then
				ReDim Attributes(0)
			Else
				ReDim Preserve Attributes(UBound(Attributes)+1)
			End If
			Attributes(UBound(Attributes)) = Array(att, val)
		End If
	End Function
End Class

' -----------------------
'  Export document class
'
Class cExportDocument
	Dim Table
	Dim Text
	Dim Line
	Dim Header
	Dim Style ' "v"(Vertical) or "h"(Horizontal)
	Dim Horizontal ' Horizontal
	Dim FldCnt

	' Class Initialize
	Private Sub Class_Initialize()
		Text = ""
		Line = ""
		Header = ""
		Style = "h"
		Horizontal = True
	End Sub

	Public Sub ChangeStyle(ToStyle)
		If LCase(ToStyle) = "v" Or LCase(ToStyle) = "h" Then
			Style = LCase(ToStyle)
		End If
		Horizontal = (Table.Export <> "xml" And (Style <> "v" Or Table.Export = "csv"))
	End Sub

	' Table Header
	Public Sub ExportTableHeader()
		Select Case Table.Export
			Case "html", "email", "word", "excel"
				Text = Text & "<table class=""ewExportTable"">"
			Case "csv"

				' No action
		End Select
	End Sub

	' Field Caption
	Public Sub ExportCaption(fld)
		Call ExportValueEx(fld, fld.ExportCaption)
	End Sub

	' Field value
	Public Sub ExportValue(fld)
		Call ExportValueEx(fld, fld.ExportValue(Table.Export, Table.ExportOriginalValue))
	End Sub

	' Field aggregate
	Public Sub ExportAggregate(fld, typ)
		If Horizontal Then
			Dim val
			val = ""
			If typ = "TOTAL" Or typ = "COUNT" Or typ = "AVERAGE" Then
				val = Language.Phrase(typ) & ": " & fld.ExportValue(Table.Export, Table.ExportOriginalValue)
			End If
			Call ExportValueEx(fld, val)
		End If
	End Sub

	' Export a value (caption, field value, or aggregate)
	Public Sub ExportValueEx(fld, val)
		Select Case Table.Export
			Case "html", "email", "word", "excel"
				Text = Text & "<td" & ew_IIf(EW_EXPORT_CSS_STYLES, fld.CellStyles, "") & ">"
				If Table.Export = "excel" And fld.FldDataType = EW_DATATYPE_STRING And IsNumeric(val) Then
					Text = Text & "=""" & val & """"
				Else
					Text = Text & val
				End If
				Text = Text & "</td>"
			Case "csv"
				If Line <> "" Then Line = Line & ","
				Line = Line & """" & Replace(val & "", """", """""") & """"
		End Select
	End Sub

	' Begin a row
	Public Sub BeginExportRow(rowcnt)
		FldCnt = 0
		If Horizontal Then
			Select Case Table.Export
				Case "html", "email", "word", "excel"
					If rowcnt = -1 Then
						Table.CssClass = "ewExportTableFooter"
					ElseIf rowcnt = 0 Then
						Table.CssClass = "ewExportTableHeader"
					Else
						Table.CssClass = ew_IIf(rowcnt Mod 2 = 1, "ewExportTableRow", "ewExportTableAltRow")
					End If
					Text = Text & "<tr" & ew_IIf(EW_EXPORT_CSS_STYLES, Table.RowStyles, "") & ">"
				Case "csv"
					Line = ""
			End Select
		End If
	End Sub

	' End a row
	Public Sub EndExportRow()
		If Horizontal Then
			Select Case Table.Export
				Case "html", "email", "word", "excel"
					Text = Text & "</tr>"
				Case "csv"
					Line = Line & vbCrLf
					Text = Text & Line
			End Select
		End If
	End Sub

	' Empty line
	Public Sub ExportEmptyLine()
		Select Case Table.Export
			Case "html", "email", "word", "excel"
				Text = Text & "<br>&nbsp;"
		End Select
	End Sub

	' Export a field
	Public Sub ExportField(fld)
		If Horizontal Then
			Call ExportValue(fld)
		Else ' Vertical, export as a row
			FldCnt = FldCnt + 1
			Dim tdcaption, tdvalue
			tdcaption = "<td"
			Select Case Table.Export
				Case "html", "email", "word", "excel"
					tdcaption = tdcaption & " class=""ewTableHeader"""
			End Select
			tdcaption = tdcaption & ">"
			fld.CellCssClass = ew_IIf(FldCnt Mod 2 = 1, "ewExportTableRow", "ewTableAltRow")
			tdvalue = "<td" & ew_IIf(EW_EXPORT_CSS_STYLES, fld.CellStyles, "") & ">"
			Text = Text & "<tr>" & tdcaption & fld.ExportCaption & "</td>" & tdvalue & fld.ExportValue(Table.Export, Table.ExportOriginalValue) & "</td></tr>"
		End If
	End Sub

	' Table Footer
	Public Sub ExportTableFooter()
		Select Case Table.Export
			Case "html", "email", "word", "excel"
				Text = Text & "</table>"
		End Select
	End Sub

	Public Sub ExportHeaderAndFooter()
		Dim Charset, Header
		Charset = ew_IIf(EW_CHARSET <> "", ";charset=" & EW_CHARSET, "") ' Charset used in header
		Select Case Table.Export
			Case "html", "email", "word", "excel"
				Header = "<html><head>" & vbCrLf
				Header = Header & "<meta http-equiv=""Content-Type"" content=""text/html" & Charset & """>" & vbCrLf
				If EW_EXPORT_CSS_STYLES Then
					Header = Header & "<style>" & ew_LoadFile(EW_PROJECT_STYLESHEET_FILENAME) & "</style>" & vbCrLf
				End If
				Header = Header & "</" & "head>" & vbCrLf & "<body>" & vbCrLf
				Text = Header & Text & "</body></html>"
		End Select
	End Sub
End Class

' --------------------
'  XML document class
'
Class cXMLDocument
	Dim Encoding
	Dim RootTagName
	Dim SubTblName
	Dim RowTagName
	Dim XmlDoc
	Dim XmlTbl
	Dim XmlSubTbl
	Dim XmlRow
	Dim XmlFld

	' Class Initialize
	Private Sub Class_Initialize()
		Encoding = ""
		RootTagName = "table"
		RowTagName = "row"
		Set XmlDoc = ew_CreateXmlDom()
	End Sub

	Public Sub AddRoot(rootname)
		If rootname <> "" Then
			RootTagName = ew_XmlTagName(rootname)
		End If
		Set XmlTbl = XmlDoc.CreateElement(RootTagName)
		XmlDoc.AppendChild(XmlTbl)
	End Sub

	' Add row
	Public Sub AddRow(tablename, rowname)
		If rowname <> "" Then
			RowTagName = ew_XmlTagName(rowname)
		End If
		Set XmlRow = XmlDoc.CreateElement(RowTagName)
		If tablename = "" Then
			If Not IsEmpty(XmlTbl) Then
				XmlTbl.AppendChild(XmlRow)
			End If
		Else
			If SubTblName = "" Then
				SubTblName = ew_XmlTagName(tablename)
				Set XmlSubTbl = XmlDoc.CreateElement(SubTblName)
				XmlTbl.AppendChild(XmlSubTbl)
			End If
			If Not IsEmpty(XmlSubTbl) Then
				XmlSubTbl.AppendChild(XmlRow)
			End If
		End If
	End Sub

	' Add field
	Public Sub AddField(Name, Value)
		Set XmlFld = XmlDoc.CreateElement(ew_XmlTagName(Name))
		Call XmlRow.AppendChild(XmlFld)
		Call XmlFld.AppendChild(XmlDoc.CreateTextNode(Value & ""))
	End Sub

	' XML
	Public Function XML()
		XML = XmlDoc.XML
	End Function

	' Output
	Public Sub Output()
		Dim PI
		If Response.Buffer Then Response.Clear
		Response.ContentType = "text/xml"
		PI = "<?xml version=""1.0"""
		If Encoding <> "" Then PI = PI & " encoding=""" & Encoding & """"
		PI = PI & " ?>"
		Response.Write PI & XmlDoc.XML
	End Sub

	' Output XML for debug
	Public Sub Print()
		If Response.Buffer Then Response.Clear
		Response.ContentType = "text/plain"
		Response.Write Server.HTMLEncode(XmlDoc.XML)
	End Sub

	' Terminate
	Private Sub Class_Terminate()
		Set XmlFld = Nothing
		Set XmlRow = Nothing
		Set XmlTbl = Nothing
		Set XmlDoc = Nothing
	End Sub
End Class 

'
'  XML document class (end)
' --------------------------
'
' ---------------------
'  Email class (begin)
'
Class cEmail

	' Class properties
	Dim Sender ' Sender
	Dim Recipient ' Recipient
	Dim Cc ' Cc
	Dim Bcc ' Bcc
	Dim Subject ' Subject
	Dim Format ' Format
	Dim Content ' Content
	Dim Charset ' Charset
	Dim SendErrNumber ' Send error number
	Dim SendErrDescription ' Send error description

	' Method to load email from template
	Public Sub Load(fn)
		Dim sWrk, sHeader, arrHeader
		Dim sName, sValue
		Dim i, j
		sWrk = ew_LoadTxt(fn) ' Load text file content
		sWrk = Replace(sWrk, vbCrLf, vbLf) ' Convert to Lf
		sWrk = Replace(sWrk, vbCr, vbLf) ' Convert to Lf
		If sWrk <> "" Then

			' Locate Header & Mail Content
			i = InStr(sWrk, vbLf&vbLf)
			If i > 0 Then
				sHeader = Mid(sWrk, 1, i)
				Content = Mid(sWrk, i+2)
				arrHeader = Split(sHeader, vbLf)
				For j = 0 to UBound(arrHeader)
					i = InStr(arrHeader(j), ":")
					If i > 0 Then
						sName = Trim(Mid(arrHeader(j), 1, i-1))
						sValue = Trim(Mid(arrHeader(j), i+1))
						Select Case LCase(sName)
							Case "subject"
								Subject = sValue
							Case "from"
								Sender = sValue
							Case "to"
								Recipient = sValue
							Case "cc"
								Cc = sValue
							Case "bcc"
								Bcc = sValue
							Case "format"
								Format = sValue
						End Select
					End If
				Next
			End If
		End If
	End Sub

	' Method to replace sender
	Public Sub ReplaceSender(ASender)
		Sender = Replace(Sender, "<!--$From-->", ASender)
	End Sub

	' Method to replace recipient
	Public Sub ReplaceRecipient(ARecipient)
		Recipient = Replace(Recipient, "<!--$To-->", ARecipient)
	End Sub

	' Method to add Cc email
	Public Sub AddCc(ACc)
		If ACc <> "" Then
			If Cc <> "" Then Cc = Cc & ";"
			Cc = Cc & ACc
		End If
	End Sub

	' Method to add Bcc email
	Public Sub AddBcc(ABcc)
		If ABcc <> "" Then
			If Bcc <> "" Then Bcc = Bcc & ";"
			Bcc = Bcc & ABcc
		End If
	End Sub

	' Method to replace subject
	Public Sub ReplaceSubject(ASubject)
		Subject = Replace(Subject, "<!--$Subject-->", ASubject)
	End Sub

	' Method to replace content
	Public Sub ReplaceContent(Find, ReplaceWith)
		Content = Replace(Content, Find, ReplaceWith)
	End Sub

	' Method to send email
	Public Function Send
		Send = ew_SendEmail(Sender, Recipient, Cc, Bcc, Subject, Content, Format, Charset)
		If Not Send Then
			SendErrNumber = Hex(gsEmailErrNo) ' Send error number
			SendErrDescription = gsEmailErrDesc ' Send error description
		Else
			SendErrNumber = 0
			SendErrDescription = ""
		End If
	End Function

	' Show object as string
	Public Function AsString()
		AsString = "{" & _
			"Sender: " & Sender & ", " & _
			"Recipient: " & Recipient & ", " & _
			"Cc: " & Cc & ", " & _
			"Bcc: " & Bcc & ", " & _
			"Subject: " & Subject & ", " & _
			"Format: " & Format & ", " & _
			"Content: " & Content & ", " & _
			"Charset: " & Charset & _
			"}"
	End Function
End Class

'
'  Email class (end)
' -------------------
'
' -------------------------------------
'  Pager classes and functions (begin)
'
' Function to create numeric pager
Function ew_NewNumericPager(FromIndex, PageSize, RecordCount, Range)
	Set ew_NewNumericPager = New cNumericPager
	ew_NewNumericPager.FromIndex = CLng(FromIndex)
	ew_NewNumericPager.PageSize = CLng(PageSize)
	ew_NewNumericPager.RecordCount = CLng(RecordCount)
	ew_NewNumericPager.Range = CLng(Range)
	ew_NewNumericPager.Init
End Function

' Function to create next prev pager
Function ew_NewPrevNextPager(FromIndex, PageSize, RecordCount)
	Set ew_NewPrevNextPager = New cPrevNextPager
	ew_NewPrevNextPager.FromIndex = CLng(FromIndex)
	ew_NewPrevNextPager.PageSize = CLng(PageSize)
	ew_NewPrevNextPager.RecordCount = CLng(RecordCount)
	ew_NewPrevNextPager.Init
End Function

' Class for Pager item
Class cPagerItem
	Dim Start, Text, Enabled
End Class

' Class for Numeric pager
Class cNumericPager
	Dim Items()
	Dim Count, FromIndex, ToIndex, RecordCount, PageSize, Range
	Dim FirstButton, PrevButton, NextButton, LastButton, ButtonCount
	Dim Visible

	' Class Initialize
	Private Sub Class_Initialize()
		Set FirstButton = New cPagerItem
		Set PrevButton = New cPagerItem
		Set NextButton = New cPagerItem
		Set LastButton = New cPagerItem
		Visible = True
	End Sub

	' Method to init pager
	Public Sub Init()
		If FromIndex > RecordCount Then FromIndex = RecordCount
		ToIndex = FromIndex + PageSize - 1
		If ToIndex > RecordCount Then ToIndex = RecordCount
		Count = -1
		ReDim Items(0)
		SetupNumericPager()
		Redim Preserve Items(Count)

		' Update button count
		ButtonCount = Count + 1
		If FirstButton.Enabled Then ButtonCount = ButtonCount + 1
		If PrevButton.Enabled Then ButtonCount = ButtonCount + 1
		If NextButton.Enabled Then ButtonCount = ButtonCount + 1
		If LastButton.Enabled Then ButtonCount = ButtonCount + 1
	End Sub

	' Add pager item
	Private Sub AddPagerItem(StartIndex, Text, Enabled)
		Count = Count + 1
		If Count > UBound(Items) Then
			Redim Preserve Items(UBound(Items)+10)
		End If
		Dim Item
		Set Item = New cPagerItem
		Item.Start = StartIndex
		Item.Text = Text
		Item.Enabled = Enabled
		Set Items(Count) = Item
	End Sub

	' Setup pager items
	Private Sub SetupNumericPager()
		Dim Eof, x, y, dx1, dx2, dy1, dy2, ny, HasPrev, TempIndex
		If RecordCount > PageSize Then
			Eof = (RecordCount < (FromIndex + PageSize))
			HasPrev = (FromIndex > 1)

			' First Button
			TempIndex = 1
			FirstButton.Start = TempIndex
			FirstButton.Enabled = (FromIndex > TempIndex)

			' Prev Button
			TempIndex = FromIndex - PageSize
			If TempIndex < 1 Then TempIndex = 1
			PrevButton.Start = TempIndex
			PrevButton.Enabled = HasPrev

			' Page links
			If HasPrev Or Not Eof Then
				x = 1
				y = 1
				dx1 = ((FromIndex-1)\(PageSize*Range))*PageSize*Range + 1
				dy1 = ((FromIndex-1)\(PageSize*Range))*Range + 1
				If (dx1+PageSize*Range-1) > RecordCount Then
					dx2 = (RecordCount\PageSize)*PageSize + 1
					dy2 = (RecordCount\PageSize) + 1
				Else
					dx2 = dx1 + PageSize*Range - 1
					dy2 = dy1 + Range - 1
				End If
				While x <= RecordCount
					If x >= dx1 And x <= dx2 Then
						Call AddPagerItem(x, y, FromIndex<>x)
						x = x + PageSize
						y = y + 1
					ElseIf x >= (dx1-PageSize*Range) And x <= (dx2+PageSize*Range) Then
						If x+Range*PageSize < RecordCount Then
							Call AddPagerItem(x, y & "-" & (y+Range-1), True)
						Else
							ny = (RecordCount-1)\PageSize + 1
							If ny = y Then
								Call AddPagerItem(x, y, True)
							Else
								Call AddPagerItem(x, y & "-" & ny, True)
							End If
						End If
						x = x + Range*PageSize
						y = y + Range
					Else
						x = x + Range*PageSize
						y = y + Range
					End If
				Wend
			End If

			' Next Button
			NextButton.Start = FromIndex + PageSize
			TempIndex = FromIndex + PageSize
			NextButton.Start = TempIndex
			NextButton.Enabled = Not Eof

			' Last Button
			TempIndex = ((RecordCount-1)\PageSize)*PageSize + 1
			LastButton.Start = TempIndex
			LastButton.Enabled = (FromIndex < TempIndex)
		End If
	End Sub

    ' Terminate
	Private Sub Class_Terminate()
		Set FirstButton = Nothing
		Set PrevButton = Nothing
		Set NextButton = Nothing
		Set LastButton = Nothing
		For Each Item In Items
			Set Item = Nothing
		Next
		Erase Items
	End Sub
End Class

' Class for PrevNext pager
Class cPrevNextPager
	Dim FirstButton, PrevButton, NextButton, LastButton
	Dim CurrentPage, PageSize, PageCount, FromIndex, ToIndex, RecordCount
	Dim Visible

	' Class Initialize
	Private Sub Class_Initialize()
		Set FirstButton = New cPagerItem
		Set PrevButton = New cPagerItem
		Set NextButton = New cPagerItem
		Set LastButton = New cPagerItem
		Visible = True
	End Sub

	' Method to init pager
	Public Sub Init()
		Dim TempIndex
		If PageSize > 0 Then
			CurrentPage = (FromIndex-1)\PageSize + 1
			PageCount = (RecordCount-1)\PageSize + 1
			If FromIndex > RecordCount Then FromIndex = RecordCount
			ToIndex = FromIndex + PageSize - 1
			If ToIndex > RecordCount Then ToIndex = RecordCount

			' First Button
			TempIndex = 1
			FirstButton.Start = TempIndex
			FirstButton.Enabled = (TempIndex <> FromIndex)

			' Prev Button
			TempIndex = FromIndex - PageSize
			If TempIndex < 1 Then TempIndex = 1
			PrevButton.Start = TempIndex
			PrevButton.Enabled = (TempIndex <> FromIndex)

			' Next Button
			TempIndex = FromIndex + PageSize
			If TempIndex > RecordCount Then TempIndex = FromIndex
			NextButton.Start = TempIndex
			NextButton.Enabled = (TempIndex <> FromIndex)

			' Last Button
			TempIndex = ((RecordCount-1)\PageSize)*PageSize + 1
			LastButton.Start = TempIndex
			LastButton.Enabled = (TempIndex <> FromIndex)
		End If
	End Sub

	' Terminate
	Private Sub Class_Terminate()
	Set FirstButton = Nothing
		Set PrevButton = Nothing
		Set NextButton = Nothing
		Set LastButton = Nothing
	End Sub
End Class

'
'  Pager classes and functions (end)
' -----------------------------------
'
' -------------
'  Field class
'
Class cField
	Dim TblName ' Table name
	Dim TblVar ' Table var
	Dim FldName ' Field name
	Dim FldVar ' Field variable name
	Dim FldExpression ' Field expression (used in SQL)
	Dim FldIsVirtual ' Virtual field
	Dim FldVirtualExpression ' Virtual field expression (used in ListSQL)
	Dim FldForceSelection ' Autosuggest force selection
	Dim VirtualValue ' Virtual field value
	Dim TooltipValue ' Field tooltip value
	Dim TooltipWidth ' Field tooltip width
	Dim FldType ' Field type
	Dim FldDataType ' Field data type
	Dim FldBlobType ' For Oracle only
	Dim FldViewTag ' View Tag
	Dim FldIsDetailKey ' Detail key
	Dim Visible ' Visible
	Dim Disabled ' Disabled
	Dim ReadOnly ' Read only
	Dim TruncateMemoRemoveHtml ' Remove Html from Memo field

	Public Property Get FldCaption() ' Field caption
		FldCaption = Language.FieldPhrase(TblVar, Mid(FldVar,3), "FldCaption")
	End Property

	Public Property Get FldTitle() ' Field title
		FldTitle = Language.FieldPhrase(TblVar, Mid(FldVar,3), "FldTitle")
	End Property

	Public Property Get FldAlt() ' Field alt
		FldAlt = Language.FieldPhrase(TblVar, Mid(FldVar,3), "FldAlt")
	End Property
	Dim FldDefaultErrMsg

	Public Property Get FldErrMsg() ' Field err msg
		FldErrMsg = Language.FieldPhrase(TblVar, Mid(FldVar,3), "FldErrMsg")
		If FldErrMsg = "" Then FldErrMsg = FldDefaultErrMsg & " - " & FldCaption
	End Property

	' Field tag caption
	Public Function FldTagCaption(i)
		FldTagCaption = Language.FieldPhrase(TblVar, Mid(FldVar,3), "FldTagCaption" & i)
	End Function

	' Reset attributes for field object
	Public Sub ResetAttrs()
		CssStyle = ""
		CssClass = ""
		CellCssStyle = ""
		CellCssClass = ""
		CellAttrs.Clear()
		EditAttrs.Clear()
		ViewAttrs.Clear()
		LinkAttrs.Clear()
	End Sub
	Dim FldDateTimeFormat ' Date time format
	Dim CssStyle ' Css style
	Dim CssClass ' Css class
	Dim ImageAlt ' Image alt
	Dim ImageWidth ' Image width
	Dim ImageHeight ' Image height
	Dim ImageResize ' Image resize
	Dim ViewCustomAttributes ' View custom attributes
	Dim CellAttrs ' Cell attributes
	Dim EditAttrs ' Edit attributes
	Dim ViewAttrs ' View attributes

	' View Attributes
	Public Property Get ViewAttributes()
		Dim sAtt, Attr, Value, i
		Dim sStyle, sClass
		sAtt = ""
		sStyle = CssStyle
		If ViewAttrs.Exists("style") Then
			Value = ViewAttrs.Item("style")
			If Trim(Value) <> "" Then
				sStyle = sStyle & " " & Value
			End If
		End If
		sClass = CssClass
		If ViewAttrs.Exists("class") Then
			Value = ViewAttrs.Item("class")
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
		If Trim(ImageAlt) <> "" Then
			sAtt = sAtt & " alt=""" & Trim(ImageAlt) & """"
		End If
		If CLng(ImageWidth) > 0 And (Not ImageResize Or (ImageResize And CLng(ImageHeight) <= 0)) Then
			sAtt = sAtt & " width=""" & CInt(ImageWidth) & """"
		End If
		If CLng(ImageHeight) > 0 And (Not ImageResize Or (ImageResize And CLng(ImageWidth) <= 0)) Then
			sAtt = sAtt & " height=""" & CInt(ImageHeight) & """"
		End If
		For i = 0 to UBound(ViewAttrs.Attributes)
			Attr = ViewAttrs.Attributes(i)(0)
			Value = ViewAttrs.Attributes(i)(1)
			If Attr <> "style" And Attr <> "class" And Value <> "" Then
				sAtt = sAtt & " " & Attr & "=""" & Value & """"
			End If
		Next
		If Trim(ViewCustomAttributes) <> "" Then
			sAtt = sAtt & " " & Trim(ViewCustomAttributes) 
		End If
		ViewAttributes = sAtt
	End Property
	Dim EditCustomAttributes ' Edit custom attributes

	' Edit Attributes
	Public Property Get EditAttributes()
		Dim sAtt, Attr, Value, i
		Dim sStyle, sClass
		sAtt = ""
		sStyle = CssStyle
		If EditAttrs.Exists("style") Then
			Value = EditAttrs.Item("style")
			If Trim(Value) <> "" Then
				sStyle = sStyle & " " & Value
			End If
		End If
		sClass = CssClass
		If EditAttrs.Exists("class") Then
			Value = EditAttrs.Item("class")
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
		For i = 0 to UBound(EditAttrs.Attributes)
			Attr = EditAttrs.Attributes(i)(0)
			Value = EditAttrs.Attributes(i)(1)
			If Attr <> "style" And Attr <> "class" And Value <> "" Then
				sAtt = sAtt & " " & Attr & "=""" & Value & """"
			End If
		Next
		If Trim(EditCustomAttributes) <> "" Then
			sAtt = sAtt & " " & Trim(EditCustomAttributes) 
		End If
		If Not EditAttrs.Exists("disabled") And Disabled Then
			sAtt = sAtt & " disabled=""disabled"""
		End If
		If Not EditAttrs.Exists("readonly") And ReadOnly Then
			sAtt = sAtt & " readonly=""readonly"""
		End If
		EditAttributes = sAtt
	End Property
	Dim CustomMsg ' Custom message
	Dim CellCssClass ' Cell CSS class
	Dim CellCssStyle ' Cell CSS style
	Dim CellCustomAttributes ' Cell custom attributes

	' Cell Styles
	Public Property Get CellStyles()
		Dim sAtt, Value
		Dim sStyle, sClass
		sAtt = ""
		sStyle = CellCssStyle
		If CellAttrs.Exists("style") Then
			Value = CellAttrs.Item("style")
			If Trim(Value) <> "" Then
				sStyle = sStyle & " " & Value
			End If
		End If
		sClass = CellCssClass
		If CellAttrs.Exists("class") Then
			Value = CellAttrs.Item("class")
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
		CellStyles = sAtt
	End Property

	' Cell Attributes
	Public Property Get CellAttributes()
		Dim sAtt, Attr, Value, i
		sAtt = CellStyles
		For i = 0 to UBound(CellAttrs.Attributes)
			Attr = CellAttrs.Attributes(i)(0)
			Value = CellAttrs.Attributes(i)(1)
			If Attr <> "style" And Attr <> "class" And Attr <> "" And Value <> "" Then
				sAtt = sAtt & " " & Attr & "=""" & Value & """"
			End If
		Next
		If Trim(CellCustomAttributes) <> "" Then
			sAtt = sAtt & " " & Trim(CellCustomAttributes) ' Cell custom attributes
		End If
		CellAttributes = sAtt
	End Property
	Dim LinkCustomAttributes ' Link custom attributes
	Dim LinkAttrs ' Link attributes

	' Link attributes
	Public Property Get LinkAttributes()
		Dim sAtt, Attr, Value, sHref, i
		sAtt = ""
		sHref = Trim(HrefValue)
		For i = 0 to UBound(LinkAttrs.Attributes)
			Attr = LinkAttrs.Attributes(i)(0)
			Value = LinkAttrs.Attributes(i)(1)
			If Trim(Value) <> "" Then
				If Attr = "href" Then
					sHref = sHref & " " & Value
				Else
					sAtt = sAtt & " " & Attr & "=""" & Trim(Value) & """"
				End If
			End If
		Next
		If sHref <> "" Then
			sAtt = sAtt & " href=""" & Trim(sHref) & """"
		End If
		If Trim(LinkCustomAttributes) <> "" Then
			sAtt = sAtt & " " & Trim(LinkCustomAttributes)
		End If
		LinkAttributes = sAtt
	End Property

	' Sort Attributes
	Dim Sortable

	Public Property Get Sort()
		Sort = Session(EW_PROJECT_NAME & "_" & TblVar & "_" & EW_TABLE_SORT & "_" & FldVar)
	End Property

	Public Property Let Sort(v)
		If Session(EW_PROJECT_NAME & "_" & TblVar & "_" & EW_TABLE_SORT & "_" & FldVar) <> v Then
			Session(EW_PROJECT_NAME & "_" & TblVar & "_" & EW_TABLE_SORT & "_" & FldVar) = v
		End If
	End Property

	Public Function ReverseSort()
		If Sort = "ASC" Then
			ReverseSort = "DESC"
		Else
			ReverseSort = "ASC"
		End If
	End Function
	Dim MultiUpdate ' Multi update
	Dim OldValue ' Old Value
	Dim ConfirmValue ' Confirm Value
	Dim CurrentValue ' Current value
	Dim ViewValue ' View value
	Dim EditValue ' Edit value
	Dim EditValue2 ' Edit value 2 (search)
	Dim HrefValue ' Href value
	Dim HrefValue2 ' Href value 2 (confirm page UPLOAD control)

	' List View value
	Public Property Get ListViewValue()
		If FldDataType = EW_DATATYPE_XML Then
			ListViewValue = ViewValue & "&nbsp;"
		ElseIf Trim(ViewValue & "") = "" Then
			ListViewValue = "&nbsp;"
		Else
			Dim regEx, Result
			Set regEx = New RegExp
			regEx.IgnoreCase = True
			regEx.Global = True

'			regEx.Pattern = "<[^>]+>" ' Remove all HTML Tags
'			regEx.Pattern = "</?(b|p|span)[^>]*[^>]*?>" ' Remove empty <b>/<p>/<span> tags

			regEx.Pattern = "<[^img][^>]*>" ' Remove all except non-empty image tag
			Result = regEx.Replace(ViewValue & "", "")
			Set regEx = Nothing
			If Trim(Result) = "" Then
				ListViewValue = "&nbsp;"
			Else
				ListViewValue = ViewValue
			End If
		End If
	End Property

	' Export Caption
	Public Property Get ExportCaption()
		If EW_EXPORT_FIELD_CAPTION Then
			ExportCaption = FldCaption
		Else
			ExportCaption = FldName
		End If
	End Property

	' Export Value
	Public Property Get ExportValue(Export, Original)
		If Original Then
			ExportValue = CurrentValue
		Else
			ExportValue = ViewValue
		End If
		If Export = "xml" Then
			If IsNull(ExportValue) Then ExportValue = "<Null>"
		End If
	End Property

	' Form value
	Private m_FormValue

	Public Property Get FormValue()
		FormValue = m_FormValue
	End Property

	Public Property Let FormValue(v)
		m_FormValue = v
		CurrentValue = m_FormValue
	End Property

	' QueryString value
	Private m_QueryStringValue

	Public Property Get QueryStringValue()
		QueryStringValue = m_QueryStringValue
	End Property

	Public Property Let QueryStringValue(v)
		m_QueryStringValue = v
		CurrentValue = m_QueryStringValue
	End Property

	' Database Value
	Dim m_DbValue

	Public Property Get DbValue()
		DbValue = m_DbValue
	End Property

	Public Property Let DbValue(v)
		m_DbValue = v
		CurrentValue = m_DbValue
	End Property

	' Set up database value
	Public Sub SetDbValue(rs, value, default, skip)
		Dim bSkipUpdate
		bSkipUpdate = skip Or Not Visible Or Disabled
		If bSkipUpdate Then Exit Sub
		Select Case FldType
			Case 2, 3, 16, 17, 18, 19, 20, 21 ' Int
				If IsNumeric(value) Then
					m_DbValue = CLng(value)
				Else
					m_DbValue = default
				End If
			Case 5, 6, 14, 131, 139 ' Double
				If IsNumeric(value) Then
					m_DbValue = CDbl(value)
				Else
					m_DbValue = default
				End If
			Case 4 ' Single
				If IsNumeric(value) Then
					m_DbValue = CSng(value)
				Else
					m_DbValue = default
				End If
			Case 7, 133, 134, 135, 145, 146 ' Date
				If IsDate(value) Then
					m_DbValue = CDate(value)
				ElseIf ew_IsDate(value) Then
					m_DbValue = value
				Else
					m_DbValue = default
				End If
			Case 201, 203, 129, 130, 200, 202 ' String
				m_DbValue = Trim(value)
				If EW_REMOVE_XSS Then m_DbValue = ew_RemoveXSS(m_DbValue)
				If m_DbValue = "" Then m_DbValue = default
			Case 128, 204, 205 ' Binary
				If IsNull(value) Then
					m_DbValue = default
				Else
					m_DbValue = value
				End If
			Case 72 ' GUID
				Dim RE
				Set RE = New RegExp
				RE.Pattern = "^(\{{1}([0-9a-fA-F]){8}-([0-9a-fA-F]){4}-([0-9a-fA-F]){4}-([0-9a-fA-F]){4}-([0-9a-fA-F]){12}\}{1})$"
				If RE.Test(Trim(value)) Then
					m_DbValue = Trim(value)
				Else
					m_DbValue = default
				End If
				Set RE = Nothing
			Case Else
				m_DbValue = value
		End Select
		rs(FldName) = m_DbValue
	End Sub

	' Session Value
	Public Property Get SessionValue()
		SessionValue = Session(EW_PROJECT_NAME & "_" & TblVar & "_" & FldVar & "_SessionValue")
	End Property

	Public Property Let SessionValue(v)
		Session(EW_PROJECT_NAME & "_" & TblVar & "_" & FldVar & "_SessionValue") = v
	End Property
	Dim Count ' Count
	Dim Total ' Total
	Dim TrueValue
	Dim FalseValue

	' AdvancedSearch Object
	Private m_AdvancedSearch

	Public Property Get AdvancedSearch()
		If Not IsObject(m_AdvancedSearch) Then Set m_AdvancedSearch = New cAdvancedSearch
		Set AdvancedSearch = m_AdvancedSearch
	End Property

	' Upload Object
	Private m_Upload

	Public Property Get Upload()
		If Not IsObject(m_Upload) Then
			Set m_Upload = New cUpload
			m_Upload.TblVar = TblVar
			m_Upload.FldVar = FldVar
		End If
		Set Upload = m_Upload
	End Property
	Dim UploadPath ' Upload path

	' Show object as string
	Public Function AsString()
		Dim AdvancedSearchAsString, UploadAsString
		If IsObject(m_AdvancedSearch) Then
			AdvancedSearchAsString = m_AdvancedSearch.AsString
		Else
			AdvancedSearchAsString = "{Null}"
		End If
		If IsObject(m_Upload) Then
			UploadAsString = m_Upload.AsString
		Else
			UploadAsString = "{Null}"
		End If
		AsString = "{" & _
			"FldName: " & FldName & ", " & _
			"FldVar: " & FldVar & ", " & _
			"FldExpression: " & FldExpression & ", " & _
			"FldType: " & FldType & ", " & _
			"FldDateTimeFormat: " & FldDateTimeFormat & ", " & _
			"CssStyle: " & CssStyle & ", " & _
			"CssClass: " & CssClass & ", " & _
			"ImageAlt: " & ImageAlt & ", " & _
			"ImageWidth: " & ImageWidth & ", " & _
			"ImageHeight: " & ImageHeight & ", " & _
			"ImageResize: " & ImageResize & ", " & _
			"ViewCustomAttributes: " & ViewCustomAttributes & ", " & _
			"EditCustomAttributes: " & EditCustomAttributes & ", " & _
			"CellCssStyle: " & CellCssStyle & ", " & _
			"CellCssClass: " & CellCssClass & ", " & _
			"Sort: " & Sort & ", " & _
			"MultiUpdate: " & MultiUpdate & ", " & _
			"CurrentValue: " & CurrentValue & ", " & _
			"ViewValue: " & ViewValue & ", " & _
			"EditValue: " & ValueToString(EditValue) & ", " & _
			"EditValue2: " & ValueToString(EditValue2) & ", " & _
			"HrefValue: " & HrefValue & ", " & _
			"HrefValue2: " & HrefValue2 & ", " & _
			"FormValue: " & m_FormValue & ", " & _
			"QueryStringValue: " & m_QueryStringValue & ", " & _
			"DbValue: " & m_DbValue & ", " & _
			"SessionValue: " & SessionValue & ", " & _
			"Count: " & Count & ", " & _
			"Total: " & Total & ", " & _
			"AdvancedSearch: " & AdvancedSearchAsString & ", " & _
			"Upload: " & UploadAsString & _
			"}"
	End Function

	' Value to string
	Private Function ValueToString(value)
		If IsArray(value) Then
			ValueToString = "[Array]"
		Else
			ValueToString = value
		End If
	End Function

	' Class Initialize
	Private Sub Class_Initialize()
		Count = 0
		Total = 0
		TrueValue = "1"
		FalseValue = "0"
		ImageWidth = 0
		ImageHeight = 0
		ImageResize = False
		Visible = True
		Disabled = False
		Sortable = True
		TruncateMemoRemoveHtml = False
		TooltipWidth = 0
		FldIsVirtual = False
		FldIsDetailKey = False
		Set CellAttrs = New cAttributes ' Cell attributes
		Set EditAttrs = New cAttributes ' Cell attributes
		Set ViewAttrs = New cAttributes ' Cell attributes
		Set LinkAttrs = New cAttributes ' Cell attributes
	End Sub

	' Class terminate
	Private Sub Class_Terminate
		If IsObject(m_AdvancedSearch) Then
			Set m_AdvancedSearch = Nothing
		End If
		If IsObject(m_Upload) Then
			Set m_Upload = Nothing
		End If
		Set CellAttrs = Nothing
		Set EditAttrs = Nothing
		Set ViewAttrs = Nothing
		Set LinkAttrs = Nothing
	End Sub
End Class

'
'  Field class (end)
' -------------------
' --------------------------------------
'  List option collection class (begin)
'
Class cListOptions
	Dim Items
	Dim CustomItem
	Dim Tag
	Dim Separator

	' Class initialize
	Private Sub Class_Initialize
		Set Items = Server.CreateObject("Scripting.Dictionary")
		Tag = "td"
		Separator = ""
	End Sub

	' Add and return a new option
	Public Function Add(Name)
		Set Add = New cListOption
		Add.Name = Name
		Add.Tag = Tag
		Add.Separator = Separator
		Set Add.Parent = Me
		Items.Add Items.Count, Add
	End Function

	' Load default settings
	Public Sub LoadDefault()
		CustomItem = ""
		For i = 0 to Items.Count - 1
			Items(i).Body = ""
		Next
	End Sub

	' Hide all options
	Public Sub HideAllOptions()
		Dim i
		For i = 0 to Items.Count - 1
			Items(i).Visible = False
		Next
	End Sub

	' Show all options
	Public Sub ShowAllOptions()
		Dim i
		For i = 0 to Items.Count - 1
			Items(i).Visible = True
		Next
	End Sub

	' Get Item By Name
	Public Function GetItem(Name)
		Dim i
   	For i = 0 To Items.Count - 1
      If Items.Item(i).Name = Name Then
      	Set GetItem = Items.Item(i)
      	Exit Function
      End If
   	Next
   	Set GetItem = Nothing
	End Function

	' Move item to position
	Public Sub MoveItem(Name, Pos)
		Dim i, oldpos, bfound
		If Pos < 0 Then
			Pos = 0
		ElseIf Pos >= Items.Count Then
			Pos = Items.Count - 1
		End If
		bfound = False
		For i = 0 To Items.Count - 1
			If Items.Item(i).Name = Name Then
				bfound = True
				oldpos = i
				Exit For
			End If
		Next
		If bfound And Pos <> oldpos Then
			Items.Key(oldpos) = Items.Count ' Move out of position first
			If oldpos < Pos Then ' Shuffle backward
				For i = oldpos+1 to Pos
					Items.Key(i) = i-1
				Next
			Else ' Shuffle forward
				For i = oldpos-1 to Pos Step -1
					Items.Key(i) = i+1
				Next
			End If
			Items.Key(Items.Count) = Pos ' Move to position
		End If
	End Sub

	' Render list options
	Sub Render(Part, Pos)
		Dim cnt, opt
		If CustomItem <> "" Then
			cnt = 0
			For i = 0 to Items.Count - 1
				If Items(i).Visible And ShowPos(Items(i).OnLeft, Pos) Then cnt = cnt + 1
				If Items(i).Name = CustomItem Then Set opt = Items(i)
			Next
			If IsObject(opt) And cnt > 0 Then
				If ShowPos(opt.OnLeft, Pos) Then
					Response.Write opt.Render(Part, cnt)
				Else
					Response.Write opt.Render("", cnt)
				End If
			End If
		Else
			cnt = 1
			For i = 0 to Items.Count - 1
				If Items(i).Visible And ShowPos(Items(i).OnLeft, Pos) Then Response.Write Items(i).Render(Part, cnt)
			Next
		End If
	End Sub

	Private Function ShowPos(OnLeft, Pos)
		ShowPos = (OnLeft And Pos = "left") Or (Not OnLeft And Pos = "right") Or (Pos = "")
	End Function

	' Class terminate
	Private Sub Class_Terminate
		Dim i
		For i = 0 To Items.Count - 1
			Set Items.Item(i) = Nothing
		Next
	End Sub
End Class

'
'  List option collection class (end)
' ------------------------------------
'
' ---------------------------
'  List option class (begin)
'
Class cListOption
	Dim Name
	Dim OnLeft
	Dim CssStyle
	Dim Visible
	Dim Header
	Dim Body
	Dim Footer
	Dim Tag
	Dim Separator
	Dim Parent

	' Class initialize
	Private Sub Class_Initialize
		OnLeft = False
		Visible = True
		Tag = "td"
		Separator = ""
	End Sub

	Public Sub MoveTo(Pos)
		Parent.MoveItem Name, Pos
	End Sub

	Public Function Render(Part, Colspan)
		Dim value, res, tags, tage
		If Part = "header" Then
			value = Header
		ElseIf Part = "body" Then
			value = Body
		ElseIf Part = "footer" Then
			value = Footer
		Else
			value = Part
		End If
		If value = "" And LCase(Tag) <> "td" Then
			Render = ""
			Exit Function
		End If
		res = ew_IIf(value <> "", value, "&nbsp;")
		tage = "</" & Tag & ">"
		tags = "<" & Tag
		tags = tags & " class=""aspmaker"""
		If CssStyle <> "" Then
			tags = tags & " style=""" & CssStyle & """"
		End If
		If LCase(Tag) = "td" And Colspan > 1 Then
			tags = tags & " colspan=""" & Colspan & """"
		End If
		tags = tags & ">"
		Render = tags & res & tage & Separator
	End Function

	' Convert to string
	Public Function AsString
		AsString = "{" & _
			"Name: " & Name & ", " & _
			"OnLeft: " & OnLeft & ", " & _
			"CssStyle: " & CssStyle & ", " & _
			"Visible: " & Visible & ", " & _
			"Header: " & Server.HTMLEncode(Header) & ", " & _
			"Body: " & Server.HTMLEncode(Body) & ", " & _
			"Footer: " & Server.HTMLEncode(Footer) & _
			"}"
	End Function
End Class

'
'  List option class (end)
' -------------------------
' Menu class
Class cMenu

	Public Id

	Public IsRoot

	Public ItemData

	' Init
	Private Sub Class_Initialize
		IsRoot = False
		Set ItemData = Server.CreateObject("Scripting.Dictionary") ' Data type: array of cMenuItem
	End Sub

	' Terminate
	Private Sub Class_Terminate
		Set ItemData = Nothing
	End Sub

	' Get menu item count
	Function Count()
		Count = ItemData.Count
	End Function

	' Move item to position
	Sub MoveItem(Text, Pos)
		Dim i, oldpos, bfound, Items
		Set Items = ItemData
		If Pos < 0 Then
			Pos = 0
		ElseIf Pos >= Items.Count Then
			Pos = Items.Count - 1
		End If
		bfound = False
		For i = 0 To Items.Count - 1
			If Items.Item(i).Text = Text Then
				bfound = True
				oldpos = i
				Exit For
			End If
		Next
		If bfound And Pos <> oldpos Then
			Items.Key(oldpos) = Items.Count ' Move out of position first
			If oldpos < Pos Then ' Shuffle backward
				For i = oldpos+1 to Pos
					Items.Key(i) = i-1
				Next
			Else ' Shuffle forward
				For i = oldpos-1 to Pos Step -1
					Items.Key(i) = i+1
				Next
			End If
			Items.Key(Items.Count) = Pos ' Move to position
		End If
	End Sub

	' Create a menu item
	Function NewMenuItem(id, text, url, parentid, source, target, allowed, grouptitle)
		Set NewMenuItem = New cMenuItem
		NewMenuItem.Id = id
		NewMenuItem.Text = text
		NewMenuItem.Url = url
		NewMenuItem.ParentId = parentid
		NewMenuItem.Target = target
		NewMenuItem.Source = source
		NewMenuItem.Allowed = allowed
		NewMenuItem.GroupTitle = grouptitle
	End Function

	' Add a menu item
	Sub AddMenuItem(id, text, url, parentid, source, target, allowed, grouptitle)
		Dim item, oParentMenu
		Set item = NewMenuItem(id, text, url, parentid, source, target, allowed, grouptitle)
		If Not MenuItem_Adding(item) Then
			Exit Sub
		End If
		If item.ParentId < 0 Then
			AddItem(item)
		Else
			If FindItem(item.parentid, oParentMenu) Then
				oParentMenu.AddItem(item)
			End If
		End If
	End Sub

	' Add item to internal dictionary
	Sub AddItem(item)
		ItemData.Add ItemData.Count, item
	End Sub

	' Clear all menu items
	Sub Clear()
		Dim i
		For i = 0 To ItemData.Count -1
			Set ItemData.Item(i) = Nothing
		Next
		ItemData.RemoveAll
	End Sub

	' Find item
	Function FindItem(id, out)
		Dim i, item
		FindItem = False
		For i = 0 To ItemData.Count -1
			If ItemData.Item(i).Id = id Then
				Set out = ItemData.Item(i)
				FindItem = True
				Exit Function
			ElseIf Not IsNull(ItemData.Item(i).SubMenu) Then
				FindItem = ItemData.Item(i).SubMenu.FindItem(id, out)
			End If
		Next
	End Function

	' Check if a menu item should be shown
	Function RenderItem(item)
		Dim i, subitem
		If Not IsNull(item.SubMenu) Then
			For i = 0 To item.SubMenu.ItemData.Count - 1
				If item.SubMenu.RenderItem(item.SubMenu.ItemData.Item(i)) Then
					RenderItem = True
					Exit Function
				End If
			Next
		End If
		RenderItem = (item.Allowed And item.Url <> "")
	End Function

	' Check if this menu should be rendered
	Function RenderMenu()
		Dim i
		For i = 0 To ItemData.Count - 1
			If RenderItem(ItemData.Item(i)) Then
				RenderMenu = True
				Exit Function
			End If
		Next
		RenderMenu = False
	End Function

	' Render the menu
	Function Render(ret)
		Dim str, gopen, gcnt, i, j, classfot, itemcnt, item, aclass, liclass
		Call Menu_Rendering(Me)
		If Not RenderMenu() Then Exit Function
		str = "<div"
		If Id <> "" Then
			If IsNumeric(Id) Then
				str = str & " id=""menu_" & Id & """"
			Else
				str = str & " id=""" & Id & """"
			End If
		End If
		str = str & " class=""" & ew_IIf(IsRoot, EW_MENUBAR_CLASSNAME, EW_MENU_CLASSNAME) & """>"
		str = str & "<div class=""bd" & ew_IIf(IsRoot, " first-of-type", "") & """>" & vbCrLf
		gopen = False ' Group open status
		gcnt = 0 ' Group count
		i = 0 ' Menu item count
		classfot = " class=""first-of-type"""
		itemcnt = ItemData.Count
		For j = 0 to itemcnt - 1
			Set item = ItemData.Item(j)
			If RenderItem(item) Then
				i = i + 1

				' Begin a group
				If i = 1 And Not item.GroupTitle Then
					gcnt = gcnt + 1
					str = str & "<ul " & classfot & ">" & vbCrLf
					gopen = True
				End If
				aclass = ew_IIf(IsRoot, EW_MENUBAR_ITEM_LABEL_CLASSNAME, EW_MENU_ITEM_LABEL_CLASSNAME)
				liclass = ew_IIf(IsRoot, EW_MENUBAR_ITEM_CLASSNAME, EW_MENU_ITEM_CLASSNAME)
				If item.GroupTitle And EW_MENU_ITEM_CLASSNAME <> "" Then ' Group title
					gcnt = gcnt + 1
					If i > 1 And gopen Then
						str = str & "</ul>" & vbCrLf ' End last group
						gopen = False
					End If

					' Begin a new group with title
					If item.Text <> "" Then
						str = str & "<h6" & ew_IIf(gcnt = 1, classfot, "") & ">" & item.Text & "</h6>" & vbCrLf
					End If
					str = str & "<ul" & ew_IIf(gcnt = 1, classfot, "") & ">" & vbCrLf
					gopen = True
					If Not IsNull(item.SubMenu) Then
						Dim subitem, subitemcnt, k
						subitemcnt = item.SubMenu.ItemData.Count
						For k = 0 to subitemcnt - 1
							Set subitem = item.SubMenu.ItemData.Item(k)
							If RenderItem(subitem) Then
								str = str & subitem.Render(aclass, liclass) & vbCrLf ' Create <LI>
							End If
						Next
					End If
					str = str & "</ul>" & vbCrLf ' End the group
					gopen = False
				Else ' Menu item
					If Not gopen Then ' Begin a group if no opened group
						gcnt = gcnt + 1
						str = str & "<ul" & ew_IIf(gcnt = 1, classfot, "") & ">" & vbCrLf
						gopen = True
					End If
					If IsRoot And i = 1 Then ' For horizontal menu
						liclass = liclass & " first-of-type"
					End If
					str = str & item.Render(aclass, liclass) & vbCrLf ' Create <LI>
				End If
			End If
		Next
		If gopen Then
			str = str & "</ul>" & vbCrLf ' End last group
		End If
		str = str & "</div></div>" & vbCrLf
		If ret Then ' Return as string
			Render = str
		Else
			Response.Write str ' Output
		End If
	End Function
End Class

' Menu item class
Class cMenuItem

	Public Id

	Public Text

	Public Url

	Public ParentId

	Public Source

	Public Target

	Public Allowed

	Public GroupTitle

	Public SubMenu ' Data type = cMenu

	Private Sub Class_Initialize
		Url = ""
		GroupTitle = False
		SubMenu = Null
	End Sub

	Sub AddItem(item) ' Add submenu item
		If IsNull(SubMenu) Then
			Set SubMenu = New cMenu
			SubMenu.Id = Id
		End If
		SubMenu.AddItem(item)
	End Sub

	' Render
	Function Render(aclass, liclass)

		' Create <A>
		Dim attrs, innerhtml
		attrs = Array(Array("class", aclass), Array("href", Url), Array("target", Target))
		innerhtml = ew_HtmlElement("a", attrs, Text, True)
		If Not IsNull(SubMenu) Then
			innerhtml = innerhtml & SubMenu.Render(True)
		End If

		' Create <LI>
		Render = ew_HtmlElement("li", Array(Array("class", liclass)), innerhtml, True)
	End Function

	Function AsString
		AsString = "{ Id: " & Id & ", Text: " & Text & ", Url: " & Url & ", ParentId: " & ParentId & ", Target: " & Target & ", Source: " & Source & ", Allowed: " & Allowed
		If IsNull(SubMenu) Then
			AsString = AsString & ", SubMenu: (Null)"
		Else
			AsString = AsString & ", SubMenu: (Object)"
		End If
		AsString = AsString & " }" & "<br>"
	End Function
End Class

' Menu Rendering event
Sub Menu_Rendering(Menu)

	' Change menu items here
End Sub

Function MenuItem_Adding(Item)

	'Response.Write Item.AsString
	' Return False if menu item not allowed

	MenuItem_Adding = True
End Function

' Output SCRIPT tag
Sub ew_AddClientScript(src)
	ew_AddClientScriptEx src, Null
End Sub

' Output SCRIPT tag
Sub ew_AddClientScriptEx(src, attrs)
	Dim atts
	atts = Array(Array("type", "text/javascript"), Array("src", src))
	If IsArray(attrs) Then
		atts = ew_MergeAttrs(atts, attrs)
	End If
	Response.Write ew_HtmlElement("script", atts, "", True) & vbCrLf
End Sub

' Output LINK tag
Sub ew_AddStylesheet(href)
	ew_AddStylesheetEx href, Null
End Sub

' Output LINK tag
Sub ew_AddStylesheetEx(href, attrs)
	Dim atts
	atts = Array(Array("rel", "stylesheet"), Array("type", "text/css"), Array("href", href))
	If IsArray(attrs) Then
		atts = ew_MergeAttrs(atts, attrs)
	End If
	Response.Write ew_HtmlElement("link", atts, "", False) & vbCrLf
End Sub

' Build HTML element
Function ew_HtmlElement(tagname, attrs, innerhtml, endtag)
	Dim html, i, name, attr
	html = "<" & tagname
	If IsArray(attrs) Then
		For i = 0 to UBound(attrs)
			If IsArray(attrs(i)) Then
				If UBound(attrs(i)) >= 1 Then
					name = attrs(i)(0)
					attr = attrs(i)(1)
					If attr <> "" Then
						html = html & " " & name & "=""" & ew_HtmlEncode(attr) & """"
					End If
				End If
			End If
		Next
	End If
	html = html & ">"
	If innerhtml <> "" Then
		html = html & innerhtml
	End If
	If endtag Then
		html = html & "</" & tagname & ">"
	End If
	ew_HtmlElement = html
End Function

Function ew_MergeAttrs(attrs1, attrs2)
	Dim attrs, i, cnt, idx
	cnt = 0
	If IsArray(attrs1) Then cnt = cnt + UBound(attrs1) + 1
	If IsArray(attrs2) Then cnt = cnt + UBound(attrs1) + 1
	If cnt > 0 Then
		ReDim attrs(cnt-1)
		idx = 0
		If IsArray(attrs1) Then
			For i = 0 to UBound(attrs1)
				attrs(idx) = attrs1(i)
				idx = idx + 1
			Next
		End If
		If IsArray(attrs2) Then
			For i = 0 to UBound(attrs2)
				attrs(idx) = attrs2(i)
				idx = idx + 1
			Next
		End If
	End If
	ew_MergeAttrs = attrs
End Function

' XML tag name
Function ew_XmlTagName(name)
	Dim RegEx, wrkname
	wrkname = Trim(name)
	Set RegEx = New RegExp
	RegEx.Global = True
	RegEx.IgnoreCase = True

	'RegEx.Pattern = "\A(?!XML)[a-z][\w0-9-]*"
	RegEx.Pattern = "[a-z][\w0-9-]*"
	If Not RegEx.Test(wrkname) Then
		wrkname = "_" & wrkname
	End If
	Set RegEx = Nothing
	ew_XmlTagName = wrkname
End Function
%>
<%

' -------------------------------
'  Advanced Search class (begin)
'
Class cAdvancedSearch
	Dim SearchValue ' Search value
	Dim SearchOperator ' Search operator
	Dim SearchCondition ' Search condition
	Dim SearchValue2 ' Search value 2
	Dim SearchOperator2 ' Search operator 2

	' Show object as string
	Public Function AsString()
		AsString = "{" & _
			"SearchValue: " & SearchValue & ", " & _
			"SearchOperator: " & SearchOperator & ", " & _
			"SearchCondition: " & SearchCondition & ", " & _
			"SearchValue2: " & SearchValue2 & ", " & _
			"SearchOperator2: " & SearchOperator2 & _
			"}"
	End Function
End Class

'
'  Advanced Search class (end)
' -----------------------------

%>
<%

' ----------------------
'  Upload class (begin)
'
Class cUpload
	Dim Index ' Index to handle multiple form elements

	' Class initialize
	Private Sub Class_Initialize
		Index = 0
	End Sub
	Dim TblVar ' Table variable
	Dim FldVar ' Field variable

	' Error message
	Private m_Message

	Public Property Get Message()
		Message = m_Message
	End Property
	Dim DbValue ' Value from database

	' Upload value
	Dim m_Value

	Public Property Get Value()
		Value = m_Value
	End Property

	Public Property Let Value(v)
		m_Value = v
	End Property

	' Upload action
	Private m_Action

	Public Property Get Action()
		Action = m_Action
	End Property

	' Upload file name
	Private m_FileName

	Public Property Get FileName()
		FileName = m_FileName
	End Property

	' Upload file size
	Private m_FileSize

	Public Property Get FileSize()
		FileSize = m_FileSize
	End Property

	' File content type
	Private m_ContentType

	Public Property Get ContentType()
		ContentType = m_ContentType
	End Property

	' Image width
	Private m_ImageWidth

	Public Property Get ImageWidth()
		ImageWidth = m_ImageWidth
	End Property

	' Image height
	Private m_ImageHeight

	Public Property Get ImageHeight()
		ImageHeight = m_ImageHeight
	End Property

	' Save Db value to Session
	Public Sub SaveDbToSession()
		Dim sSessionID
		sSessionID = EW_PROJECT_NAME & "_" & TblVar & "_" & FldVar & "_" & Index
		Session(sSessionID & "_DbValue") = DbValue
	End Sub

	' Restore Db value from Session
	Public Sub RestoreDbFromSession()
		Dim sSessionID
		sSessionID = EW_PROJECT_NAME & "_" & TblVar & "_" & FldVar & "_" & Index
		DbValue = Session(sSessionID & "_DbValue")
	End Sub

	' Remove Db value from Session
	Public Sub RemoveDbFromSession()
		Dim sSessionID
		sSessionID = EW_PROJECT_NAME & "_" & TblVar & "_" & FldVar & "_" & Index
		Session.Contents.Remove(sSessionID & "_DbValue")
	End Sub

	' Save Upload values to Session
	Public Sub SaveToSession()
		Dim sSessionID
		sSessionID = EW_PROJECT_NAME & "_" & TblVar & "_" & FldVar & "_" & Index
		Session(sSessionID & "_Action") = m_Action
		Session(sSessionID & "_FileSize") = m_FileSize
		Session(sSessionID & "_FileName") = m_FileName
		Session(sSessionID & "_ContentType") = m_ContentType
		Session(sSessionID & "_ImageWidth") = m_ImageWidth
		Session(sSessionID & "_ImageHeight") = m_ImageHeight
		Session(sSessionID & "_Value") = m_Value
	End Sub

	' Restore Upload values from Session
	Public Sub RestoreFromSession()
		Dim sSessionID
		sSessionID = EW_PROJECT_NAME & "_" & TblVar & "_" & FldVar & "_" & Index
		m_Action = Session(sSessionID & "_Action")
		m_FileSize = Session(sSessionID & "_FileSize")
		m_FileName = Session(sSessionID & "_FileName")
		m_ContentType = Session(sSessionID & "_ContentType")
		m_ImageWidth = Session(sSessionID & "_ImageWidth")
		m_ImageHeight = Session(sSessionID & "_ImageHeight")
		m_Value = Session(sSessionID & "_Value")
	End Sub

	' Remove Upload values from Session
	Public Sub RemoveFromSession()
		Dim sSessionID
		sSessionID = EW_PROJECT_NAME & "_" & TblVar & "_" & FldVar & "_" & Index
		Session.Contents.Remove(sSessionID & "_Action")
		Session.Contents.Remove(sSessionID & "_FileSize")
		Session.Contents.Remove(sSessionID & "_FileName")
		Session.Contents.Remove(sSessionID & "_ContentType")
		Session.Contents.Remove(sSessionID & "_ImageWidth")
		Session.Contents.Remove(sSessionID & "_ImageHeight")
		Session.Contents.Remove(sSessionID & "_Value")
		Call RemoveDbFromSession()
	End Sub

	' Function to check the file type of the uploaded file
	Private Function UploadAllowedFileExt(FileName)
		If Trim(FileName & "") = "" Then
			UploadAllowedFileExt = True
			Exit Function
		End If
		Dim Ext, Pos, arExt, FileExt
		arExt = Split(EW_UPLOAD_ALLOWED_FILE_EXT & "", ",")
		Ext = ""
		Pos = InStrRev(FileName, ".")
		If Pos > 0 Then	Ext = Mid(FileName, Pos+1)
		UploadAllowedFileExt = False
		For Each FileExt in arExt
	 		If LCase(Trim(FileExt)) = LCase(Ext) Then
				UploadAllowedFileExt = True
				Exit For
			End If
		Next
	End Function

	' Get upload file
	Public Function UploadFile()
		Dim gsFldVar, gsFldVarAction
		gsFldVar = FldVar
		gsFldVarAction = "a" & Mid(gsFldVar, 2)

		' Initialize upload value
		m_Value = Null

		' Get action
		m_Action = ObjForm.GetValue(gsFldVarAction)

		' Get and check the upload file size
		m_FileSize = ObjForm.GetUploadFileSize(gsFldVar)

		' Get and check the upload file type
		m_FileName = ObjForm.GetUploadFileName(gsFldVar)

		' Get upload file content type
		m_ContentType = ObjForm.GetUploadFileContentType(gsFldVar)

		' Get upload value
		m_Value = ObjForm.GetUploadFileData(gsFldVar)

		' Get image width and height
		m_ImageWidth = ObjForm.GetUploadImageWidth(gsFldVar)
		m_ImageHeight = ObjForm.GetUploadImageHeight(gsFldVar)
		UploadFile = True ' Normal return
	End Function

	' Resize image
	Public Function Resize(Width, Height, Interpolation)
		Dim wrkWidth, wrkHeight
		If Not IsNull(m_Value) Then
			wrkWidth = Width
			wrkHeight = Height
			If ew_ResizeBinary(m_Value, wrkWidth, wrkHeight, Interpolation) Then
				m_ImageWidth = wrkWidth
				m_ImageHeight = wrkHeight
				m_FileSize = LenB(m_Value)
			End If
		End If
	End Function

	' Save uploaded data to file (Path relative to application root)
	Public Function SaveToFile(Path, NewFileName, Overwrite)
		SaveToFile = False
		If Not IsNull(m_Value) Then
			Path = ew_UploadPathEx(True, Path)
			If Trim(NewFileName & "") = "" Then NewFileName = m_FileName
			If Overwrite Then
				SaveToFile = ew_SaveFile(Path, NewFileName, m_Value)
			Else
				SaveToFile = ew_SaveFile(Path, ew_UploadFileNameEx(Path, NewFileName), m_Value)
			End If
		End If
	End Function

	' Resize and save uploaded data to file (Path relative to application root)
	Public Function ResizeAndSaveToFile(Width, Height, Interpolation, Path, NewFileName, Overwrite)
		Dim OldValue, OldWidth, OldHeight, OldFileSize
		ResizeAndSaveToFile = False
		If Not IsNull(m_Value) Then
			OldValue = m_Value: OldWidth = m_ImageWidth: OldHeight = m_ImageHeight: OldFileSize = m_FileSize ' Save old values
			Call Resize(Width, Height, Interpolation)
			ResizeAndSaveToFile = SaveToFile(Path, NewFileName, Overwrite)
			m_Value = OldValue: m_ImageWidth = OldWidth: m_ImageHeight = OldHeight: m_FileSize = OldFileSize ' Restore old values
		End If
	End Function

	' Show object as string
	Public Function AsString()
		AsString = "{" & _
			"Index: " & Index & ", " & _
			"Message: " & m_Message & ", " & _
			"Action: " & m_Action & ", " & _
			"FileName: " & m_FileName & ", " & _
			"FileSize: " & m_FileSize & ", " & _
			"ContentType: " & m_ContentType & ", " & _
			"ImageWidth: " & m_ImageWidth & ", " & _
			"ImageHeight: " & m_ImageHeight & _
			"}"
	End Function
End Class

'
'  Upload class (end)
' --------------------

%>
<%

' ---------------------------------
'  Advanced Security class (begin)
'
Class cAdvancedSecurity
	Dim m_ArUserLevel
	Dim m_ArUserLevelPriv
	Dim m_ArUserLevelID

	' Current user level id / user level
	Dim CurrentUserLevelID
	Dim CurrentUserLevel

	' Current user id / parent user id / user id array
	Dim CurrentUserID
	Dim CurrentParentUserID
	Dim m_ArUserID

	' Class Initialize
	Private Sub Class_Initialize()

		' Init User Level
		CurrentUserLevelID = SessionUserLevelID
		If IsNumeric(CurrentUserLevelID) Then
			If CurrentUserLevelID >= -1 Then
				ReDim m_ArUserLevelID(0)
				m_ArUserLevelID(0) = CurrentUserLevelID
			End If
		End If

		' Init User ID
		CurrentUserID = SessionUserID
		CurrentParentUserID = SessionParentUserID

		' Load user level (for TablePermission_Loading event)
		Call LoadUserLevel()
	End Sub

	' Session user id
	Public Property Get SessionUserID()
		SessionUserID = Session(EW_SESSION_USER_ID) & ""
	End Property

	Public Property Let SessionUserID(v)
		Session(EW_SESSION_USER_ID) = Trim(v & "")
		CurrentUserID = Trim(v & "")
	End Property

	' Session parent user id
	Public Property Get SessionParentUserID()
		SessionParentUserID = Session(EW_SESSION_PARENT_USER_ID) & ""
	End Property

	Public Property Let SessionParentUserID(v)
		Session(EW_SESSION_PARENT_USER_ID) = Trim(v & "")
		CurrentParentUserID = Trim(v & "")
	End Property

	' Current user name
	Public Property Get CurrentUserName()
		CurrentUserName = Session(EW_SESSION_USER_NAME) & ""
	End Property

	Public Property Let CurrentUserName(v)
		Session(EW_SESSION_USER_NAME) = v
	End Property

	' Session user level id
	Public Property Get SessionUserLevelID()
		SessionUserLevelID = Session(EW_SESSION_USER_LEVEL_ID)
	End Property

	Public Property Let SessionUserLevelID(v)
		Session(EW_SESSION_USER_LEVEL_ID) = v
		CurrentUserLevelID = v
		If IsNumeric(CurrentUserLevelID) Then
			If CurrentUserLevelID >= -1 Then
				ReDim m_ArUserLevelID(0)
				m_ArUserLevelID(0) = CurrentUserLevelID
			End If
		End If
	End Property

	' Session user level value
	Public Property Get SessionUserLevel()
		SessionUserLevel = Session(EW_SESSION_USER_LEVEL)
	End Property

	Public Property Let SessionUserLevel(v)
		Session(EW_SESSION_USER_LEVEL) = v
		CurrentUserLevel = v
	End Property

	' Can add
	Public Property Get CanAdd()
		CanAdd = ((CurrentUserLevel And EW_ALLOW_ADD) = EW_ALLOW_ADD)
	End Property

	Public Property Let CanAdd(b)
		If (b) Then
			CurrentUserLevel = (CurrentUserLevel Or EW_ALLOW_ADD)
		Else
			CurrentUserLevel = (CurrentUserLevel And (Not EW_ALLOW_ADD))
		End If
	End Property

	' Can delete
	Public Property Get CanDelete()
		CanDelete = ((CurrentUserLevel And EW_ALLOW_DELETE) = EW_ALLOW_DELETE)
	End Property

	Public Property Let CanDelete(b)
		If (b) Then
			CurrentUserLevel = (CurrentUserLevel Or EW_ALLOW_DELETE)
		Else
			CurrentUserLevel = (CurrentUserLevel And (Not EW_ALLOW_DELETE))
		End If
	End Property

	' Can edit
	Public Property Get CanEdit()
		CanEdit = ((CurrentUserLevel And EW_ALLOW_EDIT) = EW_ALLOW_EDIT)
	End Property

	Public Property Let CanEdit(b)
		If (b) Then
			CurrentUserLevel = (CurrentUserLevel Or EW_ALLOW_EDIT)
		Else
			CurrentUserLevel = (CurrentUserLevel And (Not EW_ALLOW_EDIT))
		End If
	End Property

	' Can view
	Public Property Get CanView()
		CanView = ((CurrentUserLevel And EW_ALLOW_VIEW) = EW_ALLOW_VIEW)
	End Property

	Public Property Let CanView(b)
		If (b) Then
			CurrentUserLevel = (CurrentUserLevel Or EW_ALLOW_VIEW)
		Else
			CurrentUserLevel = (CurrentUserLevel And (Not EW_ALLOW_VIEW))
		End If
	End Property

	' Can list
	Public Property Get CanList()
		CanList = ((CurrentUserLevel And EW_ALLOW_LIST) = EW_ALLOW_LIST)
	End Property

	Public Property Let CanList(b)
		If (b) Then
			CurrentUserLevel = (CurrentUserLevel Or EW_ALLOW_LIST)
		Else
			CurrentUserLevel = (CurrentUserLevel And (Not EW_ALLOW_LIST))
		End If
	End Property

	' Can report
	Public Property Get CanReport()
		CanReport = ((CurrentUserLevel And EW_ALLOW_REPORT) = EW_ALLOW_REPORT)
	End Property

	Public Property Let CanReport(b)
		If (b) Then
			CurrentUserLevel = (CurrentUserLevel Or EW_ALLOW_REPORT)
		Else
			CurrentUserLevel = (CurrentUserLevel And (Not EW_ALLOW_REPORT))
		End If
	End Property

	' Can search
	Public Property Get CanSearch()
		CanSearch = ((CurrentUserLevel And EW_ALLOW_SEARCH) = EW_ALLOW_SEARCH)
	End Property

	Public Property Let CanSearch(b)
		If (b) Then
			CurrentUserLevel = (CurrentUserLevel Or EW_ALLOW_SEARCH)
		Else
			CurrentUserLevel = (CurrentUserLevel And (Not EW_ALLOW_SEARCH))
		End If
	End Property

	' Can admin
	Public Property Get CanAdmin()
		CanAdmin = ((CurrentUserLevel And EW_ALLOW_ADMIN) = EW_ALLOW_ADMIN)
	End Property

	Public Property Let CanAdmin(b)
		If (b) Then
			CurrentUserLevel = (CurrentUserLevel Or EW_ALLOW_ADMIN)
		Else
			CurrentUserLevel = (CurrentUserLevel And (Not EW_ALLOW_ADMIN))
		End If
	End Property

	' Last url
	Public Property Get LastUrl()
		LastUrl = Request.Cookies(EW_PROJECT_NAME)("lasturl")
	End Property

	' Save last url
	Public Sub SaveLastUrl()
		Dim s, q
		s = Request.ServerVariables("SCRIPT_NAME")
		q = Request.ServerVariables("QUERY_STRING")
		If q <> "" Then s = s & "?" & q
		If LastUrl = s Then s = ""
		Response.Cookies(EW_PROJECT_NAME)("lasturl") = s
	End Sub

	' Auto login
	Public Function AutoLogin()
		Dim sUsr, sPwd
		If Request.Cookies(EW_PROJECT_NAME)("autologin") = "autologin" Then
			sUsr = Request.Cookies(EW_PROJECT_NAME)("username")
			sPwd = Request.Cookies(EW_PROJECT_NAME)("password")
			sPwd = TEAdecrypt(ew_Decode(sPwd), EW_RANDOM_KEY)
			AutoLogin = ValidateUser(sUsr, sPwd, True)
		Else
			AutoLogin = False
		End If
	End Function

	' Validate user
	Public Function ValidateUser(usr, pwd, autologin)
		Dim RsUser, sFilter, sSql
		ValidateUser = False

		' Call User Custom Validate event
		If EW_USE_CUSTOM_LOGIN Then
			ValidateUser = User_CustomValidate(usr, pwd)
			If ValidateUser Then
				Session(EW_SESSION_STATUS) = "login"
			End If
		End If

		' Check other users
		If Not ValidateUser Then
				sFilter = Replace(EW_USER_NAME_FILTER, "%u", ew_AdjustSql(usr))

				' Get Sql from GetSql function in <UseTable> class, <UserTable>info.asp
				sSql = Users.GetSQL(sFilter, "")
				Set RsUser = Conn.Execute(sSql)
				If Not RsUser.Eof Then
					ValidateUser = ew_ComparePassword(RsUser("Password"), pwd)

					' Load user profile
					UserProfile.LoadProfileFromDatabase usr

					' Set up retry count from manual login
					Dim retrycount
					If Not autologin Then
						If Not ValidateUser Then
							retrycount = UserProfile.GetValue(EW_USER_PROFILE_LOGIN_RETRY_COUNT)
							retrycount = retrycount + 1
							UserProfile.SetValue EW_USER_PROFILE_LOGIN_RETRY_COUNT, retrycount
							UserProfile.SetValue EW_USER_PROFILE_LAST_BAD_LOGIN_DATE_TIME, ew_StdCurrentDateTime()
						Else
							UserProfile.SetValue EW_USER_PROFILE_LOGIN_RETRY_COUNT, 0
						End If
						UserProfile.SaveProfileToDatabase usr ' Save profile
					End If
					If ValidateUser Then
						Session(EW_SESSION_STATUS) = "login"
						Session(EW_SESSION_SYS_ADMIN) = 0 ' Non System Administrator
						CurrentUserName = RsUser("Username") ' Load user name
						If IsNull(RsUser("Permissions")) Then
							SessionUserLevelID = 0
						Else
							SessionUserLevelID = CLng(RsUser("Permissions")) ' Load user level
						End If
						Call SetUpUserLevel()

						' Call User Validated event
						Call User_Validated(RsUser)
					End If
				End If
				RsUser.Close
				Set RsUser = Nothing
		End If
		If Not ValidateUser Then Session(EW_SESSION_STATUS) = "" ' Clear login status
	End Function

	' Dynamic user level security
	' Get current user level settings from database
	Public Sub SetUpUserLevel()
		SetUpUserLevelEx() ' Load all user levels

		' User Level loaded event
		Call UserLevel_Loaded()

		' Save the user level to session variable
		SaveUserLevel()
	End Sub

	' Sub to get (all) user level settings from database
	Sub SetUpUserLevelEx()
		Dim RsUser, sSql

		 ' Get the user level definitions
		sSql = "SELECT " & EW_USER_LEVEL_ID_FIELD & ", " & EW_USER_LEVEL_NAME_FIELD & " FROM " & EW_USER_LEVEL_TABLE
		Set RsUser = Conn.Execute(sSql)
		If Not RsUser.Eof Then m_ArUserLevel = RsUser.GetRows
		RsUser.Close
		Set RsUser = Nothing

		 ' Get the user level privileges
		sSql = "SELECT " & EW_USER_LEVEL_PRIV_TABLE_NAME_FIELD & ", " & EW_USER_LEVEL_PRIV_USER_LEVEL_ID_FIELD & ", " & EW_USER_LEVEL_PRIV_PRIV_FIELD & " FROM " & EW_USER_LEVEL_PRIV_TABLE
		Set RsUser = Conn.Execute(sSql)
		If Not RsUser.Eof Then m_ArUserLevelPriv = RsUser.GetRows
		RsUser.Close
		Set RsUser = Nothing
	End Sub

	' Add user permission
	Public Sub AddUserPermission(UserLevelName, TableName, UserPermission)
		Dim UserLevelID, i

		' Get user level id from user name
		UserLevelID = ""
		If IsArray(m_ArUserLevel) Then
			For i = 0 To UBound(m_ArUserLevel, 2)
				If UserLevelName & "" = m_ArUserLevel(1, i) & "" Then
					UserLevelID = m_ArUserLevel(0, i)
					Exit For
				End If
			Next
		End If
		If IsArray(m_ArUserLevelPriv) And UserLevelID <> "" Then
			For i = 0 To UBound(m_ArUserLevelPriv, 2)
				If m_ArUserLevelPriv(0, i) & "" = TableName & "" And _
				   m_ArUserLevelPriv(1, i) & "" = UserLevelID & "" Then
					m_ArUserLevelPriv(2, i) = m_ArUserLevelPriv(2, i) Or UserPermission ' Add permission
					Exit For
				End If
			Next
		End If
	End Sub

	' Delete user permission
	Public Sub DeleteUserPermission(UserLevelName, TableName, UserPermission)
		Dim UserLevelID, i

		' Get user level id from user name
		UserLevelID = ""
		If IsArray(m_ArUserLevel) Then
			For i = 0 To UBound(m_ArUserLevel, 2)
				If UserLevelName & "" = m_ArUserLevel(1, i) & "" Then
					UserLevelID = m_ArUserLevel(0, i)
					Exit For
				End If
			Next
		End If
		If IsArray(m_ArUserLevelPriv) And UserLevelID <> "" Then
			For i = 0 To UBound(m_ArUserLevelPriv, 2)
				If m_ArUserLevelPriv(0, i) & "" = TableName & "" And _
				   m_ArUserLevelPriv(1, i) & "" = UserLevelID & "" Then
					m_ArUserLevelPriv(2, i) = m_ArUserLevelPriv(2, i) And (127-UserPermission) ' Remove permission
					Exit For
				End If
			Next
		End If
	End Sub

	' Load current user level
	Public Sub LoadCurrentUserLevel(Table)
		Call LoadUserLevel()
		SessionUserLevel = CurrentUserLevelPriv(Table)
	End Sub

	' Get current user privilege
	Private Function CurrentUserLevelPriv(TableName)
		If IsLoggedIn() Then
			CurrentUserLevelPriv = 0
			For i = 0 To UBound(m_ArUserLevelID)
				CurrentUserLevelPriv = CurrentUserLevelPriv Or GetUserLevelPrivEx(TableName, m_ArUserLevelID(i))
			Next
		Else
			CurrentUserLevelPriv = 0
		End If
	End Function

	' Get user level ID by user level name
	Public Function GetUserLevelID(UserLevelName)
		GetUserLevelID = -2
		If CStr(UserLevelName) = "Administrator" Then
			GetUserLevelID = -1
		ElseIf UserLevelName <> "" Then
			If IsArray(m_ArUserLevel) Then
				Dim i
				For i = 0 to UBound(m_ArUserLevel, 2)
					If CStr(m_ArUserLevel(1, i)) = CStr(UserLevelName) Then
						GetUserLevelID = m_ArUserLevel(0, i)
						Exit For
					End If
				Next
			End If
		End If
	End Function

	' Add user level (for use with UserLevel_Loading event)
	Sub AddUserLevel(UserLevelName)
		Dim bFound, i, UserLevelID
		If UserLevelName = "" Or IsNull(UserLevelName) Then Exit Sub
		UserLevelID = GetUserLevelID(UserLevelName)
		If Not IsNumeric(UserLevelID) Then Exit Sub
		If UserLevelID < -1 Then Exit Sub
		bFound = False
		If Not IsArray(m_ArUserLevelID) Then
			ReDim m_ArUserLevelID(0)
		Else
			For i = 0 to UBound(m_ArUserLevelID)
				If m_ArUserLevelID(i) = UserLevelID Then
					bFound = True
					Exit For
				End If
			Next
			If Not bFound Then ReDim Preserve m_ArUserLevelID(UBound(m_ArUserLevelID)+1)
		End If
		If Not bFound Then
			m_ArUserLevelID(UBound(m_ArUserLevelID)) = UserLevelID
		End If
	End Sub

	' Delete user level (for use with UserLevel_Loading event)
	Sub DeleteUserLevel(UserLevelName)
		Dim i, j, UserLevelID
		If UserLevelName = "" Or IsNull(UserLevelName) Then Exit Sub
		UserLevelID = GetUserLevelID(UserLevelName)
		If Not IsNumeric(UserLevelID) Then Exit Sub
		If UserLevelID < -1 Then Exit Sub
		If IsArray(m_ArUserLevelID) Then
			For i = 0 to UBound(m_ArUserLevelID)
				If m_ArUserLevelID(i) = UserLevelID Then
					For j = i+1 to UBound(m_ArUserLevelID)
						m_ArUserLevelID(j-1) = m_ArUserLevelID(j)
					Next
					If UBound(m_ArUserLevelID) = 0 Then
						m_ArUserLevelID = ""
					Else
						ReDim Preserve m_ArUserLevelID(UBound(m_ArUserLevelID)-1)
					End If
					Exit Sub
				End If
			Next
		End If
	End Sub

	' User level list
	Function UserLevelList()
		Dim i
		UserLevelList = ""
		If IsArray(m_ArUserLevelID) Then
			For i = 0 to UBound(m_ArUserLevelID)
				If UserLevelList <> "" Then UserLevelList = UserLevelList & ", "
				UserLevelList = UserLevelList & m_ArUserLevelID(i)
			Next
		End If
	End Function

	' User level name list
	Function UserLevelNameList()
		Dim i
		UserLevelNameList = ""
		If IsArray(m_ArUserLevelID) Then
			For i = 0 to UBound(m_ArUserLevelID)
				If UserLevelNameList <> "" Then UserLevelNameList = UserLevelNameList & ", "
				UserLevelNameList = UserLevelNameList & ew_QuotedValue(GetUserLevelName(m_ArUserLevelID(i)), EW_DATATYPE_STRING)
			Next
		End If
	End Function

	' Get user privilege based on table name and user level
	Public Function GetUserLevelPrivEx(TableName, UserLevelID)
		GetUserLevelPrivEx = 0
		If CStr(UserLevelID) = "-1" Then ' System Administrator
			If EW_USER_LEVEL_COMPAT Then
				GetUserLevelPrivEx = 31 ' Use old user level values
			Else
				GetUserLevelPrivEx = 127 ' Use new user level values (separate View/Search)
			End If
		ElseIf UserLevelID >= 0 Then
			If IsArray(m_ArUserLevelPriv) Then
				Dim i
				For i = 0 to UBound(m_ArUserLevelPriv, 2)
					If CStr(m_ArUserLevelPriv(0, i)) = CStr(TableName) And _
						CStr(m_ArUserLevelPriv(1, i)) = CStr(UserLevelID) Then
						GetUserLevelPrivEx = m_ArUserLevelPriv(2, i)
						If IsNull(GetUserLevelPrivEx) Then GetUserLevelPrivEx = 0
						If Not IsNumeric(GetUserLevelPrivEx) Then GetUserLevelPrivEx = 0
						GetUserLevelPrivEx = CLng(GetUserLevelPrivEx)
						Exit For
					End If
				Next
			End If
		End If
	End Function

	' Get current user level name
	Public Function CurrentUserLevelName()
		CurrentUserLevelName = GetUserLevelName(CurrentUserLevelID)
	End Function

	' Get user level name based on user level
	Public Function GetUserLevelName(UserLevelID)
		GetUserLevelName = ""
		If CStr(UserLevelID) = "-1" Then
			GetUserLevelName = "Administrator"
		ElseIf UserLevelID >= 0 Then
			If IsArray(m_ArUserLevel) Then
				Dim i
				For i = 0 to UBound(m_ArUserLevel, 2)
					If CStr(m_ArUserLevel(0, i)) = CStr(UserLevelID) Then
						GetUserLevelName = m_ArUserLevel(1, i)
						Exit For
					End If
				Next
			End If
		End If
	End Function

	' Sub to display all the User Level settings (for debug only)
	Public Sub ShowUserLevelInfo()
		Dim i
		If IsArray(m_ArUserLevel) Then
			Response.Write "User Levels:<br>"
			Response.Write "UserLevelId, UserLevelName<br>"
			For i = 0 To UBound(m_ArUserLevel, 2)
				Response.Write "&nbsp;&nbsp;" & m_ArUserLevel(0, i) & ", " & _
					m_ArUserLevel(1, i) & "<br>"
			Next
		Else
			Response.Write "No User Level definitions." & "<br>"
		End If
		If IsArray(m_ArUserLevelPriv) Then
			Response.Write "User Level Privs:<br>"
			Response.Write "TableName, UserLevelId, UserLevelPriv<br>"
			For i = 0 To UBound(m_ArUserLevelPriv, 2)
				Response.Write "&nbsp;&nbsp;" & m_ArUserLevelPriv(0, i) & ", " & _
					m_ArUserLevelPriv(1, i) & ", " & m_ArUserLevelPriv(2, i) & "<br>"
			Next
		Else
			Response.Write "No User Level privilege settings." & "<br>"
		End If
		Response.Write "CurrentUserLevel = " & CurrentUserLevel & "<br>"
	End Sub

	' Function to check privilege for List page (for menu items)
	Public Function AllowList(TableName)
		AllowList = CBool(CurrentUserLevelPriv(TableName) And EW_ALLOW_LIST)
	End Function

	' Function to check privilege for Add / Detail-Add
	Public Function AllowAdd(TableName)
		AllowAdd = CBool(CurrentUserLevelPriv(TableName) And EW_ALLOW_ADD)
	End Function

	' Check privilege for Edit page (for Detail-Edit)
	Public Function AllowEdit(TableName)
		AllowEdit = CBool(CurrentUserLevelPriv(TableName) And EW_ALLOW_EDIT)
	End Function

	' Check if user password expired
	Public Function IsPasswordExpired()
		IsPasswordExpired = (Session(EW_SESSION_STATUS) = "passwordexpired")
	End Function

	' Check if user is logging in (after changing password)
	Public Function IsLoggingIn()
		IsLoggingIn = (Session(EW_SESSION_STATUS) = "loggingin")
	End Function

	' Check if user is logged in
	Public Function IsLoggedIn()
		IsLoggedIn = (Session(EW_SESSION_STATUS) = "login")
	End Function

	' Check if user is system administrator
	Public Function IsSysAdmin()
		IsSysAdmin = (Session(EW_SESSION_SYS_ADMIN) = 1)
	End Function

	' Check if user is administrator
	Function IsAdmin()
		IsAdmin = IsSysAdmin
		If Not IsAdmin Then
			IsAdmin = (CurrentUserLevelID = -1)
		End If
	End Function

	' Save user level to session
	Public Sub SaveUserLevel()
		Session(EW_SESSION_AR_USER_LEVEL) = m_ArUserLevel
		Session(EW_SESSION_AR_USER_LEVEL_PRIV) = m_ArUserLevelPriv
	End Sub

	' Load user level from session
	Public Sub LoadUserLevel()
		If Not IsArray(Session(EW_SESSION_AR_USER_LEVEL)) Or Not IsArray(Session(EW_SESSION_AR_USER_LEVEL_PRIV)) Then
			Call SetupUserLevel()
			Call SaveUserLevel()
		Else
			m_ArUserLevel = Session(EW_SESSION_AR_USER_LEVEL)
			m_ArUserLevelPriv = Session(EW_SESSION_AR_USER_LEVEL_PRIV)
		End If
	End Sub

	' Function to get user email
	Public Function CurrentUserEmail()
		CurrentUserEmail = CurrentUserInfo("Email")
	End Function

	' Function to get user info
	Public Function CurrentUserInfo(fieldname)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		CurrentUserInfo = Null
	End Function

	' UserID Loading event
	Sub UserID_Loading()

		'Response.Write "UserID Loading: " & CurrentUserID & "<br>"
	End Sub

	' UserID Loaded event
	Sub UserID_Loaded()

		'Response.Write "UserID Loaded: " & UserIDList & "<br>"
	End Sub

	' User Level Loaded event
	Sub UserLevel_Loaded()

		'AddUserPermission <UserLevelName>, <TableName>, <UserPermission>
		'DeleteUserPermission <UserLevelName>, <TableName>, <UserPermission>

	End Sub

	' Table Permission Loading event
	Sub TablePermission_Loading()

		'Response.Write "Table Permission Loading: " & CurrentUserLevelID & "<br>"
	End Sub

	' Table Permission Loaded event
	Sub TablePermission_Loaded()

		'Response.Write "Table Permission Loaded: " & CurrentUserLevel & "<br>"
	End Sub

	' User Custom Validate event
	Function User_CustomValidate(usr, pwd)

		' Return FALSE to continue with default validation after event exits, or return TRUE to skip default validation
		User_CustomValidate = False
	End Function

	' User Validated event
	Sub User_Validated(rs)

		'Session("UserEmail") = rs("Email")
	End Sub

	' User PasswordExpired event
	Sub User_PasswordExpired(rs)

	  'Response.Write "User_PasswordExpired"
	End Sub
End Class

'
'  Advanced Security class (end)
' -------------------------------
' User Profile Class
Class cUserProfile

	Private Profile

	Private KeySep, FldSep

	Private TimeoutTime

	Private MaxRetryCount, RetryLockoutTime

	Private PasswordExpiryTime

	' Initialze
	Private Sub Class_Initialize
		Set Profile = CreateObject("Scripting.Dictionary")
		InitProfile
	End Sub

	' Terminate
	Private Sub Class_Terminate()
		Set Profile = Nothing
	End Sub

	' Initialize profile object
	Sub InitProfile()
		KeySep = EW_USER_PROFILE_KEY_SEPARATOR
		FldSep = EW_USER_PROFILE_FIELD_SEPARATOR

		' Max login retry
		Profile.Add EW_USER_PROFILE_LOGIN_RETRY_COUNT, 0
		Profile.Add EW_USER_PROFILE_LAST_BAD_LOGIN_DATE_TIME, ""
		MaxRetryCount = EW_USER_PROFILE_MAX_RETRY
		RetryLockoutTime = EW_USER_PROFILE_RETRY_LOCKOUT
	End Sub

	Function GetValue(Name)
		If Profile.Exists(Name) Then
			GetValue = Profile.Item(Name)
		Else
			GetValue = ""
		End If
	End Function

	Function SetValue(Name, Value)
		If Profile.Exists(Name) Then
			Profile.Item(Name) = Value
			SetValue = True
		Else
			SetValue = False
		End If
	End Function

	Function LoadProfileFromDatabase(usr)
		Dim sSql, sFilter, rswrk
		If usr = "" Or usr = EW_ADMIN_USER_NAME Then ' Ignore hard code admin
			LoadProfileFromDatabase = False
			Exit Function
		End If
		sFilter = Replace(EW_USER_NAME_FILTER, "%u", ew_AdjustSql(usr))

		' Get Sql from GetSql function in <UseTable> class, <UserTable>info.asp
		sSql = Users.GetSQL(sFilter, "")
		Set rswrk = Conn.Execute(sSql)
		If Not rswrk.Eof Then
			LoadProfile rswrk(EW_USER_PROFILE_FIELD_NAME)
			LoadProfileFromDatabase = True
		Else
			LoadProfileFromDatabase = False
		End If
		Set rswrk = Nothing
	End Function

	Sub LoadProfile(pstr)
		Dim ar, i, name, value
		Dim wrkstr
		wrkstr = CStr(pstr&"")
		If wrkstr <> "" Then
			ar = SplitProfile(wrkstr)
			If IsArray(ar) Then
				For i = 0 to UBound(ar)
					name = ar(i)(0)
					value = ar(i)(1)

'Response.Write "name: " & name & ", value: " & value & "<br>"
					If Profile.Exists(name) Then
						Profile.Item(name) = value
					End If
				Next
			End If
		End If
	End Sub

	Sub WriteProfile()
		Dim p
		For Each p In Profile
			Response.Write "Name: " & p & ", Value: " & Profile.Item(p) & "<br>"
		Next
	End Sub

	Sub ClearProfile()
		Profile.RemoveAll
	End Sub

	Sub SaveProfileToDatabase(usr)
		Dim sSql, sFilter, rswrk
		If usr = "" Or usr = EW_ADMIN_USER_NAME Then Exit Sub ' Ignore hard code admin
		sFilter = Replace(EW_USER_NAME_FILTER, "%u", ew_AdjustSql(usr))

		' Get Sql from GetSql function in <UseTable> class, <UserTable>info.asp
		sSql = Users.GetSQL(sFilter, "")
		Set rswrk = Server.CreateObject("ADODB.Recordset")
		rswrk.CursorLocation = EW_CURSORLOCATION
		rswrk.Open sSql, Conn, 1, EW_RECORDSET_LOCKTYPE
		If Not rswrk.Eof Then
			rswrk(EW_USER_PROFILE_FIELD_NAME) = ProfileToString ' Update profile
			rswrk.Update
		End If
		rswrk.Close
		Set rswrk = Nothing
	End Sub

	Function ProfileToString()
		Dim name, value, sProfileStr
		For Each name In Profile
			value = Profile.Item(name)
			If value <> "" Then
				If sProfileStr <> "" Then sProfileStr = sProfileStr & FldSep
				sProfileStr = sProfileStr & EncodeStr(name) & KeySep & EncodeStr(value)
			End If
		Next
		ProfileToString = sProfileStr
	End Function

	' Split profile
	Private Function SplitProfile(pstr)
		Dim ar, pos1, pos2, field
		pos1 = 1
		pos2 = LocateStr(pos1,pstr,FldSep)
		Do While pos2 > 0
			field = Mid(pstr,pos1,pos2-pos1)
			AddProfileItem ar,field
			pos1 = pos2+1
			pos2 = LocateStr(pos1,pstr,FldSep)
		Loop
		If pos1 < Len(pstr) Then
			AddProfileItem ar,Mid(pstr,pos1)
		End If
		SplitProfile = ar
	End Function

	' Add profile item
	Private Sub AddProfileItem(ar,field)
		Dim pos, name, value
		pos = LocateStr(1,field,KeySep)
		If pos > 0 Then
			name = DecodeStr(Mid(field,1,pos-1))
			value = DecodeStr(Mid(field,pos+1))
			If Not IsArray(ar) Then
				ReDim ar(0)
			Else
				ReDim Preserve ar(UBound(ar)+1)
			End If
			ar(UBound(ar)) = Array(name, value)
		End If
	End Sub

	' Locate string from separator (skip escaped value)
	Private Function LocateStr(pos,str,sep)
		Dim wrkpos
		wrkpos = InStr(pos,str,sep)
		Do While wrkpos > 0
			If wrkpos <= 1 Then
				LocateStr = wrkpos
				Exit Function
			ElseIf Mid(str,wrkpos-1,1) = "\" Then ' Escaped?
				wrkpos = InStr(wrkpos+1,str,sep) ' Continue to next character
			Else
				LocateStr = wrkpos
				Exit Function
			End If
		Loop
		LocateStr = -1 ' Not found
	End Function

	' Encode value ...,...=... as "...\,...\=..."
	Private Function EncodeStr(val)
		EncodeStr = Replace(val & "", "\", "\\")
		EncodeStr = Replace(EncodeStr, KeySep, "\"&KeySep)
		EncodeStr = """" & Replace(EncodeStr, FldSep, "\"&FldSep) & """"
	End Function

	' Decode value "...\,...\=..." to ...,...=...
	Private Function DecodeStr(val)
		DecodeStr = val & ""
		If Left(DecodeStr,1) = """" Then DecodeStr = Mid(DecodeStr,2)
		If Right(DecodeStr,1) = """" Then DecodeStr = Mid(DecodeStr,1,Len(DecodeStr)-1)
		DecodeStr = Replace(DecodeStr,"\"&FldSep,FldSep)
		DecodeStr = Replace(DecodeStr,"\"&KeySep,KeySep)
		DecodeStr = Replace(DecodeStr,"\\","\")
	End Function

	Function ExceedLoginRetry()
		Dim retrycount, dt
		retrycount = Profile.Item(EW_USER_PROFILE_LOGIN_RETRY_COUNT)
		dt = Profile.Item(EW_USER_PROFILE_LAST_BAD_LOGIN_DATE_TIME)
		If CLng(retrycount) >= CLng(MaxRetryCount) Then
			If DateDiff("n", CDate(dt), Now()) < RetryLockoutTime Then
				ExceedLoginRetry = True
			Else
				ExceedLoginRetry = False
				Profile.Item(EW_USER_PROFILE_LOGIN_RETRY_COUNT) = 0
			End If
		Else
			ExceedLoginRetry = False
		End If
	End Function
End Class
%>
<%

' -------------------------------------------
'  Default Request Form Object Class (begin)
'
Class cFormObj
	Dim Index ' Index to handle multiple form elements

	' Class Initialize
	Private Sub Class_Initialize
		Index = 0
	End Sub

	' Get form element name based on index
	Function GetIndexedName(name)
		If Index <= 0 Then
			GetIndexedName = name
		Else
			Dim Pos
			Pos = InStr(name, "_")
			If Pos = 2 Or Pos = 3 Then
				GetIndexedName = Mid(name, 1, Pos-1) & Index & Mid(name, Pos)
			Else
				GetIndexedName = name
			End If
		End If
	End Function

	' Has value for form element
	Function HasValue(name)
		Dim wrkname
		wrkname = GetIndexedName(name)
		HasValue = (Request.Form(wrkname).Count > 0)
	End Function

	' Get value for form element
	Function GetValue(name)
		Dim wrkname
		wrkname = GetIndexedName(name)
		If Request.Form(wrkname).Count > 0 Then

			' Special handling for key_m
			If wrkname = "key_m" Then
				If Request.Form(wrkname).Count = 1 Then
					GetValue = Request.Form(wrkname)
				Else
					Dim i, cnt, ar
					cnt = Request.Form(wrkname).Count
					Redim ar(cnt-1)
					For i = 1 to cnt
						ar(i-1) = Request.Form(wrkname)(i)
					Next
					GetValue = ar
				End If
			Else
				GetValue = Request.Form(wrkname)
			End If
		Else
			GetValue = Null
		End If
	End Function
End Class

'
'  Default Request Form Object Class (end)
' -----------------------------------------

%>
<%

' -------------------------------------
'  Default Upload Object Class (begin)
'
Class cUploadObj
	Dim rawData, separator, lenSeparator, dict
	Dim currentPos, inStrByte, tempValue, mValue, value
	Dim intDict, begPos, endPos
	Dim nameN, isValid, nameValue, midValue
	Dim rawStream
	Dim Index
	Dim hdr, hdrEndPos

	' Class Inialize
	Private Sub Class_Initialize
		Index = 0
		If Request.TotalBytes > 0 Then
			Set rawStream = Server.CreateObject("ADODB.Stream")
			rawStream.Type = 1 'adTypeBinary
			rawStream.Mode = 3 'adModeReadWrite
			rawStream.Open
			rawStream.Write Request.BinaryRead(Request.TotalBytes)
			rawStream.Position = 0
			rawData = rawStream.Read
			separator = MidB(rawData, 1, InStrB(1, rawData, ChrB(13)) - 1)
			lenSeparator = LenB(separator)
			Set dict = Server.CreateObject("Scripting.Dictionary")
			currentPos = 1
			inStrByte = 1
			tempValue = ""
			While inStrByte > 0
				inStrByte = InStrB(currentPos, rawData, separator)
				mValue = inStrByte - currentPos
				If mValue > 1 Then
					value = MidB(rawData, currentPos, mValue)
					Set intDict = Server.CreateObject("Scripting.Dictionary")
					begPos = 1 + InStrB(1, value, ChrB(34))
					endPos = InStrB(begPos + 1, value, ChrB(34))
					nameN = MidB(value, begPos, endPos - begPos)
					isValid = True
					hdrEndPos = InStrB(1, value, ChrB(13) & ChrB(10) & ChrB(13) & ChrB(10))
					hdr = MidB(value, 1, hdrEndPos - 1)
					If InStrB(1, hdr, StringToByte("Content-Type:")) > 1 Or InStrB(1, hdr, StringToByte("filename=")) > 1 Then
						begPos = 1 + InStrB(endPos + 1, value, ChrB(34))
						endPos = InStrB(begPos + 1, value, ChrB(34))
						If endPos > 0 Then
							intDict.Add "FileName", ConvertToText(rawStream, currentPos + begPos - 2, endPos - begPos, MidB(value, begPos, endPos - begPos))
							begPos = 14 + InStrB(endPos + 1, value, StringToByte("Content-Type:"))
							endPos = InStrB(begPos, value, ChrB(13))
							intDict.Add "ContentType", ConvertToText(rawStream, currentPos + begPos - 2, endPos - begPos, MidB(value, begPos, endPos - begPos))
							begPos = endPos + 4
							endPos = LenB(value)
							nameValue = MidB(value, begPos, ((endPos - begPos) - 1))
						Else
							endPos = begPos + 1
							isValid = False
						End If
					Else
						nameValue = ConvertToText(rawStream, currentPos + endPos + 3, mValue - endPos - 4, MidB(value, endPos + 5))
					End If
					If isValid = True Then
						Dim wrkname
						wrkname = ByteToString(nameN)
						If dict.Exists(wrkname) Then
							Set intDict = dict.Item(wrkname)

							' Special handling for key_m, just append to end
							If wrkname = "key_m" Then
								intDict.Item("Value") = intDict.Item("Value") & nameValue
							Else
								If Right(intDict.Item("Value"), 2) = vbCrLf Then
									intDict.Item("Value") = Left(intDict.Item("Value"), Len(intDict.Item("Value"))-2)
								End If
								intDict.Item("Value") = intDict.Item("Value") & ", " & nameValue
							End If
						Else
							intDict.Add "Value", nameValue
							intDict.Add "Name", nameN
							dict.Add wrkname, intDict
						End If
					End If
				End If
				currentPos = lenSeparator + inStrByte
			Wend
			rawStream.Close
			Set rawStream = Nothing
		End If
	End Sub

	' Get form element name based on index
	Function GetIndexedName(name)
		If Index <= 0 Then
			GetIndexedName = name
		Else
			GetIndexedName = Mid(name, 1, 1) & Index & Mid(name, 2)
		End If
	End Function

	' Has value for form element
	Function HasValue(name)
		Dim wrkname
		wrkname = GetIndexedName(name)
		If Not IsObject(dict) Then
			HasValue = False
		Else
			HasValue = dict.Exists(wrkname)
		End If
	End Function

	' Get value for form element
	Function GetValue(name)
		Dim wrkname
		Dim gv
		GetValue = Null ' default return Null
		If IsObject(dict) Then
			wrkname = GetIndexedName(name)
			If dict.Exists(wrkname) Then
				gv = CStr(dict(wrkname).Item("Value"))
				gv = Left(gv, Len(gv)-2)
				GetValue = gv

				' Special handling for key_m
				If wrkname = "key_m" Then
					If InStr(GetValue, vbCrLf) > 0 Then
						GetValue = Split(GetValue, vbCrLf)
					End If
				End If
			End If
		End If
	End Function

	' Get upload file size
	Function GetUploadFileSize(name)
		Dim wrkname
		wrkname = GetIndexedName(name)
		If dict.Exists(wrkname) Then
			GetUploadFileSize = LenB(dict(wrkname).Item("Value"))
		Else
			GetUploadFileSize = 0
		End If
	End Function

	' Get upload file name
	Function GetUploadFileName(name)
		Dim wrkname, temp, tempPos
		wrkname = GetIndexedName(name)
		If dict.Exists(wrkname) Then
			temp = dict(wrkname).Item("FileName")
			tempPos = 1 + InStrRev(temp, "\")
			GetUploadFileName = Mid(temp, tempPos)
		Else
			GetUploadFileName = ""
		End If
	End Function

	' Get file content type
	Function GetUploadFileContentType(name)
		Dim wrkname
		wrkname = GetIndexedName(name)
		If dict.Exists(wrkname) Then
			GetUploadFileContentType = dict(wrkname).Item("ContentType")
		Else
			GetUploadFileContentType = ""
		End If
	End Function

	' Get upload file data
	Function GetUploadFileData(name)
		Dim wrkname
		wrkname = GetIndexedName(name)
		If dict.Exists(wrkname) Then
			GetUploadFileData = dict(wrkname).Item("Value")
		Else
			GetUploadFileData = Null
		End If
	End Function

	' Get file image width
	Function GetUploadImageWidth(name)
		Dim wrkname
		wrkname = GetIndexedName(name)
		Dim ImageHeight
		Call GetImageDimension(GetUploadFileData(name), GetUploadImageWidth, ImageHeight)
	End Function

	' Get file image height
	Function GetUploadImageHeight(name)
		Dim wrkname
		wrkname = GetIndexedName(name)
		Dim ImageWidth
		Call GetImageDimension(GetUploadFileData(name), ImageWidth, GetUploadImageHeight)
	End Function

	' Convert length
	Private Function ConvertLength(b)
		ConvertLength = CLng(AscB(LeftB(b, 1)) + (AscB(RightB(b, 1)) * 256))
	End Function

	' Convert length 2
	Private Function ConvertLength2(b)
		ConvertLength2 = CLng(AscB(RightB(b, 1)) + (AscB(LeftB(b, 1)) * 256))
	End Function

	' Get image dimension
	Sub GetImageDimension(img, wd, ht)
		Dim sPNGHeader, sGIFHeader, sBMPHeader, sJPGHeader, sHeader, sImgType
		sImgType = "(unknown)"

		' Image headers, do not changed
		sPNGHeader = ChrB(137) & ChrB(80) & ChrB(78)
		sGIFHeader = ChrB(71) & ChrB(73) & ChrB(70)
		sBMPHeader = ChrB(66) & ChrB(77)
		sJPGHeader = ChrB(255) & ChrB(216) & ChrB(255)
		sHeader = MidB(img, 1, 3)

		' Handle GIF
		If sHeader = sGIFHeader Then
			sImgType = "GIF"
			wd = ConvertLength(MidB(img, 7, 2))
			ht = ConvertLength(MidB(img, 9, 2))

		' Handle BMP
		ElseIf LeftB(sHeader, 2) = sBMPHeader Then
			sImgType = "BMP"
			wd = ConvertLength(MidB(img, 19, 2))
			ht = ConvertLength(MidB(img, 23, 2))

		' Handle PNG
		ElseIf sHeader = sPNGHeader Then
			sImgType = "PNG"
			wd = ConvertLength2(MidB(img, 19, 2))
			ht = ConvertLength2(MidB(img, 23, 2))

		' Handle JPG
		Else
			Dim size, markersize, pos, bEndLoop
			size = LenB(img)
			pos = InStrB(img, sJPGHeader)
			If pos <= 0 Then
				wd = -1
				ht = -1
				Exit Sub
			End If
			sImgType = "JPG"
			pos = pos + 2
			bEndLoop = False
			Do While Not bEndLoop and pos < size
				Do While AscB(MidB(img, pos, 1)) = 255 and pos < size
					pos = pos + 1
				Loop
				If AscB(MidB(img, pos, 1)) < 192 or AscB(MidB(img, pos, 1)) > 195 Then
					markersize = ConvertLength2(MidB(img, pos+1, 2))
					pos = pos + markersize + 1
				Else
					bEndLoop = True
				End If
			Loop
			If Not bEndLoop Then
				wd = -1
				ht = -1
			Else
				wd = ConvertLength2(MidB(img, pos+6, 2))
				ht = ConvertLength2(MidB(img, pos+4, 2))
			End If
		End If
	End Sub

	' Convert string to byte
	Function StringToByte(toConv)
		Dim i, tempChar
		For i = 1 to Len(toConv)
			tempChar = Mid(toConv, i, 1)
			StringToByte = StringToByte & ChrB(AscB(tempChar))
		Next
	End Function

	' Convert byte to string
	Private Function ByteToString(ToConv)
		Dim i
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		For i = 1 to LenB(ToConv)
			ByteToString = ByteToString & Chr(AscB(MidB(ToConv,i,1)))
		Next
	End Function

	' Convert to text
	Function ConvertToText(objStream, iStart, iLength, binData)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If EW_UPLOAD_CHARSET <> "" Then
			Dim tmpStream
			Set tmpStream = Server.CreateObject("ADODB.Stream")
			tmpStream.Type = 1 'adTypeBinary
			tmpStream.Mode = 3 'adModeReadWrite
			tmpStream.Open
			objStream.Position = iStart
			objStream.CopyTo tmpStream, iLength
			tmpStream.Position = 0
			tmpStream.Type = 2 'adTypeText
			tmpStream.Charset = EW_UPLOAD_CHARSET
			ConvertToText = tmpStream.ReadText
			tmpStream.Close
			Set tmpStream = Nothing
		Else
			ConvertToText = ByteToString(binData)
		End If
		ConvertToText = Trim(ConvertToText & "")
	End Function

	' Class terminate
	Private Sub Class_Terminate

		' Dispose dictionary
		If IsObject(intDict) Then
			intDict.RemoveAll
			Set intDict = Nothing
		End If
		If IsObject(dict) Then
			dict.RemoveAll
			Set dict = Nothing
		End If
	End Sub
End Class

'
'  Default Upload Object Class (end)
' -----------------------------------

%>
<%

' --------------------------
'  Common functions (begin)
'
' Write HTTP header
Sub ew_Header(cache, charset)
	Dim export
	export = Request.QueryString("export") & ""
	If (cache) Or (Not cache And ew_IsHttps() And export <> "" And export <> "print") Then ' Allow cache
		Response.AddHeader "Cache-Control", "private, must-revalidate" ' // HTTP/1.1
	Else ' No cache
		Response.AddHeader "Cache-Control", "private, no-cache, no-store, must-revalidate" ' HTTP/1.1
		Response.AddHeader "Cache-Control", "post-check=0, pre-check=0"
		Response.AddHeader "Pragma", "no-cache" ' HTTP/1.0
	End If
	If charset <> "" Then
		Response.AddHeader "Content-Type", "text/html; charset=" & charset ' Charset
	End If
End Sub

' Connect to database
Sub ew_Connect()

	' Open connection to the database
	Set Conn = Server.CreateObject("ADODB.Connection")

	' Database connecting event
	Call Database_Connecting(EW_DB_CONNECTION_STRING)
	Conn.Open EW_DB_CONNECTION_STRING

	' Database connected event
	Call Database_Connected(Conn)
End Sub

' Database Connecting event
Sub Database_Connecting(Connstr)

	'Response.Write "Database Connecting"
End Sub

' Database Connected event
Sub Database_Connected(Conn)

	' Example:
	' Conn.Execute("Your SQL")

End Sub

' Check if allow add/delete row
Function ew_AllowAddDeleteRow()
	Dim ua
	ua = ew_UserAgent()
	If UBound(ua) >= 1 Then
		ew_AllowAddDeleteRow  = (ua(0) <> "IE" Or ua(1) > 5)
	Else
		ew_AllowAddDeleteRow = False
	End If
End Function

' Get browser type and version
Function ew_UserAgent()
	Dim RegEx, objMatches
	Dim useragent, browser, browser_version, ver, i
	useragent = Request.ServerVariables("HTTP_USER_AGENT")

	' Default
	browser = "Other"
	browser_version = 0
	Set RegEx = New RegExp
	RegEx.Global = True

	' MSIE
	RegEx.Pattern = "MSIE ([0-9].[0-9]{1,2})"
	Set objMatches = RegEx.Execute(useragent)
	If objMatches.Count > 0 Then
		browser = "IE"
		browser_version = objMatches(0).SubMatches(0)

	' Firefox
	Else
		RegEx.Pattern = "Firefox/([0-9\.]+)"
		Set objMatches = RegEx.Execute(useragent)
		If objMatches.Count > 0 Then
			browser = "Firefox"
			browser_version = objMatches(0).SubMatches(0)

		' Opera
		Else
			RegEx.Pattern = "Opera/([0-9].[0-9]{1,2})"
			Set objMatches = RegEx.Execute(useragent)
			If objMatches.Count > 0 Then
				browser = "Opera"
				browser_version = objMatches(0).SubMatches(0)

			' Chrome
			Else
				RegEx.Pattern = "Chrome/([0-9\.]+)"
				Set objMatches = RegEx.Execute(useragent)
				If objMatches.Count > 0 Then
					browser = "Chrome"
					browser_version = objMatches(0).SubMatches(0)

				' Safari
				Else
					RegEx.Pattern = "Safari/"
					Set objMatches = RegEx.Execute(useragent)
					If objMatches.Count > 0 Then
						RegEx.Pattern = "Version/([0-9\.]+)"
						Set objMatches = RegEx.Execute(useragent)
						If objMatches.Count > 0 Then
							browser = "Safari"
							browser_version = objMatches(0).SubMatches(0)
						End If
					End If
				End If
			End If
		End If
	End If
	ver = Split(browser_version, ".")
	ReDim ua(UBound(ver)+1)
	ua(0) = browser
	For i = 0 to UBound(ver)
		ua(i+1) = ver(i)
	Next
	ew_UserAgent = ua
End Function

' Append like operator
Function ew_Like(pat)
	ew_Like = " LIKE " & pat
End Function

' Return multi-value search sql
Function ew_GetMultiSearchSql(Fld, FldVal)
	Dim arVal, i, sVal, sSql, sWrk
	sWrk = ""
	arVal = Split(FldVal, ",")
	For i = 0 to UBound(arVal)
		sVal = Trim(arVal(i))
		If UBound(arVal) = 0 Or EW_SEARCH_MULTI_VALUE_OPTION = 3 Then
			sSql = Fld.FldExpression & " = '" & ew_AdjustSql(sVal) & "' OR " & ew_GetMultiSearchSqlPart(Fld, sVal)
		Else
			sSql = ew_GetMultiSearchSqlPart(Fld, sVal)
		End If
		If sWrk <> "" Then
			If EW_SEARCH_MULTI_VALUE_OPTION = 2 Then
				sWrk = sWrk & " AND "
			ElseIf EW_SEARCH_MULTI_VALUE_OPTION = 3 Then
				sWrk = sWrk & " OR "
			End If
		End If
		sWrk = sWrk & "(" & sSql & ")"
	Next
	ew_GetMultiSearchSql = sWrk
End Function

' Get multi search sql part
Function ew_GetMultiSearchSqlPart(Fld, FldVal)
	ew_GetMultiSearchSqlPart = Fld.FldExpression & ew_Like("'" & ew_AdjustSql(FldVal) & ", %'") & " OR " & _
		Fld.FldExpression & ew_Like("'%, " & ew_AdjustSql(FldVal) & ",%'") & " OR " & _
		Fld.FldExpression & ew_Like("'%, " & ew_AdjustSql(FldVal) & "'")
End Function

' Get search sql
Function ew_GetSearchSql(Fld, FldVal, FldOpr, FldCond, FldVal2, FldOpr2)
	Dim IsValidValue
	ew_GetSearchSql = ""
	Dim sFldExpression, lFldDataType
	sFldExpression = ew_IIf(Fld.FldIsVirtual And Not Fld.FldForceSelection, Fld.FldVirtualExpression, Fld.FldExpression)
	lFldDataType = Fld.FldDataType
	If Fld.FldIsVirtual And Not Fld.FldForceSelection Then lFldDataType = EW_DATATYPE_STRING
	If FldOpr = "BETWEEN" Then
		IsValidValue = (lFldDataType <> EW_DATATYPE_NUMBER) Or _
			(lFldDataType = EW_DATATYPE_NUMBER And IsNumeric(FldVal) And IsNumeric(FldVal2))
		If FldVal <> "" And FldVal2 <> "" And IsValidValue Then
			ew_GetSearchSql = sFldExpression & " BETWEEN " & ew_QuotedValue(FldVal, lFldDataType) & _
				" AND " & ew_QuotedValue(FldVal2, lFldDataType)
		End If
	ElseIf FldVal = EW_NULL_VALUE Or FldOpr = "IS NULL" Then
		ew_GetSearchSql = sFldExpression & " IS NULL"
	ElseIf FldVal = EW_NOT_NULL_VALUE Or FldOpr = "IS NOT NULL" Then
		ew_GetSearchSql = sFldExpression & " IS NOT NULL"
	Else
		IsValidValue = (lFldDataType <> EW_DATATYPE_NUMBER) Or _
			(lFldDataType = EW_DATATYPE_NUMBER And IsNumeric(FldVal))
		If FldVal <> "" And IsValidValue And ew_IsValidOpr(FldOpr, lFldDataType) Then
			ew_GetSearchSql = sFldExpression & ew_SearchString(FldOpr, FldVal, lFldDataType)
			If Fld.FldDataType = EW_DATATYPE_BOOLEAN And FldVal = Fld.FalseValue And FldOpr = "=" Then
				ew_GetSearchSql = "(" & ew_GetSearchSql & " OR " & sFldExpression & " IS NULL)"
			End If
		End If
		IsValidValue = (lFldDataType <> EW_DATATYPE_NUMBER) Or _
			(lFldDataType = EW_DATATYPE_NUMBER And IsNumeric(FldVal2))
		If FldVal2 <> "" And IsValidValue And ew_IsValidOpr(FldOpr2, lFldDataType) Then
			Dim sSql2
			sSql2 = sFldExpression & ew_SearchString(FldOpr2, FldVal2, lFldDataType)
			If Fld.FldDataType = EW_DATATYPE_BOOLEAN And FldVal2 = Fld.FalseValue And FldOpr2 = "=" Then
				sSql2 = "(" & sSql2 & " OR " & sFldExpression & " IS NULL)"
			End If
			If ew_GetSearchSql <> "" Then
				ew_GetSearchSql = "(" & ew_GetSearchSql & " " & ew_IIf(FldCond = "OR", "OR", "AND") & " " & sSql2 & ")"
			Else
				ew_GetSearchSql = sSql2
			End If
		End If
	End If
End Function

' Return search string
Function ew_SearchString(FldOpr, FldVal, FldType)
	If FldOpr = "LIKE" Then
		ew_SearchString = ew_Like(ew_QuotedValue("%" & FldVal & "%", FldType))
	ElseIf FldOpr = "NOT LIKE" Then
		ew_SearchString = " NOT " & ew_Like(ew_QuotedValue("%" & FldVal & "%", FldType))
	ElseIf FldOpr = "STARTS WITH" Then
		ew_SearchString = ew_Like(ew_QuotedValue(FldVal & "%", FldType))
	Else
		ew_SearchString = " " & FldOpr & " " & ew_QuotedValue(FldVal, FldType)
	End If
End Function

' Check if valid operator
Function ew_IsValidOpr(Opr, FldType)
	ew_IsValidOpr = (Opr = "=" Or Opr = "<" Or Opr = "<=" Or _
		Opr = ">" Or Opr = ">=" Or Opr = "<>")
	If FldType = EW_DATATYPE_STRING Or FldType = EW_DATATYPE_MEMO Then
		ew_IsValidOpr = ew_IsValidOpr Or Opr = "LIKE" Or Opr = "NOT LIKE" Or Opr = "STARTS WITH"
	End If
End Function

' Quoted name for table/field
Function ew_QuotedName(Name)
	ew_QuotedName = EW_DB_QUOTE_START & Replace(Name, EW_DB_QUOTE_END, EW_DB_QUOTE_END & EW_DB_QUOTE_END) & EW_DB_QUOTE_END
End Function

' Quoted value for field type
Function ew_QuotedValue(Value, FldType) 
	Select Case FldType
	Case EW_DATATYPE_STRING, EW_DATATYPE_MEMO
		ew_QuotedValue = "'" & ew_AdjustSql(Value) & "'"
	Case EW_DATATYPE_GUID
		If EW_IS_MSACCESS Then
			ew_QuotedValue = "{guid " & ew_AdjustSql(Value) & "}"
		Else
			ew_QuotedValue = "'" & ew_AdjustSql(Value) & "'"
		End If
	Case EW_DATATYPE_DATE
		If EW_IS_MSACCESS Then
			ew_QuotedValue = "#" & ew_AdjustSql(Value) & "#"
		ElseIf EW_IS_ORACLE Then
			ew_QuotedValue = "TO_DATE('" & ew_AdjustSql(Value) & "', 'YYYY/MM/DD HH24:MI:SS')"
		Else
			ew_QuotedValue = "'" & ew_AdjustSql(Value) & "'"
		End If
	Case EW_DATATYPE_BOOLEAN
		ew_QuotedValue = Value
	Case Else
		ew_QuotedValue = Value
	End Select
End Function

' Pad zeros before number
Function ew_ZeroPad(m, t)
	ew_ZeroPad = String(t - Len(m), "0") & m
End Function

' IIf function
Function ew_IIf(cond, v1, v2)
	If Not EW_DEBUG_ENABLED Then On Error Resume Next
	If CBool(cond) Then
		ew_IIf = v1
	Else
		ew_IIf = v2
	End If
End Function

' Convert different data type value
Function ew_Conv(v, t)
	Select Case t

	' adBigInt/adUnsignedBigInt
	Case 20, 21
		If IsNull(v) Then
			ew_Conv = Null
		Else
			ew_Conv = CLng(v)
		End If

	' adSmallInt/adInteger/adTinyInt/adUnsignedTinyInt/adUnsignedSmallInt/adUnsignedInt/adBinary
	Case 2, 3, 16, 17, 18, 19, 128
		If IsNull(v) Then
			ew_Conv = Null
		Else
			ew_Conv = CLng(v)
		End If

	' adSingle
	Case 4
		If IsNull(v) Then
			ew_Conv = Null
		Else
			ew_Conv = CSng(v)
		End If

	' adDouble/adCurrency/adNumeric/adVarNumeric
	Case 5, 6, 131, 139
		If IsNull(v) Then
			ew_Conv = Null
		Else
			ew_Conv = CDbl(v)
		End If
	Case Else
		ew_Conv = v
	End Select
End Function

' Function for debug
Sub ew_Trace(pfx, aMsg)
	If Not EW_DEBUG_ENABLED Then On Error Resume Next
	Dim fso, ts
	Dim sFolder, sFn
	sFolder = EW_AUDIT_TRAIL_PATH
	sFn = pfx & ".txt"
	Set fso = Server.Createobject("Scripting.FileSystemObject")
	Set ts = fso.OpenTextFile(ew_UploadPathEx(True, sFolder) & sFn, 8, True)
	ts.writeline(Date & vbTab & Time & vbTab & aMsg)
	ts.Close
	Set ts = Nothing
	Set fso = Nothing
End Sub

' Display elapsed time (in seconds)
Function ew_CalcElapsedTime(tm)
	Dim endTimer
	endTimer = Timer
	ew_CalcElapsedTime = "<div>page processing time: " & FormatNumber((endTimer-tm),2) & " seconds</div>"
End Function

' Function to compare values with special handling for null values
Function ew_CompareValue(v1, v2)
	If IsNull(v1) And IsNull(v2) Then
		ew_CompareValue = True
	ElseIf IsNull(v1) Or IsNull(v2) Then
		ew_CompareValue = False
	ElseIf VarType(v1) = 14 Or VarType(v2) = 14 Then
		ew_CompareValue = (CDbl(v1) = CDbl(v2))
	Else
		ew_CompareValue = (v1 = v2)
	End If
End Function

' Check if boolean value is TRUE
Function ew_ConvertToBool(value)
	ew_ConvertToBool = (value & "" = "1" Or LCase(value & "") = "true" Or LCase(value & "") = "y" Or LCase(value & "") = "t")
End Function

' Add message
Sub ew_AddMessage(msg, msgtoadd)
	If msgtoadd <> "" Then
		If msg <> "" Then
			msg = msg & "<br>"
		End If
		msg = msg & msgtoadd
	End If
End Sub

' Add filter
Sub ew_AddFilter(filter, newfilter)
	If Trim(newfilter) = "" Then Exit Sub
	If Trim(filter) <> "" Then
		filter = "(" & filter & ") AND (" & newfilter & ")"
	Else
		filter = newfilter
	End If
End Sub

' Adjust sql for special characters
Function ew_AdjustSql(str)
	Dim sWrk
	sWrk = Trim(str & "")
	sWrk = Replace(sWrk, "'", "''") ' Adjust for Single Quote
	sWrk = Replace(sWrk, "[", "[[]") ' Adjust for Open Square Bracket
	ew_AdjustSql = sWrk
End Function

' Build select sql based on different sql part
Function ew_BuildSelectSql(sSelect, sWhere, sGroupBy, sHaving, sOrderBy, sFilter, sSort)
	Dim sSql, sDbWhere, sDbOrderBy
	sDbWhere = sWhere
	Call ew_AddFilter(sDbWhere, sFilter)
	sDbOrderBy = sOrderBy
	If sSort <> "" Then
		sDbOrderBy = sSort
	End If
	sSql = sSelect
	If sDbWhere <> "" Then
		sSql = sSql & " WHERE " & sDbWhere
	End If
	If sGroupBy <> "" Then
		sSql = sSql & " GROUP BY " & sGroupBy
	End If
	If sHaving <> "" Then
		sSql = sSql & " HAVING " & sHaving
	End If
	If sDbOrderBy <> "" Then
		sSql = sSql & " ORDER BY " & sDbOrderBy
	End If
	ew_BuildSelectSql = sSql
End Function

' Load recordset
Function ew_LoadRecordset(SQL)
	On Error Resume Next
	Err.Clear
	Dim RsSet
	Set RsSet = Server.CreateObject("ADODB.Recordset")
	RsSet.CursorLocation = EW_CURSORLOCATION
	RsSet.Open SQL, Conn, 1, EW_RECORDSET_LOCKTYPE
	If Err.Number <> 0 Then
		Response.Write "Load recordset error. SQL: '" & SQL & "'. Description: " & Err.Description
		Response.End
	Else
		Set ew_LoadRecordset = RsSet
	End If
End Function

' Load row
Function ew_LoadRow(SQL)
	On Error Resume Next
	Err.Clear
	Dim RsRow
	Set RsRow = Server.CreateObject("ADODB.Recordset")
	RsRow.Open SQL, Conn
	If Err.Number <> 0 Then
		Response.Write "Load row error. SQL: '" & SQL & "'. Description: " & Err.Description
		Response.End
	Else
		Set ew_LoadRow = RsRow
	End If
End Function

' Note: Object "Conn" is required
' Return sql scalar value
Function ew_ExecuteScalar(SQL)
	On Error Resume Next
	Err.Clear
	ew_ExecuteScalar = Null
	If Trim(SQL & "") = "" Then Exit Function
	Dim RsExec
	Set RsExec = Conn.Execute(SQL)
	If Err.Number <> 0 Then
		Response.Write "Execute scalar error. SQL: '" & SQL & "'. Description: " & Err.Description
		Response.End
	Else
		If Not RsExec.Eof Then ew_ExecuteScalar = RsExec(0)
	End If
	RsExec.Close
	Set RsExec = Nothing
End Function

' Clone recordset
Function ew_CloneRs(RsOld)
	Dim Stream
	Dim RsClone

	' Save the recordset to the stream object
	Set Stream = Server.CreateObject("ADODB.Stream")
	RsOld.Save Stream

	' Open the stream object into a new recordset
	Set RsClone = Server.CreateObject("ADODB.Recordset")
	RsClone.Open Stream, , , 2

	' Return the cloned recordset
	Set ew_CloneRs = RsClone

	' Release the reference
	Set RsClone = Nothing
End Function

' Function to dynamically include a file
Function ew_Include(fn)
	On Error Resume Next
	Dim sIncludeText
	sIncludeText = ew_LoadFile(fn)
	If sIncludeText <> "" Then
		sIncludeText = Replace(sIncludeText, "<" & "%", "")
		sIncludeText = Replace(sIncludeText, "%" & ">", "")
		Execute sIncludeText
		ew_Include = True
	Else
		ew_Include = False
	End If
End Function

' Function to Load a Text File
Function ew_LoadTxt(fn)
	Dim fso, fobj

	' Get text file content
	ew_LoadTxt = ""
	If Trim(fn) <> "" Then
		Set fso = Server.CreateObject("Scripting.FileSystemObject")
		If fso.FileExists(Server.MapPath(fn)) Then
			Set fobj = fso.OpenTextFile(Server.MapPath(fn))
			ew_LoadTxt = fobj.ReadAll ' Read all Content
			fobj.Close
			Set fobj = Nothing
		End If
		Set fso = Nothing
	End If
End Function

' Load file content (both ASCII and UTF-8)
Function ew_LoadFile(FileName)
	On Error Resume Next
	Dim fso, FilePath
	ew_LoadFile = ""
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	If Trim(FileName) <> "" Then
		If fso.FileExists(FileName) Then
			FilePath = FileName
		Else
			FilePath = Server.MapPath(FileName)
		End If
		If fso.FileExists(FilePath) Then
			If ew_GetFileCharset(FilePath) = "UTF-8" Then
				ew_LoadFile = ew_LoadUTF8File(FilePath)
			Else
				Dim iFile, iData
				Set iFile = fso.GetFile(FilePath)
				Set iData = iFile.OpenAsTextStream
				ew_LoadFile = iData.ReadAll
				iData.Close
				Set iData = Nothing
				Set iFile = Nothing
			End If
		End If
	End If
	Set fso = Nothing
End Function

' Open UTF8 file
Function ew_LoadUTF8File(FilePath)
	On Error Resume Next
	Dim objStream
	Set objStream = Server.CreateObject("ADODB.Stream")
	With objStream
		.Type = 2
		.Mode = 3
		.Open
		.CharSet = "UTF-8"
		.LoadFromFile FilePath
		ew_LoadUTF8File = .ReadText
		.Close
	End With
End Function

' Get file charset (UTF-8 and UNICODE)
Function ew_GetFileCharset(FilePath)
	On Error Resume Next
	Dim objStream, LoadBytes
	Set objStream = Server.CreateObject("ADODB.Stream")
	With objStream
		.Type = 1
		.Mode = 3
		.Open
		.LoadFromFile FilePath
		LoadBytes = .Read(3) ' Get first 3 bytes as BOM
		.Close
	End With
	Set objStream = Nothing
	Dim FileCharset, strFileHead

	' Get hex values
	strFileHead = ew_BinToHex(LoadBytes)

	' UTF-8
	If strFileHead = "EFBBBF" Then
		ew_GetFileCharset = "UTF-8" ' UTF-8
	Else
		ew_GetFileCharset = "" ' Non UTF-8
	End If
End Function

' Get hex values
Function ew_BinToHex(vStream)
	Dim reVal, i
	reVal = 0
	For i = 1 To LenB(vStream)
		reVal = reVal * 256 + AscB(MidB(vStream, i, 1))
	Next
	ew_BinToHex = Hex(reVal)
End Function

' Write Audit Trail (login/logout)
Sub ew_WriteAuditTrailOnLogInOut(user, logtype)
	If Not EW_DEBUG_ENABLED Then On Error Resume Next
	Dim table, sKey
	table = ew_CurrentUserIP()
	sKey = ""

	' Write Audit Trail
	Dim filePfx, curDateTime, id, action, field, keyvalue, oldvalue, newvalue
	Dim i
	filePfx = "log"
	curDateTime = ew_StdCurrentDateTime()
	id = Request.ServerVariables("SCRIPT_NAME")
	action = logtype
	Call ew_WriteAuditTrail(filePfx, curDateTime, id, user, action, table, field, keyvalue, oldvalue, newvalue)
End Sub

' Write Audit Trail (insert/update/delete)
Sub ew_WriteAuditTrail(pfx, curDateTime, script, user, action, table, field, keyvalue, oldvalue, newvalue)
	If Not EW_DEBUG_ENABLED Then On Error Resume Next
	Dim fso, ts, sMsg, sFn, sFolder
	Dim bWriteHeader, sHeader
	Dim userwrk
	userwrk = user
	If userwrk = "" Then userwrk = "-1" ' assume Administrator if no user
	If Not EW_AUDIT_TRAIL_TO_DATABASE Then

		' Write audit trail to log file
		sHeader = "date/time" & vbTab & _
			"script" & vbTab & _
			"user" & vbTab & _
			"action" & vbTab & _
			"table" & vbTab & _
			"field" & vbTab & _
			"key value" & vbTab & _
			"old value" & vbTab & _
			"new value"
		sMsg = curDateTime & vbTab & _
			script & vbTab & _
			userwrk & vbTab & _
			action & vbTab & _
			table & vbTab & _
			field & vbTab & _
			keyvalue & vbTab & _
			oldvalue & vbTab & _
			newvalue
		sFolder = EW_AUDIT_TRAIL_PATH
		sFn = pfx & "_" & ew_ZeroPad(Year(Date), 4) & ew_ZeroPad(Month(Date), 2) & ew_ZeroPad(Day(Date), 2) & ".txt"
		Set fso = Server.Createobject("Scripting.FileSystemObject")
		bWriteHeader = Not fso.FileExists(ew_UploadPathEx(True, sFolder) & sFn)
		Set ts = fso.OpenTextFile(ew_UploadPathEx(True, sFolder) & sFn, 8, True)
		If bWriteHeader Then
			ts.writeline(sHeader)
		End If
		ts.writeline(sMsg)
		ts.Close
		Set ts = Nothing
		Set fso = Nothing
	Else
		Dim sAuditSql
		sAuditSql = "INSERT INTO " & ew_QuotedName(EW_AUDIT_TRAIL_TABLE_NAME) & _
			" (" & ew_QuotedName(EW_AUDIT_TRAIL_FIELD_NAME_DATETIME) & ", " & _
			ew_QuotedName(EW_AUDIT_TRAIL_FIELD_NAME_SCRIPT) & ", " & _
			ew_QuotedName(EW_AUDIT_TRAIL_FIELD_NAME_USER) & ", " & _
			ew_QuotedName(EW_AUDIT_TRAIL_FIELD_NAME_ACTION) & ", " & _
			ew_QuotedName(EW_AUDIT_TRAIL_FIELD_NAME_TABLE) & ", " & _
			ew_QuotedName(EW_AUDIT_TRAIL_FIELD_NAME_FIELD) & ", " & _
			ew_QuotedName(EW_AUDIT_TRAIL_FIELD_NAME_KEYVALUE) & ", " & _
			ew_QuotedName(EW_AUDIT_TRAIL_FIELD_NAME_OLDVALUE) & ", " & _
			ew_QuotedName(EW_AUDIT_TRAIL_FIELD_NAME_NEWVALUE) & ") " & _
			" VALUES (" & _
			ew_QuotedValue(curDateTime, EW_DATATYPE_DATE) & ", " & _
			ew_QuotedValue(script, EW_DATATYPE_STRING) & ", " & _
			ew_QuotedValue(userwrk, EW_DATATYPE_STRING) & ", " & _
			ew_QuotedValue(action, EW_DATATYPE_STRING) & ", " & _
			ew_QuotedValue(table, EW_DATATYPE_STRING) & ", " & _
			ew_QuotedValue(field, EW_DATATYPE_STRING) & ", " & _
			ew_QuotedValue(keyvalue, EW_DATATYPE_STRING) & ", " & _
			ew_QuotedValue(oldvalue, EW_DATATYPE_STRING) & ", " & _
			ew_QuotedValue(newvalue, EW_DATATYPE_STRING) & ")"

		' Response.Write sAuditSql ' uncomment to debug
		Conn.Execute(sAuditSql)
	End If
End Sub

' Function to check date format "yyyy-MM-dd HH:mm:ss.fffffff zzz"
Function ew_IsDate(ADate)
	If ADate & "" = "" Then
		ew_IsDate = False
	Else
		ew_IsDate = IsDate(ew_GetDateTimePart(ADate))
	End If
End Function

' Function to get DateTime part (remove ".fffffff zzz" from format "yyyy-MM-dd HH:mm:ss.fffffff zzz")
Function ew_GetDateTimePart(ADate)
	If IsNull(ADate) Then
		ew_GetDateTimePart = ADate
	ElseIf InStrRev(ADate,".") > 0 And InStr(ADate,":") > 0 Then
		ew_GetDateTimePart = Mid(ADate, 1, InStrRev(ADate,".")-1)
		If Not IsDate(ew_GetDateTimePart) Or InStr(ew_GetDateTimePart,":") <= 0 Then ew_GetDateTimePart = ADate
	Else
		ew_GetDateTimePart = ADate
	End If
End Function

'-------------------------------------------------------------------------------
' Functions for default date format
' ANamedFormat = 0-8, where 0-4 same as VBScript
' 5 = "yyyymmdd"
' 6 = "mmddyyyy"
' 7 = "ddmmyyyy"
' 8 = Short Date + Short Time
' 9 = "yyyymmdd HH:MM:SS"
' 10 = "mmddyyyy HH:MM:SS"
' 11 = "ddmmyyyy HH:MM:SS"
' 12 - Short Date - 2 digit year (yy/mm/dd)
' 13 - Short Date - 2 digit year (mm/dd/yy)
' 14 - Short Date - 2 digit year (dd/mm/yy)
' 15 - Short Date - 2 digit year (yy/mm/dd) + Short Time (hh:mm:ss)
' 16 - Short Date (mm/dd/yyyy) + Short Time (hh:mm:ss)
' 17 - Short Date (dd/mm/yyyy) + Short Time (hh:mm:ss)
' 99 - "HH:MM:SS"
' Format date time based on format type
Function ew_FormatDateTime(ADate, ANamedFormat)
	Dim sDate
	sDate = ew_GetDateTimePart(ADate)
	If IsDate(sDate) Then
		If ANamedFormat >= 0 And ANamedFormat <= 4 Then
			ew_FormatDateTime = FormatDateTime(sDate, ANamedFormat)
		ElseIf ANamedFormat = 5 Or ANamedFormat = 9 Then
			ew_FormatDateTime = Year(sDate) & EW_DATE_SEPARATOR & Month(sDate) & EW_DATE_SEPARATOR & Day(sDate)
		ElseIf ANamedFormat = 6 Or ANamedFormat = 10 Then
			ew_FormatDateTime = Month(sDate) & EW_DATE_SEPARATOR & Day(sDate) & EW_DATE_SEPARATOR & Year(sDate)
		ElseIf ANamedFormat = 7 Or ANamedFormat = 11 Then
			ew_FormatDateTime = Day(sDate) & EW_DATE_SEPARATOR & Month(sDate) & EW_DATE_SEPARATOR & Year(sDate)
		ElseIf ANamedFormat = 8 Then
			ew_FormatDateTime = FormatDateTime(sDate, 2)
			If Hour(sDate) <> 0 Or Minute(sDate) <> 0 Or Second(sDate) <> 0 Then
				ew_FormatDateTime = ew_FormatDateTime & " " & FormatDateTime(sDate, 4) & ":" & ew_ZeroPad(Second(sDate), 2)
			End If
		ElseIf ANamedFormat = 99 Then
			ew_FormatDateTime = ew_ZeroPad(Hour(sDate), 2) & ":" & ew_ZeroPad(Minute(sDate), 2) & ":" & ew_ZeroPad(Second(sDate), 2)
		ElseIf ANamedFormat = 12 Or ANamedFormat = 15 Then
			ew_FormatDateTime = Right(Year(sDate),2) & EW_DATE_SEPARATOR & Month(sDate) & EW_DATE_SEPARATOR & Day(sDate)
		ElseIf ANamedFormat = 13 Or ANamedFormat = 16 Then
			ew_FormatDateTime = Month(sDate) & EW_DATE_SEPARATOR & Day(sDate) & EW_DATE_SEPARATOR & Right(Year(sDate),2)
		ElseIf ANamedFormat = 14 Or ANamedFormat = 17 Then
			ew_FormatDateTime = Day(sDate) & EW_DATE_SEPARATOR & Month(sDate) & EW_DATE_SEPARATOR & Right(Year(sDate),2)
		Else
			ew_FormatDateTime = sDate
		End If
		If (ANamedFormat >= 9 And ANamedFormat <= 11) Or (ANamedFormat >= 15 And ANamedFormat <= 17) Then
				ew_FormatDateTime = ew_FormatDateTime & " " & ew_ZeroPad(Hour(sDate), 2) & ":" & ew_ZeroPad(Minute(sDate), 2) & ":" & ew_ZeroPad(Second(sDate), 2)
				If Len(ADate) > Len(sDate) Then ew_FormatDateTime = ew_FormatDateTime & Mid(ADate, Len(sDate)+1)
		End If
	Else
		ew_FormatDateTime = ADate
	End If
End Function

' Unformat date time based on format type
Function ew_UnFormatDateTime(ADate, ANamedFormat)
	ew_UnFormatDateTime = ADate ' Default return date
	Dim arDateTime, arDate, i
	ADate = Trim(ADate & "")
	While Instr(ADate, "  ") > 0
		ADate = Replace(ADate, "  ", " ")
	Wend
	arDateTime = Split(ADate, " ")
	If UBound(arDateTime) < 0 Then
		ew_UnFormatDateTime = ADate
		Exit Function
	End If
	If ANamedFormat = 0 And IsDate(ADate) Then
		ew_UnFormatDateTime = Year(arDateTime(0)) & "/" & Month(arDateTime(0)) & "/" & Day(arDateTime(0))
		If UBound(arDateTime) > 0 Then
			For i = 1 to UBound(arDateTime)
				ew_UnFormatDateTime = ew_UnFormatDateTime & " " & arDateTime(i)
			Next
		End If
	Else
		arDate = Split(arDateTime(0), EW_DATE_SEPARATOR)
		If UBound(arDate) = 2 Then
			ew_UnFormatDateTime = arDateTime(0)
			If ANamedFormat = 6 Or ANamedFormat = 10 Then ' mmddyyyy
				If ew_CheckUSDate(arDateTime(0)) Then
					ew_UnFormatDateTime = arDate(2) & "/" & arDate(0) & "/" & arDate(1)
				End If
			ElseIf (ANamedFormat = 7 Or ANamedFormat = 11) Then ' ddmmyyyy
				If ew_CheckEuroDate(arDateTime(0)) Then
					ew_UnFormatDateTime = arDate(2) & "/" & arDate(1) & "/" & arDate(0)
				End If
			ElseIf ANamedFormat = 5 Or ANamedFormat = 9 Then ' yyyymmdd
				If ew_CheckDate(arDateTime(0)) Then
					ew_UnFormatDateTime = arDate(0) & "/" & arDate(1) & "/" & arDate(2)
				End If
			ElseIf ANamedFormat = 12 Or ANamedFormat = 15 Then ' yymmdd
				If ew_CheckShortDate(arDateTime(0)) Then
					ew_UnFormatDateTime = ew_UnformatYear(arDate(0)) & "/" & arDate(1) & "/" & arDate(2)
				End If
			ElseIf ANamedFormat = 13 Or ANamedFormat = 16 Then ' mmddyy
				If ew_CheckShortUSDate(arDateTime(0)) Then
					ew_UnFormatDateTime = ew_UnformatYear(arDate(2)) & "/" & arDate(0) & "/" & arDate(1)
				End If
			ElseIf ANamedFormat = 14 Or ANamedFormat = 17 Then ' ddmmyy
				If ew_CheckShortEuroDate(arDateTime(0)) Then
					ew_UnFormatDateTime = ew_UnformatYear(arDate(2)) & "/" & arDate(1) & "/" & arDate(0)
				End If
			End If
			If UBound(arDateTime) > 0 Then
				For i = 1 to UBound(arDateTime)
					ew_UnFormatDateTime = ew_UnFormatDateTime & " " & arDateTime(i)
				Next
			End If
		Else
			ew_UnFormatDateTime = ADate
		End If
	End If
End Function

' Unformat 2 digit year to 4 digit year
Function ew_UnformatYear(yr)
	ew_UnformatYear = yr
	If Len(yr) = 2 Then
		If IsNumeric(yr) Then
			If CLng(yr) > EW_UNFORMAT_YEAR Then
				ew_UnformatYear = "19" & yr
			Else
				ew_UnformatYear = "20" & yr
			End If
		End If
	End If
End Function

' Format currency
Function ew_FormatCurrency(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits)
	On Error Resume Next
	ew_FormatCurrency = FormatCurrency(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits)
	If Err.Number <> 0 Then
		Err.Clear
		ew_FormatCurrency = Expression
	End If
End Function

' Format number
Function ew_FormatNumber(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits)
	On Error Resume Next
	ew_FormatNumber = FormatNumber(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits)
	If Err.Number <> 0 Then
		Err.Clear
		ew_FormatNumber = Expression
	End If
End Function

' Format percent
Function ew_FormatPercent(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits)
	On Error Resume Next
	ew_FormatPercent = FormatPercent(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits)
	If Err.Number <> 0 Then
		Err.Clear
		ew_FormatPercent = FormatNumber(Expression*100, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits) & "%"
		If Err.Number <> 0 Then
			Err.Clear
			ew_FormatPercent = Expression
		End If
	End If
End Function

' Encode html
Function ew_HtmlEncode(Expression)
	ew_HtmlEncode = Server.HtmlEncode(Expression & "")
End Function

' Get key value
Function ew_GetKeyValue(Key)
	If IsNull(Key) Then
		ew_GetKeyValue = ""
	ElseIf IsArray(Key) Then
		ew_GetKeyValue = Join(Key, EW_COMPOSITE_KEY_SEPARATOR)
	Else
		ew_GetKeyValue = Key
	End If
End Function

' Convert dictionary to JSON for HTML attributes
Function ew_ArrayToJsonAttr(Ar)
	Dim str, name, value, i
	str = "{"
	If IsArray(Ar) Then
		For i = 0 to UBound(Ar)
			If IsArray(Ar(i)) Then
				If UBound(Ar(i)) >= 1 Then
					name = Ar(i)(0)
					value = Ar(i)(1)
					str = str & name & ":'" & ew_JsEncode3(value) & "',"
				End If
			End If
		Next
	End If
	If Right(str,1) = "," Then str = Mid(str,1,Len(str)-1)
	str = str & "}"
	ew_ArrayToJsonAttr = str
End Function

' Generate Value Separator based on current row count
' rowcnt - zero based row count
' dispidx - zero based display index
' fld - field object
Function ew_ValueSeparator(rowidx, dispidx, fld)
	ew_ValueSeparator = ", "
End Function

' Generate View Option Separator based on current option count (Multi-Select / CheckBox)
' optidx - zero based option index
Function ew_ViewOptionSeparator(optidx)
	ew_ViewOptionSeparator = ", "
End Function

' Render repeat column table
' rowcnt - zero based row count
Function ew_RepeatColumnTable(totcnt, rowcnt, repeatcnt, rendertype)
	Dim sWrk, i
	sWrk = ""

	' Render control start
	If rendertype = 1 Then
		If rowcnt = 0 Then sWrk = sWrk & "<table class=""" & EW_ITEM_TABLE_CLASSNAME & """>"
		If (rowcnt mod repeatcnt = 0) Then sWrk = sWrk & "<tr>"
		sWrk = sWrk & "<td>"

	' Render control end
	ElseIf rendertype = 2 Then
		sWrk = sWrk & "</td>"
		If (rowcnt mod repeatcnt = repeatcnt -1) Then
			sWrk = sWrk & "</tr>"
		ElseIf rowcnt = totcnt Then
			For i = ((rowcnt mod repeatcnt) + 1) to repeatcnt - 1
				sWrk = sWrk & "<td>&nbsp;</td>"
			Next
			sWrk = sWrk & "</tr>"
		End If
		If rowcnt = totcnt Then sWrk = sWrk & "</table>"
	End If
	ew_RepeatColumnTable = sWrk
End Function

' Truncate Memo Field based on specified length, string truncated to nearest space or CrLf
Function ew_TruncateMemo(memostr, ln, removeHtml)
	Dim i, j, k
	Dim str
	If removeHtml Then
		str = ew_RemoveHtml(memostr) ' Remove Html
	Else
		str = memostr
	End If
	If Len(str) > 0 And Len(str) > ln Then
		k = 1
		Do While k > 0 And k < Len(str)
			i = InStr(k, str, " ", 1)
			j = InStr(k, str, vbCrLf, 1)
			If i < 0 And j < 0 Then ' Not able to truncate
				ew_TruncateMemo = str
				Exit Function
			Else

				' Get nearest space or CrLf
				If i > 0 And j > 0 Then
					If i < j Then
						k = i
					Else
						k = j
					End If
				ElseIf i > 0 Then
					k = i
				ElseIf j > 0 Then
					k = j
				End If

				' Get truncated text
				If k >= ln Then
					ew_TruncateMemo = Mid(str, 1, k-1) & "..."
					Exit Function
				Else
					k = k + 1
				End If
			End If
		Loop
	Else
		ew_TruncateMemo = str
	End If
End Function

' Remove Html tags from text
Function ew_RemoveHtml(str)
	Dim RegEx
	Set RegEx = New RegExp
	RegEx.Pattern = "<[^>]*>"
	RegEx.Global = True
	ew_RemoveHtml = RegEx.Replace(str & "", "")
End Function

' Send email by template
Function ew_SendTemplateEmail(sTemplate, sSender, sRecipient, sCcEmail, sBccEmail, sSubject, arContent)
	If Not EW_DEBUG_ENABLED Then On Error Resume Next
	If sSender <> "" And sRecipient <> "" Then
		Dim Email, i, cnt
		Set Email = New cEmail
		Email.Load(sTemplate)
		Email.ReplaceSender(sSender) ' Replace Sender
		Email.ReplaceRecipient(sRecipient) ' Replace Recipient
		If sCcEmail <> "" Then Email.AddCc sCcEmail ' Add Cc
		If sBccEmail <> "" Then Email.AddBcc sBccEmail ' Add Bcc
		If sSubject <> "" Then Email.ReplaceSubject(sSubject) ' Replace subject
		If IsArray(arContent) Then
			cnt = UBound(arContent) - 1
			If cnt Mod 2 = 1 Then cnt = cnt - 1
			For i = 0 to cnt Step 2
				Email.ReplaceContent arContent(i), arContent(i+1)
			Next
		End If
		ew_SendTemplateEmail = Email.Send()
		Set Email = Nothing
	Else
		ew_SendTemplateEmail = False
	End If
End Function

' Function to Send out Email
' Supports CDO, w3JMail and ASPEmail
Function ew_SendEmail(sFrEmail, sToEmail, sCcEmail, sBccEmail, sSubject, sMail, sFormat, sCharset)
	If Not EW_DEBUG_ENABLED Then On Error Resume Next
	Dim i, objMail, sServerVersion, sIISVer, EmailComponent, arrEmail, sEmail
	Dim arCDO, arASPEmail, arw3JMail, arEmailComponent
	sServerVersion = Request.ServerVariables("SERVER_SOFTWARE")
	If InStr(sServerVersion, "Microsoft-IIS") > 0 Then
		i = InStr(sServerVersion, "/")
		If i > 0 Then
			sIISVer = Trim(Mid(sServerVersion, i+1))
		End If
	End If
	arw3JMail = Array("w3JMail", "JMail.Message")
	arASPEmail = Array("ASPEmail", "Persits.MailSender")
	If sIISVer < "5.0" Then ' NT using CDONTS
		arCDO = Array("CDO", "CDONTS.NewMail")
	Else ' 2000 / XP / 2003 using CDO
		arCDO = Array("CDO", "CDO.Message")
	End If

	' Change your precedence here
	arEmailComponent = Array(arCDO, arw3JMail, arASPEmail) ' Use CDO as default

	' arEmailComponent = Array(arw3JMail, arASPEmail, arCDO)
	EmailComponent = ""
	For i = 0 to UBound(arEmailComponent)
		Err.Clear
		Set objMail = Server.CreateObject(arEmailComponent(i)(1))
		If Err.Number = 0 Then
			EmailComponent = arEmailComponent(i)(0)
			Exit For
		End If
	Next
	If EmailComponent = "" Then
		ew_SendEmail = False
		Call ew_Trace("email_err", "Unable to create email component. Error Number: " & Hex(Err.Number))
		Exit Function
	End If
	If EmailComponent = "w3JMail" Then

		' Set objMail = Server.CreateObject("JMail.Message")
		If sCharset <> "" Then objMail.Charset = sCharset
		objMail.Logging = True
		objMail.Silent = True
		objMail.From = sFrEmail
		arrEmail = Split(Replace(sToEmail, ",", ";"), ";")
		For i = 0 to UBound(arrEmail)
			sEmail = Trim(arrEmail(i))
			If sEmail <> "" Then
				objMail.AddRecipient sEmail
			End If
		Next
		arrEmail = Split(Replace(sCcEmail, ",", ";"), ";")
		For i = 0 to UBound(arrEmail)
			sEmail = Trim(arrEmail(i))
			If sEmail <> "" Then
				objMail.AddRecipientCC sEmail
			End If
		Next
		arrEmail = Split(Replace(sBccEmail, ",", ";"), ";")
		For i = 0 to UBound(arrEmail)
			sEmail = Trim(arrEmail(i))
			If sEmail <> "" Then
				objMail.AddRecipientBCC sEmail
			End If
		Next
		objMail.Subject = sSubject
		If LCase(sFormat) = "html" Then
			objMail.HTMLBody = sMail
		Else
			objMail.Body = sMail
		End If
		If EW_SMTP_SERVER_USERNAME <> "" And EW_SMTP_SERVER_PASSWORD <> "" Then
			objMail.MailServerUserName = EW_SMTP_SERVER_USERNAME
			objMail.MailServerPassword = EW_SMTP_SERVER_PASSWORD
		End If
		ew_SendEmail = objMail.Send(EW_SMTP_SERVER)
		If Not ew_SendEmail Then
			Err.Raise vbObjectError + 1, EmailComponent, objMail.Log
		End If
		Set objMail = nothing
	ElseIf EmailComponent = "ASPEmail" Then

		' Set objMail = Server.CreateObject("Persits.MailSender")
		If sCharset <> "" Then objMail.CharSet = sCharset
		objMail.From = sFrEmail
		arrEmail = Split(Replace(sToEmail, ",", ";"), ";")
		For i = 0 to UBound(arrEmail)
			sEmail = Trim(arrEmail(i))
			If sEmail <> "" Then
				objMail.AddAddress sEmail
			End If
		Next
		arrEmail = split(Replace(sCcEmail, ",", ";"), ";")
		For i = 0 to UBound(arrEmail)
			sEmail = Trim(arrEmail(i))
			If sEmail <> "" Then
				objMail.AddCC sEmail
			End If
		Next
		arrEmail = split(Replace(sBccEmail, ",", ";"), ";")
		For i = 0 to UBound(arrEmail)
			sEmail = Trim(arrEmail(i))
			If sEmail <> "" Then
				objMail.AddBcc sEmail
			End If
		Next
		If LCase(sFormat) = "html" Then
			objMail.IsHTML = True ' html
		Else
			objMail.IsHTML = False ' text
		End If
		objMail.Subject = sSubject
		objMail.Body = sMail
		objMail.Host = EW_SMTP_SERVER
		If EW_SMTP_SERVER_USERNAME <> "" And EW_SMTP_SERVER_PASSWORD <> "" Then
			objMail.Username = EW_SMTP_SERVER_USERNAME
			objMail.Password = EW_SMTP_SERVER_PASSWORD
		End If
		ew_SendEmail = objMail.Send
		Set objMail = Nothing
	ElseIf EmailComponent = "CDO" Then
		Dim objConfig, sSmtpServer, iSmtpServerPort
		If sIISVer < "5.0" Then ' NT using CDONTS

			' Set objMail = Server.CreateObject("CDONTS.NewMail")
			'***If sCharset <> "" Then objMail.BodyPart.Charset = sCharset ' Do not support charset, ignore

			objMail.From = sFrEmail
			objMail.To = Replace(sToEmail, ",", ";")
			If sCcEmail <> "" Then
				objMail.Cc = Replace(sCcEmail, ",", ";")
			End If
			If sBccEmail <> "" Then
				objMail.Bcc = Replace(sBccEmail, ",", ";")
			End If
			If LCase(sFormat) = "html" Then
				objMail.BodyFormat = 0 ' 0 means HTML format, 1 means text
				objMail.MailFormat = 0 ' 0 means MIME, 1 means text
			End If
			objMail.Subject = sSubject
			objMail.Body = sMail
			objMail.Send
			Set objMail = Nothing
		Else ' 2000 / XP / 2003 using CDO

			' Set up Configuration
			Set objConfig = Server.CreateObject("CDO.Configuration")
			objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = EW_SMTP_SERVER ' cdoSMTPServer
			objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = EW_SMTP_SERVER_PORT ' cdoSMTPServerPort
			objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
			If EW_SMTP_SERVER_USERNAME <> "" And EW_SMTP_SERVER_PASSWORD <> "" Then
				objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'cdoBasic (clear text)
				objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = EW_SMTP_SERVER_USERNAME
				objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = EW_SMTP_SERVER_PASSWORD
			End If
			objConfig.Fields.Update

			' Set up Mail
			'Set objMail = Server.CreateObject("CDO.Message")

			objMail.From = sFrEmail
			objMail.To = Replace(sToEmail, ",", ";")
			If sCcEmail <> "" Then
				objMail.Cc = Replace(sCcEmail, ",", ";")
			End If
			If sBccEmail <> "" Then
				objMail.Bcc = Replace(sBccEmail, ",", ";")
			End If
			If sCharset <> "" Then objMail.BodyPart.Charset = sCharset
			If LCase(sFormat) = "html" Then
				objMail.HtmlBody = sMail
				If sCharset <> "" Then objMail.HtmlBodyPart.Charset = sCharset
			Else
				objMail.TextBody = sMail
				If sCharset <> "" Then objMail.TextBodyPart.Charset = sCharset
			End If
			objMail.Subject = sSubject
			If EW_SMTP_SERVER <> "" And LCase(EW_SMTP_SERVER) <> "localhost" Then
				Set objMail.Configuration = objConfig ' Use Configuration
				objMail.Send
			Else
				objMail.Send ' Send without Configuration
				If Err.Number <> 0 Then
					If Hex(Err.Number) = "80040220" Then ' Requires Configuration
						Set objMail.Configuration = objConfig
						Err.Clear
						objMail.Send
					End If
				End If
			End If
			Set objMail = Nothing
			Set objConfig = Nothing
		End If
		ew_SendEmail = (Err.Number = 0)
	End If

	' Send email failed, write error to log
	If Not ew_SendEmail Then
		gsEmailErrNo = Err.Number
		gsEmailErrDesc = Err.Description
		Call ew_Trace("email_err", "***Send email failed***")
		Call ew_Trace("email_err", "Email component: " & EmailComponent)
		Call ew_Trace("email_err", "Error Number: " & Hex(gsEmailErrNo))
		Call ew_Trace("email_err", "Error Description: " & gsEmailErrDesc)
		Call ew_Trace("email_err", "From: " & sFrEmail)
		Call ew_Trace("email_err", "To: " & sToEmail)
		Call ew_Trace("email_err", "Cc: " & sCcEmail)
		Call ew_Trace("email_err", "Bcc: " & sToEmail)
		Call ew_Trace("email_err", "Subject: " & sSubject)
	End If
End Function 

' Load content at url using xmlhttp
Function ew_LoadContentFromUrl(url)

	'On Error Resume Next
	Dim http
	Set http = Server.CreateObject("MSXML2.ServerXMLHTTP")
	http.setTimeouts 20000,20000,20000,30000
	http.Open "GET", url, False
	http.send
	ew_LoadContentFromUrl = http.responseText
End Function

Function ew_FieldDataType(FldType) ' Field data type
	Select Case FldType
		Case 20, 3, 2, 16, 4, 5, 131, 139, 6, 17, 18, 19, 21 ' Numeric
			ew_FieldDataType = EW_DATATYPE_NUMBER
		Case 7, 133, 135, 146 ' Date
			ew_FieldDataType = EW_DATATYPE_DATE
		Case 134, 145 ' Time
			ew_FieldDataType = EW_DATATYPE_TIME
		Case 201, 203 ' Memo
			ew_FieldDataType = EW_DATATYPE_MEMO
		Case 129, 130, 200, 202 ' String
			ew_FieldDataType = EW_DATATYPE_STRING
		Case 11 ' Boolean
			ew_FieldDataType = EW_DATATYPE_BOOLEAN
		Case 72 ' GUID
			ew_FieldDataType = EW_DATATYPE_GUID
		Case 128, 204, 205 ' Binary
			ew_FieldDataType = EW_DATATYPE_BLOB
		Case 141 ' Xml
			ew_FieldDataType = EW_DATATYPE_XML
		Case Else
			ew_FieldDataType = EW_DATATYPE_OTHER
		End Select
End Function

' Return path of the uploaded file
'	Parameter: If PhyPath is true(1), return physical path on the server;
'	           If PhyPath is false(0), return relative URL
Function ew_UploadPathEx(PhyPath, DestPath)
	Dim Pos
	If PhyPath Then
		ew_UploadPathEx = Request.ServerVariables("APPL_PHYSICAL_PATH")
		ew_UploadPathEx = ew_IncludeTrailingDelimiter(ew_UploadPathEx, PhyPath)
		ew_UploadPathEx = ew_PathCombine(ew_UploadPathEx, Replace(DestPath, "/", "\"), PhyPath)
	Else
		ew_UploadPathEx = Request.ServerVariables("APPL_MD_PATH")
		Pos = InStr(1, ew_UploadPathEx, "Root", 1)
		If Pos > 0 Then	ew_UploadPathEx = Mid(ew_UploadPathEx, Pos+4)
		ew_UploadPathEx = ew_IncludeTrailingDelimiter(ew_UploadPathEx, PhyPath)
		ew_UploadPathEx = ew_PathCombine(ew_UploadPathEx, DestPath, PhyPath)
	End If
	ew_UploadPathEx = ew_IncludeTrailingDelimiter(ew_UploadPathEx, PhyPath)
End Function

' Change the file name of the uploaded file
Function ew_UploadFileNameEx(Folder, FileName)
	Dim OutFileName

	' By default, ewUniqueFileName() is used to get an unique file name.
	' Amend your logic here

	OutFileName = ew_UniqueFileName(Folder, FileName)

	' Return computed output file name
	ew_UploadFileNameEx = OutFileName
End Function

' Return path of the uploaded file
' returns global upload folder, for backward compatibility only
Function ew_UploadPath(PhyPath)
	ew_UploadPath = ew_UploadPathEx(PhyPath, EW_UPLOAD_DEST_PATH)
End Function

' Change the file name of the uploaded file
' use global upload folder, for backward compatibility only
Function ew_UploadFileName(FileName)
	ew_UploadFileName = ew_UploadFileNameEx(ew_UploadPath(True), FileName)
End Function

' Generate an unique file name (filename(n).ext)
Function ew_UniqueFileName(Folder, FileName)
	If FileName = "" Then FileName = ew_DefaultFileName()
	If FileName = "." Then
		Response.Write "Invalid file name: " & FileName
		Response.End
		Exit Function
	End If
	If Folder = "" Then
		Response.Write "Unspecified folder"
		Response.End
		Exit Function
	End If
	Dim Name, Ext, Pos
	Name = ""
	Ext = ""
	Pos = InStrRev(FileName, ".")
	If Pos = 0 Then
		Name = FileName
		Ext = ""
	Else
		Name = Mid(FileName, 1, Pos-1)
		Ext = Mid(FileName, Pos+1)
	End If
	Folder = ew_IncludeTrailingDelimiter(Folder, True)
	Dim fso
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	If Not fso.FolderExists(Folder) Then
		If Not ew_CreateFolder(Folder) Then
			Response.Write "Folder does not exist: " & Folder
			Set fso = Nothing
			Exit Function
		End If
	End If
	Dim Suffix, Index
	Index = 0
	Suffix = ""

	' Check to see if filename exists
	While fso.FileExists(folder & Name & Suffix & "." & Ext)
		Index = Index + 1
		Suffix = "(" & Index & ")"
	Wend
	Set fso = Nothing

	' Return unique file name
	ew_UniqueFileName = Name & Suffix & "." & Ext
End Function

' Create a default file name (yyyymmddhhmmss.bin)
Function ew_DefaultFileName
	Dim dt
	dt = Now()
	ew_DefaultFileName = ew_ZeroPad(Year(dt), 4) & ew_ZeroPad(Month(dt), 2) &  _
		ew_ZeroPad(Day(dt), 2) & ew_ZeroPad(Hour(dt), 2) & _
		ew_ZeroPad(Minute(dt), 2) & ew_ZeroPad(Second(dt), 2) & ".bin"
End Function

' Get path relative to application root
Function ew_ServerMapPath(Path)
	ew_ServerMapPath = ew_PathCombine(Request.ServerVariables("APPL_PHYSICAL_PATH"), Path, True)
End Function

' Get path relative to a base path
Function ew_PathCombine(ByVal BasePath, ByVal RelPath, ByVal PhyPath)
	Dim Path, Path2, p1, p2, Delimiter
	BasePath = ew_RemoveTrailingDelimiter(BasePath, PhyPath)
	If PhyPath Then
		Delimiter = "\"
		RelPath = Replace(RelPath, "/", "\")
	Else
		Delimiter = "/"
		RelPath = Replace(RelPath, "\", "/")
	End If
	If RelPath = "." Or RelPath = ".." Then RelPath = RelPath & Delimiter
	p1 = InStr(RelPath, Delimiter)
	Path2 = ""
	While p1 > 0
		Path = Left(RelPath, p1)
		If Path = Delimiter Or Path = "." & Delimiter Then

			' Skip
		ElseIf Path = ".." & Delimiter Then
			p2 = InStrRev(BasePath, Delimiter)
			If p2 > 0 Then BasePath = Left(BasePath, p2-1)
		Else
			Path2 = Path2 & Path
		End If
		RelPath = Mid(RelPath, p1+1)
		p1 = InStr(RelPath, Delimiter)
	Wend
	ew_PathCombine = ew_IncludeTrailingDelimiter(BasePath, PhyPath) & Path2 & RelPath
End Function

' Remove the last delimiter for a path
Function ew_RemoveTrailingDelimiter(ByVal Path, ByVal PhyPath)
	Dim Delimiter
	If PhyPath Then Delimiter = "\" Else Delimiter = "/"
	While Right(Path, 1) = Delimiter
		Path = Left(Path, Len(Path)-1)
	Wend
	ew_RemoveTrailingDelimiter = Path
End Function

' Include the last delimiter for a path
Function ew_IncludeTrailingDelimiter(ByVal Path, ByVal PhyPath)
	Dim Delimiter
	Path = ew_RemoveTrailingDelimiter(Path, PhyPath)
	If PhyPath Then Delimiter = "\" Else Delimiter = "/"
	ew_IncludeTrailingDelimiter = Path & Delimiter
End Function

' Write the paths for config/debug only
Sub ew_WriteUploadPaths
	Response.Write "Request.ServerVariables(""APPL_PHYSICAL_PATH"")=" & _
		Request.ServerVariables("APPL_PHYSICAL_PATH") & "<br>"
	Response.Write "Request.ServerVariables(""APPL_MD_PATH"")=" & _
		Request.ServerVariables("APPL_MD_PATH") & "<br>"
End Sub

' Get current page name
Function ew_CurrentPage()
	ew_CurrentPage = ew_GetPageName(Request.ServerVariables("SCRIPT_NAME"))
End Function

' Get refer page name
Function ew_ReferPage()
	ew_ReferPage = ew_GetPageName(Request.ServerVariables("HTTP_REFERER"))
End Function

' Get page name
Function ew_GetPageName(url)
	If url <> "" Then
		ew_GetPageName = url
		If InStr(ew_GetPageName, "?") > 0 Then
			ew_GetPageName = Mid(ew_GetPageName, 1, InStr(ew_GetPageName, "?")-1) ' Remove querystring first
		End If
		ew_GetPageName = Mid(ew_GetPageName, InStrRev(ew_GetPageName, "/")+1) ' Remove path
	Else
		ew_GetPageName = ""
	End If
End Function

' Check if folder exists
Function ew_FolderExists(Folder)
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	ew_FolderExists = fso.FolderExists(Folder)
	Set fso = Nothing
End Function

' Check if file exists
Function ew_FileExists(Folder, File)
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	ew_FileExists = fso.FileExists(Folder & File)
	Set fso = Nothing
End Function

' Delete file
Sub ew_DeleteFile(FilePath)
	If Not EW_DEBUG_ENABLED Then On Error Resume Next
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	If FilePath <> "" And fso.FileExists(FilePath) Then
		fso.DeleteFile(FilePath)
	End If
	Set fso = Nothing
End Sub

' Rename file
Sub ew_RenameFile(OldFilePath, NewFilePath)
	If Not EW_DEBUG_ENABLED Then On Error Resume Next
	Dim fso
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	If OldFilePath <> "" And fso.FileExists(OldFilePath) Then
		fso.MoveFile OldFilePath, NewFilePath
	End If
	Set fso = Nothing
End Sub

' Create folder
Function ew_CreateFolder(Folder)
	If Not EW_DEBUG_ENABLED Then On Error Resume Next
	ew_CreateFolder = False
	Dim fso
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	If Not fso.FolderExists(Folder) Then
		If ew_CreateFolder(fso.GetParentFolderName(Folder)) Then
			fso.CreateFolder(Folder)
			If Err.Number = 0 Then ew_CreateFolder = True
		End If
	Else
		ew_CreateFolder = True
	End If
	Set fso = Nothing
End Function

' Add an element to a position of an array
Function ew_AddItemToArray(ar, pos, aritem)
	Dim newar(), d1, d2, d3, p
	Dim i, j
	If not IsArray(aritem) Then
		ew_AddItemToArray = ar
		Exit Function
	End If
	d3 = UBound(aritem)
	If not IsArray(ar) Then
		Redim newar(d3,0)
		For i = 0 to d3
			newar(i,0) = aritem(i)
		Next
		ew_AddItemToArray = newar
		Exit Function
	Else
		d1 = UBound(ar,1)
		d2 = UBound(ar,2)
		p = pos
		If p < 0 Then p = 0 ' add at front
		If p > d2 Then p = d2 ' add at end
		Redim newar(d1, d2+1)

		' Copy item before p
		For j = 0 to p-1
			For i = 0 to d1
				newar(i,j) = ar(i,j)
			Next
		Next

		' Copy new item
		For i = 0 to d1
			If i <= d3 Then
				newar(i,p) = aritem(i)
			Else
				newar(i,p) = "" ' Initialize to empty string
			End If
		Next

		' Copy the rest
		For j = p to d2
			For i = 0 to d1
				newar(i,j+1) = ar(i,j)
			Next
		Next
	End If
	ew_AddItemToArray = newar
End Function

' Remove an element from a position of an array
Function ew_RemoveItemFromArray(ar, pos)
	Dim newar(), d1, d2
	Dim i, j
	ew_RemoveItemFromArray = Null
	If IsArray(ar) Then
		d1 = UBound(ar,1)
		d2 = UBound(ar,2)
		If pos < 0 Or pos > d2 Then
			ew_RemoveItemFromArray = ar
			Exit Function
		End If
		If d2 = 0 Then
			ew_RemoveItemFromArray = Null
		Else
			Redim newar(d1, d2-1)

			' Copy items before pos
			For j = 0 to pos-1
				For i = 0 to d1
					newar(i,j) = ar(i,j)
				Next
			Next

			' Copy items after pos
			For j = pos+1 to d2
				For i = 0 to d1
					newar(i,j-1) = ar(i,j)
				Next
			Next
			ew_RemoveItemFromArray = newar
		End If
	End If
End Function

' Functions for Export
Function ew_ExportHeader(ExpType)
	Select Case ExpType
		Case "html", "email"
			ew_ExportHeader = "<table class=""ewExportTable"">"
			If EW_EXPORT_CSS_STYLES Then
				ew_ExportHeader = "<style>" & ew_LoadFile(EW_PROJECT_STYLESHEET_FILENAME) & "</style>" & ew_ExportHeader
			End If
		Case "word", "excel"
			ew_ExportHeader = "<table>"
		Case "csv"
			ew_ExportHeader = ""
	End Select
End Function

Function ew_ExportFooter(ExpType)
	Select Case ExpType
		Case "html", "email", "word", "excel"
			ew_ExportFooter = "</table>"
		Case "csv"
			ew_ExportFooter = ""
	End Select
End Function

Sub ew_ExportAddValue(str, val, ExpType, Attr)
	Select Case ExpType
		Case "html", "email", "word", "excel"
			str = str & "<td"
			If Attr <> "" Then str = str & " " & Attr
			str = str & ">" & val & "</td>"
		Case "csv"
			If str <> "" Then str = str & ","
			str = str & """" & Replace(val & "", """", """""") & """"
	End Select
End Sub

Function ew_ExportLine(str, ExpType, Attr)
	Select Case ExpType
		Case "html", "email", "word", "excel"
			ew_ExportLine = "<tr"
			If Attr <> "" Then ew_ExportLine = ew_ExportLine & " " & Attr
			ew_ExportLine = ew_ExportLine & ">" & str & "</tr>"
		Case "csv"
			ew_ExportLine = str & vbCrLf
	End Select
End Function

Function ew_ExportField(cap, val, ExpType, Attr)
	Dim sTD
	sTD = "<td"
	If Attr <> "" Then sTD = sTD & " " & Attr
	sTD = sTD & ">"
	ew_ExportField = "<tr>" & sTD & cap & "</td>" & sTD & val & "</td></tr>"
End Function

' Check if field exists in recordset
Function ew_FieldExistInRs(rs, fldname)
	Dim fld
	For Each fld in rs.Fields
		If fld.name = fldname then
			ew_FieldExistInRs = True
			Exit Function
    End If
	Next
	ew_FieldExistInRs = False
End Function

' Calculate field hash
Function ew_GetFldHash(value, fldtype)
	ew_GetFldHash = MD5(ew_GetFldValueAsString(value, fldtype))
End Function

' Get field value as string
Function ew_GetFldValueAsString(value, fldtype)
	If IsNull(value) Then
		ew_GetFldValueAsString = ""
	Else
		If fldtype = 128 Or fldtype = 204 Or fldtype = 205 Then ' Binary
			If EW_BLOB_FIELD_BYTE_COUNT > 0 Then
				ew_GetFldValueAsString = ew_ByteToString(LeftB(value,EW_BLOB_FIELD_BYTE_COUNT))
			Else
				ew_GetFldValueAsString = ew_ByteToString(value)
			End If
		Else
			ew_GetFldValueAsString = CStr(value)
		End If
	End If
End Function

' Convert byte to string
Function ew_ByteToString(b)
	Dim i
	For i = 1 to LenB(b)
		ew_ByteToString = ew_ByteToString & Chr(AscB(MidB(b,i,1)))
	Next
End Function

' Write global debug message
Function ew_DebugMsg()
	Dim msg
	msg = gsDebugMsg
	gsDebugMsg = ""
	If msg <> "" Then
		ew_DebugMsg = "<p>" & msg & "</p>"
	Else
		ew_DebugMsg = ""
	End If
End Function

' Write global debug message
Sub ew_SetDebugMsg(v)
	Call ew_AddMessage(gsDebugMsg, v)
End Sub

'
'  Common functions (end)
' ------------------------
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

' Highlight value based on basic search / advanced search keywords
Function ew_Highlight(name, src, bkw, bkwtype, akw)
	Dim i, x, y, outstr, kwlist, kw, kwstr
	Dim wrksrc, xx, yy
	outstr = ""
	If Len(src) > 0 And (Len(bkw) > 0 Or Len(akw) > 0) Then
		xx = 1
		yy = InStr(xx, src, "<", 1)
		If yy <= 0 Then yy = Len(src)+1
		Do While yy > 0
			If (yy > xx) Then
				wrksrc = Mid(src, xx, yy-xx)
			kwstr = Trim(bkw)
			If Len(akw) > 0 Then
				If Len(kwstr) > 0 Then kwstr = kwstr & " "
				kwstr = kwstr & Trim(akw)
			End If
			kwlist = Split(kwstr, " ")
			x = 1
			Call ew_GetKeyword(wrksrc, kwlist, x, y, kw)
			Do While y > 0
				outstr = outstr & Mid(wrksrc, x, y-x) & _
					"<span id=""" & name & """ name=""" & name & """ class=""ewHighlightSearch"">" & _
					Mid(wrksrc, y, Len(kw)) & "</span>"
				x = y + Len(kw)
				Call ew_GetKeyword(wrksrc, kwlist, x, y, kw)
			Loop
			outstr = outstr & Mid(wrksrc, x)
				xx = xx + Len(wrksrc)
			End If
			If xx < len(src) Then
				yy = InStr(xx, src, ">", 1)
				If yy > 0 Then
					outstr = outstr & Mid(src, xx, yy-xx+1)
					xx = yy + 1
					yy = InStr(xx, src, "<", 1)
					If yy <= 0 Then yy = Len(src)+1
				Else
					outstr = outstr & Mid(src, xx)
					yy = -1
				End If
			Else
				yy = -1
			End If
		Loop
	Else
		outstr = src
	End If
	ew_Highlight = outstr
End Function

' Get keyword
Sub ew_GetKeyword(src, kwlist, x, y, kw)
	Dim i, thisy, thiskw, wrky, wrkkw
	thisy = -1
	thiskw = ""
	For i = 0 to UBound(kwlist)
		wrkkw = Trim(kwlist(i))
		If wrkkw <> "" Then
			wrky = InStr(x, src, wrkkw, EW_HIGHLIGHT_COMPARE)
			If wrky > 0 Then
				If thisy = -1 Then
					thisy = wrky
					thiskw = wrkkw
				ElseIf wrky < thisy Then
					thisy = wrky
					thiskw = wrkkw
				End If
			End If
		End If
	Next
	y = thisy
	kw = thiskw
End Sub

' Set attribute
Sub ew_SetAttr(Attrs, Key, Value)
	If Not (Attrs Is Nothing) And Key <> "" And Value <> "" Then
		Attrs.AddAttribute Key, Value, True
	End If
End Sub

' Set up key
Sub ew_AddKey(Ar, Key, Value)
	If Key & "" <> "" And Value & "" <> "" Then
		If Not IsArray(Ar) Then
			ReDim Ar(0)
		Else
			ReDim Preserve Ar(UBound(Ar)+1)
		End If
		Ar(UBound(Ar)) = Array(Key, Value)
	End If
End Sub

' Get array position
Function ew_GetArPos(Ar, Name)
	Dim i
	If IsArray(Ar) Then
		For i = 0 to UBound(Ar,2)
			If Ar(0,i) = Name Then
				ew_GetArPos = i
				Exit Function
			End If
		Next
		i = UBound(Ar,2)+1
		ReDim Preserve Ar(1,i)
	Else
		i = 0
		ReDim Ar(1,i)
	End If
	ew_GetArPos = i
End Function

' Set array value
Sub ew_SetArVal(Ar, Name, Val)
	Dim idx, wrkname
	idx = ew_GetArPos(Ar, Name)
	wrkname = Name
	If wrkname = "" Then wrkname = idx
	Ar(0,idx) = wrkname
	Ar(1,idx) = Val
End Sub

' Set array object
Sub ew_SetArObj(Ar, Name, Obj)
	Dim idx, wrkname
	idx = ew_GetArPos(Ar, Name)
	wrkname = Name
	If wrkname = "" Then wrkname = idx
	Ar(0,idx) = wrkname
	Set Ar(1,idx) = Obj
End Sub

' Encrypt password
Function ew_EncryptPassword(input)
	ew_EncryptPassword = MD5(input)
End Function

' Compare password
Function ew_ComparePassword(pwd, input)
	If EW_CASE_SENSITIVE_PASSWORD Then
		If EW_ENCRYPTED_PASSWORD Then
			ew_ComparePassword = (pwd = ew_EncryptPassword(input))
		Else
			ew_ComparePassword = (pwd = input)
		End If
	Else
		If EW_ENCRYPTED_PASSWORD Then
			ew_ComparePassword = (pwd = ew_EncryptPassword(LCase(input)))
		Else
			ew_ComparePassword = (LCase(pwd) = LCase(input))
		End If
	End If
End Function

' Check empty string
Function ew_EmptyStr(value)
	Dim str
	str = CStr(value & "")
	str = Replace(str, "&nbsp;", "")
	ew_EmptyStr = (Trim(str) = "")
End Function

' Check empty file
Function ew_Empty(value)
	ew_Empty = IsEmpty(value) Or IsNull(value)
End Function
%>
<%

' Functions for backward compatibilty
' Get current user name
Function CurrentUserName()
	If Not EW_DEBUG_ENABLED Then On Error Resume Next
	If IsObject(Security) Then
		CurrentUserName = Security.CurrentUserName
	Else
		CurrentUserName = Session(EW_SESSION_USER_NAME) & ""
	End If
End Function

' Get current user ID
Function CurrentUserID()
	If Not EW_DEBUG_ENABLED Then On Error Resume Next
	If IsObject(Security) Then
		CurrentUserID = Security.CurrentUserID
	Else
		CurrentUserID = Session(EW_SESSION_USER_ID) & ""
	End If
End Function

' Get current parent user ID
Function CurrentParentUserID()
	If Not EW_DEBUG_ENABLED Then On Error Resume Next
	If IsObject(Security) Then
		CurrentParentUserID = Security.CurrentParentUserID
	Else
		CurrentParentUserID = Session(EW_SESSION_PARENT_USER_ID) & ""
	End If
End Function

' Get current user level
Function CurrentUserLevel()
	If Not EW_DEBUG_ENABLED Then On Error Resume Next
	If IsObject(Security) Then
		CurrentUserLevel = Security.CurrentUserLevelID
	Else
		CurrentUserLevel = Session(EW_SESSION_USER_LEVEL_ID)
	End If
End Function

' Get current user level list
Function CurrentUserLevelList()
	If Not EW_DEBUG_ENABLED Then On Error Resume Next
	If IsObject(Security) Then
		CurrentUserLevelList = Security.UserLevelList
	Else
		CurrentUserLevelList = Session(EW_SESSION_USER_LEVEL_ID) & ""
	End If
End Function

' Get Current user info
Function CurrentUserInfo(fldname)
	If IsObject(Security) Then
		CurrentUserInfo = Security.CurrentUserInfo(fldname)
		Exit Function
	ElseIf Not IsEmpty(EW_USER_TABLE) And Not IsSysAdmin() Then
		Dim user
		user = CurrentUserName()
		If user <> "" Then
			CurrentUserInfo = ew_ExecuteScalar("SELECT " & ew_QuotedName(fldname) & " FROM " & EW_USER_TABLE & " WHERE " & Replace(EW_USER_NAME_FILTER, "%u", ew_AdjustSql(user)))
			Exit Function
		End If
	End If
	CurrentUserInfo = Null
End Function

' Get current page ID
Function CurrentPageID()
	If Not IsEmpty(Page) Then
		CurrentPageID = Page.PageID
		Exit Function
	ElseIf Not IsEmpty(EW_PAGE_ID) Then
		CurrentPageID = EW_PAGE_ID
		Exit Function
	End If
	CurrentPageID = ""
End Function

' Allow list
Function AllowList(TableName)
	If Not EW_DEBUG_ENABLED Then On Error Resume Next
	If IsObject(Security) Then
		AllowList = Security.AllowList(TableName)
	Else
		AllowList = True
	End If
End Function

' Allow add
Function AllowAdd(TableName)
	If Not EW_DEBUG_ENABLED Then On Error Resume Next
	If IsObject(Security) Then
		AllowAdd = Security.AllowAdd(TableName)
	Else
		AllowAdd = True
	End If
End Function

' Is Password Expired
Function IsPasswordExpired()
	If Not EW_DEBUG_ENABLED Then On Error Resume Next
	If IsObject(Security) Then
		IsPasswordExpired = Security.IsPasswordExpired
	Else
		IsPasswordExpired = (Session(EW_SESSION_STATUS) = "passwordexpired")
	End If
End Function

' Is Logging In
Function IsLoggingIn()
	If Not EW_DEBUG_ENABLED Then On Error Resume Next
	If IsObject(Security) Then
		IsLoggingIn = Security.IsLoggingIn
	Else
		IsLoggingIn = (Session(EW_SESSION_STATUS) = "loggingin")
	End If
End Function

' Is Logged In
Function IsLoggedIn()
	If Not EW_DEBUG_ENABLED Then On Error Resume Next
	If IsObject(Security) Then
		IsLoggedIn = Security.IsLoggedIn
	Else
		IsLoggedIn = (Session(EW_SESSION_STATUS) = "login")
	End If
End Function

' Is System Admin
Function IsSysAdmin()
	If Not EW_DEBUG_ENABLED Then On Error Resume Next
	If IsObject(Security) Then
		IsSysAdmin = Security.IsSysAdmin
	Else
		IsSysAdmin = (Session(EW_SESSION_SYS_ADMIN) = 1)
	End If
End Function

' Get current page object
Function CurrentPage()
	If Not (Page Is Nothing) Then
		Set CurrentPage = Page
	Else
		Set CurrentPage = Nothing
	End If
End Function

' Get current table object
Function CurrentTable()
	If Not (Table Is Nothing) Then
		Set CurrentTable = Table
	Else
		Set CurrentTable = Nothing
	End If
End Function

' Get current master table object
Function CurrentMasterTable()
	Dim tbl
	Set tbl = CurrentTable()
	If Not (tbl Is Nothing) Then
		Set CurrentMasterTable = tbl.MasterTable
	Else
		Set CurrentMasterTable = Nothing
	End If
End Function
%>
<%

' Get server variable by name
Function ew_GetServerVariable(Name)
	ew_GetServerVariable = Request.ServerVariables(Name)
End Function

' Get user IP
Function ew_CurrentUserIP()
	ew_CurrentUserIP = ew_GetServerVariable("REMOTE_HOST")
End Function

' Get current host name, e.g. "www.mycompany.com"
Function ew_CurrentHost()
	ew_CurrentHost = ew_GetServerVariable("HTTP_HOST")
End Function

' Get current date in default date format
Function ew_CurrentDate()
	ew_CurrentDate = Date
	Select Case EW_DEFAULT_DATE_FORMAT
		Case 5, 9, 12, 15
			ew_CurrentDate = ew_FormatDateTime(ew_CurrentDate, 5)
		Case 6, 10, 13, 16
			ew_CurrentDate = ew_FormatDateTime(ew_CurrentDate, 6)
		Case 7, 11, 14, 17
			ew_CurrentDate = ew_FormatDateTime(ew_CurrentDate, 7)
	End Select
	If EW_DATE_SEPARATOR <> "/" Then ew_CurrentDate = Replace(ew_CurrentDate, EW_DATE_SEPARATOR, "/")
End Function

' Get current time in hh:mm:ss format
Function ew_CurrentTime()
	Dim DT
	DT = Now()
	ew_CurrentTime = ew_ZeroPad(Hour(DT), 2) & ":" & _
		ew_ZeroPad(Minute(DT), 2) & ":" & ew_ZeroPad(Second(DT), 2)
End Function

' Get current date in default date format with
' Current time in hh:mm:ss format
Function ew_CurrentDateTime()
	ew_CurrentDateTime = ew_CurrentDate() & " " & ew_CurrentTime()
End Function

' Get current date in standard format (yyyy/mm/dd)
Function ew_StdCurrentDate()
	ew_StdCurrentDate = ew_StdDate(Date)
End Function

' Get date in standard format (yyyy/mm/dd)
Function ew_StdDate(dt)
	ew_StdDate = ew_ZeroPad(Year(dt), 4) & "/" & ew_ZeroPad(Month(dt), 2) & "/" & ew_ZeroPad(Day(dt), 2)
End Function

' Get current date and time in standard format (yyyy/mm/dd hh:mm:ss)
Function ew_StdCurrentDateTime()
	ew_StdCurrentDateTime = ew_StdDateTime(Now)
End Function

' Get date/time in standard format (yyyy/mm/dd hh:mm:ss)
Function ew_StdDateTime(dt)
	ew_StdDateTime = ew_ZeroPad(Year(dt), 4) & "/" & ew_ZeroPad(Month(dt), 2) & "/" & ew_ZeroPad(Day(dt), 2) & " " & _
		ew_ZeroPad(Hour(dt), 2) & ":" & ew_ZeroPad(Minute(dt), 2) & ":" & ew_ZeroPad(Second(dt), 2)
End Function

' Remove XSS
Function ew_RemoveXSS(val)
	Dim regEx, search, ra, i, j, Found, val_before, pattern, replacement

	' Handle null value
	If IsNull(val) Then
		ew_RemoveXSS = val
		Exit Function
	End If

	' Remove all non-printable characters. CR(0a) and LF(0b) and TAB(9) are allowed 
	' This prevents some character re-spacing such as <java\0script> 
	' Note that you have to handle splits with \n, \r, and \t later since they *are* allowed in some inputs

	Set regEx = New RegExp ' Create regular expression.
	regEx.IgnoreCase = True
	regEx.Global = True
	regEx.Pattern = "([\x00-\x08][\x0b-\x0c][\x0e-\x20])"
	val = regEx.Replace(val & "", "")

	' Straight replacements, the user should never need these since they're normal characters 
	' This prevents like <IMG SRC=&#X40&#X61&#X76&#X61&#X73&#X63&#X72&#X69&#X70&#X74&#X3A&#X61&#X6C&#X65&#X72&#X74&#X28&#X27&#X58&#X53&#X53&#X27&#X29> 

	search = "abcdefghijklmnopqrstuvwxyz"
	search = search & "ABCDEFGHIJKLMNOPQRSTUVWXYZ" 
	search = search & "1234567890!@#$%^&*()" 
	search = search & "~`"";:?+/={}[]-_|'\"
	For i = 1 To Len(search)

		' ;? matches the ;, which is optional 
		' 0{0,7} matches any padded zeros, which are optional and go up to 8 chars 
		' &#x0040 @ search for the hex values

		regEx.Pattern = "(&#[x|X]0{0,8}" & Hex(Asc(Mid(search, i, 1))) & ";?)" ' With a ;
		val = regEx.Replace(val, Mid(search, i, 1))

		' &#00064 @ 0{0,7} matches '0' zero to seven times
		regEx.Pattern = "(&#0{0,8}" & Asc(Mid(search, i, 1)) & ";?)" ' With a ;
		val = regEx.Replace(val, Mid(search, i, 1))
	Next

	' Now the only remaining whitespace attacks are \t, \n, and \r 
	ra = EW_XSS_ARRAY
	Found = True ' Keep replacing as long as the previous round replaced something 
	Do While Found
		val_before = val
		For i = 0 To UBound(ra)
			pattern = ""
			For j = 1 To Len(ra(i))
				If j > 1 Then
					pattern = pattern & "("
					pattern = pattern & "(&#[x|X]0{0,8}([9][a][b]);?)?"
					pattern = pattern & "|(&#0{0,8}([9][10][13]);?)?"
					pattern = pattern & ")?"
				End If
				pattern = pattern & Mid(ra(i), j, 1)
			Next
			replacement = Mid(ra(i), 1, 2) & "<x>" & Mid(ra(i), 3) ' Add in <> to nerf the tag
			regEx.Pattern = pattern
			val = regEx.Replace(val, replacement) ' Filter out the hex tags
			If val_before = val Then

				' No replacements were made, so exit the loop
				Found = False
			End If
		Next
	Loop
	ew_RemoveXSS = val
End Function

' Copy file
Function ew_CopyFile(src, dest)
	On Error Resume Next
	Dim fso
	Set fso = Server.Createobject("Scripting.FileSystemObject")
	If fso.FileExists(src) Then
		fso.CopyFile src, dest, True
		ew_CopyFile = (Err.Number = 0)
	Else
		ew_CopyFile = False
	End If
	Set fso = Nothing
End Function
%>
<%

' ---------------------------
'  Get upload object (begin)
'
Function ew_GetUploadObj()
		Set ew_GetUploadObj = New cUploadObj
End Function

'
'  Get upload object (end)
' -------------------------

%>
<%

' Save binary to file
Function ew_SaveFile(folder, fn, filedata)
	If Not EW_DEBUG_ENABLED Then On Error Resume Next
	Dim oStream
	ew_SaveFile = False
	If Not ew_SaveFileByComponent(folder, fn, filedata) Then
		If ew_CreateFolder(folder) Then
			Set oStream = Server.CreateObject("ADODB.Stream")
			oStream.Type = 1 ' 1=adTypeBinary
			oStream.Open
			oStream.Write ew_ConvertToBinary(filedata)
			oStream.SaveToFile folder & fn, 2 ' 2=adSaveCreateOverwrite
			oStream.Close
			Set oStream = Nothing
			If Err.Number = 0 Then ew_SaveFile = True
		End If
	End If
End Function

' Convert raw to binary
Function ew_ConvertToBinary(rawdata)
	Dim oRs
	Set oRs = Server.CreateObject("ADODB.Recordset")

	' Create field in an empty RecordSet
	Call oRs.Fields.Append("Blob", 205, LenB(rawdata)) ' Add field with type adLongVarBinary
	Call oRs.Open()
	Call oRs.AddNew()
	Call oRs.Fields("Blob").AppendChunk(rawdata & ChrB(0))
	Call oRs.Update()

	' Save Blob Data
	ew_ConvertToBinary = oRs.Fields("Blob").GetChunk(LenB(rawdata))

	' Close RecordSet
	Call oRs.Close()
	Set oRs = Nothing
End Function
%>
<%

' Resize binary to thumbnail
Function ew_ResizeBinary(filedata, width, height, interpolation)
	ew_ResizeBinary = False ' No resize
End Function

' Resize file to thumbnail file
Function ew_ResizeFile(fn, tn, width, height, interpolation)
	If Not EW_DEBUG_ENABLED Then On Error Resume Next
	Dim fso

	' Just copy across
	Set fso = Server.Createobject("Scripting.FileSystemObject")
	If fso.FileExists(fn) Then
		fso.CopyFile fn, tn, True
	End If
	Set fso = Nothing
	ew_ResizeFile = True
End Function

' Resize file to binary
Function ew_ResizeFileToBinary(fn, width, height, interpolation)
	If Not EW_DEBUG_ENABLED Then On Error Resume Next
	Dim oStream, fso
	ew_ResizeFileToBinary = Null

	' Return file content in binary
	Set fso = Server.Createobject("Scripting.FileSystemObject")
	If fso.FileExists(fn) Then
		Set oStream = Server.CreateObject("ADODB.Stream")
		oStream.Type = 1 ' 1=adTypeBinary
		oStream.Open
		oStream.LoadFromFile fn
		ew_ResizeFileToBinary = oStream.Read
		oStream.Close
		Set oStream = Nothing
	End If
	Set fso = Nothing
End Function

' Save file by component
Function ew_SaveFileByComponent(folder, fn, filedata)
	ew_SaveFileByComponent = False
End Function
%>
<script language="JScript" runat="server">
// Server-side JScript functions for ASPMaker 7+ (Requires script engine 5.5.+)
// encrytion key
EW_RANDOM_KEY = 'zLb0kZF6o3yggBY0';
function ew_Encode(str) {	
	return encodeURIComponent(str);
}
function ew_Decode(str) {	
	return decodeURIComponent(str);	
}
// JavaScript implementation of Block TEA by Chris Veness
// http://www.movable-type.co.uk/scripts/TEAblock.html
//
// TEAencrypt: Use Corrected Block TEA to encrypt plaintext using password
//            (note plaintext & password must be strings not string objects)
//
// Return encrypted text as string
//
function TEAencrypt(plaintext, password)
{
    if (plaintext.length == 0) return('');  // nothing to encrypt
    // 'escape' plaintext so chars outside ISO-8859-1 work in single-byte packing, but  
    // keep spaces as spaces (not '%20') so encrypted text doesn't grow too long, and 
    // convert result to longs
    var v = strToLongs(escape(plaintext).replace(/%20/g,' '));
    if (v.length == 1) v[1] = 0;  // algorithm doesn't work for n<2 so fudge by adding nulls
    var k = strToLongs(password.slice(0,16));  // simply convert first 16 chars of password as key
    var n = v.length;
    var z = v[n-1], y = v[0], delta = 0x9E3779B9;
    var mx, e, q = Math.floor(6 + 52/n), sum = 0;
    while (q-- > 0) {  // 6 + 52/n operations gives between 6 & 32 mixes on each word
        sum += delta;
        e = sum>>>2 & 3;
        for (var p = 0; p < n-1; p++) {
            y = v[p+1];
            mx = (z>>>5 ^ y<<2) + (y>>>3 ^ z<<4) ^ (sum^y) + (k[p&3 ^ e] ^ z)
            z = v[p] += mx;
        }
        y = v[0];
        mx = (z>>>5 ^ y<<2) + (y>>>3 ^ z<<4) ^ (sum^y) + (k[p&3 ^ e] ^ z)
        z = v[n-1] += mx;
    }
    // note use of >>> in place of >> due to lack of 'unsigned' type in JavaScript 
    return escCtrlCh(longsToStr(v));
}
//
// TEAdecrypt: Use Corrected Block TEA to decrypt ciphertext using password
//
function TEAdecrypt(ciphertext, password)
{
    if (ciphertext.length == 0) return('');
    var v = strToLongs(unescCtrlCh(ciphertext));
    var k = strToLongs(password.slice(0,16)); 
    var n = v.length;
    var z = v[n-1], y = v[0], delta = 0x9E3779B9;
    var mx, e, q = Math.floor(6 + 52/n), sum = q*delta;
    while (sum != 0) {
        e = sum>>>2 & 3;
        for (var p = n-1; p > 0; p--) {
            z = v[p-1];
            mx = (z>>>5 ^ y<<2) + (y>>>3 ^ z<<4) ^ (sum^y) + (k[p&3 ^ e] ^ z)
            y = v[p] -= mx;
        }
        z = v[n-1];
        mx = (z>>>5 ^ y<<2) + (y>>>3 ^ z<<4) ^ (sum^y) + (k[p&3 ^ e] ^ z)
        y = v[0] -= mx;
        sum -= delta;
    }
    var plaintext = longsToStr(v);
    // strip trailing null chars resulting from filling 4-char blocks:
    if (plaintext.search(/\0/) != -1) plaintext = plaintext.slice(0, plaintext.search(/\0/));
    return unescape(plaintext);
}
// supporting functions
function strToLongs(s) {  // convert string to array of longs, each containing 4 chars
    // note chars must be within ISO-8859-1 (with Unicode code-point < 256) to fit 4/long
    var l = new Array(Math.ceil(s.length/4))
    for (var i=0; i<l.length; i++) {
        // note little-endian encoding - endianness is irrelevant as long as 
        // it is the same in longsToStr() 
        l[i] = s.charCodeAt(i*4) + (s.charCodeAt(i*4+1)<<8) + 
               (s.charCodeAt(i*4+2)<<16) + (s.charCodeAt(i*4+3)<<24);
    }
    return l;  // note running off the end of the string generates nulls since 
}              // bitwise operators treat NaN as 0
function longsToStr(l) {  // convert array of longs back to string
    var a = new Array(l.length);
    for (var i=0; i<l.length; i++) {
        a[i] = String.fromCharCode(l[i] & 0xFF, l[i]>>>8 & 0xFF, 
                                   l[i]>>>16 & 0xFF, l[i]>>>24 & 0xFF);
    }
    return a.join('');  // use Array.join() rather than repeated string appends for efficiency
}
function escCtrlCh(str) {  // escape control chars which might cause problems with encrypted texts
    return str.replace(/[\0\n\v\f\r!]/g, function(c) { return '!' + c.charCodeAt(0) + '!'; });
}
function unescCtrlCh(str) {  // unescape potentially problematic nulls and control characters
    return str.replace(/!\d\d?!/g, function(c) { return String.fromCharCode(c.slice(1,-1)); });
}
</script>
<script language="JScript" src="js/ewvalidator.js" runat="server"></script>
