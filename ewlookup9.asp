<%@ CodePage="65001" LCID="2057" EnableSessionState="False" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="ewcfg9.asp"-->
<!--#include file="aspfn9.asp"-->
<%
Call ew_Header(False, "utf-8")
Dim lookup
Set lookup = New clookup
lookup.Page_Main()
Set lookup = Nothing

' Page class for lookup
Class clookup
	Dim QS, Sql, Value, arValue, LnkFldType

	' Page ID
	Public Property Get PageID()
		PageID = "lookup"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "lookup"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
	End Property

	' Main
	Sub Page_Main()
		On Error Resume Next
		QS = Split(Request.Querystring, "&")
		If IsArray(QS) Then
			If UBound(QS) > 0 Then
				Sql = GetValue("s")
				Sql = TEAdecrypt(Sql, EW_RANDOM_KEY)
				If Sql <> "" Then

					' Get the filter values (for "IN")
					Value = ew_AdjustSql(GetValue("f"))
					If Value <> "" Then
						arValue = Split(Value, ",")
						LnkFldType = GetValue("lft") ' Link field data type
						If IsNumeric(LnkFldType) Then LnkFldType = CInt(LnkFldType)
						For ari = 0 To UBound(arValue)
							arValue(ari) = ew_QuotedValue(arValue(ari), LnkFldType)
						Next
						Sql = Replace(Sql, "{filter_value}", Join(arValue, ","))
					End If

					' get the query value (for "LIKE" or "=")
					Value = ew_AdjustSql(GetValue("q"))
					If Value <> "" Then
						If (InStr(Sql, "LIKE '{query_value}%'") > 0) Then
							Sql = Replace(Sql, "LIKE '{query_value}%'", ew_Like("'" & Value & "%'"))
						Else
							Sql = Replace(Sql, "{query_value}", Value)
						End If
					End If
					GetLookupValues(Sql)
				End If
			End If
		Else
			Err.raise 500
		End If
	End Sub

	' Get raw querystring value
	Function GetValue(Key)
		Dim kv, i
		For i = 0 To UBound(QS)
			kv = Split(QS(i), "=")
			If (kv(0) = Key) Then
				GetValue = ew_Decode(kv(1))
				Exit Function
			End If
		Next
		GetValue = ""
	End Function

	' Get values from database
	Sub GetLookupValues(Sql)
		On Error Resume Next

		' Connect to database
		Dim Rs, RsArr, str, i, j
		Call ew_Connect()
		Set Rs = Conn.Execute(Sql)
		If Not Rs.EOF Then
			RsArr = Rs.GetRows
		End If

		' Close database
		Rs.Close
		Set Rs = Nothing
		Conn.Close
		Set Conn = Nothing

		' Output
		If IsArray(RsArr) Then
			For j = 0 To UBound(RsArr, 2)
				For i = 0 To UBound(RsArr, 1)
					str = RsArr(i, j) & ""
					str = RemoveDelimiters(str)
					Response.Write str & EW_FIELD_DELIMITER
				Next
				Response.Write EW_RECORD_DELIMITER
			Next
		End If
	End Sub

	' Process values
	Function RemoveDelimiters(s)
		Dim wrkstr
		wrkstr = s
		If Len(wrkstr) > 0 Then
			wrkstr = Replace(wrkstr, vbCr, " ")
			wrkstr = Replace(wrkstr, vbLf, " ")
			wrkstr = Replace(wrkstr, EW_RECORD_DELIMITER, "")
			wrkstr = Replace(wrkstr, EW_FIELD_DELIMITER, "")
		End If
		RemoveDelimiters = wrkstr
	End Function
End Class
%>
