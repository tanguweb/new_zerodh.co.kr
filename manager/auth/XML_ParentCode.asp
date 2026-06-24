<!-- #include virtual = "/Include/Config.asp" -->
<%
Dim objADO, SQL, i
Dim NO_DEPTH, CD_PFORM
Dim aryCD_PFORM

NO_DEPTH = trim(Request("NO_DEPTH"))
CD_PFORM = trim(Request("CD_PFORM"))

'상위메뉴 정보
If CD_PFORM = "" Then
	Set objADO = new clsADO

	objADO.setConString(m_DB)
	SQL = ""
	SQL = SQL & " SELECT CD_FORM, NM_FORM			      "
	SQL = SQL & " FROM T_MENU							  "
	SQL = SQL & " WHERE NO_DEPTH = '" & NO_DEPTH - 1 & "' "

	objADO.setSql(SQL)
	aryCD_PFORM = objADO.getArrRs()

	Set objADO = Nothing 
End If

		Response.Charset="utf-8"
		Response.Write "<?xml version=""1.0"" encoding=""utf-8"" ?>" & vbcrlf
		Response.Write "<root>" &vbcrlf


'상위메뉴 정보
If IsArray(aryCD_PFORM) Then
		Response.Write "	<CD_PFORM_Infos>" &vbcrlf
	For i = 0 To UBound(aryCD_PFORM,2)
		Response.Write "	<CD_PFORM_Info>" &vbcrlf
		Response.Write "		<CD_PFORM>"&aryCD_PFORM(0,i)&"</CD_PFORM>" &vbcrlf 
		Response.Write "		<NM_FORM>"&aryCD_PFORM(1,i)&"</NM_FORM>" &vbcrlf 
		Response.Write "	</CD_PFORM_Info>" &vbcrlf
	Next
Response.Write "	</CD_PFORM_Infos>" &vbcrlf
End If

Response.Write "</root>"

%>