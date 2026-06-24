<!-- #include virtual = "/Include/Config.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<%
	'########################################################################################
	'#	File		: /manager/code/code_Proc.assp
	'#  Create		: 2010.09.08
	'#	Info		: 코드테이블 저장
	'#	Update		: 
	'#	Update Memo	: 
	'########################################################################################

	'---------------------------------------------------
	' 변수선언 및 셋팅
	'---------------------------------------------------
	Dim MODE, CD_GUBUN, CD_CODE, NM_CODE, NO_ORDER, YN_USE, nm_string, RTN_STR, ary_CODE(7)
	Dim iloop, jloop, kloop
	Dim objADO, Rs, SQL

	MODE		= ReqF("MODE")
	RTN_STR		= ReqF("RTN_STR")
	CD_GUBUN	= ReqF("CD_GUBUN")
	CD_CODE		= ReqF("CD_CODE")
	nm_string	= ReqF("nm_string")

	NM_CODE		= ReqF("NM_CODE_" & nm_string)
	NO_ORDER	= ReqF("NO_ORDER_" & nm_string)
	YN_USE		= ReqF("YN_USE_" & nm_string)

for iloop = 1 to 7
	ary_CODE(iloop-1) = ReqF("NM_CODE" & iloop & "_" & nm_string)
next

	'---------------------------------------------------
	' 데이터 
	'---------------------------------------------------	

	Set objADO = new clsADO

	' 수정
	If MODE = "U" Then
		SQL = SQL & "UPDATE T_CODE SET													"
		SQL = SQL & " NM_CODE = '" & NM_CODE & "'										"

for iloop = 1 to 7
		SQL = SQL & " ,NM_CODE"&iloop&" = '" & ary_CODE(iloop-1) & "'					"
next

		SQL = SQL & " ,NO_ORDER = '" & NO_ORDER & "'									"
		SQL = SQL & " ,YN_USE = '" & YN_USE & "'										"
		SQL = SQL & "WHERE CD_GUBUN= '" & CD_GUBUN & "' AND CD_CODE = '" & CD_CODE & "'	"

	' 신규
	ElseIf MODE = "I" Then
		SQL = SQL & "	SELECT CD_GUBUN FROM T_CODE WHERE CD_GUBUN= '" & CD_GUBUN & "' AND CD_CODE = '" & CD_CODE & "'	"
		objADO.setSQL(SQL)
		Set Rs = objADO.getRs()
		
		If Not Rs.EOF Then
			Response.Write "	<script language='javascript'>"
			Response.Write "		alert('이미 등록되어진 코드구분입니다');"
			Response.Write "		history.back();"
			Response.Write "	</script>"
			Response.End
		End If

		SQL = ""
		SQL = SQL & "	INSERT INTO T_CODE (CD_GUBUN,CD_CODE,NM_CODE,NM_CODE1,NM_CODE2,NM_CODE3,NM_CODE4,NM_CODE5,NM_CODE6,NM_CODE7,NO_ORDER,YN_USE)"
		SQL = SQL & "	VALUES ('" & CD_GUBUN & "','" & CD_CODE & "','" & NM_CODE & "'"
for iloop=0 to 6		
		SQL = SQL & "			,'" & ary_CODE(iloop) & "'"
next
		SQL = SQL & "			,'" & NO_ORDER & "','" & YN_USE & "')"

	End If

	objADO.setSQL(SQL)
	objADO.ExecuteQuery()

	Set objADO = Nothing

	Response.Redirect RTN_STR
%>