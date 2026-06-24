<!-- #include virtual = "/Include/Config.asp" -->
<!-- #include virtual = "/function/fn_AdminLoginCheck.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<%
	'########################################################################################
	'#	File		: /manager/auth/adminList_ifm_Proc.assp
	'#  Create		: 조영준 / 2011.01.10
	'#	Info		: 사용자별 사용 메뉴 관리
	'#	Update		: 
	'#	Update Memo	: 
	'########################################################################################

	'---------------------------------------------------
	' 변수선언 및 셋팅
	'---------------------------------------------------
	Dim objADO, Rs, SQL
	Dim iloop, jloop, kloop
	Dim MODE, RTN_STR, CD_ADMINID
	Dim CHECK_MENU	'체크된 메뉴값
	Dim ArrFormList 'CD_FORM(메뉴 코드) 리스트
	
	MODE		 = ReqF("MODE")				'MODE (I:등록,U:수정,D:삭제)
	RTN_STR		 = ReqF("RTN_STR")			'리턴 URL
	CD_ADMINID = ReqF("CD_ADMINID")		'회원 고유 아이디


	'// BEGIN : 메뉴 코드 리스트 가져오기
	Set objADO = new clsADO
		objADO.setConString(m_DB)

		SQL = ""
		SQL = SQL & " SELECT CD_FORM FROM T_MENU ORDER BY NO_ORDER ASC"

		objADO.setSql(SQL)
		ArrFormList = objADO.getArrRs()
	Set objADO = Nothing
	'// END : 메뉴 코드 리스트 가져오기



	'// 메뉴 코드 리스트의 수 만큼 배열 설정(사용중인 리스트 최대)
	ReDim arrCHECK_MENU(UBound(ArrFormList,2) + 1)



	'// 메뉴코드로 POST 값을 CHECK_MENU에 저장 -> @ 으로 Split 한 값을 arrCHECK_MENU 배열에 저장
	For iloop = 0 To UBound(ArrFormList,2)
		'// Form 값 있는 것만 배열에 저장
		If Req(ArrFormList(0,iloop)) <> "" And Req(ArrFormList(0,iloop)) <> Empty Then 
			CHECK_MENU = Split(Req(ArrFormList(0,iloop)),"@")
			arrCHECK_MENU(iloop) = CHECK_MENU(0)
		End If 
	Next


	'// Request.Form 값 확인
'	response.write "MODE : " & MODE & "<br>"
'	response.write "RTN_STR : " & RTN_STR & "<br>"
'	response.write "CD_ADMINID : " & CD_ADMINID & "<br>"
'	For iloop = 0 To UBound(ArrFormList,2) 
'	response.write arrCHECK_MENU(iloop) & "<br>"
'	Next
'	response.End


	If CD_ADMINID = "" Then 
		response.write "<script>alert('사용자를 선택해주세요.');history.back();</script>"
		response.End
	End If 


	If MODE = "" Or RTN_STR = "" Then 
		response.write "<script>alert('필요한 인수가 부족합니다.');history.back();</script>"
		response.End
	End If 



	'---------------------------------------------------
	' 데이터 
	'---------------------------------------------------	

	Set objADO = new clsADO
	
		objADO.setConString(m_DB)

		'// 사용자 고유키(CD_ADMINID)로 T_AUTH 테이블(사용자별 메뉴 리스트)에서 삭제 
		SQL = "" 
		SQL = SQL & " DELETE T_AUTH							  "
		SQL = SQL & " WHERE CD_ADMINID = '" & CD_ADMINID & "' "
		
		objADO.setSQL(SQL)
		objADO.ExecuteQuery()

		'// 쿼리문 초기화 후 '메뉴 코드','사용자 고유키' 로 저장
		SQL = ""
		For iloop = 0 To UBound(ArrFormList,2)
			If arrCHECK_MENU(iloop) <> "" And arrCHECK_MENU(iloop) <> Empty Then 
				SQL = SQL & "	INSERT INTO T_AUTH (CD_FORM, CD_ADMINID, YN_VIEW, YN_INS, YN_EDIT, YN_DEL, YN_PRINT, YN_EXCEL)		"
				SQL = SQL & "	VALUES ('" & arrCHECK_MENU(iloop) & "','" & CD_ADMINID & "',1,1,1,1,1,1) "
			End If 
		Next 

		If SQL <> "" Then 
			objADO.setSQL(SQL)
			objADO.ExecuteQuery()
		End If 

	Set objADO = Nothing

	Response.write "<script>parent.document.location.href='"&RTN_STR&"';alert('저장되었습니다.');</script>"
%>