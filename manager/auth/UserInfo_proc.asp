<!-- #include virtual = "/Include/Config.asp" -->
<!-- #include virtual = "/function/fn_AdminLoginCheck.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<%	
	'########################################################################################
	'#	File		: /auth/UserInfo_proc.asp
	'#  Create		: 조영준 / 2011.08.04
	'#	Info		: 관리자 비밀번호 등록
	'#	Update		: 
	'#	Update Memo	: 
	'########################################################################################
	Dim objADO, Rs, SQL
	Dim NM_ADMINPW : NM_ADMINPW = ReqF("NM_ADMINPW")
	Dim NM_ADMINPW1 : NM_ADMINPW1 = ReqF("NM_ADMINPW1")
	Dim NM_ADMINPW2 : NM_ADMINPW2 = ReqF("NM_ADMINPW2")

	Set objADO = new clsADO
		SQL = ""
		SQL = SQL & " SELECT NM_ADMINPW FROM T_ADMIN WHERE CD_ADMINID = '" & m_UserId & "' "
		objADO.setSql(SQL)
		Set Rs = objADO.getRs()

		If Not Rs.EOF Then 
			If Rs(0) <> NM_ADMINPW Then 
				Response.Write "<script>alert('기존 비밀번호가 다릅니다.');parent.document.getElementById('NM_ADMINPW').select();</script>"
				Response.End 	
			Else
				SQL = ""
				SQL = SQL & " UPDATE T_ADMIN															 "
				SQL = SQL & " SET NM_ADMINPW = '" & NM_ADMINPW2 & "'									 "
				SQL = SQL & " WHERE CD_ADMINID = '" & m_UserId & "' AND NM_ADMINPW = '" & NM_ADMINPW & "'"
				objADO.setSql(SQL)
				objADO.ExecuteQuery()

				Response.Write "<script>alert('비밀번호가 수정되었습니다.');parent.document.location.reload();</script>"
				Response.End 
			End If 
		End If 

	Set objADO = Nothing
%>