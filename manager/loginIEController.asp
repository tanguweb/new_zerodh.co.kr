<!-- #include virtual = "/Include/Config.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<%
	'########################################################################################
	'#	File		: /manager/loginIEController.asp
	'#  Create		: domoyosi / 2015.06.30
	'#	Info		: 어드민 로그인처리페이지
	'#	Update		: 
	'#	Update Memo	: 
	'########################################################################################

	Dim userid, password, primaryid, usernm
	Dim objADO, SQL, aryLoginInfo

	userid = ReqF("userid")
	password = ReqF("password")

	'response.write userid & " || " & password
	'response.end 
	
	Set objADO = new clsADO
	
	SQL = ""
	SQL = SQL & "SELECT																	"
	SQL = SQL & "	CD_ADMINID, NM_ADMINPW, NM_NAME, NO_TEL, NM_EMAIL, YN_LOGIN, YN_USE	"
	SQL = SQL & "FROM T_ADMIN															"
	SQL = SQL & "WHERE YN_USE = 'Y' AND YN_LOGIN = 'Y' 									"
	SQL = SQL & "	AND CD_ADMINID = '" & userid & "'									"			
	SQL = SQL & "	AND NM_ADMINPW = '" & lcase(password) & "'							"

	objADO.setConString(m_DB)
	objADO.setSql(SQL)
	aryLoginInfo = objADO.getArrRs()

	Set objADO = Nothing

	If IsArray(aryLoginInfo) Then
		Response.Cookies("member")					= ""
		Response.Cookies("member")("m_UserID")		= fn_AESEnc(aryLoginInfo(0,0))
		'Response.Cookies("member")("m_UserNM")		= fn_AESEnc(aryLoginInfo(2,0))
		Response.Cookies("member")("m_UserNM")		= aryLoginInfo(2,0)
		Response.Cookies("member").Path = "/"
		Response.Cookies("member").domain = CookieDomain

		Response.Write "<script>document.location.href='/manager/index.asp';</script>"
		Response.End
	Else
		Response.Write "<script>alert('회원정보가 일치하지 않습니다.');history.back();</script>"
		Response.End 
	End If 
%>