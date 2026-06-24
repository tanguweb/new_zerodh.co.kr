<!--METADATA TYPE= "typelib"  NAME= "ADODB Type Library" FILE="C:\Program Files\Common Files\SYSTEM\ADO\msado15.dll"  -->
<%@CodePage="65001" Language="VBScript"%>
<%

	'########################################################################################
	'#	File		: /Include/Config.asp
	'#  Create		: 조영준 / 2011.08.02
	'#	Info		: Config.asp
	'#	Update		:
	'#	Update Memo	:
	'########################################################################################
	Option Explicit
	Response.Expires = -100
	Response.buffer = True
	Session.CodePage = 65001
	Response.CharSet = "utf-8"
	Response.ContentType="text/html"

	'// P3P 규약으로 도메인 포워딩 시 쿠키 적용 안되는 거 해결 헤더 입니다.
	Response.AddHeader "P3P", "CP=NOI CURa ADMa DEVa TAIa OUR DELa BUS IND PHY ONL UNI COM NAV INT DEM PRE"

	Dim m_DB : m_DB	= "Provider=SQLOLEDB.1;Password=dgzero123!!;User ID=dgzero2026;Data Source=sql19-003.cafe24.com,1433;Persist Security Info=True;Initial Catalog=dgzero2026"
%>
<!-- #include virtual = "/class/ADO_Class.asp" -->
<!-- #include virtual = "/function/fn_base.asp" -->
<!-- #include virtual = "/function/fn_Enc.asp" -->
<!-- #include virtual = "/function/fn_AES.asp" -->

<%
	'============================================================================================================================
	' 기본 변수 선언
	'============================================================================================================================
	Const m_AESkey = "rlawnsghdWkd"

	'## 쿠키 도메인
	Dim CookieDomain : CookieDomain = Request.ServerVariables("SERVER_NAME")

	'## 홈페이지 도메인
	'Dim m_SiteDomain : m_SiteDomain = "http://dgzero2026.cafe24.com/"
	Dim m_SiteDomain : m_SiteDomain = "http://" & Request.ServerVariables("SERVER_NAME") & "/"

	'Dim m_MobileDomain : m_MobileDomain = "http://dgzero2026.com/m"
	Dim m_MobileDomain : m_MobileDomain = "http://" & Request.ServerVariables("SERVER_NAME") & "/m"

	Dim m_RefererCheckHomepage : m_RefererCheckHomepage = "dgzero2026"

	'## 홈페이지 사용 변수
	Dim m_ServicePrimaryID, m_ServiceUserID, m_ServiceUserNM	'회원 ID
	Dim m_UserID	'어드민 ID
	Dim m_UserNM
	Dim m_IPAddr	'ip주소

	'어드민페이지 관리자 ID 쿠키
	If Request.Cookies("member")("m_UserID") = "" Then
		m_UserID	= ""
		m_UserNM	= ""
	Else
		m_UserID	= fn_AESDec(Request.Cookies("member")("m_UserID"))
		m_UserNM	= Request.Cookies("member")("m_UserNM")
	End If

	'프론트페이지 회원 ID 쿠키
	'If Request.Cookies("user")("m_ServicePrimaryID") = "" Then
	'	m_ServicePrimaryID	= ""
	'	m_ServiceUserID	= ""
	'	m_ServiceUserNM	= ""
	'Else
	'	m_ServicePrimaryID	= fn_AESDec(Request.Cookies("user")("m_ServicePrimaryID"))
	'	m_ServiceUserID	= fn_AESDec(Request.Cookies("user")("m_ServiceUserID"))
	'	m_ServiceUserNM	= Request.Cookies("user")("m_ServiceUserNM")
	'End If

	m_IPAddr = Request.ServerVariables("REMOTE_ADDR")

	Dim FileSaveRoot : FileSaveRoot = "F:\HOME\dgzero2026\www\file"
	Dim FileSaveRoot_Board : FileSaveRoot_Board = FileSaveRoot & "\board\"
	Dim FileSaveRoot_Img : FileSaveRoot_Img = FileSaveRoot & "\img\"
	Dim FileSave_Customer : FileSave_Customer = FileSaveRoot & "\customer\"
	Dim FileSave_Mail : FileSave_Mail = FileSaveRoot & "\mail\"	'메일 (첨부파일)
	Dim FileSave_Customer_for_smarteditor : FileSave_Customer_for_smarteditor = "/file/customer"


	'## SMTP
	Dim CONFIG_SMTP_SENDUSING : CONFIG_SMTP_SENDUSING = 2	' 1:내부, 2:외부
	Dim CONFIG_SMTP_SMTPSERVERPICKUPDIRECTORY : CONFIG_SMTP_SMTPSERVERPICKUPDIRECTORY = "C:\Inetpub\mailroot\Pickup"
	Dim CONFIG_SMTP_SMTPSERVER : CONFIG_SMTP_SMTPSERVER = "mw-002.cafe24.com"	' "localhost", "mw-002.cafe24.com"
	Dim CONFIG_SMTP_SMTPSERVERPORT : CONFIG_SMTP_SMTPSERVERPORT = 25
	Dim CONFIG_SMTP_SMTPCONNECTIONTIMEOUT : CONFIG_SMTP_SMTPCONNECTIONTIMEOUT = 30
	Dim CONFIG_SMTP_SMTPAUTHENTICATE : CONFIG_SMTP_SMTPAUTHENTICATE = 1
	Dim CONFIG_SMTP_SENDUSERNAME : CONFIG_SMTP_SENDUSERNAME = "webmaster@"	' 도메인 미설정
	Dim CONFIG_SMTP_SENDPASSWORD : CONFIG_SMTP_SENDPASSWORD = ""

	Dim m_naverMap_apiKey : m_naverMap_apiKey = ""

	'## 카페24 SMS
	Dim m_cafe24_sms_user_id : m_cafe24_sms_user_id = "cwoodentsms"
	Dim m_cafe24_sms_send_phone1 : m_cafe24_sms_send_phone1 = "010"
	Dim m_cafe24_sms_send_phone2 : m_cafe24_sms_send_phone2 = "4148"
	Dim m_cafe24_sms_send_phone3 : m_cafe24_sms_send_phone3 = "9021"
	Dim m_cafe24_sms_secure : m_cafe24_sms_secure = "9c6fd54e66737eb551cb4a575ff24b9c"

	Const m_client_domain = "http://topplant.kr"	' 도메인
	Const m_client_name_s = "탑플란트"	' 의료기관명칭 또는 상호명
	Const m_client_name = "탑플란트치과"	' 의료기관명칭 또는 상호명
	Const m_client_en = "TOPPLANT DENTAL CLINIC"	' 의료기관명칭 또는 상호명
	Const m_client_address1 = "전라북도 익산시 남중동 375-201"	' 기업 주소
	Const m_client_address2 = ""	' 기업 주소
	Const m_client_owner = "한수일"	' 기업 대표자명
	Const m_client_business_no = "628-98-00047"	' 기업 사업자등록번호
	Const m_client_tel = "063.851.2275"	' 기업 연락처
	Const m_client_fax = ""	' 기업 연락처
	Const m_client_kakao = "https://pf.kakao.com/_WAmZxl"	' 카톡주소
	Const m_client_utube = ""	' 네이버블로그
	Const m_client_naver_blog = "https://blog.naver.com/fade803742"	' 네이버블로그
	Const m_client_insta = ""	' 인스타그램
	Const m_client_facebook = ""	' 페이스북
	Const m_client_naver = "https://naver.me/xbnuoHSf"	' 네이버예약
	Const m_main_color = "#ff823e"	' 메인컬러


	Const m_client_site_copyright = "COPYRIGHTⓒ TOPPLANT DENTAL CLINIC. ALL RIGHTS RESERVED."	' 기업 연락처
	Const m_client_counsel_pw = "dgzero123!!"	'퀵 상담 디폴트 비밀번호
	Const m_client_counsel_regist_email_title = "탑플란트치과"	' 상담글 등록시 발송하는 이메일의 제목에 포함되는 상호
	Const m_client_counsel_regist_sms_title = "탑플란트치과"	' 상담글 등록시 발송하는 SMS에 포함되는 상호

	Const m_client_counsel_footer_send_email = false	' 푸터 상담 등록 메일발송 여부(True: 발송, False: 발송 안함)
	Const m_client_counsel_footer_send_sms = true	' 푸터 상담 등록 SMS발송 여부(True: 발송, False: 발송 안함)

	Const m_client_counsel_quick_send_email = false	' 퀵메뉴 상담 등록 메일발송 여부(True: 발송, False: 발송 안함)
	Const m_client_counsel_quick_send_sms = true	' 퀵메뉴 상담 등록 SMS발송 여부(True: 발송, False: 발송 안함)

	Const m_client_counsel_online_send_email = false	' 온라인 상담 등록 메일발송 여부(True: 발송, False: 발송 안함)
	Const m_client_counsel_online_send_sms = true	' 온라인 상담 등록 SMS발송 여부(True: 발송, False: 발송 안함)

	Const m_client_counsel_online_answer_email = false	' 온라인 상담 답변 메일발송 여부(True: 발송, False: 발송 안함)
	Const m_client_counsel_online_answer_sms = false	' 온라인 상담 답변 SMS발송 여부(True: 발송, False: 발송 안함)

	Dim p___domain : p___domain = "dgzero2026.cafe24.com" ' "www.topplant.kr"
%>
<% 'SSL 도메인 붙이고 주석 해제 <!-- #include virtual = "/function/fn_Redirect.asp" --> %>