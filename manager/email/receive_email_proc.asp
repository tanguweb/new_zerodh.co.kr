<!-- #include virtual = "/Include/Config.asp" -->
<!-- #include virtual = "/function/fn_AdminLoginCheck.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<%	
	'########################################################################################
	'#	File		: 
	'#  Create		: 
	'#	Info		: 
	'#	Update		: 
	'#	Update Memo	: 
	'########################################################################################
	
	Dim objADO, Rs, SQL
	Dim adoCmd, Result
	
	Dim MODE : MODE = ReqF("MODE")
	Dim CD_CODE
	Dim NM_CODE
	Dim NM_CODE1
	Dim NM_CODE2
	Dim NM_CODE3
	Dim YN_USE

	If MODE = "I" Then 
		CD_CODE = Year(now) & right("0"&Month(now),2) & right("0"&Day(now),2) & right("0"& Hour(now),2) & right("0"&Minute(now),2) & right("0"&Second(now),2)
		NM_CODE = CD_CODE
		NM_CODE1 = ReqF("NM_CODE1")
		NM_CODE2 = ReqF("NM_CODE2")
		NM_CODE3 = ReqF("NM_CODE3")
		YN_USE = ReqF("YN_USE") 
	ElseIf MODE = "U" Then 
		CD_CODE = ReqF("update_CD_CODE")
		NM_CODE1 = ReqF("NM_CODE1_" & CD_CODE)
		NM_CODE2 = ReqF("NM_CODE2_" & CD_CODE)
		NM_CODE3 = ReqF("NM_CODE3_" & CD_CODE)
		YN_USE = ReqF("YN_USE_" & CD_CODE)
	ElseIf MODE = "D" Then 
		CD_CODE = ReqF("update_CD_CODE")
	End If 

	If MODE = "I" Then 
		
		Set objADO = new clsADO
		SQL = ""
		SQL = SQL & "	INSERT INTO T_CODE (CD_GUBUN,CD_CODE,NM_CODE,NM_CODE1,NM_CODE2,NM_CODE3,NM_CODE4,NM_CODE5,NM_CODE6,NM_CODE7,NO_ORDER,YN_USE)"
		SQL = SQL & "	VALUES ('COUNSEL_MAIL_RECV','"&CD_CODE&"','"&NM_CODE&"','"&NM_CODE1&"','"&NM_CODE2&"','"&NM_CODE3&"','','','','',100,'"&YN_USE&"')"
		objADO.setSQL(SQL)
		objADO.ExecuteQuery()

		Set objADO = Nothing
		
		Response.Write "<script>alert('저장되었습니다.');parent.document.location.reload();</script>"

	ElseIf MODE = "U" Then 
		
		Set objADO = new clsADO
		SQL = ""
		SQL = SQL & "	UPDATE T_CODE	"
		SQL = SQL & "	SET	"
		SQL = SQL & "		NM_CODE1 = '" & NM_CODE1 & "' "
		SQL = SQL & "		, NM_CODE2 = '" & NM_CODE2  & "' "
		SQL = SQL & "		, NM_CODE3 = '" & NM_CODE3 & "' "
		SQL = SQL & "		, YN_USE = '" & YN_USE & "' "
		SQL = SQL & "	WHERE CD_GUBUN = 'COUNSEL_MAIL_RECV' AND CD_CODE = '"&CD_CODE&"' "
		objADO.setSQL(SQL)
		objADO.ExecuteQuery()
		Set objADO = Nothing
		
		Response.Write "<script>alert('수정되었습니다.');parent.document.location.reload();</script>"

	ElseIf MODE = "D" Then 
		
		Set objADO = new clsADO
		SQL = ""
		SQL = SQL & "	DELETE T_CODE	"
		SQL = SQL & "	WHERE CD_GUBUN = 'COUNSEL_MAIL_RECV' AND CD_CODE = '"&CD_CODE&"' "
		objADO.setSQL(SQL)
		objADO.ExecuteQuery()
		Set objADO = Nothing

		Response.Write "<script>alert('삭제되었습니다.');parent.document.location.reload();</script>"

	End If 

	'Response.Write "<script>alert('저장되었습니다.');parent.document.location.reload();</script>"
	Response.End 
%>