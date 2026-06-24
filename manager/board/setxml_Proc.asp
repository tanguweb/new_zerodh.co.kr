<!-- #include virtual = "/Include/Config.asp" -->
<%
	'########################################################################################
	'#	File		: /manager/board/setxml_Proc.asp
	'#  Create		: / 2010.10.12
	'#	Info		: setxml_Proc.asp
	'#	Update		: 
	'#	Update Memo	: 
	'########################################################################################

	'---------------------------------------------------
	' 변수선언 및 셋팅
	'---------------------------------------------------
	Dim tablenm, CD_BOARDCD
	Dim NM_XMLPATH, NM_XMLQUERY, NM_XMLFILE
	Dim YN_XMLDISPLAY
	Dim iloop, jloop, kloop
	Dim objADO, Rs, SQL, AryXMLList, juuun

	tablenm			= ReqF("tablenm")
	CD_BOARDCD		= ReqF("sCD_BOARDCD")

	NM_XMLPATH		= ReqF("NM_XMLPATH")
	NM_XMLFILE		= ReqF("NM_XMLFILE")
	NM_XMLQUERY		= ReqF("NM_XMLQUERY")
	NM_XMLQUERY		= replace(NM_XMLQUERY,"'","''")

	YN_XMLDISPLAY	= ReqF("YN_XMLDISPLAY")

'	response.write "tablenm : " & tablenm & "<br>"
'	response.write "CD_BOARDCD : " & CD_BOARDCD & "<br>"
'	response.write "NM_XMLPATH : " & NM_XMLPATH & "<br>"
'	response.write "NM_XMLQUERY : " & NM_XMLQUERY & "<br>"
'	response.write "YN_XMLDISPLAY : " & YN_XMLDISPLAY & "<br>"
'	response.end

	if tablenm = "T_BOARDMAST" then 
		if CD_BOARDCD = "" or NM_XMLPATH="" or NM_XMLQUERY="" or NM_XMLFILE="" then 
			response.write "<script>alert('잘못된 접근입니다.');history.back();</script>"
			response.end
		end if 
	elseif tablenm = "T_BOARDDETAIL" then 
		if YN_XMLDISPLAY = "" then 
			response.write "<script>alert('잘못된 접근입니다.');history.back();</script>"
			response.end
		end if 
	else
		response.write "<script>alert('잘못된 접근입니다.');history.back();</script>"
		response.end
	end if 

	'---------------------------------------------------
	' 데이터 
	'---------------------------------------------------	

	Set objADO = new clsADO

	if tablenm = "T_BOARDMAST" then 
		
		SQL = ""
		SQL = SQL & "	UPDATE T_BOARDMAST SET							"
		SQL = SQL & "	NM_XMLPATH = '"&NM_XMLPATH&"'					"
		SQL = SQL & "	,NM_XMLQUERY = '"&NM_XMLQUERY&"'				"
		SQL = SQL & "	,NM_XMLFILE = '"&NM_XMLFILE&"'					"
		SQL = SQL & "	WHERE CD_BOARDCD = '"&CD_BOARDCD&"'				"
		objADO.setSQL(SQL)
		objADO.ExecuteQuery()

	elseif tablenm = "T_BOARDDETAIL" then 

		SQL = ""
		SQL = SQL & "	UPDATE T_BOARDDETAIL SET						"
		SQL = SQL & "	YN_XMLDISPLAY = 'N'								"
		SQL = SQL & "	,YN_USE = 'Y'									"
		SQL = SQL & "	WHERE CD_BOARDCD = '"&CD_BOARDCD&"'				"
		SQL = SQL & "													"
		SQL = SQL & "	UPDATE T_BOARDDETAIL SET						"
		SQL = SQL & "	YN_XMLDISPLAY = 'Y'								"
		SQL = SQL & "	,YN_USE = 'Y'									"
		SQL = SQL & "	WHERE CD_BOARDID IN ("&YN_XMLDISPLAY&")			"
		objADO.setSQL(SQL)
		objADO.ExecuteQuery()

		SQL = ""
		SQL = SQL & "	SELECT CD_BOARDCD, NM_BOARDNM, NM_XMLPATH, NM_XMLQUERY, NM_XMLFILE		"
		SQL = SQL & "	FROM T_BOARDMAST 														"
		SQL = SQL & "	WHERE CD_BOARDCD = '"&CD_BOARDCD&"' 									"
		SQL = SQL & "	ORDER BY NM_BOARDNM ASC													"
		objADO.setSql(SQL)
		AryXMLList = objADO.getArrRs()

	end if 
	
	Set objADO = Nothing

	if tablenm = "T_BOARDMAST" then
		response.write "<script>alert('저장되었습니다.');document.location.href='setxml.asp?sCD_BOARDCD="&CD_BOARDCD&"';</script>"
	elseif tablenm = "T_BOARDDETAIL" then 
		Set juuun = Server.CreateObject("XMLParser.XMLParser") 
		juuun.ConnectionString = m_DB_XMLLoad
		juuun.Query = AryXMLList(3,0)
		juuun.FilePath = FileSaveDefault & AryXMLList(2,0) & AryXMLList(4,0)
		juuun.SaveXML
		Set juuun = Nothing
		response.write "<script>alert('XML배포가 처리되었습니다.');document.location.href='setxml.asp?sCD_BOARDCD="&CD_BOARDCD&"';</script>"
	end if 
	response.end 
%>