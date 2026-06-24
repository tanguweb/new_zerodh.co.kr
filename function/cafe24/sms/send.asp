<%
Dim sBASE_64_CHARACTERS_cafe24_sms, sBASE_64_CHARACTERSansi_cafe24_sms
sBASE_64_CHARACTERS_cafe24_sms = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
sBASE_64_CHARACTERSansi_cafe24_sms	= strUnicode2Ansi_cafe24_sms(sBASE_64_CHARACTERS_cafe24_sms)

Function strUnicodeLen_cafe24_sms(asContents)
	Dim asContents1	: asContents1 ="a"	& asContents
	Dim Len1		: Len1=Len(asContents1)
	Dim K		: K=0
	Dim I, Asc1

	For I=1 To Len1
		Asc1	= asc(mid(asContents1,I,1))
		IF Asc1 < 0 Then Asc1	= 65536 + Asc1
		IF Asc1 > 255 Then
			K	= K + 2
		ELSE
			K	= K + 1
		End IF
	Next

	strUnicodeLen_cafe24_sms = K - 1
End Function

Function strUnicode2Ansi_cafe24_sms(asContents)
	Dim Len1			: Len1 = Len(asContents)
	Dim I, VarCHAR, VarASC, VarHEX, VarLOW, VarHIGH

	strUnicode2Ansi_cafe24_sms	= ""

	For I = 1 to Len1
		VarCHAR	= Mid(asContents,I,1)
		VarASC	= Asc(VarCHAR)
		IF VarASC < 0 Then VarASC = VarASC + 65536
		IF VarASC > 255 Then
			VarHEX		= Hex(VarASC)
			VarLOW		= Left(VarHEX,2)
			VarHIGH		= Right(VarHEX,2)
			strUnicode2Ansi_cafe24_sms	= strUnicode2Ansi_cafe24_sms & ChrB("&H" & VarLOW ) & ChrB("&H" & VarHIGH )
		Else
			strUnicode2Ansi_cafe24_sms	= strUnicode2Ansi_cafe24_sms & ChrB(VarASC)
		End IF
	Next
End Function

Function strAnsi2Unicode_cafe24_sms(asContents)
	Dim Len1			: Len1		= LenB(asContents)
	Dim VarCHAR, VarASC, I

	strAnsi2Unicode_cafe24_sms	= ""

	IF Len1=0 Then	Exit Function

	For I=1 To Len1
		VarCHAR	= MidB(asContents,I,1)
		VarASC	= AscB(VarCHAR)
		IF VarASC > 127 Then
			strAnsi2Unicode_cafe24_sms	= strAnsi2Unicode_cafe24_sms & Chr(AscW(MidB(asContents, I+1,1) & VarCHAR))
			I = I + 1
		Else
			strAnsi2Unicode_cafe24_sms	= strAnsi2Unicode_cafe24_sms & Chr(VarASC)
		End IF
	Next

End function

Function Base64encode_cafe24_sms(asContents)
	Dim lnPosition
	Dim lsResult
	Dim Char1
	Dim Char2
	Dim Char3
	Dim Char4
	Dim Byte1
	Dim Byte2
	Dim Byte3
	Dim SaveBits1
	Dim SaveBits2
	Dim lsGroupBinary
	Dim lsGroup64
	Dim M3, M4, Len1, Len2

	Len1			=LenB(asContents)

	IF Len1 < 1 Then
		Base64encode_cafe24_sms	= ""
		Exit Function
	End IF

	M3=Len1 Mod 3

	IF M3 > 0 Then asContents = asContents & String(3 - M3, ChrB(0))

	IF m3 > 0 Then
		Len1	= Len1 + (3 - M3)
		Len2	= Len1 - 3
	Else
		Len2	= Len1
	End IF

	lsResult	= ""

	For lnPosition = 1 To Len2 Step 3
		lsGroup64	= ""
		lsGroupBinary	= MidB(asContents, lnPosition, 3)

		Byte1		= AscB(MidB(lsGroupBinary, 1, 1))	: SaveBits1	= Byte1 And 3
		Byte2		= AscB(MidB(lsGroupBinary, 2, 1))	: SaveBits2	= Byte2 And 15
		Byte3		= AscB(MidB(lsGroupBinary, 3, 1))

		Char1		= MidB(sBASE_64_CHARACTERSansi_cafe24_sms, ((Byte1 And 252) \ 4) + 1, 1)
		Char2		= MidB(sBASE_64_CHARACTERSansi_cafe24_sms, (((Byte2 And 240) \ 16) Or (SaveBits1 * 16) And &HFF) + 1, 1)
		Char3		= MidB(sBASE_64_CHARACTERSansi_cafe24_sms, (((Byte3 And 192) \ 64) Or (SaveBits2 * 4) And &HFF) + 1, 1)
		Char4		= MidB(sBASE_64_CHARACTERSansi_cafe24_sms, (Byte3 And 63) + 1, 1)
		lsGroup64	= Char1 & Char2 & Char3 & Char4

		lsResult		= lsResult & lsGroup64
	Next

	IF M3 > 0 Then
		lsGroup64	= ""
		lsGroupBinary	= MidB(asContents, Len2 + 1, 3)

		Byte1		= AscB(MidB(lsGroupBinary, 1, 1))	: SaveBits1	= Byte1 And 3
		Byte2		= AscB(MidB(lsGroupBinary, 2, 1))	: SaveBits2	= Byte2 And 15
		Byte3		= AscB(MidB(lsGroupBinary, 3, 1))

		Char1		= MidB(sBASE_64_CHARACTERSansi_cafe24_sms, ((Byte1 And 252) \ 4) + 1, 1)
		Char2		= MidB(sBASE_64_CHARACTERSansi_cafe24_sms, (((Byte2 And 240) \ 16) Or (SaveBits1 * 16) And &HFF) + 1, 1)
		Char3		= MidB(sBASE_64_CHARACTERSansi_cafe24_sms, (((Byte3 And 192) \ 64) Or (SaveBits2 * 4) And &HFF) + 1, 1)

		IF M3=1 Then
			lsGroup64	= Char1 & Char2 & ChrB(61) & ChrB(61)
		Else
			lsGroup64	= Char1 & Char2 & Char3 & ChrB(61)
		End IF

		lsResult		= lsResult & lsGroup64
	End IF

	Base64encode_cafe24_sms = lsResult
End Function

Function Base64decode_cafe24_sms(asContents)
	   Dim lsResult
	   Dim lnPosition
	   Dim lsGroup64, lsGroupBinary
	   Dim Char1, Char2, Char3, Char4
	   Dim Byte1, Byte2, Byte3
	   Dim M4, Len1, Len2

	   Len1	   = LenB(asContents)
	   M4	   = Len1 Mod 4

	   IF Len1 < 1 Or M4 > 0 Then
		   Base64decode_cafe24_sms = ""
		   Exit Function
	   End IF

	   IF MidB(asContents, Len1, 1) = ChrB(61) Then	   M4 = 3
	   IF MidB(asContents, Len1-1, 1) = ChrB(61) Then	   M4 = 2

	   IF M4 = 0 Then
		   Len2	   = Len1
	   Else
		   Len2	   = Len1 - 4
	   End IF

	   For lnPosition = 1 To Len2 Step 4
		   lsGroupBinary	   = ""
		   lsGroup64	   = MidB(asContents, lnPosition, 4)

		   Char1	   	   = InStrB(sBASE_64_CHARACTERSansi_cafe24_sms, MidB(lsGroup64, 1, 1)) - 1
		   Char2	   	   = InStrB(sBASE_64_CHARACTERSansi_cafe24_sms, MidB(lsGroup64, 2, 1)) - 1
		   Char3	   	   = InStrB(sBASE_64_CHARACTERSansi_cafe24_sms, MidB(lsGroup64, 3, 1)) - 1
		   Char4	   	   = InStrB(sBASE_64_CHARACTERSansi_cafe24_sms, MidB(lsGroup64, 4, 1)) - 1

		   Byte1	   	   = ChrB(((Char2 And 48) \ 16) Or (Char1 * 4) And &HFF)
		   Byte2	   	   = lsGroupBinary & ChrB(((Char3 And 60) \ 4) Or (Char2 * 16) And &HFF)
		   Byte3	   	   = ChrB((((Char3 And 3) * 64) And &HFF) Or (Char4 And 63))
		   lsGroupBinary	   = Byte1 & Byte2 & Byte3

		   lsResult	   	   = lsResult & lsGroupBinary
	   Next

	   IF M4 > 0 Then
		   lsGroupBinary	   = ""
		   lsGroup64	   = MidB(asContents, Len2 + 1, M4) & ChrB(65)
		   IF M4=2 Then
			   lsGroup64	   = lsGroup64 & chrB(65)
		   End IF
		   Char1	   = InStrB(sBASE_64_CHARACTERSansi_cafe24_sms, MidB(lsGroup64, 1, 1)) - 1
		   Char2	   = InStrB(sBASE_64_CHARACTERSansi_cafe24_sms, MidB(lsGroup64, 2, 1)) - 1
		   Char3	   = InStrB(sBASE_64_CHARACTERSansi_cafe24_sms, MidB(lsGroup64, 3, 1)) - 1
		   Char4	   = InStrB(sBASE_64_CHARACTERSansi_cafe24_sms, MidB(lsGroup64, 4, 1)) - 1

		   Byte1	   = ChrB(((Char2 And 48) \ 16) Or (Char1 * 4) And &HFF)
		   Byte2	   = lsGroupBinary & ChrB(((Char3 And 60) \ 4) Or (Char2 * 16) And &HFF)
		   Byte3	   = ChrB((((Char3 And 3) * 64) And &HFF) Or (Char4 And 63))

		   IF M4=2 Then
			   lsGroupBinary	   = Byte1
		   elseIF M4=3 Then
			   lsGroupBinary	   = Byte1 & Byte2
		   end IF

		   lsResult	   	   	   = lsResult & lsGroupBinary
	   End IF

	   Base64decode_cafe24_sms	   	   	   = lsResult
End Function


Function sendSMS_cafe24 (p_sms_msg, p_sms_rphone, p_sms_testflag)


	'수신자 휴대폰 문자열 치환
	p_sms_rphone = Replace(p_sms_rphone," ","")
	'p_sms_rphone = Replace(p_sms_rphone,"-","")
	p_sms_rphone = Replace(p_sms_rphone,".","-")
	p_sms_rphone = Replace(p_sms_rphone,"/","-")
	
	Dim p_sms_rphone_prefix : p_sms_rphone_prefix = Left(p_sms_rphone,3)		'수신자 휴대폰 앞자리

	Dim p_is_valid_prefix : p_is_valid_prefix = False	'수신자 휴대폰 앞자리 유효성

	If p_sms_rphone_prefix = "010" Or p_sms_rphone_prefix = "011" Or p_sms_rphone_prefix = "016" Or p_sms_rphone_prefix = "017" Or p_sms_rphone_prefix = "018" Or p_sms_rphone_prefix = "019"  Then
		p_is_valid_prefix = True 
	End If 

	If p_sms_msg <> "" and p_sms_rphone <> "" And p_is_valid_prefix = True Then 

		If p_sms_testflag <> "Y" Then 
			p_sms_testflag = ""
		End If 

		' 원본 소스는 카페24 에서 확인할 수 있음.
		Dim sms_url
			sms_url = "https://sslsms.cafe24.com/sms_sender.php" ' SMS 요청 URL

		Dim user_id	    : user_id	= m_cafe24_sms_user_id   'SMS 아이디 (peterssms)		' config.asp 에 선언할 것
		Dim secure	    : secure	= m_cafe24_sms_secure	'"인증후 사용하세요"  '인증키
		Dim encoderurl  : encoderurl = "Y" '리턴 URL을 encode 해서 받을지를 결정합니다. (사용:Y, 사용안함:N, Y가 아닐 경우 변수를 여러개 넘겨받을 없습니다.)
		Dim subject     : subject="" '제목(LMS일 경우만)
		'Dim msg	        : msg	= request.Form("msg")
		Dim msg	        : msg	= p_sms_msg
		'Dim rphone	    : rphone	= request.Form("rphone")
		Dim rphone	    : rphone	= p_sms_rphone
		'Dim sphone1	    : sphone1	= request.Form("sphone1")	
		'Dim sphone2	    : sphone2	= request.Form("sphone2")
		'Dim sphone3	    : sphone3	= request.Form("sphone3")
		Dim sphone1	    : sphone1	= m_cafe24_sms_send_phone1	' config.asp 에 선언할 것
		Dim sphone2	    : sphone2	= m_cafe24_sms_send_phone2	' config.asp 에 선언할 것
		Dim sphone3	    : sphone3	= m_cafe24_sms_send_phone3	' config.asp 에 선언할 것
		Dim rdate	    : rdate	= ""	'request.Form("rdate")
		Dim reserveTime	: reserveTime	= ""	'request.Form("rtime")
		Dim mode	    : mode	= "1"  '// base64 사용시 반드시 모드값을 1로 주셔야 합니다.
		Dim rtime	    : rtime	= ""	'request.Form("rtime")
		Dim returnurl	: returnurl	= ""	'request.Form("returnurl")
		Dim testflag	: testflag	= p_sms_testflag	'request.Form("testflag")
		Dim destination	: destination	= ""	'request.Form("destination")
		Dim repeatFlag	: repeatFlag	= ""	'request.Form("repeatFlag")
		Dim repeatNum	: repeatNum	= ""	'request.Form("repeatNum")
		Dim repeatTime	: repeatTime	= ""	'request.Form("repeatTime")
		Dim actionFlag  : actionFlag = ""	'request("action")
		Dim nointeractive  : nointeractive = "1"	'request("nointeractive")  '성공시 대화 상자를 사용 하지 않게 합니다.
		Dim smsType	: smsType	= "S"	'request.Form("smsType") 'LMS 사용시 L
		IF smsType="L" Then
			'subject = request.Form("subject")
			subject =  strAnsi2Unicode_cafe24_sms(Base64encode_cafe24_sms(strUnicode2Ansi_cafe24_sms(subject)))
		End IF

'		user_id  = strAnsi2Unicode_cafe24_sms(Base64encode_cafe24_sms(strUnicode2Ansi_cafe24_sms(user_id)))
'		secure  = strAnsi2Unicode_cafe24_sms(Base64encode_cafe24_sms(strUnicode2Ansi_cafe24_sms(secure)))
'		msg  = strAnsi2Unicode_cafe24_sms(Base64encode_cafe24_sms(strUnicode2Ansi_cafe24_sms(msg)))
'		rphone  = strAnsi2Unicode_cafe24_sms(Base64encode_cafe24_sms(strUnicode2Ansi_cafe24_sms(rphone)))
'		sphone1  = strAnsi2Unicode_cafe24_sms(Base64encode_cafe24_sms(strUnicode2Ansi_cafe24_sms(sphone1)))
'		sphone2  = strAnsi2Unicode_cafe24_sms(Base64encode_cafe24_sms(strUnicode2Ansi_cafe24_sms(sphone2)))
'		sphone3  = strAnsi2Unicode_cafe24_sms(Base64encode_cafe24_sms(strUnicode2Ansi_cafe24_sms(sphone3)))
'		rdate  = strAnsi2Unicode_cafe24_sms(Base64encode_cafe24_sms(strUnicode2Ansi_cafe24_sms(rdate)))
'		reserveTime  = strAnsi2Unicode_cafe24_sms(Base64encode_cafe24_sms(strUnicode2Ansi_cafe24_sms(reserveTime)))
'		mode  = strAnsi2Unicode_cafe24_sms(Base64encode_cafe24_sms(strUnicode2Ansi_cafe24_sms(mode)))
'		rtime  = strAnsi2Unicode_cafe24_sms(Base64encode_cafe24_sms(strUnicode2Ansi_cafe24_sms(rtime)))
'		returnurl  = strAnsi2Unicode_cafe24_sms(Base64encode_cafe24_sms(strUnicode2Ansi_cafe24_sms(returnurl)))
'		testflag  = strAnsi2Unicode_cafe24_sms(Base64encode_cafe24_sms(strUnicode2Ansi_cafe24_sms(testflag)))
'		destination  = strAnsi2Unicode_cafe24_sms(Base64encode_cafe24_sms(strUnicode2Ansi_cafe24_sms(destination)))
'		repeatFlag  = strAnsi2Unicode_cafe24_sms(Base64encode_cafe24_sms(strUnicode2Ansi_cafe24_sms(repeatFlag)))
'		repeatNum  = strAnsi2Unicode_cafe24_sms(Base64encode_cafe24_sms(strUnicode2Ansi_cafe24_sms(repeatNum)))
'		repeatTime  = strAnsi2Unicode_cafe24_sms(Base64encode_cafe24_sms(strUnicode2Ansi_cafe24_sms(repeatTime)))
'		smsType  = strAnsi2Unicode_cafe24_sms(Base64encode_cafe24_sms(strUnicode2Ansi_cafe24_sms(smsType)))

		Dim sendurl : sendurl = "http://" & Request.ServerVariables("server_name") & request.ServerVariables("PATH_INFO")


		' 발송하기
		Dim PostData
		PostData = "lang="
		PostData = PostData & "&sendurl=" & sendurl
		PostData = PostData & "&user_id=" & user_id
		PostData = PostData & "&secure=" & secure
		PostData = PostData & "&subject=" & subject
		PostData = PostData & "&msg=" & msg
		PostData = PostData & "&rphone=" & rphone
		PostData = PostData & "&sphone1=" & sphone1
		PostData = PostData & "&sphone2=" & sphone2
		PostData = PostData & "&sphone3=" & sphone3
		PostData = PostData & "&rdate=" & rdate
		PostData = PostData & "&rtime=" & rtime
		PostData = PostData & "&reserveTime=" & reserveTime
		PostData = PostData & "&mode=" & mode
		PostData = PostData & "&returnurl=" & returnurl
		PostData = PostData & "&testflag=" & testflag
		PostData = PostData & "&destination=" & destination
		PostData = PostData & "&repeatFlag=" & repeatFlag
		PostData = PostData & "&repeatNum=" & repeatNum
		PostData = PostData & "&repeatTime=" & repeatTime
		PostData = PostData & "&nointeractive=" & nointeractive
		PostData = PostData & "&encoderurl=" & encoderurl
		PostData = PostData & "&smsType=" & smsType

		'Response.Write("PostData : " & PostData & "<br />")
		'sendSMS_cafe24 = msg
		'Response.End 
		
		
		Dim ServerXmlHttp
		Dim tmpResult

		Set ServerXmlHttp = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
		ServerXmlHttp.open "POST", sms_url
		ServerXmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		ServerXmlHttp.setRequestHeader "Content-Length", Len(PostData)

		ServerXmlHttp.send PostData
		If ServerXmlHttp.status = 200 Then
			tmpResult = ServerXmlHttp.responseText
		Else
			tmpResult = "Connection Fail"
			' Handle missing response or other errors here
		End If

		Set ServerXmlHttp = Nothing
		
		'Response.Write("tmpResult : " & tmpResult & "<br />")

		Dim rMsg    : rMsg	=  split(tmpResult , ",")
		Dim Result    : Result	=  rMsg (0)
		Dim Count    : Count	=  rMsg (1)
		Dim alert      : alert  = ""
		SELECT CASE Result
			CASE "Test Success!"
				alert = "테스트성공"
				alert = alert & " 잔여건수는 "+ Count+"건 입니다."
			CASE "success"
				alert = "문자 정상적으로 전송되었습니다."
				alert = alert & " 잔여건수 : "+ Count+"건 "
			CASE "reserved"
				alert = "예약되었습니다."
				alert = alert & " 잔여건수는 "+ Count+"건 입니다."
			CASE "3205"
				alert = "잘못된 번호형식입니다."
			CASE "0044"
				alert = "스팸문자는 보낼 수 없습니다."
			CASE Else
				alert = "[Error]"+Result
		END Select

		If nointeractive="1" Then
			If Result  = "Test Success!" Or Result  = "success" Or Result  = "reserved" Then
			Else
				'Response.Write("<script>alert('" + alert + "')</script>")
				'Response.Write("Result : " & Result & ", alert : " & alert & "<br />")
			End if
		Else
			'Response.Write("<script>alert('" + alert + "')</script>")
			'Response.Write("Result : " & Result & ", alert : " & alert & "<br />")
		End if
		IF returnurl <> "" Then
			'Response.Write("<script>location.href='"+request.Form("returnurl")+"?result=" +tmpResult+"';</script>")
			'Response.Write("Result : " & Result & ", tmpResult : " & tmpResult & "<br />")
		ELSE
			'Response.Write("<script>history.go(-1);</script>")
			'Response.Write("Result : " & Result & ", alert : " & alert & "<br />")
		END If
		
		'sendSMS_cafe24 = Result & alert & tmpResult
		sendSMS_cafe24 = Result
	Else 
		sendSMS_cafe24 = ""	

	End If 
		

End Function 



%>
                