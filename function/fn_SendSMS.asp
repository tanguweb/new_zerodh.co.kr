<%
'##################################################################################################################################################
'#	File		: /function/fn_SendSMS.asp
'#  Create		: 조영준 / 2011.08.05
'#	Info		: SMS 발송 함수(카페24용)
'#	Update		: 
'#	Update Memo	: 
'#				  
'#	[변수]									[파라미터]			[설명]
'#	-----------------------------------------------------------------------------------------------------------------------------------------------
'#	Dim msg	        	 : msg				= msg_				'// 전송메세지
'#	Dim rphone	    	 : rphone			= rphone_			'// 받는 번호(010-1111-2222 , "-" 포함 / 여러명일 경우 ",(콤마)"로 구분 / ※ destination 변수 사용시에는 "" 으로 비워주세요 )
'#	Dim sphone1	    	 : sphone1			= sphone1_			'// 보내는번호 1 (02)
'#	Dim sphone2	    	 : sphone2			= sphone2_			'// 보내는번호 2 (123)
'#	Dim sphone3	    	 : sphone3			= sphone3_			'// 보내는번호 3 (1234)
'#	Dim rdate	    	 : rdate			= rdate_			'// 예약날짜 예)20090909
'#	Dim reserveTime		 : reserveTime		= rtime_			'// 예약시간 예)173000 ,오후 5시 30분,예약시간은 최소 10분 이상으로 설정.
'#	Dim mode			 : mode				= "1"				'// base64 사용시 반드시 모드값을 1로 주셔야 합니다.
'#	Dim rtime			 : rtime			= rtime_			'// 예약시간 예)173000 ,오후 5시 30분,예약시간은 최소 10분 이상으로 설정.
'#	Dim returnurl_sms : returnurl_sms	= returnurl_sms_	'// 리턴주소
'#	Dim testflag		 : testflag			= testflag_			'// 예) 테스트시: Y, 테스트가 아닐 경우 입력하지 마세요.
'#	Dim destination		 : destination		= destination_		'// 이름삽입번호 예) 010-000-0000|홍길동 (여러명일 경우 ",(콤마)"로 구분 / msg값에 “{name}” 이라는 문구를 입력 후 전송하시면 됩니다.)
'#																	예)
'#																	<input type="type" name="destination" value="010-000-0000|홍길동,010-000-0000|김영희">
'#																	<input type="type" name="msg" value="{name}님, 주문하신 물품이 배송되었습니다.">
'#	Dim repeatFlag		 : repeatFlag		= repeatFlag_		'// 반복설정 예) 1~10회 가능.
'#	Dim repeatNum		 : repeatNum		= repeatNum_		'// 반복횟수 예) 분마다
'#	Dim repeatTime		 : repeatTime		= repeatTime_		'// 전송간격 예)15분 이상부터 가능
'#	Dim actionFlag		 : actionFlag		= action_			'// go
'#	Dim nointeractive	 : nointeractive	= nointeractive_	'// 예) 사용할 경우 : 1, 성공시 대화상자(alert)를 생략.
'#
'##################################################################################################################################################

	Public Function fn_SendSMS( msg_ , rphone_ , sphone1_ , sphone2_ , sphone3_ , rdate_ , rtime_ , returnurl_sms_ , returnurl_ , actionFlag_)

		'################################## 사용 샘플 ##################################

		Dim sms_url
			sms_url = "http://sslsms.cafe24.com/sms_sender.php" ' SMS 요청 URL

		Dim user_id	    : user_id	= "elevendent1"   'SMS 아이디
		Dim secure	    : secure	= "79ae9113575b7d8322866fd29ee2259f"  '인증키
		Dim encoderurl  : encoderurl = "Y" '리턴 URL을 encode 해서 받을지를 결정합니다. (사용:Y, 사용안함:N, Y가 아닐 경우 변수를 여러개 넘겨받을 없습니다.)
	'	Dim msg	        : msg	= request.Form("msg")
	'	Dim rphone	    : rphone	= request.Form("rphone")
	'	Dim sphone1	    : sphone1	= request.Form("sphone1")
	'	Dim sphone2	    : sphone2	= request.Form("sphone2")
	'	Dim sphone3	    : sphone3	= request.Form("sphone3")
	'	Dim rdate	    : rdate	= request.Form("rdate")
	'	Dim reserveTime	: reserveTime	= request.Form("rtime")
	'	Dim mode	    : mode	= "1"  '// base64 사용시 반드시 모드값을 1로 주셔야 합니다.
	'	Dim rtime	    : rtime	= request.Form("rtime")
	'	Dim returnurl_sms	: returnurl_sms	= request.Form("returnurl_sms")
	'	Dim testflag	: testflag	= request.Form("testflag")
	'	Dim destination	: destination	= request.Form("destination")
	'	Dim repeatFlag	: repeatFlag	= request.Form("repeatFlag")
	'	Dim repeatNum	: repeatNum	= request.Form("repeatNum")
	'	Dim repeatTime	: repeatTime	= request.Form("repeatTime")
	'	Dim actionFlag  : actionFlag = request("action")
	'	Dim nointeractive  : nointeractive = request("nointeractive")  '성공시 대화 상자를 사용 하지 않게 합니다.
		
'		Response.Write "=== 파라미터 ==== <br />"
'		Response.Write "msg : " & 	msg_   & "<br />"
'		Response.Write "rphone : " & 	rphone_  & "<br />"
'		Response.Write "sphone1 : " & 	sphone1_ & "<br />"
'		Response.Write "sphone2 : " & 	sphone2_ & "<br />"
'		Response.Write "sphone3 : " & 	sphone3_ & "<br />"
'		Response.Write "rdate : " & 	rdate_   & "<br />"
'		Response.Write "rtime : " & 	rtime_   & "<br />"
'		Response.Write "returnurl_sms : " & 	returnurl_sms_  & "<br />"
'		Response.Write "actionFlag : " & 	actionFlag_ & "<br />"


		Dim msg	        	 : msg				= msg_
		Dim rphone	    	 : rphone			= rphone_
		Dim sphone1	    	 : sphone1			= sphone1_
		Dim sphone2	    	 : sphone2			= sphone2_
		Dim sphone3	    	 : sphone3			= sphone3_
		Dim rdate	    	 : rdate			= rdate_
		Dim reserveTime		 : reserveTime		= rtime_
		Dim mode			 : mode				= "1"  '// base64 사용시 반드시 모드값을 1로 주셔야 합니다.
		Dim rtime			 : rtime			= rtime_
		Dim returnurl_sms	 : returnurl_sms	= returnurl_sms_
		'Dim testflag		 : testflag			= "Y"	'// 테스트 발송시
		Dim testflag		 : testflag			= ""	'// 리얼 발송
		Dim destination		 
		Dim repeatFlag		 
		Dim repeatNum		 
		Dim repeatTime		 
		Dim actionFlag		 : actionFlag		= actionFlag_
		Dim nointeractive	 '성공시 대화 상자를 사용 하지 않게 합니다.

'		Response.Write "=== 암호화 전 ==== <br />"
'		Response.Write "user_id : " & user_id & "<br />"
'		Response.Write "secure : " & 	secure & "<br />"
'		Response.Write "msg : " & 	msg   & "<br />"
'		Response.Write "rphone : " & 	rphone  & "<br />"
'		Response.Write "sphone1 : " & 	sphone1 & "<br />"
'		Response.Write "sphone2 : " & 	sphone2 & "<br />"
'		Response.Write "sphone3 : " & 	sphone3 & "<br />"
'		Response.Write "rdate : " & 	rdate   & "<br />"
'		Response.Write "reserveTime : " & 	reserveTime & "<br />"
'		Response.Write "mode : " & 	mode   & "<br />"
'		Response.Write "rtime : " & 	rtime   & "<br />"
'		Response.Write "returnurl_sms : " & 	returnurl_sms  & "<br />"
'		Response.Write "testflag : " & 	testflag   & "<br />"
'		Response.Write "destination : " & 	destination & "<br />"
'		Response.Write "repeatFlag : " & 	repeatFlag & "<br />"
'		Response.Write "repeatNum : " & 	repeatNum  & "<br />"
'		Response.Write "repeatTime : " & 	repeatTime & "<br />"
'		Response.Write "actionFlag : " & 	actionFlag & "<br />"
'		Response.Write "nointeractive : " & 	nointeractive & "<br /><br />"

'		user_id  = strAnsi2Unicode(Base64encode_sms(strUnicode2Ansi(user_id)))
'		secure  = strAnsi2Unicode(Base64encode_sms(strUnicode2Ansi(secure)))
'		msg  = strAnsi2Unicode(Base64encode_sms(strUnicode2Ansi(msg)))
'		rphone  = strAnsi2Unicode(Base64encode_sms(strUnicode2Ansi(rphone)))
'		sphone1  = strAnsi2Unicode(Base64encode_sms(strUnicode2Ansi(sphone1)))
'		sphone2  = strAnsi2Unicode(Base64encode_sms(strUnicode2Ansi(sphone2)))
'		sphone3  = strAnsi2Unicode(Base64encode_sms(strUnicode2Ansi(sphone3)))
'		rdate  = strAnsi2Unicode(Base64encode_sms(strUnicode2Ansi(rdate)))
'		reserveTime  = strAnsi2Unicode(Base64encode_sms(strUnicode2Ansi(reserveTime)))
'		mode  = strAnsi2Unicode(Base64encode_sms(strUnicode2Ansi(mode)))
'		rtime  = strAnsi2Unicode(Base64encode_sms(strUnicode2Ansi(rtime)))
'		returnurl_sms  = strAnsi2Unicode(Base64encode_sms(strUnicode2Ansi(returnurl_sms)))
'		testflag  = strAnsi2Unicode(Base64encode_sms(strUnicode2Ansi(testflag)))
'		destination  = strAnsi2Unicode(Base64encode_sms(strUnicode2Ansi(destination)))
'		repeatFlag  = strAnsi2Unicode(Base64encode_sms(strUnicode2Ansi(repeatFlag)))
'		repeatNum  = strAnsi2Unicode(Base64encode_sms(strUnicode2Ansi(repeatNum)))
'		repeatTime  = strAnsi2Unicode(Base64encode_sms(strUnicode2Ansi(repeatTime)))
 

		
'		Response.Write "=== 암호화 후 ==== <br />"
'		Response.Write "user_id : " & user_id & "<br />"
'		Response.Write "secure : " & 	secure & "<br />"
'		Response.Write "msg : " & 	msg   & "<br />"
'		Response.Write "rphone : " & 	rphone  & "<br />"
'		Response.Write "sphone1 : " & 	sphone1 & "<br />"
'		Response.Write "sphone2 : " & 	sphone2 & "<br />"
'		Response.Write "sphone3 : " & 	sphone3 & "<br />"
'		Response.Write "rdate : " & 	rdate   & "<br />"
'		Response.Write "reserveTime : " & 	reserveTime & "<br />"
'		Response.Write "mode : " & 	mode   & "<br />"
'		Response.Write "rtime : " & 	rtime   & "<br />"
'		Response.Write "returnurl_sms : " & 	returnurl_sms  & "<br />"
'		Response.Write "testflag : " & 	testflag   & "<br />"
'		Response.Write "destination : " & 	destination & "<br />"
'		Response.Write "repeatFlag : " & 	repeatFlag & "<br />"
'		Response.Write "repeatNum : " & 	repeatNum  & "<br />"
'		Response.Write "repeatTime : " & 	repeatTime & "<br />"
'		Response.Write "actionFlag : " & 	actionFlag & "<br />"
'		Response.Write "nointeractive : " & 	nointeractive & "<br /><br />"
'		Response.End 
		

'		user_id  = strAnsi2Unicode(Base64decode_sms(strUnicode2Ansi(user_id)))
'		secure  = strAnsi2Unicode(Base64decode_sms(strUnicode2Ansi(secure)))
'		msg  = strAnsi2Unicode(Base64decode_sms(strUnicode2Ansi(msg)))
'		rphone  = strAnsi2Unicode(Base64decode_sms(strUnicode2Ansi(rphone)))
'		sphone1  = strAnsi2Unicode(Base64decode_sms(strUnicode2Ansi(sphone1)))
'		sphone2  = strAnsi2Unicode(Base64decode_sms(strUnicode2Ansi(sphone2)))
'		sphone3  = strAnsi2Unicode(Base64decode_sms(strUnicode2Ansi(sphone3)))
'		rdate  = strAnsi2Unicode(Base64decode_sms(strUnicode2Ansi(rdate)))
'		reserveTime  = strAnsi2Unicode(Base64decode_sms(strUnicode2Ansi(reserveTime)))
'		mode  = strAnsi2Unicode(Base64decode_sms(strUnicode2Ansi(mode)))
'		rtime  = strAnsi2Unicode(Base64decode_sms(strUnicode2Ansi(rtime)))
'		returnurl_sms  = strAnsi2Unicode(Base64decode_sms(strUnicode2Ansi(returnurl_sms)))
'		testflag  = strAnsi2Unicode(Base64decode_sms(strUnicode2Ansi(testflag)))
'		destination  = strAnsi2Unicode(Base64decode_sms(strUnicode2Ansi(destination)))
'		repeatFlag  = strAnsi2Unicode(Base64decode_sms(strUnicode2Ansi(repeatFlag)))
'		repeatNum  = strAnsi2Unicode(Base64decode_sms(strUnicode2Ansi(repeatNum)))
'		repeatTime  = strAnsi2Unicode(Base64decode_sms(strUnicode2Ansi(repeatTime)))

'		Response.Write "=== 복호화 후 ==== <br />"
'		Response.Write "user_id : " & user_id & "<br />"
'		Response.Write "secure : " & 	secure & "<br />"
'		Response.Write "msg : " & 	msg   & "<br />"
'		Response.Write "rphone : " & 	rphone  & "<br />"
'		Response.Write "sphone1 : " & 	sphone1 & "<br />"
'		Response.Write "sphone2 : " & 	sphone2 & "<br />"
'		Response.Write "sphone3 : " & 	sphone3 & "<br />"
'		Response.Write "rdate : " & 	rdate   & "<br />"
'		Response.Write "reserveTime : " & 	reserveTime & "<br />"
'		Response.Write "mode : " & 	mode   & "<br />"
'		Response.Write "rtime : " & 	rtime   & "<br />"
'		Response.Write "returnurl_sms : " & 	returnurl_sms  & "<br />"
'		Response.Write "testflag : " & 	testflag   & "<br />"
'		Response.Write "destination : " & 	destination & "<br />"
'		Response.Write "repeatFlag : " & 	repeatFlag & "<br />"
'		Response.Write "repeatNum : " & 	repeatNum  & "<br />"
'		Response.Write "repeatTime : " & 	repeatTime & "<br />"
'		Response.Write "actionFlag : " & 	actionFlag & "<br />"
'		Response.Write "nointeractive : " & 	nointeractive & "<br /><br />"
'		Response.End 

'		Dim sendurl : sendurl = "http://" & Request.ServerVariables("server_name") & request.ServerVariables("PATH_INFO")
		Dim sendurl : sendurl = "http://" & Request.ServerVariables("HTTP_HOST") & request.ServerVariables("PATH_INFO")

'		Response.Write sendurl
'		Response.End 
	%>
		<form id='frmSendsms' name='frmSendsms' method='post' action="<%=sms_url%>">
		<input type="hidden" name="lang" id="lang" value="asp">
		<input type="hidden" name="sendurl" id="sendurl" value="<%=sendurl%>">
		<input type="hidden" name="user_id" id="user_id" value="<%=user_id%>">
		<input type="hidden" name="secure" id="secure" value="<%=secure%>">
		<input type="hidden" name="msg" id="msg" value="<%=msg%>">
		<input type="hidden" name="rphone" id="rphone" value="<%=rphone%>">
		<input type="hidden" name="sphone1" id="sphone1" value="<%=sphone1%>">
		<input type="hidden" name="sphone2" id="sphone2" value="<%=sphone2%>">
		<input type="hidden" name="sphone3" id="sphone3" value="<%=sphone3%>">
		<input type="hidden" name="rdate" id="rdate" value="<%=rdate%>">
		<input type="hidden" name="rtime" id="rtime" value="<%=rtime%>">
		<input type="hidden" name="reserveTime" id="reserveTime" value="<%=reserveTime%>">
		<input type="hidden" name="mode" id="mode" value="<%=mode%>">
		<input type="hidden" name="returnurl" id="returnurl" value="<%=returnurl_sms%>">
		<input type="hidden" name="testflag" id="testflag" value="<%=testflag%>">
		<input type="hidden" name="destination" id="destination" value="<%=destination%>">
		<input type="hidden" name="repeatFlag" id="repeatFlag" value="<%=repeatFlag%>">
		<input type="hidden" name="repeatNum" id="repeatNum" value="<%=repeatNum%>">
		<input type="hidden" name="repeatTime" id="repeatTime" value="<%=repeatTime%>">
		<input type="hidden" name="nointeractive" id="nointeractive" value="<%=nointeractive%>">
		<input type="hidden" name="encoderurl" id="encoderurl" value="<%=encoderurl%>">
		</form>
	<%
		' 발송하기
		If actionFlag="go" Then
			response.write "<script type='text/javascript'>var frm = document.getElementById('frmSendsms');frm.submit();</script>"
			Response.End


		' 발송결과
		ElseIf actionFlag="result" Then
			Dim tmpResult   : tmpResult  = request("result")
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
				CASE "-105"
					alert = "예약시간은 최소 15분 이상으로 설정해주세요."
				CASE Else
					alert = "[Error]"+Result
			END Select

			'리턴 URL을 encode 해서 받았을 경우 다시 decode해준다.
			'returnurl_sms = request("returnurl_sms")
			returnurl_sms = returnurl_

			If encoderurl="Y" Then
				returnurl_sms = strAnsi2Unicode(Base64decode_sms(strUnicode2Ansi(returnurl_sms)))
			End If

			If nointeractive="1" Then
				If Result  = "Test Success!" Or Result  = "success" Or Result  = "reserved" Then
				Else
					Response.Write("<script>alert('" + alert + "')</script>")
				End if
			Else
				Response.Write("<script>alert('" + alert + "')</script>")
			End if
			Response.Write("<script>location.href='"+returnurl_sms+"&rtn="+Result+"';</script>")
		End if

	End Function


	'################################## 사용 샘플(카페24 제공 base24 암호화) ##################################

	Dim sBASE_64_CHARACTERS_sms, sBASE_64_CHARACTERSansi_sms
	sBASE_64_CHARACTERS_sms = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
	sBASE_64_CHARACTERSansi_sms	= strUnicode2Ansi(sBASE_64_CHARACTERS_sms)

	Function strUnicodeLen(asContents)
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

		strUnicodeLen = K - 1
	End Function

	Function strUnicode2Ansi(asContents)
		Dim Len1			: Len1 = Len(asContents)
		Dim I, VarCHAR, VarASC, VarHEX, VarLOW, VarHIGH

		strUnicode2Ansi	= ""

		For I = 1 to Len1
			VarCHAR	= Mid(asContents,I,1)
			VarASC	= Asc(VarCHAR)
			IF VarASC < 0 Then VarASC = VarASC + 65536
			IF VarASC > 255 Then
				VarHEX		= Hex(VarASC)
				VarLOW		= Left(VarHEX,2)
				VarHIGH		= Right(VarHEX,2)
				strUnicode2Ansi	= strUnicode2Ansi & ChrB("&H" & VarLOW ) & ChrB("&H" & VarHIGH )
			Else
				strUnicode2Ansi	= strUnicode2Ansi & ChrB(VarASC)
			End IF
		Next
	End Function

	Function strAnsi2Unicode(asContents)
		Dim Len1			: Len1		= LenB(asContents)
		Dim VarCHAR, VarASC, I

		strAnsi2Unicode	= ""

		IF Len1=0 Then	Exit Function

		For I=1 To Len1
			VarCHAR	= MidB(asContents,I,1)
			VarASC	= AscB(VarCHAR)
			IF VarASC > 127 Then
				strAnsi2Unicode	= strAnsi2Unicode & Chr(AscW(MidB(asContents, I+1,1) & VarCHAR))
				I = I + 1
			Else
				strAnsi2Unicode	= strAnsi2Unicode & Chr(VarASC)
			End IF
		Next

	End function

	Function Base64encode_sms(asContents)
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
			Base64encode_sms	= ""
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

			Char1		= MidB(sBASE_64_CHARACTERSansi_sms, ((Byte1 And 252) \ 4) + 1, 1)
			Char2		= MidB(sBASE_64_CHARACTERSansi_sms, (((Byte2 And 240) \ 16) Or (SaveBits1 * 16) And &HFF) + 1, 1)
			Char3		= MidB(sBASE_64_CHARACTERSansi_sms, (((Byte3 And 192) \ 64) Or (SaveBits2 * 4) And &HFF) + 1, 1)
			Char4		= MidB(sBASE_64_CHARACTERSansi_sms, (Byte3 And 63) + 1, 1)
			lsGroup64	= Char1 & Char2 & Char3 & Char4

			lsResult		= lsResult & lsGroup64
		Next

		IF M3 > 0 Then
			lsGroup64	= ""
			lsGroupBinary	= MidB(asContents, Len2 + 1, 3)

			Byte1		= AscB(MidB(lsGroupBinary, 1, 1))	: SaveBits1	= Byte1 And 3
			Byte2		= AscB(MidB(lsGroupBinary, 2, 1))	: SaveBits2	= Byte2 And 15
			Byte3		= AscB(MidB(lsGroupBinary, 3, 1))

			Char1		= MidB(sBASE_64_CHARACTERSansi_sms, ((Byte1 And 252) \ 4) + 1, 1)
			Char2		= MidB(sBASE_64_CHARACTERSansi_sms, (((Byte2 And 240) \ 16) Or (SaveBits1 * 16) And &HFF) + 1, 1)
			Char3		= MidB(sBASE_64_CHARACTERSansi_sms, (((Byte3 And 192) \ 64) Or (SaveBits2 * 4) And &HFF) + 1, 1)

			IF M3=1 Then
				lsGroup64	= Char1 & Char2 & ChrB(61) & ChrB(61)
			Else
				lsGroup64	= Char1 & Char2 & Char3 & ChrB(61)
			End IF

			lsResult		= lsResult & lsGroup64
		End IF

		Base64encode_sms = lsResult
	End Function

	Function Base64decode_sms(asContents)
		   Dim lsResult
		   Dim lnPosition
		   Dim lsGroup64, lsGroupBinary
		   Dim Char1, Char2, Char3, Char4
		   Dim Byte1, Byte2, Byte3
		   Dim M4, Len1, Len2

		   Len1	   = LenB(asContents)
		   M4	   = Len1 Mod 4

		   IF Len1 < 1 Or M4 > 0 Then
			   Base64decode_sms = ""
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

			   Char1	   	   = InStrB(sBASE_64_CHARACTERSansi_sms, MidB(lsGroup64, 1, 1)) - 1
			   Char2	   	   = InStrB(sBASE_64_CHARACTERSansi_sms, MidB(lsGroup64, 2, 1)) - 1
			   Char3	   	   = InStrB(sBASE_64_CHARACTERSansi_sms, MidB(lsGroup64, 3, 1)) - 1
			   Char4	   	   = InStrB(sBASE_64_CHARACTERSansi_sms, MidB(lsGroup64, 4, 1)) - 1

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
			   Char1	   = InStrB(sBASE_64_CHARACTERSansi_sms, MidB(lsGroup64, 1, 1)) - 1
			   Char2	   = InStrB(sBASE_64_CHARACTERSansi_sms, MidB(lsGroup64, 2, 1)) - 1
			   Char3	   = InStrB(sBASE_64_CHARACTERSansi_sms, MidB(lsGroup64, 3, 1)) - 1
			   Char4	   = InStrB(sBASE_64_CHARACTERSansi_sms, MidB(lsGroup64, 4, 1)) - 1

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

		   Base64decode_sms	   	   	   = lsResult
	End Function
%>