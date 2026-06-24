
<%

	' ##### SQL INJECTION 방지 #####	
	Public Function Req(RName)
		Req = RemoveInjection(Request(RName))
	End Function

	Public Function ReqF(RName)
		ReqF = RemoveInjection(Request.Form(RName))
	End Function

	Public Function ReqQ(RName)
		ReqQ = RemoveInjection(Request.QueryString(RName))
	End Function

	Public Function RemoveInjection(Value)
		'RemoveInjection = Replace(Replace(Replace(Replace(Replace(Value,"--",""),"UPDATE ",""),"INSERT ",""),"DELETE ",""),"EXECUTE ","")

		Dim BadChars, newChars, tmpChars, regEx, i
		BadChars = Array( _
			"select(.*)(from|with|by){1}", _
			"insert(.*)(into|values){1}", _
			"update(.*)set", _
			"delete(.*)(frm|with){1}", _
			"drop(.*)(from|aggre|role|assem|key|cert|cont|credential|data|endpoint|event|fulltext|function|index|login|type|schema|procedure|que|remote|role|route|sign|stat|syno|table|trigger|user|view|xml){1}", _
			"alert(.*)(application|assem|key|author|cert|credential|data|endpoint|fulltext|function|index|login|type|schema|procedure|que|remote|role|route|serv|table|user|view|xml){1}", _
			"xp_", "sp_", "restore\s", "grant\s", "revoke\s", "dbcc", "dump", "use\s", "set\s", "truncate\s", "backup\s", "load\s", "save\s", "shutdown", _
			"cast(.*)\(", "convert(.*)\(", "execute\s", "updatetext", "writetext", "reconfigure", _
			"/\*", "\*/", ";", "\-\-", "\[", "\]", "char(.*)\(", "nchar(.*)\("  )

		newChars = Value

		For i = 0 To UBound(BadChars)
			Set regEx = New RegExp ' 정규식 객체의 인스턴스 생성
				regEx.Pattern = BadChars(i)
				regEx.IgnoreCase = True
				regEx.Global = True
				newChars = regEx.Replace(newChars,"")
			Set regEx = Nothing
		Next

		'newChars = Replace(newChARS, "'", "''")

		RemoveInjection = newChars
	End Function



	' ##### 초기값 세팅 #####	
	Function getDefault(CheckValue, DefaultValue)
		If isNull(CheckValue) Or CheckValue = "" Then
			getDefault = DefaultValue
		Else
			getDefault = CheckValue
		End If
	End Function

	' ##### 표시값이 0인경우의 화면 변경 #####	
	Function getZeroValue(CheckValue, toValue)
		If CheckValue = "0" Then
			getZeroValue = toValue
		Else
			getZeroValue = CheckValue
		End If
	End Function

	' ##### 표시값 확인 #####	
	Function getBoolean(CheckValue, tValue, fValue)
		If CheckValue = True Then
			getBoolean = tValue
		Else
			getBoolean = fValue
		End If
	End Function





	' ##### 금액단위 표시 (억/천만/백만) #####	
	Function getWon(CheckValue, Div, Gubun) ' (백만단위숫자, 백만단위 만들기 위해 나누기, 구분(1:억,2:천,3:억천,4:억천백)
		ReturnValue = ""
		CheckValue = Trim(cStr(CheckValue))
		If Len(CheckValue) = 0 Then
			ReturnValue = "0"
		ElseIf cLng(CheckValue) = 0 Then
			ReturnValue = "0"
		Else
			If isNumeric(Div) Then
				If cLng(Div) > 0 Then
					CheckValue = cStr(Int(cLng(CheckValue) / Int(Div)))
				End If
			End If
			Select Case Gubun
				Case 1		' 2132 -> 21.32억
					ReturnValue = cStr(cLng(CheckValue) / 100) & "억"

				Case 2		' 2132 -> 213.2천만
					ReturnValue = cStr(cLng(CheckValue) / 10) & "천"
					ReturnValue = ReturnValue & "만"

				Case 3		' 2132 -> 21억3천2백
					' 억단위 확인
					If Len(CheckValue) >= 3 Then
						ReturnValue = Left(CheckValue, Len(CheckValue)-2) & "억"	' 21억
						CheckValue = cStr(cLng(Right(CheckValue,2)) / 10)			' 3.2
						CheckValue = Int(cLng(CheckValue))							' 3 (백만단위 절사)
					End If
					If Len(CheckValue) >= 2 Then
						If cLng(CheckValue) > 0 Then
						ReturnValue = ReturnValue & CheckValue & "천"					' 21억3천만
						End If
					Else
						If cLng(CheckValue) > 0 Then
						ReturnValue = ReturnValue & CheckValue & "백"
						End If
					End If
					If inStr(CheckValue, "천") > 0 Or inStr(CheckValue,"백") > 0 Then
					ReturnValue = ReturnValue & "만"
					End If

				Case 4		' 2132
					If Len(CheckValue) >= 3 Then
						ReturnValue = Left(CheckValue, Len(CheckValue)-2) & "억"	' 21억
						CheckValue = Right(CheckValue,2)							' 32
					End If
					If Len(CheckValue) >= 2 Then
						If cLng(CheckValue) > 0 Then
						ReturnValue = ReturnValue & Left(CheckValue,1) & "천"		' 3천
						End If
						CheckValue = Right(CheckValue,1)							'2
					End If
					If cLng(CheckValue) > 0 Then
					ReturnValue = ReturnValue & CheckValue & "백"					' 21억3천2백만
					End If
					If inStr(CheckValue, "천") > 0 Or inStr(CheckValue,"백") > 0 Then
					ReturnValue = ReturnValue & "만"
					End If

				Case 5		' 요거는 백만단위로 단위 보여주는거
					ReturnValue = FormatNumber(CheckValue,0) & "백만"
				Case 6		' 요거는 단위 안보여주는거
					ReturnValue = FormatNumber(CheckValue,0)
				Case 7		' 2132 -> 21.32 (단위 없이 억단위)
					ReturnValue = cStr(cLng(CheckValue) / 100)
				Case Else
					ReturnValue = CheckValue
			End Select
		End If
		getWon = ReturnValue
	End Function



	' ##### 초기값 세팅 함수 #####	
	Function getNumeric(CheckValue, ToValue)
		CheckValue = Replace(CheckValue,",","")
		If isNumeric(CheckValue) Then
			getNumeric = CheckValue
		Else
			getNumeric = ToValue
		End If
	End Function



	' ##### Html 관련 함수 #####	
	Function getText(CheckValue)
		CheckValue = replace(CheckValue, "&" , "&amp;")
		CheckValue = replace(CheckValue, "<", "&lt;")
		CheckValue = replace(CheckValue, ">", "&gt;")
		CheckValue = replace(CheckValue, "'", "&quot;")
		CheckValue = replace(CheckValue, vbcrlf, "<br>")

		getText = CheckValue
	End Function

	Function getHtml(CheckValue)
		If CheckValue <> "" Then
			CheckValue = replace(CheckValue, "<br>" , vbcrlf)
			CheckValue = replace(CheckValue, "&lt;", "<")
			CheckValue = replace(CheckValue, "&gt;", ">")
			CheckValue = replace(CheckValue, "&quot;","'")
			CheckValue = replace(CheckValue, "&amp;", "&" )
		Else 
			
		End If 
		
		getHtml = CheckValue
	End Function
	
	Function getHtml_editor(CheckValue)
		CheckValue = replace(CheckValue, "<br>" , vbcrlf)
		CheckValue = replace(CheckValue, "<br />" , vbcrlf)
		CheckValue = replace(CheckValue, "<br/>" , vbcrlf)
		CheckValue = replace(CheckValue, "<BR>" , vbcrlf)
		CheckValue = replace(CheckValue, "<BR />" , vbcrlf)
		CheckValue = replace(CheckValue, "<BR/>" , vbcrlf)
		CheckValue = replace(CheckValue, "<p>" , "")
		CheckValue = replace(CheckValue, "<P>" , "")
		CheckValue = replace(CheckValue, "</p>" , vbcrlf)
		CheckValue = replace(CheckValue, "</P>" , vbcrlf)
		CheckValue = replace(CheckValue, "&lt;", "<")
		CheckValue = replace(CheckValue, "&gt;", ">")
		CheckValue = replace(CheckValue, "&quot;","'")
		CheckValue = replace(CheckValue, "&amp;", "&" )
		getHtml_editor = CheckValue
	End Function


	' ##### 쿠키 관련 함수 #####
	Sub setCookie(Group, Key, Value)
		If CoolieDomail <> "" Then
			Response.Cookies(Group).Domain = CoolieDomail
		End If
		Response.Cookies(Group)(Key) = Value
	End Sub

	Function getCookie(Key)
		getCookie = Request.Cookies(Group)(Key)
	End Function



	' ##### URL 전송시 & 처리를 위한 문장 #####
	Function ReturnUrl(CheckValue, Bool)
		If Bool = True Then
			ReturnUrl = Replace(CheckValue, "&", "||")
		Else
			ReturnUrl = Replace(CheckValue, "||", "&")
		End If
	End Function


	' ##### 폼체크시 Div로 보이지 않거나 한 부분에 대한 체크로 Proc 호출시 기본값을 저장한다 #####
	Function ReturnNull(CheckValue, str)
		If isNull(CheckValue) = True Or CheckValue = "" Then
			CheckValue = str
		End If
		ReturnNull = CheckValue
	End Function


	' ##### 받은 문자열을 지정한 길이만큼 추가문자를 붙여준다 #####
	Function ChangeLength(CheckValue, Char, ToLength)	'("1", "0", 2) -> "02"
		Dim j
		If Len(CheckValue) < ToLength Then
			For j = 1 To Cint(ToLength)-Len(CheckValue)
				CheckValue = Char & CheckValue
			Next
		End If
		ChangeLength = CheckValue
	End Function

	Function CheckLength(strString, intLength)
		If Len(strString) > intLength-2 Then
			CheckLength = Left(strString,intLength-2) & "..."
		Else
			CheckLength = strString
		End If
	End Function

	Function CheckLength2(strString, intLength, appendString)
		If Len(strString) > intLength Then
			CheckLength2 = Left(strString,intLength) & appendString
		Else
			CheckLength2 = strString
		End If
	End Function


	' ##### 검색결과시 검색어를 빨간색으로 진하게 보여준다 #####
	Function HighLight(CheckValue, Char)
		HighLight = Replace(Ucase(CheckValue), Ucase(Char), "<b><font color=red>" & Ucase(Char) & "</font></b>")
	End Function


	' ##### 날짜 형태로 바꿔준다 (YYYY-MM-DD) #####
	Function ToDate(CheckValue)
		If Len(CheckValue) = 8 Then
			ToDate = Left(CheckValue,4) & "-" & Mid(CheckValue,5,2) & "-" & Right(CheckValue,2)
		Else
			ToDate = Date
		End If
	End Function

	Function WeekDate(CheckValue)
		Dim rt
		Select Case Weekday(ToDate(CheckValue))
			Case 1
				rt = "<font color='red'>(일)</font>"
			Case 2
				rt = "<font color='black'>(월)</font>"
			Case 3
				rt = "<font color='black'>(화)</font>"
			Case 4
				rt = "<font color='black'>(수)</font>"
			Case 5
				rt = "<font color='black'>(목)</font>"
			Case 6
				rt = "<font color='black'>(금)</font>"
			Case 7
				rt = "<font color='blue'>(토)</font>"
			Case Else
				rt = ""
		End Select
		WeekDate = rt
	End Function


	' #### HTML Editor 출력 ######################
	Sub Editor(strName, strWidth, strHeight, DefaultValue)
		Dim oFCKeditor
		Set oFCKeditor = New FCKeditor
		oFCKeditor.width = strWidth
		oFCKeditor.height = strHeight
		oFCKeditor.Value = DefaultValue
		oFCKeditor.Config("UseBROnCarriageReturn") = True
		oFCKeditor.Create strName
	End Sub

	' #### HTML 제거 content = StripHTML(content) ######################
	Function StripHTML(oSource)

		dim Result_Text

		Result_text = ReplaceText(oSource," ( )+"," ")

		Result_text = Replace(Result_text,"=" & vbcrlf,"")
		'Result_text = Replace(Result_text,";" & vblrcf,"")
		Result_text = Replace(Result_text,";" & vbcrlf,"")

		' Remove the header (prepare first by clearing attributes)

	   ' head 태그 안의 모든 내용을 지운다
		Result_text = ReplaceText(Result_text,"<( )*head([^>])*>","<head>")
		Result_text = ReplaceText(Result_text,"(<( )*(/)( )*head( )*>)","</head>")
		Result_text = ReplaceText(Result_text,"(<head>)[\s\S]*(</head>)","")

		' remove all scripts (prepare first by clearing attributes)

	   ' script 태그 안의 모든 내용을 지운다
		Result_text = ReplaceText(Result_text,"<( )*script([^>])*?>","<script>")
		Result_text = ReplaceText(Result_text,"(<( )*(/)( )*?script()*>)","</script>")
		Result_text = ReplaceText(Result_text,"(<script>)([^(<script>\.</script>)])*?(</script>)","")
		Result_text = ReplaceText(Result_text,"(<script>)[\s\S]*?(</script>)","")

		' remove all styles (prepare first by clearing attributes)

	   ' style 태그 안의 모든 내용을 지운다
		Result_text = ReplaceText(Result_text,"<( )*style([^>])*?>","<style>")
		Result_text = ReplaceText(Result_text,"(<( )*(/)( )*?style( )*>)","</style>")
		Result_text = ReplaceText(Result_text,"(<style>)[\s\S]*?(</style>)","")

		' remove all object (prepare first by clearing attributes)

	   ' object 태그 안의 모든 내용을 지운다
		Result_text = ReplaceText(Result_text,"<( )*object([^>])*?>","<object>")
		Result_text = ReplaceText(Result_text,"(<( )*(/)( )*?object( )*>)","</object>")
		Result_text = ReplaceText(Result_text,"(<object>)[\s\S]*?(</object>)","")

		' Remove the link (prepare first by clearing attributes)

	   ' link 태그 안의 모든 내용을 지운다
		Result_text = ReplaceText(Result_text,"<( )*link([^>])*>","<link>")
		Result_text = ReplaceText(Result_text,"(<( )*(/)( )*link( )*>)","</link>")
		Result_text = ReplaceText(Result_text,"(<link>)[\s\S]*(</link>)","")    

	   ' 자바스크립트 함수 치환

		Result_text = ReplaceText(Result_text,"onclick=","xonclick=")    
		Result_text = ReplaceText(Result_text,"onmouseover=","xonmouseover=")    
		Result_text = ReplaceText(Result_text,"onmouseout=","xonmouseout=")    
		Result_text = ReplaceText(Result_text,"onchange=","xonchange=")    
		Result_text = ReplaceText(Result_text,"href=""javascript","href=""xjavascript")    

	   ' span 및 div 태그의 속성을 제거

		Result_text = ReplaceText(Result_text,"<( )*span([^>])*?>","<span>")
		Result_text = ReplaceText(Result_text,"(<( )*(/)( )*?span( )*>)","</span>")

		Result_text = ReplaceText(Result_text,"<( )*div([^>])*?>","<div>")
		Result_text = ReplaceText(Result_text,"(<( )*(/)( )*?div( )*>)","</div>")   

	 

	   ' input 태그를 지운다

		Result_text = ReplaceText(Result_text,"<( )*input([^>])*?>","")
				
		' Remove remaining tags like <a>, links, images, comments etc - anything thats enclosed inside < >

	   ' 허용태그 이외의 태그 제거


		Result_text = ReplaceText(Result_text,"<[^(image|a|div|span|table|tr|td|li|p)]*?>","")
		
		' Thats it.
		StripHTML = Result_Text
	End Function

	 
	' 위에꺼 이어지는 함수
	Function ReplaceText(str1, patrn, replStr)
	   Dim regEx
	   Set regEx = New RegExp
	   with regEx
		  .Pattern = patrn
		  .IgnoreCase = True
		  .Global = True
	   end with
	   ReplaceText = regEx.Replace(str1, replStr)
	End Function 


	' HTML 태그 지우는 정규식 함수
	Function SplitTag(str)
		Dim RegularExpressionObject
		Set RegularExpressionObject = new RegExp
		with RegularExpressionObject
			.Pattern = "<[^>]+>"	' 태그만
			.IgnoreCase = True		' 대소문자 구분
			.Global = True			' 모든 패턴 전역 검색
		End with

		SplitTag = RegularExpressionObject.Replace(str,"")
		Set RegularExpressionObject = Nothing
	End Function

	'// 천단위 콤마 찍기
	Function setComma(num)
		Dim returnValue
		Dim temp

		temp = cstr(num)

		returnValue = ""
		
		Do While Len(temp) > 0
			If Len(returnValue) > 0 Then returnValue = "," & returnValue End If

			If Len(temp) > 3 then
				returnValue = Right(temp, 3) & returnValue 
				temp = Left(temp, Len(temp)-3)
			Else
				returnValue = temp & returnValue
				temp = ""
			End If

		loop

		setComma = returnValue
	End Function

	'// 임의의 문자열을 Length 만큼 생성한다.
	'// 문자열 생성범위 (a-zA-Z0-9)
	Function getRandomString(Lenth)
		Dim RanNum
		Dim RanStr : RanStr = Null
		Dim i

		Randomize

		For i=1 to Lenth
			Do
				RanNum = Round(Rnd * 1000, 0)
					if(RanNum>=48 and _
						RanNum<=122 and _
						(RanNum<=90 or RanNum>=97) and _
						(RanNum<=57 or RanNum>=65)) then
							Exit Do
					end if
			Loop
		RanStr = RanStr & Chr(RanNum)
		Next

		getRandomString = RanStr:
	End Function

	'---------------------------------------------------------------------------------------------------------
	' @Description: User Agent 로 접속한 기기가 모바일인지 체크한다.
	' @Return: True (모바일), False (모바일 아님)
	' @Author: domoyosi
	'---------------------------------------------------------------------------------------------------------
	Function isMobile() 
		Dim httpRegEx, httpMatches 
		Set httpRegEx = New RegExp
		With httpRegEx
			.Pattern = "(iphone|ipod|iemobile|mobile|lgtelecom|ppc|blackberry|sch-|sph-|lg-|canu|im-|ev-|nokia)"
			.IgnoreCase = True  
			.Global = False
		End With
		Set httpMatches = httpRegEx.Execute(Request.ServerVariables("HTTP_USER_AGENT"))
		If httpMatches.count > 0 Then
			isMobile = True
		Else
			isMobile = False
			Set httpRegEx = Nothing
			Set httpMatches = Nothing
		End If
	End Function 
	
	'---------------------------------------------------------------------------------------------------------
	' @Description: 접속한 기기가 모바일인지 데스크탑인지 구분한다.
	' @Return: MOBILE (모바일), DESKTOP (데스크탑)
	' @Author: domoyosi
	'---------------------------------------------------------------------------------------------------------
	Function getDeviceType() 		
		If isMobile() = True Then
			getDeviceType = "MOBILE"
		Else
			getDeviceType = "DESKTOP"
		End If
	End Function 
	
	'---------------------------------------------------------------------------------------------------------
	' @Description: 접속한 기기가 모바일인지 데스크탑인지 구분한다. (한글버전)
	' @Return: 모바일, 데스크탑
	' @Author: domoyosi
	'---------------------------------------------------------------------------------------------------------
	Function getDeviceTypeKor(deviceType)
		If deviceType = "MOBILE" Then
			getDeviceTypeKor = "모바일"
		ElseIf deviceType = "DESKTOP" Then 
			getDeviceTypeKor = "데스크탑"
		End If
	End Function 

	'---------------------------------------------------------------------------------------------------------
	' @Description: 유입 경로의 미디어를 확인한다.
	' @Return: naver, daum, facebook, ...
	' @Author: domoyosi
	'---------------------------------------------------------------------------------------------------------
	Function getMediaType() 
		Dim tempRef : tempRef = LCase(Request.ServerVariables("HTTP_REFERER"))
		Dim tempUA : tempUA = LCase(Request.ServerVariables("HTTP_USER_AGENT"))

		Dim media
		Dim refererDomain : refererDomain = ""
		Dim refererMedia : refererMedia = ""
		Dim refererMediaArray(15)
		refererMediaArray(0) = m_RefererCheckHomepage
		refererMediaArray(1) = "naver"
		refererMediaArray(2) = "daum"
		refererMediaArray(3) = "tistory"
		refererMediaArray(4) = "nate"
		refererMediaArray(5) = "zum"
		refererMediaArray(6) = "www.google"
		refererMediaArray(7) = "admong"	'애드몽
		refererMediaArray(8) = "story.kakao.com"	'카카오스토리 모바일 앱(PC용 카카오스토리에서는 user agent 로 구분이 불가하다. 2016.03.29 기준)
		refererMediaArray(9) = "facebook"
		refererMediaArray(10) = "twitter"
		refererMediaArray(11) = "instagram"
		refererMediaArray(12) = "tumblr"
		refererMediaArray(13) = "linkedin"
		refererMediaArray(14) = "pinterest"
		refererMediaArray(15) = "plus.url.google.com"	'Google+

		'refererMediaArray(?) = "추천인"	'고민중..(만약한다면 쿼리스트링에 추천인 파라미터 여부 체크로 해야할듯)


		'If Request.ServerVariables("HTTP_REFERER") <> "" Then 
		If tempRef <> "" Then 

			'referer 를 소문자로 변경.
			Dim lowercaseReferer : lowercaseReferer = LCase(tempRef)

			'referer 를 / 으로 split.
			Dim refererArray : refererArray = Split(lowercaseReferer, "/")
			
			'referer 에서 도메인 추출.
			If IsArray(refererArray) Then 
				If UBound(refererArray) >= 2 Then
					refererDomain = refererArray(2)
				End If
			End If

			'referer의 도메인으로 media 추출.
			If refererDomain <> "" Then
				
				For Each media In refererMediaArray
					'Response.Write " - Check Media (" & media & "): <br />"

					If InStr(refererDomain, media) > 0 Then
						'Response.Write "- Check Media InStr(" & media & "): " & InStr(refererDomain, media) & "<br />"
						refererMedia = media
						Exit For
					End If 
				Next 
				
				'media 배열에서 검색되지 않은 경우 '기타' 분류.
				If refererMedia = "" Then
					refererMedia = "etc"	'기타
				ElseIf refererMedia = m_RefererCheckHomepage Then 
					refererMedia = "homepage"	'홈페이지
				End If 

			End If 
		Else 
			'referer 가 없는 경우 홈페이지로 직접 접속.
			refererMedia = "homepage"
		End If 

		'user agent 에 특정 모바일 앱에서 링크를 타고 넘어왔는지 확인
		If InStr(tempUA, "kakaotalk") > 0 Then
			'user agent 에 'KAKAOTALK' 가 있으면 모바일 카톡에서 접속한 것(PC용 카카오톡에서는 user agent 로 구분이 불가하다. 2016.03.29 기준)
			refererMedia = "kakaotalk"
		'ElseIf InStr(tempUA, "kakaostory") > 0 Then 
		ElseIf InStr(tempUA, "fb_iab") > 0 Then 
			'페이스북 모바일 앱에서 링크로 접속시 FB_IAB 가 user agent 에 있다.
			refererMedia = "facebook"
		End If

		getMediaType = refererMedia
	End Function 

	'---------------------------------------------------------------------------------------------------------
	' @Description: 유입 경로의 미디어의 한글명을 반환한다.
	' @Return: 홈페이지, 네이버, 다음, ...
	' @Author: domoyosi
	'---------------------------------------------------------------------------------------------------------
	Function getMediaTypeKor(mediaType) 
		If mediaType = m_RefererCheckHomepage Then 
			getMediaTypeKor = "홈페이지"
		ElseIf mediaType = "homepage" Then 
			getMediaTypeKor = "홈페이지"
		ElseIf mediaType = "etc" Then 
			getMediaTypeKor = "기타"
		ElseIf mediaType = "naver" Then 
			getMediaTypeKor = "네이버"
		ElseIf mediaType = "daum" Then 
			getMediaTypeKor = "다음"
		ElseIf mediaType = "tistory" Then 
			getMediaTypeKor = "티스토리"
		ElseIf mediaType = "nate" Then 
			getMediaTypeKor = "네이트"
		ElseIf mediaType = "zum" Then 
			getMediaTypeKor = "줌"
		ElseIf mediaType = "www.google" Then 
			getMediaTypeKor = "구글"
		ElseIf mediaType = "admong" Then 
			getMediaTypeKor = "애드몽"
		ElseIf mediaType = "story.kakao.com" Then 
			getMediaTypeKor = "카카오스토리"
		ElseIf mediaType = "facebook" Then 
			getMediaTypeKor = "페이스북"
		ElseIf mediaType = "twitter" Then 
			getMediaTypeKor = "트위터"
		ElseIf mediaType = "instagram" Then 
			getMediaTypeKor = "인스타그램"
		ElseIf mediaType = "tumblr" Then 
			getMediaTypeKor = "텀블러"
		ElseIf mediaType = "linkedin" Then 
			getMediaTypeKor = "링크드인"
		ElseIf mediaType = "pinterest" Then 
			getMediaTypeKor = "핀터레스트"
		ElseIf mediaType = "plus.url.google.com" Then 
			getMediaTypeKor = "Google+"
		ElseIf mediaType = "kakaotalk" Then 
			getMediaTypeKor = "카카오톡"
		Else 
			getMediaTypeKor = "unknown"
		End If 
	End Function 
	
	'---------------------------------------------------------------------------------------------------------
	' @Description: Referer 를 html 에 출력시 tag 를 생성한다.
	' @Return: referer 가 있을 경우만 a 태그로 출력한다.
	' @Author: domoyosi
	'---------------------------------------------------------------------------------------------------------
	Function convertReferer2Html(referer)
		If referer = "" Then 
			convertReferer2Html = "직접"
		Else
			convertReferer2Html = "<a href=" & referer & " target=""_blank"">" & referer & "</a></div>"
		End If 
	End Function 

	Function maskingString(value)
		
		Dim returnString
		Dim stringLength
		Dim i

		If value <> "" Then 
			stringLength = Len(value)
			
			If stringLength > 2 Then
				
				returnString = returnString & Left(value, 1)
				For i = 1 To stringLength-2
					returnString = returnString & "*"
				Next 
				returnString = returnString & Right(value, 1)

			ElseIf stringLength = 2 Then 
				returnString = Left(value, 1) & "*"
			Else
				returnString = value & "*"
			End If 
		
		Else
			maskingString = ""
		End If

		maskingString = returnString
	End Function 

	Function replaceBRtagToOther(CheckValue, ReplaceValue)
		CheckValue = replace(CheckValue, "<br>" , ReplaceValue)
		CheckValue = replace(CheckValue, "<br />" , ReplaceValue)
		CheckValue = replace(CheckValue, "<br/>" , ReplaceValue)
		CheckValue = replace(CheckValue, "<BR>" , ReplaceValue)
		CheckValue = replace(CheckValue, "<BR />" , ReplaceValue)
		CheckValue = replace(CheckValue, "<BR/>" , ReplaceValue)
		replaceBRtagToOther = CheckValue
	End Function

	Function replaceEndPtagToOther(CheckValue, ReplaceValue)
		CheckValue = replace(CheckValue, "</p>" , ReplaceValue)
		CheckValue = replace(CheckValue, "</P>" , ReplaceValue)
		replaceEndPtagToOther = CheckValue
	End Function

	'XSS 출력 필터 함수
	'XSS 필터 함수
	'$str - 필터링할 출력값
	'$avatag - 허용할 태그 리스트 예)  $avatag = "p,br"
	Function clearXSS(strString, avatag)
		'XSS 필터링
		strString = replace(strString, "<", "&lt;")
		strString = replace(strString, ">", "&gt;")
		strString = replace(strString, "\0", "")

		'허용할 태그 변환
		avatag = replace(avatag, " ", "")       '공백 제거
		If (avatag <> "") Then
			taglist = split(avatag, ",")

			for each p in taglist
				strString = replace(strString, "<"&p&" ", "<"&p&" ", 1, -1, 1)
				strString = replace(strString, "<"&p&">", "<"&p&">", 1, -1, 1)
				strString = replace(strString, "</"&p&" ", "</"&p&" ", 1, -1, 1)
			next
		End If

		clearXSS = strString
	End Function

	
	Function clearXSSdefault(strString)
		clearXSSdefault = clearXSS(strString, "")
	End Function

	Function addHyphen(fmemtel)
		Dim t1, t2, t3, t4
		Select Case Len(fmemtel)
		Case 8    '1588-xxxx
			t1 = Mid(fmemtel,1,4)
			t2 = Mid(fmemtel,5,4)
			'response.write t1 & "-" &t2
			addHyphen = t1 & "-" &t2
		Case 9  '02-xxx-xxxx
			t1 = Mid(fmemtel,1,2)
			t2 = Mid(fmemtel,3,3)
			t3 = Mid(fmemtel,6,4)
			'response.write t1 & "-" &t2 & "-" &t3
			addHyphen = t1 & "-" &t2 & "-" &t3
		Case 10  '휴대전화 010-xxx-xxxx
			If Mid(fmemtel,1,2) = "01" Then  '휴대전화 010-xxx-xxxx
				t1 = Mid(fmemtel,1,3)
				t2 = Mid(fmemtel,4,3)
				t3 = Mid(fmemtel,7,4)
				'response.write t1 & "-" &t2 & "-" &t3
				addHyphen = t1 & "-" &t2 & "-" &t3
			Else  '일반전화
				If Mid(fmemtel,1,2) = "02" Then
					t1 = Mid(fmemtel,1,2)
					t2 = Mid(fmemtel,3,4)
					t3 = Mid(fmemtel,7,4)
					'response.write t1 & "-" &t2 & "-" &t3
					addHyphen = t1 & "-" &t2 & "-" &t3
				Else
					t1 = Mid(fmemtel,1,3)
					t2 = Mid(fmemtel,4,3)
					t3 = Mid(fmemtel,7,4)
					'response.write t1 & "-" &t2 & "-" &t3
					addHyphen = t1 & "-" &t2 & "-" &t3
				End If
			End If
		Case 11  'xxx-xxxx-xxxx(휴대전화,070)
			t1 = Mid(fmemtel,1,3)
			t2 = Mid(fmemtel,4,4)
			t3 = Mid(fmemtel,8,4)
			'response.write t1 & "-" &t2 & "-" &t3
			addHyphen = t1 & "-" &t2 & "-" &t3
		Case Else
			'response.write fmemtel
			addHyphen = fmemtel
		End Select
	End Function

	Function getImageUrlFirst(strString)
		Dim ObjRegExp, Matches, RetStr, Match
		'On Error Resume Next   
		Set ObjRegExp = New RegExp
		ObjRegExp.Pattern = "[^='|""]*\.(gif|jpg|bmp)"	' 정규 표현식 패턴
		ObjRegExp.Global = True	' 문자열 전체를 검색함
		ObjRegExp.IgnoreCase = True	' 대.소문자 구분 안함

		Set Matches = ObjRegExp.Execute(strString)

		RetStr = ""
		For Each Match in Matches 'Matches 컬렉션을 반복
			'RetStr = RetStr & "<br>" & Replace(Match, Value, "<", "&lt;") & vbcrlf
			'RetStr = RetStr & "<br>" & Match
			RetStr = Match
		Next
		Set ObjRegExp = Nothing

		getImageUrlFirst = RetStr

		
	End Function

	Function checkFilename(filename)
		
		Dim real_file1_filename : real_file1_filename = Mid(filename, 1, InstrRev(filename,".")-1)

		Dim invalid_idx : invalid_idx = 0
		invalid_idx = invalid_idx + InStr(real_file1_filename,".")
		invalid_idx = invalid_idx + InStr(real_file1_filename,"/")
		invalid_idx = invalid_idx + InStr(real_file1_filename,"&")
		invalid_idx = invalid_idx + InStr(real_file1_filename,"%")
		invalid_idx = invalid_idx + InStr(real_file1_filename," ")
		invalid_idx = invalid_idx + InStr(real_file1_filename,",")
		invalid_idx = invalid_idx + InStr(real_file1_filename,"?")
		invalid_idx = invalid_idx + InStr(real_file1_filename,"@")
		invalid_idx = invalid_idx + InStr(real_file1_filename,"$")
		invalid_idx = invalid_idx + InStr(real_file1_filename,"!")
		invalid_idx = invalid_idx + InStr(real_file1_filename,"^")
		invalid_idx = invalid_idx + InStr(real_file1_filename,"*")
		invalid_idx = invalid_idx + InStr(real_file1_filename,"(")
		invalid_idx = invalid_idx + InStr(real_file1_filename,")")
		invalid_idx = invalid_idx + InStr(real_file1_filename,"+")
		invalid_idx = invalid_idx + InStr(real_file1_filename,";")
		invalid_idx = invalid_idx + InStr(real_file1_filename,"#")
		invalid_idx = invalid_idx + InStr(real_file1_filename,"=")
		invalid_idx = invalid_idx + InStr(real_file1_filename,"'")
		
		invalid_idx = invalid_idx + InStr(filename,"&")
		invalid_idx = invalid_idx + InStr(filename,"%")
		invalid_idx = invalid_idx + InStr(filename," ")
		invalid_idx = invalid_idx + InStr(filename,",")
		invalid_idx = invalid_idx + InStr(filename,"?")
		invalid_idx = invalid_idx + InStr(filename,"@")
		invalid_idx = invalid_idx + InStr(filename,"$")
		invalid_idx = invalid_idx + InStr(filename,"!")
		invalid_idx = invalid_idx + InStr(filename,"^")
		invalid_idx = invalid_idx + InStr(filename,"*")
		invalid_idx = invalid_idx + InStr(filename,"(")
		invalid_idx = invalid_idx + InStr(filename,")")
		invalid_idx = invalid_idx + InStr(filename,"+")
		invalid_idx = invalid_idx + InStr(filename,";")
		invalid_idx = invalid_idx + InStr(filename,"#")
		invalid_idx = invalid_idx + InStr(filename,"=")
		invalid_idx = invalid_idx + InStr(filename,"'")
		
		'Response.Write "invalid_idx " & invalid_idx & "<br />"

		If invalid_idx > 0 Then 
			checkFilename = False
		Else
			checkFilename = True
		End If

	End Function
%>