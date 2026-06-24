<%
' 오휘숙 전통부각 프론트 기본
Sub PageNavigation_front(Fmurl, Fparam, Ftn, FPgcount, FStartPage,  Fpage,  FSetSize,  FPrevPage,  FNextPage)
	Dim intLoop_i
	Dim f_i

	If Fparam <> "" Then 
		Fparam = "&"&Fparam
	End If 
	
	Fpage = Int(Fpage)

	If Ftn > 0 Then 
		If Fpage <> 1 Then
			Response.Write "<li><a href='"&Fmurl&"?page=1"&Fparam&"' title='맨앞 페이지로 이동'><!--맨앞--></a></li> "
		Else 
			Response.Write "<li><a href='javascript:void(0)' title='맨앞 페이지로 이동'><!--맨앞--></a></li> "
		End If

		if cLng(FPrevPage) > 0 then
			Response.Write "<li><a href='"&Fmurl&"?page="&FPrevPage&Fparam&"' title='이전페이지로 이동'><!--이전--></a></li> "
		else
			Response.Write "<li><a href='javascript:void(0)' title='이전페이지로 이동'><!--이전--></a></li> "
		end if

		intLoop_i = FSetSize+FStartPage - 1

		for f_i=FStartPage To intLoop_i
			if f_i > FPgcount then
				exit for
			end If

			'Response.Write "<!-- Fpage: " & Fpage & ", f_i: " & f_i & ", ==" & (Int(Fpage)=f_i) & "-->"

			if Fpage=f_i then
				Response.Write "<li class='active'><a href='javascript:void(0)'>"&f_i&"</a></li> "
			Else
				Response.Write "<li><a href='"&Fmurl&"?page="& f_i &Fparam&"'>"&f_i&"</a></li> "
			end if
		Next

		if cLng(FNextPage) > 0 then
			Response.Write "<li><a href='"&Fmurl&"?page="&FNextPage&Fparam&"' title='다음페이지로 이동'><!--다음--></a></li> "
		else
			Response.Write "<li><a href='javascript:void(0)' title='다음페이지로 이동'><!--다음--></a></li> "
		end if

		If Fpage <> FPgcount Then
			Response.Write "<li><a href='"&Fmurl&"?page="&FPgcount&Fparam&"' title='맨끝 페이지로 이동'><!--맨끝--></a></li> "
		Else
			Response.Write "<li><a href='javascript:void(0)' title='맨끝 페이지로 이동'><!--맨끝--></a></li> "
		End If
	End If 
End Sub

' 기본(using bootstrap style)
Sub PageNavigation(Fmurl, Fparam, Ftn, FPgcount, FStartPage,  Fpage,  FSetSize,  FPrevPage,  FNextPage)
	Dim intLoop_i
	Dim f_i

	If Fparam <> "" Then 
		Fparam = "&"&Fparam
	End If 

	If Ftn > 0 Then 
		Response.Write "<nav>"
		Response.Write "	<ul class='pagination'>"

		If Fpage <> 1 Then
			Response.Write "		<li>"
			Response.Write "			<a href='"&Fmurl&"?page=1"&Fparam&"' aria-label='First' title='첫 페이지로'>"
			Response.Write "				<span aria-hidden='true'>&laquo;</span>"
			Response.Write "			</a>"
			Response.Write "		</li>"
		End If

		if cLng(FPrevPage) > 0 then
			Response.Write "		<li>"
			Response.Write "			<a href='"&Fmurl&"?page="&FPrevPage&Fparam&"' aria-label='Previous' title='이전 "&FSetSize&"개 페이지로'>"
			Response.Write "				<span aria-hidden='true'>&lsaquo;</span>"
			Response.Write "			</a>"
			Response.Write "		</li>"
		else
			Response.Write "		<li class='disabled'>"
			Response.Write "			<a href='javascript:void(0)' aria-label='Previous' title='이전 "&FSetSize&"개 페이지로'>"
			Response.Write "				<span aria-hidden='true'>&lsaquo;</span>"
			Response.Write "			</a>"
			Response.Write "		</li>"
		end if

		intLoop_i = FSetSize+FStartPage - 1
		for f_i=FStartPage To intLoop_i
			if f_i > FPgcount then
				exit for
			end if
			if Fpage=f_i then
				Response.Write "		<li class='active'><a href='javascript:void(0)'>"&f_i&"</a></li>"
			Else
				Response.Write "		<li><a href='"&Fmurl&"?page="& f_i &Fparam&"'>"&f_i&"</a></li>"
			end if
		Next

		if cLng(FNextPage) > 0 then
			Response.Write "		<li>"
			Response.Write "			<a href='"&Fmurl&"?page="&FNextPage&Fparam&"' aria-label='Next' title='다음 "&FSetSize&"개 페이지로'>"
			Response.Write "				<span aria-hidden='true'>&rsaquo;</span>"
			Response.Write "			</a>"
			Response.Write "		</li>"
		else
			Response.Write "		<li>"
			Response.Write "			<a href='javascript:void(0)' aria-label='Next' title='다음 "&FSetSize&"개 페이지로'>"
			Response.Write "				<span aria-hidden='true'>&rsaquo;</span>"
			Response.Write "			</a>"
			Response.Write "		</li>"
		end if

		If Fpage <> FPgcount Then
			Response.Write "		<li>"
			Response.Write "			<a href='"&Fmurl&"?page="&FPgcount&Fparam&"' aria-label='Last' title='마지막 페이지로'>"
			Response.Write "				<span aria-hidden='true'>&raquo;</span>"
			Response.Write "			</a>"
			Response.Write "		</li>"
		End If
		
		Response.Write "	</ul>"
		Response.Write "</nav>"

	End If 
End Sub

' 이미지등록 네비
Sub PageNavigation_Img (Fmurl, Fparam, Ftn, FPgcount, FStartPage,  Fpage,  FSetSize,  FPrevPage,  FNextPage, i_Fpage, i_FPrevPage, i_NextPage, i_Epage)
	Dim intLoop_i
	Dim f_i

	If Fparam <> "" Then 
		Fparam = "&"&Fparam
	End If 

	If Ftn > 0 Then 
		If Fpage <> 1 Then
			Response.Write "<a href='"&Fmurl&"?page=1"&Fparam&"' onfocus='blur()'><img src='"&i_Fpage&"' class='vat' style=' vertical-align:middle' alt='첫 페이지로' border='0'/></a> "
		End If

		if cLng(FPrevPage) > 0 then
			Response.Write "<a href='"&Fmurl&"?page="&FPrevPage&Fparam&"' onfocus='blur()'><img src='"&i_FPrevPage&"' class='vat' style=' vertical-align:middle' alt='이전 페이지로' border='0' /></a> "
		else
			Response.Write "<a href='#none' onfocus='blur()'><img src='"&i_FPrevPage&"' class='vat' style=' vertical-align:middle' alt='이전 페이지로' border='0' /></a> "
		end if

		intLoop_i = FSetSize+FStartPage - 1
		Response.Write "<span class='num'>"
		for f_i=FStartPage To intLoop_i
			if f_i > FPgcount then
				exit for
			end if
			if Fpage=f_i then
				Response.Write " | <b> "&f_i&" </b>"
			Else
				Response.Write " | <a href='"&Fmurl&"?page="& f_i &Fparam&"' onfocus='this.blur()' class='store'> "&f_i&" </a> "
			end if
		Next
		Response.Write " | "
		Response.Write "</span>"


		if cLng(FNextPage) > 0 then
			Response.Write " <a href='"&Fmurl&"?page="&FNextPage&Fparam&"' onfocus='blur()'><img src='"&i_NextPage&"' class='vat' style=' vertical-align:middle' alt='다음 페이지로' border='0' /></a> "
		else
			Response.Write " <a href='#none' onfocus='blur()'><img src='"&i_NextPage&"' class='vat' style=' vertical-align:middle' alt='다음 페이지로' border='0' /></a> "
		end if

		If Fpage <> FPgcount Then
			Response.Write " <a href='"&Fmurl&"?page="&FPgcount&Fparam&"' onfocus='blur()'><img src='"&i_Epage&"' class='vat' style=' vertical-align:middle' alt='마지막 페이지로' border='0' /></a> "
		End If
	End If 
End Sub

' 이미지등록 네비_V2
Sub PageNavigation_Img_V2 (Fmurl, Fparam, Ftn, FPgcount, FStartPage,  Fpage,  FSetSize,  FPrevPage,  FNextPage, i_FPrevPage, i_NextPage)
	Dim intLoop_i
	Dim f_i

	If Fparam <> "" Then 
		Fparam = "&"&Fparam
	End If 
	
	If Fpage <> 1 Then 
		 FPrevPage = Fpage - 1
	End If 
	
	If Ftn > 0 Then 
		if cLng(FPrevPage) > 0 then
			Response.Write "<a href='"&Fmurl&"?page="&FPrevPage&Fparam&"' onfocus='blur()'><img src='"&i_FPrevPage&"' class='vat' style=' vertical-align:middle' alt='이전 페이지로' border='0' /></a> "
		else
			Response.Write "<a href='#none' onfocus='blur()'><img src='"&i_FPrevPage&"' class='vat' style=' vertical-align:middle' alt='이전 페이지로' border='0' /></a> "
		end if

		if cLng(FNextPage) > 0 then
			Response.Write " <a href='"&Fmurl&"?page="&FNextPage&Fparam&"' onfocus='blur()'><img src='"&i_NextPage&"' class='vat' style=' vertical-align:middle' alt='다음 페이지로' border='0' /></a> "
		else
			Response.Write " <a href='#none' onfocus='blur()'><img src='"&i_NextPage&"' class='vat' style=' vertical-align:middle' alt='다음 페이지로' border='0' /></a> "
		end if
	End If 
End Sub

' 이미지등록 네비
Sub PageNavigation_Img_V3 (Fmurl, Fparam, Ftn, FPgcount, FStartPage,  Fpage,  FSetSize,  FPrevPage,  FNextPage, i_Fpage, i_FPrevPage, i_NextPage, i_Epage)
	Dim intLoop_i
	Dim f_i

	If Fparam <> "" Then 
		Fparam = "&"&Fparam
	End If 

	If Ftn > 0 Then 
		If Fpage <> 1 Then
			Response.Write "<a href='"&Fmurl&"?page=1"&Fparam&"' onfocus='blur()'><img src='"&i_Fpage&"' class='vat' style='vertical-align:middle' alt='첫 페이지로' border='0'/></a> "
		End If

		if cLng(FPrevPage) > 0 then
			Response.Write "<a href='"&Fmurl&"?page="&FPrevPage&Fparam&"' onfocus='blur()'><img src='"&i_FPrevPage&"' class='vat' style='vertical-align:middle' alt='이전 페이지로' border='0' /></a> "
		else
			Response.Write "<a href='#none' onfocus='blur()'><img src='"&i_FPrevPage&"' class='vat' style=' vertical-align:middle' alt='이전 페이지로' border='0' /></a> "
		end if

		intLoop_i = FSetSize+FStartPage - 1
		Response.Write "<span class='num'>"
		for f_i=FStartPage To intLoop_i
			if f_i > FPgcount then
				exit for
			end if
			if Fpage=f_i then
				Response.Write " / <b> "&f_i&" </b>"
			Else
				Response.Write " / <a href='"&Fmurl&"?page="& f_i &Fparam&"' onfocus='this.blur()' class='store'> "&f_i&" </a> "
			end if
		Next
		Response.Write " / "
		Response.Write "</span>"


		if cLng(FNextPage) > 0 then
			Response.Write " <a href='"&Fmurl&"?page="&FNextPage&Fparam&"' onfocus='blur()'><img src='"&i_NextPage&"' class='vat' style=' vertical-align:middle' alt='다음 페이지로' border='0' /></a> "
		else
			Response.Write " <a href='#none' onfocus='blur()'><img src='"&i_NextPage&"' class='vat' style=' vertical-align:middle' alt='다음 페이지로' border='0' /></a> "
		end if

		If Fpage <> FPgcount Then
			Response.Write " <a href='"&Fmurl&"?page="&FPgcount&Fparam&"' onfocus='blur()'><img src='"&i_Epage&"' class='vat' style=' vertical-align:middle' alt='마지막 페이지로' border='0' /></a> "
		End If
	End If 
End Sub

' 이미지등록 네비_CSS 적용(페이지 URL, 파라미터, 총 페이지 수, 페이지 카운트, 시작 페이지, 현재 페이지, 페이지 세팅 수, 이전 페이지, 다음 페이지, 선택 페이지 CSS, 선택 안된 페이지 CSS,페이지 이미지 4개)
Sub PageNavigation_Img_Class (Fmurl, Fparam, Ftn, FPgcount, FStartPage,  Fpage,  FSetSize,  FPrevPage,  FNextPage, Fclass_selected, Fclass_unselected, i_Fpage, i_FPrevPage, i_NextPage, i_Epage)
	Dim intLoop_i
	Dim f_i

	If Fparam <> "" Then 
		Fparam = "&"&Fparam
	End If 

	If Ftn > 0 Then 
		If Fpage <> 1 Then
			Response.Write "<a href='"&Fmurl&"?page=1"&Fparam&"' onfocus='blur()'><img src='"&i_Fpage&"' class='vat' style=' vertical-align:middle' alt='첫 페이지로' border='0'/></a> "
		End If

		if cLng(FPrevPage) > 0 then
			Response.Write "<a href='"&Fmurl&"?page="&FPrevPage&Fparam&"' onfocus='blur()'><img src='"&i_FPrevPage&"' class='vat' style=' vertical-align:middle' alt='이전 페이지로' border='0' /></a> "
		else
			Response.Write "<a href='#none' onfocus='blur()'><img src='"&i_FPrevPage&"' class='vat' style=' vertical-align:middle' alt='이전 페이지로' border='0' /></a> "
		end if

		intLoop_i = FSetSize+FStartPage - 1
		Response.Write "<span class='num'>"
		for f_i=FStartPage To intLoop_i
			if f_i > FPgcount then
				exit for
			end if
			if Fpage=f_i then
				Response.Write " | <b class='" & Fclass_selected & "' style='text-decoration:none;'> "&f_i&" </b>"
			Else
				Response.Write " | <a href='"&Fmurl&"?page="& f_i &Fparam&"' onfocus='this.blur()' class='" & Fclass_unselected & "' style='text-decoration:none;'> "&f_i&" </a> "
			end if
		Next
		Response.Write " | "
		Response.Write "</span>"


		if cLng(FNextPage) > 0 then
			Response.Write " <a href='"&Fmurl&"?page="&FNextPage&Fparam&"' onfocus='blur()'><img src='"&i_NextPage&"' class='vat' style=' vertical-align:middle' alt='다음 페이지로' border='0' /></a> "
		else
			Response.Write " <a href='#none' onfocus='blur()'><img src='"&i_NextPage&"' class='vat' style=' vertical-align:middle' alt='다음 페이지로' border='0' /></a> "
		end if

		If Fpage <> FPgcount Then
			Response.Write " <a href='"&Fmurl&"?page="&FPgcount&Fparam&"' onfocus='blur()'><img src='"&i_Epage&"' class='vat' style=' vertical-align:middle' alt='마지막 페이지로' border='0' /></a> "
		End If
	End If 
End Sub


' Css & 텍스트 네비(리뉴얼된 메이폴에서 사용)
Sub PageNavigation_Css_Text (Fmurl, Fparam, Ftn, FPgcount, FStartPage,  Fpage,  FSetSize,  FPrevPage,  FNextPage, css_Fpage, css_FPrevPage, css_NextPage, css_Epage, txt_Fpage, txt_FPrevPage, txt_NextPage, txt_Epage)
	Dim intLoop_i
	Dim f_i

	If Fparam <> "" Then 
		Fparam = "&"&Fparam
	End If 

	If Ftn > 0 Then 
		If Fpage <> 1 Then
			Response.Write "<a href='"&Fmurl&"?page=1"&Fparam&"' class='"&css_Fpage&"' onfocus='blur()'><span></span><span></span>" & txt_Fpage & "</a> "
		End If

		if cLng(FPrevPage) > 0 then
			Response.Write "<a href='"&Fmurl&"?page="&FPrevPage&Fparam&"' class='"&css_FPrevPage&"' onfocus='blur()'><span></span>" & txt_FPrevPage & "</a> "
		else
			Response.Write "<a href='#none' onfocus='blur()' class='"&css_FPrevPage&"'><span></span>" & txt_FPrevPage & "</a> "
		end if

		intLoop_i = FSetSize+FStartPage - 1
		Response.Write "<span>"
		for f_i=FStartPage To intLoop_i
			if f_i > FPgcount then
				exit for
			end if
			if Fpage=f_i then
				Response.Write "  <strong> "&f_i&" </strong>"
			Else
				Response.Write "  <a href='"&Fmurl&"?page="& f_i &Fparam&"' onfocus='this.blur()' class='store'> "&f_i&" </a> "
			end if
		Next
		Response.Write "  "
		Response.Write "</span>"


		if cLng(FNextPage) > 0 then
			Response.Write " <a href='"&Fmurl&"?page="&FNextPage&Fparam&"' onfocus='blur()' class='"&css_NextPage&"'>" & txt_NextPage & "<span></span></a> "
		else
			Response.Write " <a href='#none' onfocus='blur()' class='"&css_NextPage&"'>" & txt_NextPage & "<span></span></a> "
		end if

		If Fpage <> FPgcount Then
			Response.Write " <a href='"&Fmurl&"?page="&FPgcount&Fparam&"' onfocus='blur()' class='"&css_Epage&"'>" & txt_Epage & "<span></span><span></span></a> "
		End If
	End If 
End Sub

' 이미지등록 네비
Sub PageNavigation_Img_V4 (Fmurl, Fparam, Ftn, FPgcount, FStartPage,  Fpage,  FSetSize,  FPrevPage,  FNextPage, i_Fpage, i_FPrevPage, i_NextPage, i_Epage)
	Dim intLoop_i
	Dim f_i

	If Fparam <> "" Then 
		Fparam = "&"&Fparam
	End If 

	If Ftn > 0 Then 
		If Fpage <> 1 Then
			Response.Write "<a href='"&Fmurl&"?page=1"&Fparam&"'><img src='"&i_Fpage&"' class='vat' style='margin-bottom:1px; vertical-align:middle' alt='첫 페이지로' border='0'/></a> "
		End If

		if cLng(FPrevPage) > 0 then
			Response.Write "<a href='"&Fmurl&"?page="&FPrevPage&Fparam&"'><img src='"&i_FPrevPage&"' class='vat' style='margin-bottom:1px; vertical-align:middle' alt='이전 페이지로' border='0' /></a> "
		else
			Response.Write "<a href='#none'><img src='"&i_FPrevPage&"' class='vat' style='margin-bottom:1px; vertical-align:middle' alt='이전 페이지로' border='0' /></a> "
		end if

		intLoop_i = FSetSize+FStartPage - 1
		Response.Write "<span class='num'>"
		for f_i=FStartPage To intLoop_i
			if f_i > FPgcount then
				exit for
			end if
			if Fpage=f_i then
				Response.Write " | <b> "&f_i&" </b>"
			Else
				Response.Write " | <a href='"&Fmurl&"?page="& f_i &Fparam&"' class='store'> "&f_i&" </a> "
			end if
		Next
		Response.Write " | "
		Response.Write "</span>"


		if cLng(FNextPage) > 0 then
			Response.Write " <a href='"&Fmurl&"?page="&FNextPage&Fparam&"'><img src='"&i_NextPage&"' class='vat' style='margin-bottom:1px; vertical-align:middle' alt='다음 페이지로' border='0' /></a> "
		else
			Response.Write " <a href='#none'><img src='"&i_NextPage&"' class='vat' style='margin-bottom:1px; vertical-align:middle' alt='다음 페이지로' border='0' /></a> "
		end if

		If Fpage <> FPgcount Then
			Response.Write " <a href='"&Fmurl&"?page="&FPgcount&Fparam&"'><img src='"&i_Epage&"' class='vat' style='margin-bottom:1px; vertical-align:middle' alt='마지막 페이지로' border='0' /></a> "
		End If
	End If 
End Sub


' 프레디 아이스쇼 이벤트 페이징 (특수기호 사용)
Sub PageNavigation_specialChar (Fmurl, Fparam, Ftn, FPgcount, FStartPage,  Fpage,  FSetSize,  FPrevPage,  FNextPage)
	Dim intLoop_i
	Dim f_i

	If Fparam <> "" Then 
		Fparam = "&"&Fparam
	End If 

	If Ftn > 0 Then 
		If Fpage <> 1 Then
			'Response.Write "<a href='"&Fmurl&"?page=1"&Fparam&"'>◀◀</a> "
		End If

		if cLng(FPrevPage) > 0 then
			Response.Write "<a href='"&Fmurl&"?page="&FPrevPage&Fparam&"' title='이전 페이지로'>◀</a> "
		else
			Response.Write "<a href='#none'>◀</a> "
		end if

		intLoop_i = FSetSize+FStartPage - 1
		Response.Write "&nbsp;"
		for f_i=FStartPage To intLoop_i
			if f_i > FPgcount then
				exit for
			end if
			if Fpage=f_i then
				Response.Write " | <b> "&f_i&" </b>"
			Else
				Response.Write " | <a href='"&Fmurl&"?page="& f_i &Fparam&"'> "&f_i&" </a> "
			end if
		Next
		Response.Write " | "
		Response.Write "&nbsp;"


		if cLng(FNextPage) > 0 then
			Response.Write " <a href='"&Fmurl&"?page="&FNextPage&Fparam&"'  title='다음 페이지로'>▶</a> "
		else
			Response.Write " <a href='#none'>▶</a> "
		end if

		If Fpage <> FPgcount Then
			'Response.Write " <a href='"&Fmurl&"?page="&FPgcount&Fparam&"'>▶▶</a> "
		End If
	End If 
End Sub
%>