<%
	' ---------------------------------------------------------
	' 관리자 체크 파일 (홈페이지와 관리자 페이지 로그인은 별도)
	' ---------------------------------------------------------
	If m_UserID = "" Or m_UserID = Null Then 
		Response.Write "<Script Language='JavaScript'>"
		Response.Write "	parent.location.href='/manager/Login.asp';"
		Response.Write "</Script>"
		Response.End
	End If
%>