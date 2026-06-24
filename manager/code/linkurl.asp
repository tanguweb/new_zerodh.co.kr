<!-- #include virtual = "/Include/Config.asp" -->
<%
	Dim objADO, Rs, SQL

	Set objADO = new clsADO
	SQL = SQL & "	SELECT *				"
	SQL = SQL & "	FROM T_LINKURL				"
	SQL = SQL & "	ORDER BY DT_INSYSDATE DESC 	"

	objADO.setSql(SQL)
	Set Rs = objADO.getRs()


%>
<HTML>
 <HEAD>
  <TITLE> New Document </TITLE>
  <link rel="stylesheet" type="text/css" href="/css/admin.css" />
	<Script Language="javascript" src="/js/Common.js"></Script>
	<Script Language = "JavaScript">
	function fn_go()
	{
		var frm = document.frm01;
	
		

		frm01.submit();
	}
	function fn_edt(CD_URLKEY)
	{
		var frm = document.frm01;
		frm01.action = "linkurl_edtproc.asp?CD_URLKEY="+CD_URLKEY;
		frm01.submit();
	}
	function fn_log(CD_URLKEY)
	{
		window.open("linklog.asp?CD_URLKEY="+CD_URLKEY);
	}
	</Script>
 </HEAD>

 <BODY>
 <table border="0" cellpadding="0" cellspacing="0" width="100%">
 <tr>
 <td>
  <table style="BORDER-RIGHT: #cfcfdf 1px solid; BORDER-TOP: #cfcfdf 1px solid; BORDER-LEFT: #cfcfdf 1px solid; BORDER-BOTTOM: #cfcfdf 1px solid" cellSpacing="0" cellPadding="1" width="100%">
	<tr bgColor="#76A4C0" height="25">
		<td align="center" colSpan="3">&nbsp;&nbsp;<font color="white"><B>링크관리</B></font></td>
	</tr>
	</table>
</td>
</tr>

<tr>
<td>
	<form name="frm01" action="linkurl_proc.asp" method="post">
	<table cellSpacing="1" cellPadding="1" width="100%" align="left" bgColor="#a9a9a9" border="0">
	<tr align="center" bgColor="#eaeaea" height="25">
		<td align="center" width="80">KEY길이</td>
		<td align="center" width="20" bgColor="#ffffff"><input type="text" name="NO_LEN" value="8" size="2"></td>
		<td align="center" width="80">URL</td>
		<td align="left" bgColor="#ffffff"><input type="text" name="CD_RURL" value="/" size="40"></td>
		<td align="center" width="80">메모</td>
		<td align="left" bgColor="#ffffff"><input type="text" name="NM_MEMO" value="" size="40"></td>
		<td align="center" width="80">시작일</td>
		<td align="center" width="100" bgColor="#ffffff"><input type="text" name="DT_FROM" value="<%=DATE%>" size="10" style="behavior:url(/js/htc_calendar.htc)"></td>
		<td align="center" width="80">종료일</td>
		<td align="center" width="100" bgColor="#ffffff"><input type="text" name="DT_TO" value="<%=DATE%>" size="10" style="behavior:url(/js/htc_calendar.htc)"></td>
		<td align="center" width="100">로그여부</td>
		<td align="center" width="100" bgColor="#ffffff">
			<select name="YN_LOG">
				<option value="N">아니요</option>
				<option value="Y">예</option>
			</select>
		</td>
		<td align="center" width="80"><input type="button" value="추가" onClick="javascript:fn_go();"></td>
	</tr>
	</table>
	<table><tr height="30"><td>&nbsp;</td></tr></table>
	<table cellSpacing="1" cellPadding="1" width="100%" align="left" bgColor="#a9a9a9" border="0">
	<tr align="center" bgColor="#eaeaea" height="30">
		<td align="center">KEY</td>
		<td align="center">URL</td>
		<td align="center">메모</td>
		<td align="center">시작일</td>
		<td align="center">종료일</td>
		<td align="center">사용여부</td>
		<td align="center">Log여부</td>
		<td align="center">등록일</td>
		<td align="center">수정</td>
		<td align="center">로그보기</td>
	</tr>
	<%If Not Rs.EOF Then%>
	<%Do Until Rs.EOF%>
	<tr align="center" bgColor="#ffffff" height="25">
		<td align="left"><a href="http://<%=Request.ServerVariables("HTTP_HOST")%>/?<%=Rs("CD_URLKEY")%>" target="_blank"><%=Rs("CD_URLKEY")%></a></td>
		<td align="left"><input type="text" size="70" name="CD_RURL_<%=Rs("CD_URLKEY")%>" value="<%=Rs("CD_RURL")%>"></td>
		<td align="center"><input type="text" size="40" name="NM_MEMO_<%=Rs("CD_URLKEY")%>" value="<%=Rs("NM_MEMO")%>"></td>
		<td align="center"><input type="text" size="8" name="DT_FROM_<%=Rs("CD_URLKEY")%>" value="<%=Rs("DT_FROM")%>"></td>
		<td align="center"><input type="text" size="8" name="DT_TO_<%=Rs("CD_URLKEY")%>" value="<%=Rs("DT_TO")%>"></td>
		<td align="center">
			<select name="YN_USE_<%=Rs("CD_URLKEY")%>">
				<option value="N" <% If Rs("YN_USE") = "N" Then Response.Write " selected " End If%>>아니요</option>
				<option value="Y" <% If Rs("YN_USE") = "Y" Then Response.Write " selected " End If%>>예</option>
			</select>
		</td>
		<td align="center">
			<select name="YN_LOG_<%=Rs("CD_URLKEY")%>">
				<option value="N" <% If Rs("YN_LOG") = "N" Then Response.Write " selected " End If%>>아니요</option>
				<option value="Y" <% If Rs("YN_LOG") = "Y" Then Response.Write " selected " End If%>>예</option>
			</select>
		</td>
		<td align="center"><%=Rs("DT_INSYSDATE")%></td>
		<td><input type="button" value="수정" onClick="javascript:fn_edt('<%=Rs("CD_URLKEY")%>');"></td>
		<td><input type="button" value="로그" onClick="javascript:fn_log('<%=Rs("CD_URLKEY")%>');"></td>
	</tr>
	<%
	Rs.MoveNext
	Loop
	End If
	%>
	<table>
</tr>
</td>
</table>
	</form>
 </BODY>
</HTML>
<%
Set objADO = Nothing
%>