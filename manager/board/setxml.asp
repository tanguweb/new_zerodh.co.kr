<!-- #include virtual = "/Include/Config.asp" -->
<%
	'########################################################################################
	'#	File		: /manager/board/setxml.asp
	'#  Create		: / 2010.10.12
	'#	Info		: setxml.asp
	'#	Update		: 
	'#	Update Memo	: 
	'########################################################################################
	Dim objADO, AryXMLList, AryXMLDetail, SQL, iloop, jloop
	Dim CD_BOARDCD, NM_BOARDNM, NM_XMLPATH, NM_XMLQUERY, NM_XMLFILE
	Dim sCD_BOARDCD : sCD_BOARDCD = ReqQ("sCD_BOARDCD")

	Set objADO = new clsADO
	SQL = ""
	SQL = SQL & "	SELECT CD_BOARDCD, NM_BOARDNM, NM_XMLPATH, NM_XMLQUERY, NM_XMLFILE		"
	SQL = SQL & "	FROM T_BOARDMAST 														"
	SQL = SQL & "	WHERE YN_XML = 'Y' AND YN_USE = 'Y' 									"
	SQL = SQL & "	ORDER BY NM_BOARDNM ASC													"
	objADO.setSql(SQL)
	AryXMLList = objADO.getArrRs()
	
	If IsArray(AryXMLList) Then
		for iloop=0 to ubound(AryXMLList,2)
			if sCD_BOARDCD =cstr(AryXMLList(0,iloop)) then
				CD_BOARDCD	= AryXMLList(0,iloop)
				NM_BOARDNM	= AryXMLList(1,iloop)
				NM_XMLPATH	= AryXMLList(2,iloop)
				NM_XMLQUERY	= AryXMLList(3,iloop)
				NM_XMLFILE	= AryXMLList(4,iloop)
			end if 
		next 
	else
		Response.Write "<script>alert('XML배포를 설정할 게시판이 없습니다.');history.back();</script>"
		Response.End
	End If

	SQL = ""
	SQL = SQL & "	SELECT BD.CD_BOARDID, BD.NM_TITLE, BD.YN_XMLDISPLAY, BD.DT_INSYSDATE		"
	SQL = SQL & "	FROM T_BOARDDETAIL AS BD													"
	SQL = SQL & "	WHERE BD.YN_USE = 'Y' AND BD.CD_BOARDCD = '"&CD_BOARDCD&"'					"
	SQL = SQL & "	ORDER BY BD.CD_BOARDID DESC													"
	objADO.setSql(SQL)
	AryXMLDetail = objADO.getArrRs()

	Set objADO = Nothing
%>
<HTML>
<HEAD>
	<TITLE> XML 관리 </TITLE>
	<link rel="stylesheet" type="text/css" href="/css/admin.css" />
	<Script Language="javascript" src="/js/Common.js"></Script>
	<Script Language = "JavaScript">
		function fn_goxml()
		{
			var sCD_BOARDCD = document.getElementById("sCD_BOARDCD").value;
			document.location.href="setxml.asp?sCD_BOARDCD="+sCD_BOARDCD;
		}
		function fn_setXML(tablenm)
		{
			var NM_XMLPATH = document.getElementById("NM_XMLPATH");
			var NM_XMLFILE = document.getElementById("NM_XMLFILE");
			var NM_XMLQUERY = document.getElementById("NM_XMLQUERY");
			
			if (NM_XMLPATH.value == "") { alert("XML경로를 입력하세요."); NM_XMLPATH.focus(); return; }
			if (NM_XMLFILE.value == "") { alert("XML파일명을 입력하세요."); NM_XMLFILE.focus(); return; }
			if (NM_XMLQUERY.value == "") { alert("쿼리를 입력하세요."); NM_XMLQUERY.focus(); return; }

			document.getElementById("tablenm").value = tablenm;
			document.frm_XML.submit();
		}
	</Script>
</HEAD>
<BODY>
<form id="frm_XML" name="frm_XML" method="post" action="setxml_Proc.asp">
<input type="hidden" id="tablenm" name="tablenm">
	<table style="BORDER-RIGHT: #cfcfdf 1px solid; BORDER-TOP: #cfcfdf 1px solid; BORDER-LEFT: #cfcfdf 1px solid; BORDER-BOTTOM: #cfcfdf 1px solid" cellSpacing="0" cellPadding="1" width="100%">
	<tr bgColor="#76A4C0" height="25">
		<td align="center" colSpan="3">&nbsp;&nbsp;<font color="white"><B>XML 관리</B></font></td>
	</tr>
	<tr height="30"><%'=FileSaveDefault%>
		<td>&nbsp;* XML저장경로는 <font color="red"><%=Request.ServerVariables("APPL_PHYSICAL_PATH")%></font>&nbsp;이후 경로를 입력해주세요.</td>
		<td align="right">
			<Select name="sCD_BOARDCD" id="sCD_BOARDCD" onchange="fn_goxml();">
				<option value="">:: 선택하세요 ::</option>
<%
if isarray(AryXMLList) then
	for iloop=0 to ubound(AryXMLList,2)
%>
				<option value="<%=AryXMLList(0,iloop)%>" <%=getBoolean(sCD_BOARDCD=cstr(AryXMLList(0,iloop)),"selected","")%>><%=AryXMLList(1,iloop)%></option>
<%
	next
end if 
%>
			</Select>&nbsp;
		</td>
		<td width="100">&nbsp;</td>
	</tr>
	</table>

	<br>

	<table>
	<tr>
		<td>
			<table cellSpacing="1" cellPadding="1" width="100%" height="100%" align="left" bgColor="#a9a9a9" border="0">
			<tr align="center" bgColor="#eaeaea" height="30">
				<td align="center" width="70">게시판<br>코드</td>
				<td align="center" width="300">게시판명</td>
				<td align="center">XML 경로</td>
				<td align="center">XML 파일명</td>
				<td align="center" width="50">저장</td>
			</tr>
			<tr align="center" bgColor="#ffffff" height="30">
				<td align="center" rowspan="3"><%=CD_BOARDCD%></td>
				<td align="center"><%=NM_BOARDNM%></td>
				<td align="left">&nbsp;<input type="text" id="NM_XMLPATH" name="NM_XMLPATH" value="<%=NM_XMLPATH%>" size="40" <%=getBoolean(m_userID = "sysadmin","","ReadOnly")%>></td>
				<td align="center">&nbsp;<input type="text" id="NM_XMLFILE" name="NM_XMLFILE" value="<%=NM_XMLFILE%>" size="15" <%=getBoolean(m_userID = "sysadmin","","ReadOnly")%>></td>
				<td align="center"><%if sCD_BOARDCD<>"" then %><input type="button" value="저장" onclick="fn_setXML('T_BOARDMAST');"><%end if %></td>
			</tr>
			<tr align="center" bgColor="#eaeaea" height="30">
				<td align="center" colspan="4">쿼리 - XML노드설정</td>
			</tr>
			<tr align="center" bgColor="#ffffff" height="30">
				<td align="center" colspan="4"><textarea cols="95" rows="30" wrap="soft" id="NM_XMLQUERY" name="NM_XMLQUERY" style="background: #f2eef4; padding: 3px; scrollbar-face-color: #ffffff; scrollbar-shadow-color: #cccccc; scrollbar-highlight-color: #cccccc; scrollbar-3dlight-color: #ffffff; scrollbar-darkshadow-color: #ffffff; scrollbar-track-color: #ffffff; scrollbar-arrow-color: #cccccc; border:1px solid #c1c1c1; overflow:auto" <%=getBoolean(m_userID = "sysadmin","","ReadOnly")%>><%=NM_XMLQUERY%></textarea></td>
			</tr>
			</table>
		</td>
		<td width="10"></td>
		<td>
			<table cellSpacing="1" cellPadding="1" width="100%" height="100%" align="left" bgColor="#a9a9a9" border="0">
			<tr align="center" bgColor="#eaeaea" height="30">
				<td align="center" width="90">XML<br>전시여부</td>
				<td align="left" width="160">등록일</td>
				<td align="center" width="500">제목</td>
				<td align="center"><%if isarray(AryXMLDetail) then %><input type="button" value="XML배포" onclick="fn_setXML('T_BOARDDETAIL');"><%end if %></td>
			</tr>
<%
If isarray(AryXMLDetail) then 
	for iloop=0 to ubound(AryXMLDetail,2)
%>
			<tr height="25" bgcolor="#ffffff" onMouseOver="javascript:this.style.backgroundColor='#D4D4D4'" onMouseOut="javascript:this.style.backgroundColor=''">
				<td align="center"><input type="checkbox" id="YN_XMLDISPLAY" name="YN_XMLDISPLAY" value="<%=AryXMLDetail(0,iloop)%>" <%=getBoolean(AryXMLDetail(2,iloop)="Y","checked","")%>></td>
				<td align="center"><%=AryXMLDetail(3,iloop)%></td>
				<td align="center"><%=AryXMLDetail(1,iloop)%></td>
				<td align="center"></td>
			</tr>
<%
	Next
%>

<%
End If 
%>
			</table>
		</td>
	</tr>
	</table>
</form>
 </BODY>
</HTML>