<!-- #include virtual = "/Include/Config.asp" -->
<!-- #include virtual = "/function/fn_CodeInfo.asp" -->
<!-- #include virtual = "/function/fn_AdminLoginCheck.asp" -->
<!-- #include virtual = "/function/fn_XMLpath.asp" -->
<!-- #Include Virtual = "/FCKeditor/fckeditor.asp" -->
<%
	Dim objADO, SQL, arrBoard, i, CD_BOARDCD

	CD_BOARDCD = ReqQ("CD_BOARDCD")

	Set objADO = new clsADO
	SQL = ""
	SQL = SQL & "	SELECT"
	SQL = SQL & "	CD_BOARDCD, NM_BOARDNM,"
	SQL = SQL & "	NO_COUNT = ISNULL((SELECT COUNT(*) FROM T_BOARDDETAIL AA WHERE AA.CD_BOARDCD = A.CD_BOARDCD AND AA.YN_USE = 'Y' ),0)"
	SQL = SQL & "	FROM T_BOARDMAST A"
	SQL = SQL & "	WHERE YN_USE = 'Y' AND CD_BOARDTYPE = (SELECT CD_BOARDTYPE FROM T_BOARDMAST WHERE CD_BOARDCD = " & CD_BOARDCD & ")"
	SQL = SQL & "	ORDER BY NO_VIEWORDER, NM_BOARDNM"

	objADO.setSql(SQL)
	arrBoard = objADO.getArrRs()

%>
<HTML>
 <HEAD>
  <TITLE> 게시판 선택 </TITLE>
	<link rel="stylesheet" type="text/css" href="/css/admin.css" />
	<Script Language="javascript" src="/js/Common.js"></Script>
	<script Language="javascript">
		function fn_select(CD_BOARDCD, NM_BOARDNM)
		{
			var frm = opener.document.all;
			frm.btnSave.value = "[  " + NM_BOARDNM + "  ]" + "게시판으로 복사합니다";
			//?CD_BOARDCD=1024&CD_BOARDID=0&CD_BOARDKEY=0&NO_BOARD_DEPTH=1&NO_BOARD_SEQ=1
			frm.MODE.value="I";
			frm.CD_BOARDCD.value=CD_BOARDCD;
			frm.CD_BOARDID.value="0";
			frm.CD_BOARDKEY.value="0";
			frm.NO_BOARD_DEPTH.value="1";
			frm.NO_BOARD_SEQ.value = "1";

			self.close();
		}
	</script>
 </HEAD>

 <BODY>
	<font color="red">* 첨부파일은 복사되지 않습니다</font><br>
	<table cellSpacing="1" cellPadding="1" width="100%" align="left" bgColor="#a9a9a9" border="0">
	<tr align="center" bgColor="#eaeaea" height="30">
		<td align="center" width="40">선택</td>
		<td align="center">게시판 번호</td>
		<td align="center">게시판 제목</td>
		<td align="center">글 갯수</td>
	</tr>
	<%
	If IsArray(arrBoard) Then
	For i = LBound(arrBoard, 2) To UBound(arrBoard,2)
	%>
	<tr align="center" bgColor="#ffffff" height="30">
		<td align="center" width="40">
			<input type="button" value="선택" onClick="javascript:fn_select('<%=arrBoard(0,i)%>', '<%=arrBoard(1,i)%>');">
		</td>
		<td align="center"><%=arrBoard(0,i)%></td>
		<td align="center">
			<%If CStr(CD_BOARDCD) = CStr(arrBoard(0,i)) Then%>
			<font color="#0000FF">
			<%Else%>
			<font color="#000000">
			<%End If%>
			<%=arrBoard(1,i)%>
			</font>
		</td>
		<td align="center"><%=FormatNumber(arrBoard(2,i),0)%></td>
	</tr>
	<%
	Next
	End If
	%>
 </BODY>
</HTML>
