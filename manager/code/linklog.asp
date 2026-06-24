<!-- #include virtual = "/Include/Config.asp" -->
<%
	Dim objADO, Rs, SQL, CD_URLKEY, i, j
	
	CD_URLKEY = Req("CD_URLKEY")

	Set objADO = new clsADO
	SQL = SQL & "	SELECT *				"
	SQL = SQL & "	FROM "
	SQL = SQL & "	DB_LOG.dbo.T_LINKLOG				"
	SQL = SQL & "	WHERE CD_URLKEY = '" & CD_URLKEY & "'"
	SQL = SQL & "	ORDER BY DT_DATETIME DESC 	"

	objADO.setSql(SQL)
	Set Rs = objADO.getRs()


%>
<HTML>
 <HEAD>
  <TITLE> Link Log </TITLE>
  <link rel="stylesheet" type="text/css" href="/css/admin.css" />
	<Script Language="javascript" src="/js/Common.js"></Script>
 </HEAD>

 <BODY>
 <table border="0" cellpadding="0" cellspacing="0" width="100%">
 <tr>
 <td>
  <table style="BORDER-RIGHT: #cfcfdf 1px solid; BORDER-TOP: #cfcfdf 1px solid; BORDER-LEFT: #cfcfdf 1px solid; BORDER-BOTTOM: #cfcfdf 1px solid" cellSpacing="0" cellPadding="1" width="100%">
	<tr bgColor="#76A4C0" height="25">
		<td align="center" colSpan="3">&nbsp;&nbsp;<font color="white"><B>로그 보기 [<%=CD_URLKEY%>]</B></font></td>
	</tr>
	</table>
</td>
</tr>

<tr>
<td>
	<table cellSpacing="1" cellPadding="1" width="100%" align="left" bgColor="#a9a9a9" border="0">
	<tr align="center" bgColor="#eaeaea" height="25">
		<%For i = 0 To Rs.Fields.Count - 1%>
		<td align="center"><%=Rs.Fields(i).Name%></td>
		<%Next%>
	</tr>
	<%If Not Rs.EOF Then%>
	<%Do Until Rs.EOF%>
	<tr height="25" bgColor="#FFFFFF">
		<%For i = 0 To Rs.Fields.Count - 1%>
		<td><%=Rs.Fields(i).value%></td>
		<%Next%>
	</tr>
	<%
	Rs.MoveNext
	Loop
	End If
	%>
</tr>
</td>
</table>

 </BODY>
</HTML>
<%
Set objADO = Nothing
%>