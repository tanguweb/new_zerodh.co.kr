<%
	Dim help_CD_BOARDCD, help_NM_TITLE ,help_NM_CONTENTS, help_NM_FIELD_1
	Dim aryBoardHelper, help_SQL, help_objADO
	
	help_CD_BOARDCD = "1000"

	Set help_objADO = new clsADO
	
	help_SQL = ""
	help_SQL = help_SQL & " SELECT TOP 1 BD2.CD_BOARDCD, BD2.NM_TITLE ,BD2.NM_CONTENTS AS [Help], BD2.NM_FIELD_1 AS [CD_BOARDCD]"
	help_SQL = help_SQL & " FROM T_BOARDDETAIL AS BD1																			"
	help_SQL = help_SQL & " 	INNER JOIN T_BOARDDETAIL AS BD2																	"
	help_SQL = help_SQL & " 	ON BD2.NM_FIELD_1 = '" & CD_BOARDCD & "' AND BD2.CD_BOARDCD = '" & help_CD_BOARDCD & "'			"
	help_SQL = help_SQL & " WHERE BD2.CD_BOARDCD = '" & help_CD_BOARDCD & "'													"

	help_objADO.setSql(help_SQL)
	aryBoardHelper = help_objADO.getArrRs()

	Set help_objADO = Nothing	
%>
<%
	If IsArray(aryBoardHelper) Then 
		help_NM_TITLE = aryBoardHelper(1,0)
		help_NM_CONTENTS = aryBoardHelper(2,0)
		help_NM_FIELD_1 = aryBoardHelper(3,0)
%>
<tr id="adminHelper" style="display:none;" bgcolor="#FFFFFF" >
	<td colspan="6" style="padding:20 0 20 0"><%=help_NM_CONTENTS%></td>
</tr>
<%
	Else 
%>
&nbsp;
<%	End If %>