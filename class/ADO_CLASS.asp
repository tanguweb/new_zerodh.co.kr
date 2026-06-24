<%

'//	-------------------------------------------------
'//	Class : clsADO
'//	Desc : ADO 관련 Class, 걍 급한대로 함 맹글었으..
'//	Auth : 장경훈
'//	Date : 몰라
'//	-------------------------------------------------

'On Error Resume Next

Class clsADO
	'########### 멤버변수 선언
	private classCon			'// Connection Object
	private classRs				'// RecordSet Object
	private classStr			'// Connection String (Initialize에서 설정)
	private classQuery			'// 실행 쿼리
	Private classArray			'// 쿼리 실형결과
	private classTrace			'// 클래스 Trace 출력여부
	Private classShowColumns	'// Column Collection 출력 여부
	private classLastEffected	'// 실행된 레코드 수
	private classTranMode		'// 현재 트랜잭션 상태
	private classAutoTran		'// 자동 트랜잭션 모드

	'########### 클래스 초기화 (변수 초기화)
	Private Sub Class_Initialize
		classCon = null
		classRs = null
		'classStr = "Provider=SQLOLEDB.1;Password=peter;User ID=peters;Data Source=127.0.0.1,1433;Persist Security Info=True;Initial Catalog=DB_TEST"
		classStr = m_DB
		classQuery = ""			'// 연결 문자열
		classTrace = False		'// 실행과정 표시 여부
		classShowColumns = True	'// Columng명을 출력할지 여부
		classLastEffected = 0	'// 마지막 실행결과 Row 수
		classTranMode = ""		'// 현재 트랜잭션 상태 (중복실행 및 시작없는 종료 방지)
		classAutoTran = False	'// 자동 트랜잭션 모드 (설정시 해당 객체의 Connection Object에 설정)

		If classTrace = True Then : Response.Write "<hr><font color='blue' size='2'>[ Class Initialize ]</font><p>"
	End Sub

	'########### 클래스 종료 (자원 리턴)
	Private Sub Class_Terminate
		CloseRs()
		CloseDB()
		Set classRs = Nothing
		Set classCon = Nothing

		If classTrace = True Then : Response.Write "<font color='blue' size='2'>[ Class Terminate ]</font><p><hr>"
	End Sub

	'########### Sub 멤버 함수 선언


	Public Sub OpenDB()
		If IsObject(classCon) = False Then
			Set classCon = Server.CreateObject("ADODB.Connection")
		End If

		If classCon.State = 0 Then
			classCon.Open classStr
			If classAutoTran = True Then
				StartTran()
			End If
			If classTrace = True Then : Response.Write "<font color='blue' size='2'>　　[ Connection Open : " & classCon.State & " ]</font><p>"
		End If
	End Sub

	Public Sub OpenRs()
		OpenDB()
		CloseRs()
		Dim i
		If IsObject(classRs) = False Then
			Set classRs = Server.CreateObject("ADODB.RecordSet")
		End If
		If Trim(classQuery) <> "" Then
 
			classRs.Open classQuery, classCon, 1
			classLastEffected = classRs.RecordCount
			If classTrace = True Then : Response.Write "<font color='blue' size='2'>　　　　　　[ RecordSet Open : <font color='black'>" & classQuery & "</font> ]</font><p>"
			If classTrace = True And classShowColumns = True Then
				Response.Write "<font color='blue' size='2'>　　　　　　[ Fields Collection ] <font color='black'>"
				For i = 0 To classRs.Fields.Count - 1
					Response.Write "(" & i & ")" & classRs.Fields(i).Name & ", "
				Next
				Response.Write "</font></font><p>"
			End If
		Else
			classLastEffected = 0
			If classTrace = True Then : Response.Write "<font color='blue' size='2'>　　　　　　[ RecordSet Open : <font color='red'>SQL is Empty</font> ]</font><p>"
		End If

		If classTrace = True Then : Response.Write "<font color='blue' size='2'>　　　　　　[ Effected Rows : " & classLastEffected & " ]</font><p>"
	End Sub

	Public Sub CloseDB()
		CloseRs()
		If classTranMode = "Start" Then
			EndTran()
		End If
		If classCon.State <> 0 Then
			classCon.Close()
			If classTrace = True Then : Response.Write "<font color='blue' size='2'>　　[ Connection Close : " & classCon.State & " ]</font><p>"
		End If
	End Sub

	Public Sub CloseRs()
		If IsObject(classRs) Then
			If classRs.State <> 0 Then
				classRs.Close
				If classTrace = True Then : Response.Write "<font color='blue' size='2'>　　　　　　[ RecordSet Close : " & classRs.State & " ]</font><p>"
			End If
		End If
	End Sub

	Public Sub ExecuteQuery()
		OpenDB()
		If Trim(classQuery) <> "" Then
			classCon.Execute classQuery, classLastEffected
			If classTrace = True Then : Response.Write "<font color='blue' size='2'>　　　　　　[ Execute Query : <font color='black'>" & classQuery & "</font> ]</font><p>"
		Else
			classLastEffected = 0
			If classTrace = True Then : Response.Write "<font color='blue' size='2'>　　　　　　[ Execute Query : <font color='red'>not SQL</font> ]</font><p>"
		End If
		If classTrace = True Then : Response.Write "<font color='blue' size='2'>　　　　　　[ Effected Rows : " & classLastEffected & " ]</font><p>"
	End Sub

	

	Public Sub StartTran()
		OpenDB()
		classCon.BeginTrans
		If classTranMode <> "Start" Then
			classTranMode = "Start"
			If classTrace = True Then : Response.Write "<font color='red' size='2'>　　　　　　[ Transaction Start ]</font><p>"
		End If
	End Sub

	Public Sub EndTran()
		If classTranMode = "Start" Then
			If classCon.Errors.Count = 0 Then
				classCon.CommitTrans
				classTranMode = "Commit"
				If classTrace = True Then : Response.Write "<font color='red' size='2'>　　　　　　[ Transaction Commit ]</font><p>"
			Else
				classCon.RollbackTrans
				classTranMode = "Rollback-" & classCon.Errors.Count
				If classTrace = True Then : Response.Write "<font color='red' size='2'>　　　　　　[ Transaction Rollback ]</font><p>"
			End If
		End If
	End Sub


	'########### Function 멤버 함수 선언



	Public Function getRs()
		OpenRs()
		Set getRs = classRs
	End Function
	
	Public Function getArrRs()
		OpenRs()
		If classRs.State <> 0 Then
			If classRs.EOF Then
				classArray = null
			Else
				classArray = classRs.GetRows()
			End If
		Else
			classArray = null
		End If
		getArrRs = classArray
	End Function

	Public Function getArrRs2()
		OpenRs()
		Dim i
		If classRs.State <> 0 Then
			If classRs.EOF Then
				classArray = null
			Else
				classArray = classRs.GetRows()

				Redim Preserve classArray(UBound(classArray,1), UBound(classArray,2)+1)
				For i = 0 To classRs.Fields.Count -1
					classArray(i, UBound(classArray,2)) = classRs(i).Name
				Next
			End If
		Else
			classArray = null
		End If
		getArrRs2 = classArray
	End Function


	Public Function getExec()
		ExecuteQuery()
		getExec = classLastEffected
	End Function

	Public Function getLastEffected()
		getLastEffected = classLastEffected
	End Function

	Public Function getTranStart()
		StartTran()
		getTranStart = classTranMode
	End Function

	Public Function getTranEnd()
		EndTran()
		getTranEnd = classTranMode
	End Function

	Public Function getAutoTran()
		getAutoTran = classAutoTran
	End Function



	Public Sub setSql(strSql)
		classQuery = strSql
		If classTrace = True Then : Response.Write "<font color='blue' size='2'>　　　　[ Set Query : <font color='black'>" & classQuery & "</font> ]</font><p>"
	End Sub

	Public Sub setSqlAdd(strSql)
		classQuery = classQuery & " " & strSql
		If Trace = True Then : Response.Write "<font color='blue' size='2'>　　　　[ Set Query Add : <font color='black'>" & classQuery & "</font> ]</font><p>"
	End Sub

	Public Sub setTrace(boolTrace)
		classTrace = boolTrace
		If classTrace = True Then : Response.Write "<font color='blue' size='2'>　　[ Set Trace : " & classTrace & " ]</font><p>"
	End Sub

	Public Sub setAutoTran(boolTran)
		classAutoTran = boolTran
		If classTrace = True Then : Response.Write "<font color='blue' size='2'>　　[ Set AutoTran : " & classAutoTran & " ]</font><p>"
	End Sub

	Public Sub setConString(constr)
		classStr = constr
		If classTrace = True Then : Response.Write "<font color='blue' size='2'>　　[ Set ConnectionString : <font color='black'>" & classStr & "</font> ]</font><p>"
	End Sub



	


End Class


%>