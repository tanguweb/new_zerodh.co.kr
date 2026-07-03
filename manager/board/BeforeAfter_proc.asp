<!-- #include virtual = "/Include/Config.asp" -->
<!-- #include virtual = "/function/fn_AdminLoginCheck.asp" -->
<!-- #include virtual = "/function/fn_FileSave.asp" -->
<!-- #include virtual = "/function/fn_XMLpath.asp" -->
<!-- #include virtual = "/manager/inc/inc.seogen.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<%
	'########################################################################################
	'#	File		: /manager/board/BoardDetail_Proc.asp
	'#  Create		: / 2010.09.09
	'#	Info		: 게시판등록
	'#	Update		: 
	'#	Update Memo	: 
	'########################################################################################
	
	Dim CD_BOARDCD : CD_BOARDCD = 1004
	Dim conn, cmd, rs, iloop
	Dim sql
	Dim sql_update_board_key
	Dim sql_insert_before_after_img
	Dim dbErrorCount : dbErrorCount = 0
	Dim CD_BOARDID, YN_DISPLAY, NO_CNTVIEW, NM_TITLE, NM_CONTENTS, NM_FIELD_1
	
	Dim UploadForm
	Dim arr_before_after_image
	Dim arr_save_before_image, arr_save_after_image
	Dim arr_BEFORE_IMG, arr_AFTER_IMG
	Dim arr_hd_BEFORE_IMG, arr_hd_AFTER_IMG

	Set UploadForm = Server.CreateObject("DEXT.FileUpload")
	UploadForm.DefaultPath = FileSaveRoot_Board
	
	YN_DISPLAY = UploadForm("YN_DISPLAY")
	NO_CNTVIEW = UploadForm("NO_CNTVIEW")
	NM_TITLE = UploadForm("NM_TITLE")
	NM_CONTENTS = UploadForm("NM_CONTENTS")
	NM_FIELD_1 = UploadForm("NM_FIELD_1")

	'Response.write "arr_before_after_image : [" &  UploadForm("arr_before_after_image") & "]"
	arr_before_after_image = Split(UploadForm("arr_before_after_image"), ",")
	'Response.write "UBound(arr_before_after_image) : [" &  UBound(arr_before_after_image) & "]"
	
	ReDim arr_save_before_image(UBound(arr_before_after_image))
	ReDim arr_save_after_image(UBound(arr_before_after_image))

	'Response.write "BEFORE_IMG : [" &  UploadForm("BEFORE_IMG") & "]"	'파일명, 파일명
	'Response.write "hd_BEFORE_IMG : [" &  UploadForm("hd_BEFORE_IMG") & "]"	'파일명, 파일명
	arr_BEFORE_IMG = Split(UploadForm("BEFORE_IMG"), ", ")
	arr_hd_BEFORE_IMG = Split(UploadForm("hd_BEFORE_IMG"), ", ")
	'Response.write "Ubound(arr_BEFORE_IMG) : [" &  Ubound(arr_BEFORE_IMG) & "]"	'파일명, 파일명
	arr_AFTER_IMG = Split(UploadForm("AFTER_IMG"), ", ")
	arr_hd_AFTER_IMG = Split(UploadForm("hd_AFTER_IMG"), ", ")
	

	For iloop = 0 To UBound(arr_before_after_image)
		
		If UploadForm("BEFORE_IMG_" & iloop) <> "" Then 
			arr_save_before_image(iloop) = fn_FileSave (UploadForm, ("BEFORE_IMG_" & iloop), FileSaveRoot_Board)
		Else 
			arr_save_before_image(iloop) = UploadForm("hd_BEFORE_IMG_" & iloop)
		End If

		If UploadForm("AFTER_IMG_" & iloop) <> "" Then 
			arr_save_after_image(iloop) = fn_FileSave (UploadForm, ("AFTER_IMG_" & iloop), FileSaveRoot_Board)
		Else 
			arr_save_after_image(iloop) = UploadForm("hd_AFTER_IMG_" & iloop)
		End If
		
		'Response.write "("&iloop&")[" & arr_save_before_image(iloop) & "]" & ", [" & arr_save_after_image(iloop) & "]"
	Next 

	'Response.End
	
	sql = sql & "SET NOCOUNT ON;	"
	sql = sql & "INSERT INTO T_BOARDDETAIL (CD_BOARDKEY, NO_BOARD_DEPTH,NO_BOARD_SEQ,CD_BOARDCD,YN_DISPLAY	"
	sql = sql & "							,NO_VIEWCNT,NO_VIEWORDER,NO_DOWNLOAD,YN_USE,YN_ADMIN	"
	sql = sql & "							,NM_TITLE,NM_CONTENTS	"
	sql = sql & "							,NM_FIELD_1,NM_FIELD_2,NM_FIELD_3,NM_FIELD_4,NM_FIELD_5	"
	sql = sql & "							,NM_FIELD_6,NM_FIELD_7,NM_FIELD_8,NM_FIELD_9,NM_FIELD_10	"
	sql = sql & "							,NM_FIELD_11,NM_FIELD_12,NM_FIELD_13,NM_FIELD_14,NM_FIELD_15	"
	sql = sql & "							,CD_INUSER,NO_INIPADDR,DT_INSYSDATE	"
	sql = sql & "							,CD_MDUSER,NO_MDIPADDR,DT_MDSYSDATE	"
	sql = sql & "							,HTTP_REFERER,HTTP_USER_AGENT,HTTP_DEVICE_TYPE,HTTP_MEDIA_TYPE)	"
	'sql = sql & "VALUES					(@CD_BOARDKEY, 1, 1, 1004, ?	"
	sql = sql & "VALUES					(0, 1, 1, " & CD_BOARDCD & ", ?	"
	sql = sql & "						, ? , 100, 0, 'Y', 'Y'	"
	sql = sql & "						, ?, ?	"
	sql = sql & "						, ?, '', '', '', ''	"
	sql = sql & "						, '', '', '', '', ''	"
	sql = sql & "						, '', '', '', '', ''	"
	sql = sql & "						, ?, ?, GETDATE()	"
	sql = sql & "						, ?, ?, GETDATE()	"
	sql = sql & "						, '', '', '', '');	"
	sql = sql & "SELECT SCOPE_IDENTITY();	"

	sql_update_board_key = sql_update_board_key & "UPDATE T_BOARDDETAIL SET CD_BOARDKEY = CD_BOARDID WHERE CD_BOARDID = ?"

	sql_insert_before_after_img = sql_insert_before_after_img & "INSERT INTO T_BEFORE_AFTER_IMG (CD_BOARDID, CD_BEFOREAFTER_ID, BEFORE_IMG, AFTER_IMG)	"
	sql_insert_before_after_img = sql_insert_before_after_img & "VALUES (?, ?, ?, ?)	"
	
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open m_DB
	
	On Error Resume Next
	
	conn.BeginTrans

	Set cmd = server.createobject("ADODB.Command")
	cmd.ActiveConnection = conn
	cmd.CommandType = adCmdText
	cmd.CommandTimeout = 900

	cmd.CommandText = sql
	cmd.Parameters.Append cmd.CreateParameter("@YN_DISPLAY", adChar, adParamInput, 1, YN_DISPLAY)
	cmd.Parameters.Append cmd.CreateParameter("@NO_VIEWCNT", adInteger, adParamInput, , NO_CNTVIEW)
	cmd.Parameters.Append cmd.CreateParameter("@NM_TITLE", adVarchar, adParamInput, 500, NM_TITLE)
	cmd.Parameters.Append cmd.CreateParameter("@NM_CONTENTS", adVarchar, adParamInput, 100000, NM_CONTENTS)
	cmd.Parameters.Append cmd.CreateParameter("@NM_FIELD_1", adVarchar, adParamInput, 1000, NM_FIELD_1)
	cmd.Parameters.Append cmd.CreateParameter("@CD_USER", adVarchar, adParamInput, 15, m_UserID)
	cmd.Parameters.Append cmd.CreateParameter("@NO_IPADDR", adVarchar, adParamInput, 20, m_IPAddr)
	cmd.Parameters.Append cmd.CreateParameter("@CD_USER", adVarchar, adParamInput, 15, m_UserID)
	cmd.Parameters.Append cmd.CreateParameter("@NO_IPADDR", adVarchar, adParamInput, 20, m_IPAddr)
	Set rs = cmd.Execute()
	dbErrorCount = dbErrorCount + conn.errors.count
	'Response.write "cmd.error.count: [" & conn.errors.count & "]"

	'Response.write "new cd_boardid: [" & rs(0) & "]"
	CD_BOARDID = rs(0)



	Set cmd = server.createobject("ADODB.Command")
	cmd.ActiveConnection = conn
	cmd.CommandType = adCmdText
	cmd.CommandTimeout = 900
	cmd.CommandText = sql_update_board_key
	cmd.Parameters.Append cmd.CreateParameter("@CD_BOARDID", adInteger, adParamInput, , CD_BOARDID)
	cmd.Execute()
	dbErrorCount = dbErrorCount + conn.errors.count



	For iloop = 0 To UBound(arr_before_after_image)
		Set cmd = server.createobject("ADODB.Command")
		cmd.ActiveConnection = conn
		cmd.CommandType = adCmdText
		cmd.CommandTimeout = 900
		cmd.CommandText = sql_insert_before_after_img
		cmd.Parameters.Append cmd.CreateParameter("@CD_BOARDID", adInteger, adParamInput, , CD_BOARDID)
		cmd.Parameters.Append cmd.CreateParameter("@CD_BEFOREAFTER_ID", adInteger, adParamInput, , arr_before_after_image(iloop))
		cmd.Parameters.Append cmd.CreateParameter("@BEFORE_IMG", adVarchar, adParamInput, 1000, arr_save_before_image(iloop))
		cmd.Parameters.Append cmd.CreateParameter("@AFTER_IMG", adVarchar, adParamInput, 1000, arr_save_after_image(iloop))
		cmd.Execute()
		dbErrorCount = dbErrorCount + conn.errors.count
	Next 

	If dbErrorCount > 0 Then 
		conn.RollbackTrans
	Else 		
		conn.CommitTrans
	End If

	conn.close

	Set cmd = Nothing
	Set conn = Nothing


	If dbErrorCount = 0 Then
		' --- SEO: 치료전후(1004) 등록 성공 시 sitemap.xml / rss.xml 자동 재생성 ---
		On Error Resume Next
		Call SeoGenSitemap()
		Call SeoGenRss()
		On Error GoTo 0
		Response.write "<script>alert('게시글이 저장되었습니다.');document.location.href='BeforeAfterList.asp?CD_BOARDCD="&CD_BOARDCD&"';</script>"
	Else
		Response.write "<script>alert('게시글 저장에 실패하였습니다.');document.location.href='BeforeAfterList.asp?CD_BOARDCD="&CD_BOARDCD&"';</script>"
	End If 
	Response.End
%>