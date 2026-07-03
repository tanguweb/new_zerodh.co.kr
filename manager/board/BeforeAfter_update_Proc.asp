<!-- #include virtual = "/Include/Config.asp" -->
<!-- #include virtual = "/function/fn_AdminLoginCheck.asp" -->
<!-- #include virtual = "/function/fn_FileSave.asp" -->
<!-- #include virtual = "/function/fn_XMLpath.asp" -->
<!-- #include virtual = "/manager/inc/inc.seogen.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<%	
	Dim CD_BOARDCD : CD_BOARDCD = 1004
	Dim conn, cmd, rs, iloop
	Dim sql, sql_delete_before_after_image, sql_insert_before_after_image
	Dim dbErrorCount : dbErrorCount = 0
	Dim CD_BOARDID, YN_DISPLAY, NO_CNTVIEW, NM_TITLE, NM_CONTENTS, NM_FIELD_1
	Dim UploadForm
	Dim arr_before_after_image
	Dim arr_save_before_image, arr_save_after_image
	Dim arr_BEFORE_IMG, arr_AFTER_IMG
	Dim arr_hd_BEFORE_IMG, arr_hd_AFTER_IMG

	Set UploadForm = Server.CreateObject("DEXT.FileUpload")
	UploadForm.DefaultPath = FileSaveRoot_Board
	
	CD_BOARDID = UploadForm("CD_BOARDID")
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
	
	Dim update_img_idx : update_img_idx = 0

	For iloop = 0 To UBound(arr_before_after_image)

		update_img_idx = arr_before_after_image(iloop)

		If UploadForm("BEFORE_IMG_" & update_img_idx) <> "" Then 
			arr_save_before_image(iloop) = fn_FileSave (UploadForm, ("BEFORE_IMG_" & update_img_idx), FileSaveRoot_Board)
		Else 
			arr_save_before_image(iloop) = UploadForm("hd_BEFORE_IMG_" & update_img_idx)
		End If

		If UploadForm("AFTER_IMG_" & update_img_idx) <> "" Then 
			arr_save_after_image(iloop) = fn_FileSave (UploadForm, ("AFTER_IMG_" & update_img_idx), FileSaveRoot_Board)
		Else 
			arr_save_after_image(iloop) = UploadForm("hd_AFTER_IMG_" & update_img_idx)
		End If
		
		'Response.write "("&iloop&")[" & arr_save_before_image(iloop) & "]" & ", [" & arr_save_after_image(iloop) & "]<br />"
	Next 










	sql = "UPDATE T_BOARDDETAIL SET NM_TITLE = ?, NM_CONTENTS = ?, YN_DISPLAY = ?, NM_FIELD_1 = ?, NO_VIEWCNT = ? WHERE CD_BOARDCD = ? AND CD_BOARDID = ?;	"

	sql_delete_before_after_image = "DELETE FROM T_BEFORE_AFTER_IMG WHERE CD_BOARDID = ?;"

	sql_insert_before_after_image = "INSERT INTO T_BEFORE_AFTER_IMG (CD_BOARDID, CD_BEFOREAFTER_ID, BEFORE_IMG, AFTER_IMG) VALUES (?, ?, ?, ?); "

	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open m_DB
	
	On Error Resume Next
	
	conn.BeginTrans

	Set cmd = server.createobject("ADODB.Command")
	cmd.ActiveConnection = conn
	cmd.CommandType = adCmdText
	cmd.CommandTimeout = 900
	cmd.CommandText = sql
	cmd.Parameters.Append cmd.CreateParameter("@NM_TITLE", adVarchar, adParamInput, 500, NM_TITLE)
	cmd.Parameters.Append cmd.CreateParameter("@NM_CONTENTS", adVarchar, adParamInput, 100000, NM_CONTENTS)
	cmd.Parameters.Append cmd.CreateParameter("@YN_DISPLAY", adChar, adParamInput, 1, YN_DISPLAY)
	cmd.Parameters.Append cmd.CreateParameter("@NM_FIELD_1", adVarchar, adParamInput, 1000, NM_FIELD_1)
	cmd.Parameters.Append cmd.CreateParameter("@NO_VIEWCNT", adInteger, adParamInput, , NO_CNTVIEW)
	cmd.Parameters.Append cmd.CreateParameter("@CD_BOARDCD", adInteger, adParamInput, , CD_BOARDCD)
	cmd.Parameters.Append cmd.CreateParameter("@CD_BOARDID", adInteger, adParamInput, , CD_BOARDID)
	cmd.Execute()
	dbErrorCount = dbErrorCount + conn.errors.count
	

	Set cmd = server.createobject("ADODB.Command")
	cmd.ActiveConnection = conn
	cmd.CommandType = adCmdText
	cmd.CommandTimeout = 900
	cmd.CommandText = sql_delete_before_after_image
	cmd.Parameters.Append cmd.CreateParameter("@CD_BOARDID", adInteger, adParamInput, , CD_BOARDID)
	cmd.Execute()
	dbErrorCount = dbErrorCount + conn.errors.count

	For iloop = 0 To UBound(arr_before_after_image)
		Set cmd = server.createobject("ADODB.Command")
		cmd.ActiveConnection = conn
		cmd.CommandType = adCmdText
		cmd.CommandTimeout = 900
		cmd.CommandText = sql_insert_before_after_image
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
		' --- SEO: 치료전후(1004) 수정 성공 시 sitemap.xml / rss.xml 자동 재생성 ---
		On Error Resume Next
		Call SeoGenSitemap()
		Call SeoGenRss()
		On Error GoTo 0
		Response.write "<script>alert('게시글이 수정되었습니다.');document.location.href='BeforeAfterList.asp?CD_BOARDCD="&CD_BOARDCD&"';</script>"
	Else
		Response.write "<script>alert('게시글 수정에 실패하였습니다.');document.location.href='BeforeAfterList.asp?CD_BOARDCD="&CD_BOARDCD&"';</script>"
	End If 
	Response.End
%>