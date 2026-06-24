<!-- #include virtual = "/Include/Config.asp" -->
<!-- #include virtual = "/function/fn_AdminLoginCheck.asp" -->
<!-- #include virtual = "/function/fn_FileSave.asp" -->
<!-- #include virtual = "/function/fn_XMLpath.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<%	
	Dim CD_BOARDCD : CD_BOARDCD = 1004
	Dim conn, cmd, rs, iloop
	Dim sql
	Dim dbErrorCount : dbErrorCount = 0
	Dim CD_BOARDID
	Dim UploadForm

	Set UploadForm = Server.CreateObject("DEXT.FileUpload")
	UploadForm.DefaultPath = FileSaveRoot_Board
	
	CD_BOARDID = UploadForm("CD_BOARDID")
	
	sql = sql & "DELETE FROM T_BOARDDETAIL WHERE CD_BOARDCD = ? AND CD_BOARDID = ?;	"
	sql = sql & "DELETE FROM T_BEFORE_AFTER_IMG WHERE CD_BOARDID = ?;	"
	
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open m_DB
	
	On Error Resume Next
	
	conn.BeginTrans

	Set cmd = server.createobject("ADODB.Command")
	cmd.ActiveConnection = conn
	cmd.CommandType = adCmdText
	cmd.CommandTimeout = 900

	cmd.CommandText = sql
	cmd.Parameters.Append cmd.CreateParameter("@CD_BOARDCD", adInteger, adParamInput, , CD_BOARDCD)
	cmd.Parameters.Append cmd.CreateParameter("@CD_BOARDID", adInteger, adParamInput, , CD_BOARDID)
	cmd.Parameters.Append cmd.CreateParameter("@CD_BOARDID", adInteger, adParamInput, , CD_BOARDID)
	cmd.Execute()
	dbErrorCount = dbErrorCount + conn.errors.count
	
	If dbErrorCount > 0 Then 
		conn.RollbackTrans
	Else 		
		conn.CommitTrans
	End If

	conn.close

	Set cmd = Nothing
	Set conn = Nothing


	If dbErrorCount = 0 Then
		Response.write "<script>alert('게시글이 삭제되었습니다.');document.location.href='BeforeAfterList.asp?CD_BOARDCD="&CD_BOARDCD&"';</script>"
	Else
		Response.write "<script>alert('게시글 삭제에 실패하였습니다.');document.location.href='BeforeAfterList.asp?CD_BOARDCD="&CD_BOARDCD&"';</script>"
	End If 
	Response.End
%>