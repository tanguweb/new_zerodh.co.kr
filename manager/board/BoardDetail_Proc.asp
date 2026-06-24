<!-- #include virtual = "/Include/Config.asp" -->
<!-- #include virtual = "/function/fn_AdminLoginCheck.asp" -->
<!-- #include virtual = "/function/fn_FileSave.asp" -->
<!-- #include virtual = "/function/fn_XMLpath.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<%
	'########################################################################################
	'#	File		: /manager/board/BoardDetail_Proc.asp
	'#  Create		: / 2010.09.09
	'#	Info		: 게시판등록
	'#	Update		: 
	'#	Update Memo	: 
	'########################################################################################
	

	Dim adoCmd, Result1, Result1_ID, Result1_DEPTH, Result1_SEQ, Result2, Rtn_string
	Dim MODE, CD_BOARDCD, CD_BOARDID, CD_BOARDKEY, NO_BOARD_DEPTH, NO_BOARD_SEQ, YN_USE, NM_TITLE, NM_CONTENTS, NO_CNTVIEW, YN_DISPLAY, NO_VIEWORDER, NO_DOWNLOAD, YN_XML, NM_XMLPATH
	Dim ary_FIELD(16), ary_FIELDTYPE(16)
	Dim UploadForm, arrFileNames(16), arrFileSize(16), arrFileType(16), arrFileTypechk, i, j, iCount, FilePath, FileNameOnly, FileExt, FileName
	Dim HTTP_REFERER, HTTP_USER_AGENT, HTTP_DEVICE_TYPE, HTTP_MEDIA_TYPE

	Set UploadForm = Server.CreateObject("DEXT.FileUpload")
	UploadForm.DefaultPath = FileSaveRoot_Board

	Mode			= UploadForm("MODE")
	CD_BOARDCD		= UploadForm("CD_BOARDCD")
	CD_BOARDID		= UploadForm("CD_BOARDID")
	CD_BOARDKEY		= UploadForm("CD_BOARDKEY")
	NO_BOARD_DEPTH	= UploadForm("NO_BOARD_DEPTH")
	NO_BOARD_SEQ	= UploadForm("NO_BOARD_SEQ")
	YN_USE			= UploadForm("YN_USE")
	NM_TITLE		= UploadForm("NM_TITLE")
	NM_CONTENTS		= UploadForm("NM_CONTENTS")
	NO_CNTVIEW		= UploadForm("NO_CNTVIEW")
	YN_DISPLAY		= UploadForm("YN_DISPLAY")
	NO_VIEWORDER	= UploadForm("NO_VIEWORDER")
	NO_DOWNLOAD		= UploadForm("NO_DOWNLOAD")
	YN_XML			= UploadForm("YN_XML")
	NM_XMLPATH		= UploadForm("NM_XMLPATH")

for i=1 to 15
	
	ary_FIELDTYPE(i)	= UploadForm("CD_FIELDTYPE_" & i)
	ary_FIELD(i)		= UploadForm("NM_FIELD_" & i)
	
next

'	If Mode <> "D" And NM_CONTENTS = "" Then 
'		response.write "<script>alert('내용을 입력하세요.');history.back();</script>"
'		response.End 
'	End If 

	If Mode <> "D" And NM_TITLE = "" Then 
		response.write "<script>alert('제목을 입력하세요.');history.back();</script>"
		response.End 
	End If 

	For j = 1 To 15
		If ary_FIELD(j) <> "" and ary_FIELDTYPE(j) = "FILE" Then
			arrFileNames(j) = fn_FileSave (UploadForm, "NM_FIELD_"&(j), getBoolean(YN_XML="Y" and NM_XMLPATH<>"",FileSaveRoot & NM_XMLPATH,FileSaveRoot_Board))
		End If		
	Next

	for i=1 to 15
		if ary_FIELDTYPE(i)="FILE" and UploadForm("NM_FIELD_" & i)="" and UploadForm("hd_NM_FIELD_" & i)<>"" then 
			arrFileNames(i)		= UploadForm("hd_NM_FIELD_" & i)
		end if 
	next
	
	HTTP_REFERER = UploadForm("HTTP_REFERER")
	HTTP_USER_AGENT = UploadForm("HTTP_USER_AGENT")
	HTTP_DEVICE_TYPE = UploadForm("HTTP_DEVICE_TYPE")
	HTTP_MEDIA_TYPE = UploadForm("HTTP_MEDIA_TYPE")

	' 게시글 저장
	SET adoCmd	=	Server.CreateObject("ADODB.Command")
		with adoCmd
		.ActiveConnection	= m_DB
		.CommandType		= adCmdStoredProc
		.CommandText		= "SPA_BOARDDETAIL_EDIT"
		.Parameters.Append	.CreateParameter("@MODE"				, adChar,		adParamInput,	1, Mode)
		.Parameters.Append	.CreateParameter("@CD_BOARDID"			, adInteger,	adParamInput,	 , CD_BOARDID)
		.Parameters.Append	.CreateParameter("@CD_BOARDKEY"			, adInteger,	adParamInput,	 , CD_BOARDKEY)
		.Parameters.Append	.CreateParameter("@NO_BOARD_DEPTH"		, adInteger,	adParamInput,    , NO_BOARD_DEPTH)
		.Parameters.Append	.CreateParameter("@NO_BOARD_SEQ"		, adInteger,	adParamInput,    , NO_BOARD_SEQ)
		.Parameters.Append	.CreateParameter("@CD_BOARDCD"			, adInteger,	adParamInput,    , CD_BOARDCD)
		.Parameters.Append	.CreateParameter("@YN_DISPLAY"			, adchar,		adParamInput,	1, YN_DISPLAY)
		.Parameters.Append	.CreateParameter("@NO_VIEWCNT"			, adInteger,	adParamInput,    , NO_CNTVIEW)
		.Parameters.Append	.CreateParameter("@NO_VIEWORDER"		, adInteger,	adParamInput,	 , NO_VIEWORDER)
		.Parameters.Append	.CreateParameter("@NO_DOWNLOAD"			, adInteger,	adParamInput,	 , NO_DOWNLOAD)
		.Parameters.Append	.CreateParameter("@YN_USE"				, adchar,		adParamInput,	1, YN_USE)
		.Parameters.Append	.CreateParameter("@NM_TITLE"			, adVarchar,	adParamInput,	500, NM_TITLE)
		.Parameters.Append	.CreateParameter("@NM_CONTENTS"			, adVarchar,	adParamInput,	100000, NM_CONTENTS)

If Mode = "D" Then 
	for i=1 to 15
		.Parameters.Append	.CreateParameter("@NM_FIELD_"&i			, adVarchar,	adParamInput,1000, "")
	next		
Else 
	for i=1 to 15
		if  ary_FIELDTYPE(i) = "FILE" Then
			.Parameters.Append	.CreateParameter("@NM_FIELD_"&i		, adVarchar,	adParamInput,1000, arrFileNames(i))
		else 
			.Parameters.Append	.CreateParameter("@NM_FIELD_"&i		, adVarchar,	adParamInput,1000, ary_FIELD(i))
		end if 
	next		
End If 

		.Parameters.Append	.CreateParameter("@CD_USER"				, adVarchar,	adParamInput,	15, m_UserID)
		.Parameters.Append	.CreateParameter("@NO_IPADDR"			, adVarchar,	adParamInput,	20, m_IPAddr)		
		.Parameters.Append	.CreateParameter("@HTTP_REFERER"		, adVarchar,	adParamInput,	5000, HTTP_REFERER)
		.Parameters.Append	.CreateParameter("@HTTP_USER_AGENT"		, adVarchar,	adParamInput,	5000, HTTP_USER_AGENT)
		.Parameters.Append	.CreateParameter("@HTTP_DEVICE_TYPE"	, adVarchar,	adParamInput,	200, HTTP_DEVICE_TYPE)
		.Parameters.Append	.CreateParameter("@HTTP_MEDIA_TYPE"		, adVarchar,	adParamInput,	200, HTTP_MEDIA_TYPE)
		.Parameters.Append  .CreateParameter("@Result"				, adVarChar,	adParamoutput,  20	)
		.Parameters.Append  .CreateParameter("@Result_ID"			, adInteger,	adParamoutput  	)
		.Execute
		Result1			= .Parameters("@Result")
		Result1_ID		= .Parameters("@Result_ID")
		end with
	SET adoCmd = Nothing

	' 첨부파일 저장
'If Result1 = "SUCCESS" Then
'For i=1 to 10
'	If ary_FIELD(i) <> "" and ary_FIELDTYPE(i) = "FILE" Then
'	SET adoCmd	=	Server.CreateObject("ADODB.Command")
'		with adoCmd
'		.ActiveConnection	= m_DB_InTheF
'		.CommandType		= adCmdStoredProc
'		.CommandText		= "USP_BOARDFILE_EDIT"
'		.Parameters.Append	.CreateParameter("@MODE"			, adChar,		adParamInput,	1, Mode)
'		.Parameters.Append	.CreateParameter("@CD_BOARDID"		, adInteger,	adParamInput,	 , getBoolean(Mode="I",Result1_ID,CD_BOARDID))
'		.Parameters.Append	.CreateParameter("@NO_BOARD_DEPTH"	, adInteger,	adParamInput,    , getBoolean(Mode="I",Result1_DEPTH,NO_BOARD_DEPTH))
'		.Parameters.Append	.CreateParameter("@NO_BOARD_SEQ"	, adInteger,	adParamInput,    , getBoolean(Mode="I",Result1_SEQ,NO_BOARD_SEQ))
'		.Parameters.Append	.CreateParameter("@CD_FILENO"		, adInteger,	adParamInput,	1, i)
'		.Parameters.Append	.CreateParameter("@NM_FILE"			, adVarchar,	adParamInput,   500, arrFileNames(i))
'		.Parameters.Append	.CreateParameter("@NO_FILESIZE"		, adInteger,	adParamInput,	 , arrFileSize(i))
'		.Parameters.Append	.CreateParameter("@NM_MIMETYPE"		, adVarchar,	adParamInput,	10, arrFileType(i))
'		.Parameters.Append	.CreateParameter("@YN_USE"			, adchar,		adParamInput,	1, "Y")
'		.Parameters.Append	.CreateParameter("@CD_USER"			, adVarchar,	adParamInput,	15, m_PrimaryID)
'		.Parameters.Append	.CreateParameter("@NO_IPADDR"		, adVarchar,	adParamInput,	20, m_IPAddr)
'		.Parameters.Append  .CreateParameter("@Result"			, adVarChar,	adParamoutput,  20	)
'		.Execute
'		Result2 = Result2 & "||" & .Parameters("@Result")
'		end with
'	SET adoCmd = Nothing
'	end if 
'next
'end if 

'	if Result1 = "SUCCESS" and instr(Result2,"ERROR") = 0 then 
	if Result1 = "SUCCESS" then 
		If MODE = "I" Then	' 등록
			Rtn_string = "<script>alert('게시글이 저장되었습니다.');document.location.href='BoardDetail.asp?CD_BOARDCD="&CD_BOARDCD&"';</script>" 
		ElseIf MODE = "U" Then	' 수정
			Rtn_string = "<script>alert('게시글이 수정되었습니다.');document.location.href='BoardDetail.asp?CD_BOARDCD="&CD_BOARDCD&"';</script>" 
		ElseIf MODE = "D" Then	' 삭제
			Rtn_string = "<script>alert('게시글이 삭제되었습니다.');document.location.href='BoardDetail.asp?CD_BOARDCD="&CD_BOARDCD&"';</script>" 
		End If 
		response.write Rtn_string 
		response.end 
	else
		response.write "<script>alert('게시글이 정상적으로 등록되지 못하였습니다.');history.back();</script>"
		response.end 
	end if 
		
%>