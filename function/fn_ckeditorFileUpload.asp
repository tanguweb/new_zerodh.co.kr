<!-- #include virtual = "/Include/Config.asp" -->
<%
    Dim funcNum
    Dim CKEditor
    Dim langCode
    Dim fileUrl
    Dim message
	Dim Upload
	Dim upload_filename
	Dim img_filesize

    ' 파일 중복을 제거 하기 위해 고정 사이트 만큼 특정 문자를 채워 주는 함수
    Public Function LeftFillString ( strValue, fillChar, makeLength )
        Dim strRet
        Dim strLen, diff, i
        
        strRet  = ""
        strLen  = Len(strValue)
        diff    = CInt(makeLength) - strLen
        
        if diff > 0 then
            for i=1 to diff
                strRet = strRet & CStr(fillChar)
            next
        end if
        
        LeftFillString = strRet & CStr(strValue)
    End Function

 '유니크한 파일명 만들기
 Public Function MakeUniqueFileName( strPrename )
  Dim strFilename
  Dim dtNow
  dtNow = now()
  Randomize()
  strFilename = strPrename
  strFilename = strFilename & Year(dtNow)
  strFilename = strFilename & LeftFillString( Month(dtNow),   "0", 2 )
  strFilename = strFilename & LeftFillString( Day(dtNow),     "0", 2 )
  strFilename = strFilename & "_"
  strFilename = strFilename & LeftFillString( Hour(dtNow),    "0", 2 )
  strFilename = strFilename & LeftFillString( Minute(dtNow),  "0", 2 )
  strFilename = strFilename & LeftFillString( Second(dtNow),  "0", 2 )
  strFilename = strFilename & "_"  
  strFilename = strFilename & LeftFillString ( Int(Rnd * 1000000), "0", 7 )
  MakeUniqueFileName = strFilename
 End Function

 ' 파일 확장자 추출
 Function GetFileExt( strFilename )
  Dim strExt
  Dim nPos
  nPos = InStrRev( strFilename, ".", -1, 1 ) '// Text Compare
  if nPos > 0 then
   strExt = Mid( strFilename, nPos+1 )
  end if
  GetFileExt = strExt
 End Function

    ' 변수들은 위에서 말한 개발자 가이드 문서에서 뽑았습니다.
    ' Required: anonymous function number as explained above.
    funcNum = Request("CKEditorFuncNum")
    'Optional: instance name (might be used to load specific configuration file or anything else)
    CKEditor = Request("CKEditor")
    ' Optional: might be used to provide localized messages
    langCode = Request("langCode")
    ' Check the $_FILES array and save the file. Assign the correct path to some variable ($url).
    'fileUrl = ""
    ' Usually you will assign here something only if file could not be uploaded.
    'message = "성공적으로 파일 업로드"

    ' DEXT Upload를 사용하고 있습니다.
    Set Upload = Server.CreateObject("DEXT.FileUpload")
    'Upload.DefaultPath = Server.MapPath ("/data/ckeditor")
	Upload.DefaultPath = Server.MapPath ("/file")
    Dim folderPath
    'folderPath = Server.MapPath("/data/ckeditor/images/")
	folderPath = Server.MapPath("/file/ck")
    upload_filename = Upload("upload").Filename
    if IsNull(Upload("upload")) or Upload("upload").FileLen <= 0 then
        upload_filename = ""
        img_filesize = 0
        message = "업로드 파일이 존재하지 않습니다."
    else
        upload_filename = MakeUniqueFileName("upload") & "." & GetFileExt(upload_filename)
        img_filesize = Upload("upload").FileLen
        if img_filesize > 0 then
            ' 확장자 체크 해야 하는데 귀찮아서
            ' Request("type") 으로 Images / Flash 구별해서 확장자 추출함수인 GetFileExt(upload_filename) 이걸로 체크하면 되겠죠
            ' 이미지는 jpg, gif, png..   플래쉬는 swf 파일만.. ㅎㅎㅎ
            ' 만성피로와 귀찮음으로 그냥 넘어가기.. 
            Upload("upload").SaveAs folderPath & "\" & upload_filename
            message = "정상적으로 파일을 업로드했습니다."
        else
            message = "업로드 파일 사이즈가 0입니다."
        end if
    end if
    
    fileUrl = "/file/ck/" & upload_filename
%>
<script type="text/javascript">
  // 가장 중요한 부분인것 같군요
  // ckeditor의 순번과 유효한 파일 경로만 넘기면 자동으로 이미지나 플래쉬 속성 변경 탭으로 이동합니다.
  window.parent.CKEDITOR.tools.callFunction(<%=funcNum %>, '<%=fileUrl %>', '<%=message %>');
//  history.go(-1);
</script>