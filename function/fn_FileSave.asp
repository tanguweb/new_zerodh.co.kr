<%
	' 파일저장
	' 사용예시
	' Dim UploadForm
	' Set UploadForm = Server.CreateObject("DEXT.FileUpload")
	' juuun = fn_FileSave (UploadForm, 필드명, 저장경로)

	Function fn_FileSave (f_UploadForm___, p_FileName___, p_FileSavePath___)

		Dim f_FilePath___, f_FileNameOnly___, f_FileExt___, f_FileName___, i

		f_UploadForm___.DefaultPath = p_FileSavePath___

		f_FileName___ = f_UploadForm___(p_FileName___).FileName
		'f_FilePath___ = f_UploadForm___.DefaultPath & "\" & f_FileNameOnly___
		f_FilePath___ = f_UploadForm___.DefaultPath & "\" & f_FileName___


'		Response.WRite("<br>"&f_FileName___&"<br>")
'		Response.WRite("<br>"&f_FilePath___&"<br>")

		If f_FileName___ <> "" Then
		
			If f_UploadForm___.FileExists(f_FilePath___) Then
				
'				Response.WRite("<br>파일 이미 있음<br>")

				If InStrRev(f_FileName___, ".") <> 0 Then
					f_FileNameOnly___ = Left(f_FileName___, InStrRev(f_FileName___, ".") - 1)
					f_FileExt___ = Mid(f_FileName___, InStrRev(f_FileName___, "."))
				Else
					f_FileNameOnly___ = f_FileName___
					f_FileExt___ = ""
				End If
				i = 1
				Do While (1)
					f_FilePath___ =  f_UploadForm___.DefaultPath & "\" & f_FileNameOnly___ & "[" & i & "]" & f_FileExt___
					If Not f_UploadForm___.FileExists(f_FilePath___) Then
						Exit Do
					End If
					i = i + 1
				Loop
				f_FileName___ = f_FileNameOnly___ & "[" & i & "]" & f_FileExt___
			End If
			
'			Response.WRite("<br>"&f_FilePath___&"<br>")
'			Response.WRite("<br>"&p_FileName___&"<br>")
'			Response.WRite("<br>"&f_UploadForm___(p_FileName___)&"<br>")

			f_UploadForm___(p_FileName___).SaveAs f_FilePath___

		Else
			f_FileName___ = ""		
		End If
		
		fn_FileSave = f_FileName___

	End Function
%>