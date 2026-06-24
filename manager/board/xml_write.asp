<!-- #include virtual = "/Include/Config.asp" -->
<%
Dim juuun, Query, FilePath
Dim CD_BOARDCD : CD_BOARDCD = ReqQ("CD_BOARDCD")

Select Case CD_BOARDCD
	Case 1010	' joinus(카다로그)
		Query = " SELECT 'image/'+NM_FIELD_4 AS [mainImage], 'image/thumb/'+NM_FIELD_1 AS [thumbURL], 'image/big/'+NM_FIELD_2 AS [bigURL], 'image/large/'+NM_FIELD_3 AS [largeURL], NM_FIELD_5 AS [imgInfo1], NM_FIELD_6 AS [imgInfo2], NM_FIELD_7 AS [imgInfo3], NM_FIELD_8 AS [imgInfo4] FROM T_BOARDDETAIL WHERE CD_BOARDCD='1010' AND YN_USE='Y' AND YN_DISPLAY = 'Y' "
		FilePath = "D:\Project\[inthef]www.inthef.co.kr\joinus\_swf\style\cadalog\xml\imgInfo.XML"
	Case 1011	' joinus(코디제안)
		Query = " SELECT 'flash_image/listimg/'+NM_FIELD_1 AS [listURL], 'flash_image/bigimg/'+NM_FIELD_2 AS [bigURL], NM_FIELD_3 AS [l_subject], NM_FIELD_4 AS [l_name], NM_FIELD_5 AS [l_num], NM_FIELD_6 AS [l_color], '' AS [l_colorm], '' AS [l_colormm], NM_FIELD_7 AS [r_subject], NM_FIELD_8 AS [r_name], NM_FIELD_9 AS [r_num], NM_FIELD_10 AS [r_color], '' AS [r_colorm], '' AS [r_colormm] FROM T_BOARDDETAIL WHERE CD_BOARDCD='1011' AND YN_USE='Y' AND YN_DISPLAY = 'Y' "
		FilePath = "D:\Project\[inthef]www.inthef.co.kr\joinus\_swf\style\coordi\xml\codi.XML"
	Case 1012	' compagna(카다로그)
		Query = " SELECT 'image/'+NM_FIELD_4 AS [mainImage], 'image/thumb/'+NM_FIELD_1 AS [thumbURL], 'image/big/'+NM_FIELD_2 AS [bigURL], 'image/large/'+NM_FIELD_3 AS [largeURL], NM_FIELD_5 AS [imgInfo1], NM_FIELD_6 AS [imgInfo2], NM_FIELD_7 AS [imgInfo3], NM_FIELD_8 AS [imgInfo4] FROM T_BOARDDETAIL WHERE CD_BOARDCD='1012' AND YN_USE='Y' AND YN_DISPLAY = 'Y' "
		FilePath = "D:\Project\[inthef]www.inthef.co.kr\compagna\_swf\style\cadalog\xml\imgInfo.XML"
	Case 1013	' compagna(코디제안)
		Query = " SELECT 'flash_image/listimg/'+NM_FIELD_1 AS [listURL], 'flash_image/bigimg/'+NM_FIELD_2 AS [bigURL], NM_FIELD_3 AS [l_subject], NM_FIELD_4 AS [l_name], NM_FIELD_5 AS [l_num], NM_FIELD_6 AS [l_color], '' AS [l_colorm], '' AS [l_colormm], NM_FIELD_7 AS [r_subject], NM_FIELD_8 AS [r_name], NM_FIELD_9 AS [r_num], NM_FIELD_10 AS [r_color], '' AS [r_colorm], '' AS [r_colormm] FROM T_BOARDDETAIL WHERE CD_BOARDCD='1013' AND YN_USE='Y' AND YN_DISPLAY = 'Y' "
		FilePath = "D:\Project\[inthef]www.inthef.co.kr\compagna\_swf\style\coordi\xml\codi.XML"
	Case Else
		Query = ""
		FilePath = ""
End Select

if Query <> "" and FilePath <> "" then 
	Set juuun = Server.CreateObject("XMLParser.XMLParser") 
	juuun.ConnectionString = m_DB_XMLLoad
	'juuun.Query = "SELECT TOP 10 * FROM T_CODE"
	'juuun.FilePath = "D:\Project\[inthef]www.inthef.co.kr\manager\board\TEST_Juuun2.XML"
	juuun.Query = Query
	juuun.FilePath = FilePath
	juuun.SaveXML
	Set juuun = Nothing

	response.write "<script>alert('처리되었습니다.');this.close();</script>"
	response.end 
else
	response.write "<script>alert('정상적인 접근이 아닙니다. 다시한번 확인해 주세요.');this.close();</script>"
	response.end 
end if 
%>