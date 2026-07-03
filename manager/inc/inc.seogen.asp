<%
'########################################################################################
'#  File    : /manager/inc/inc.seogen.asp
'#  Info    : sitemap.xml / rss.xml 생성 공용 모듈
'#            - 수동 생성: sitemap_gen.asp / rss_gen.asp
'#            - 자동 재생성: BoardDetail_Proc.asp / BoardDetail_display_proc.asp
'#                          BeforeAfter_proc.asp / _update_Proc.asp / _delete_Proc.asp (치료전후 1004)
'#  사용    : Config.asp(clsADO) 가 include 된 페이지에서 본 파일 include 후
'#            iCnt = SeoGenSitemap()        ' sitemap.xml 생성, 동적 게시글 수 반환
'#            iCnt = SeoGenRss()            ' rss.xml 생성, item 수 반환
'#            If SeoIsContentBoard(cd) ...  ' 해당 게시판이 콘텐츠 게시판인지
'########################################################################################

'======================================================================================
'  ★★★ 사이트별 설정 — 새 사이트 적용 시 이 구역 + SeoGenSitemap()의 [정적 페이지] 만 교체 ★★★
'======================================================================================
Const SEO_SITE      = "https://zerodh.co.kr"
Const SEO_RSS_TITLE = "대구 제로치과병원 ㅣ 대구치과 ㅣ 대구수면치과 ㅣ 대구임플란트"
Const SEO_RSS_DESC  = "대구치과 대구수면치과 대구임플란트 치료정보 · 치료전후 · 제로치과칼럼 · 자주묻는질문 · 공지사항"

' 콘텐츠 게시판: code / view 파일 / 라벨   (상담 등 비공개·비콘텐츠 게시판은 넣지 않는다)
'   ※ 코드/파일명은 사이트마다 다르므로 반드시 grep "CD_BOARDCD" html/board/*.html 로 확인할 것
'   ※ 1009=column·video 공유(메뉴엔 column만 노출), 1010=faq·pr 공유(메뉴엔 faq만 노출)
'   ※ 1002=news·event 공유(event는 view 페이지 없음 → news_view만 등록)
'   ※ 1004=치료전후(before/after) — 등록/수정/삭제는 BeforeAfter proc 3종에서 재생성 연동
Function SeoBoardMap()
    Dim a(3, 2)
    a(0, 0) = "1002" : a(0, 1) = "news_view.html"   : a(0, 2) = "공지사항"
    a(1, 0) = "1009" : a(1, 1) = "column_view.html" : a(1, 2) = "제로치과칼럼"
    a(2, 0) = "1010" : a(2, 1) = "faq_view.html"    : a(2, 2) = "자주묻는질문"
    a(3, 0) = "1004" : a(3, 1) = "bf_view.html"     : a(3, 2) = "치료전후"
    SeoBoardMap = a
End Function
'======================================================================================


'======================================================================================
'  공용 헬퍼
'======================================================================================
Function seoToday()
    seoToday = Year(Now) & "-" & Right("0" & Month(Now), 2) & "-" & Right("0" & Day(Now), 2) & "T00:00:00+09:00"
End Function

' 해당 게시판코드가 콘텐츠 게시판(SeoBoardMap)에 속하는지 — proc 자동재생성 게이팅용
Function SeoIsContentBoard(cd)
    Dim m, i
    m = SeoBoardMap()
    SeoIsContentBoard = False
    For i = 0 To UBound(m, 1)
        If m(i, 0) = CStr(cd) Then
            SeoIsContentBoard = True
            Exit Function
        End If
    Next
End Function

' UTF-8 BOM 없이 파일 저장
Sub seoSaveUtf8NoBom(sVirtualPath, sContent)
    Dim sPath, oStW, oStB
    sPath = Server.MapPath(sVirtualPath)

    Set oStW = CreateObject("ADODB.Stream")
    oStW.Type = 2 : oStW.Charset = "UTF-8" : oStW.Open
    oStW.WriteText sContent
    oStW.Position = 0 : oStW.Type = 1
    oStW.Read(3)  ' BOM 제거

    Set oStB = CreateObject("ADODB.Stream")
    oStB.Type = 1 : oStB.Open
    oStB.Write oStW.Read
    oStB.SaveToFile sPath, 2   ' adSaveCreateOverWrite
    oStB.Close : oStW.Close
    Set oStW = Nothing : Set oStB = Nothing
End Sub

Function seoUrlEntry(sUrl, sLastmod, sPriority)
    seoUrlEntry = "<url>" & vbCrLf & _
                  "  <loc>" & sUrl & "</loc>" & vbCrLf & _
                  "  <lastmod>" & sLastmod & "</lastmod>" & vbCrLf & _
                  "  <priority>" & sPriority & "</priority>" & vbCrLf & _
                  "</url>" & vbCrLf
End Function

'======================================================================================
'  sitemap.xml 생성/저장. 반환: 동적 게시글 URL 수
'======================================================================================
Function SeoGenSitemap()

    Dim sToday : sToday = seoToday()
    Dim sXml
    sXml = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
    sXml = sXml & "<urlset" & vbCrLf
    sXml = sXml & "      xmlns=""http://www.sitemaps.org/schemas/sitemap/0.9""" & vbCrLf
    sXml = sXml & "      xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""" & vbCrLf
    sXml = sXml & "      xsi:schemaLocation=""http://www.sitemaps.org/schemas/sitemap/0.9" & vbCrLf
    sXml = sXml & "            http://www.sitemaps.org/schemas/sitemap/0.9/sitemap.xsd"">" & vbCrLf
    sXml = sXml & vbCrLf

    '================== [정적 페이지] 사이트별 — menu.html 기준(주석 아닌 항목만) ==================
    '-- 메인
    sXml = sXml & "<!-- ===================== 메인 ===================== -->" & vbCrLf
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/main/", sToday, "1.00")

    '-- 병원소개
    sXml = sXml & vbCrLf & "<!-- ===================== 병원소개 ===================== -->" & vbCrLf
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/introduce/life.html",    sToday, "0.80")
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/introduce/doctor.html",  sToday, "0.90")
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/introduce/space.html",   sToday, "0.80")
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/introduce/time.html",    sToday, "0.80")
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/introduce/contact.html", sToday, "0.80")

    '-- 수면&디지털센터
    sXml = sXml & vbCrLf & "<!-- ===================== 수면&디지털센터 ===================== -->" & vbCrLf
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/sleep/sedation.html", sToday, "0.80")
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/sleep/lab.html",      sToday, "0.80")
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/sleep/oneday.html",   sToday, "0.80")

    '-- 임플란트
    sXml = sXml & vbCrLf & "<!-- ===================== 임플란트 ===================== -->" & vbCrLf
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/implant/computer.html",  sToday, "0.80")
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/implant/full.html",      sToday, "0.80")
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/implant/sinus.html",     sToday, "0.80")
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/implant/bone.html",      sToday, "0.80")
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/implant/denture.html",   sToday, "0.80")
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/implant/immediate.html", sToday, "0.80")
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/implant/prf.html",       sToday, "0.80")
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/implant/uv.html",        sToday, "0.80")

    '-- 심미치료
    sXml = sXml & vbCrLf & "<!-- ===================== 심미치료 ===================== -->" & vbCrLf
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/beauty/layer.html",         sToday, "0.80")
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/beauty/resin.html",         sToday, "0.80")
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/beauty/whitening.html",     sToday, "0.80")
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/beauty/gum_whitening.html", sToday, "0.80")
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/beauty/tooth_form.html",    sToday, "0.80")
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/beauty/gum_form.html",      sToday, "0.80")

    '-- 일반치료
    sXml = sXml & vbCrLf & "<!-- ===================== 일반치료 ===================== -->" & vbCrLf
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/general/cavity.html",     sToday, "0.80")
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/general/nerve.html",      sToday, "0.80")
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/general/prosthetic.html", sToday, "0.80")
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/general/wisdom.html",     sToday, "0.80")
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/general/scaling.html",    sToday, "0.80")
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/general/tmj.html",        sToday, "0.80")
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/general/botox.html",      sToday, "0.80")
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/general/bruxism.html",    sToday, "0.80")

    '-- 소아치료
    sXml = sXml & vbCrLf & "<!-- ===================== 소아치료 ===================== -->" & vbCrLf
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/pediatric/ortho.html",      sToday, "0.80")
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/pediatric/endo.html",       sToday, "0.80")
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/pediatric/mesiodens.html",  sToday, "0.80")
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/pediatric/prevention.html", sToday, "0.80")

    '-- 커뮤니티 (게시판 목록 + 안내 정적페이지)
    sXml = sXml & vbCrLf & "<!-- ===================== 커뮤니티 ===================== -->" & vbCrLf
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/board/bf.html",         sToday, "0.80")
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/board/news.html",       sToday, "0.80")
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/introduce/fee.html",    sToday, "0.80")
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/introduce/price.html",  sToday, "0.80")
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/board/faq.html",        sToday, "0.80")
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/board/column.html",     sToday, "0.80")
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/board/event.html",      sToday, "0.80")

    '-- 기타 (푸터·상담폼 링크)
    sXml = sXml & vbCrLf & "<!-- ===================== 기타 ===================== -->" & vbCrLf
    sXml = sXml & seoUrlEntry(SEO_SITE & "/html/board/personal_info.html", sToday, "0.40")
    '================================ [정적 페이지] 끝 ================================

    '-- 게시판별 view 페이지 (DB 조회) — SeoBoardMap() 단일 소스 사용
    Dim m : m = SeoBoardMap()
    Dim objADO, SQL, aryRows, iRow, bLoop, sLastmod, iCount, dtVal
    iCount = 0

    For bLoop = 0 To UBound(m, 1)
        Set objADO = New clsADO
        SQL = ""
        SQL = SQL & "SELECT CD_BOARDID, DT_INSYSDATE "
        SQL = SQL & "FROM T_BOARDDETAIL "
        SQL = SQL & "WHERE CD_BOARDCD = '" & m(bLoop, 0) & "' "
        SQL = SQL & "  AND YN_USE = 'Y' AND YN_DISPLAY = 'Y' "
        SQL = SQL & "  AND NO_BOARD_DEPTH = 1 AND NO_BOARD_SEQ = 1 "
        SQL = SQL & "ORDER BY CD_BOARDID DESC "
        objADO.setSql(SQL)
        aryRows = objADO.getArrRs()
        Set objADO = Nothing

        If IsArray(aryRows) Then
            sXml = sXml & vbCrLf & "<!-- ===================== " & m(bLoop, 2) & " 게시글 (" & m(bLoop, 1) & " / " & m(bLoop, 0) & ") ===================== -->" & vbCrLf
            For iRow = 0 To UBound(aryRows, 2)
                sLastmod = ""
                If Trim("" & aryRows(1, iRow)) <> "" Then
                    dtVal = CDate(aryRows(1, iRow))
                    sLastmod = Year(dtVal) & "-" & Right("0" & Month(dtVal), 2) & "-" & Right("0" & Day(dtVal), 2) & "T00:00:00+09:00"
                Else
                    sLastmod = sToday
                End If
                sXml = sXml & seoUrlEntry( _
                    SEO_SITE & "/html/board/" & m(bLoop, 1) & "?CD_BOARDID=" & aryRows(0, iRow), _
                    sLastmod, "0.80")
                iCount = iCount + 1
            Next
        End If
    Next

    sXml = sXml & vbCrLf & "</urlset>" & vbCrLf

    Call seoSaveUtf8NoBom("/sitemap.xml", sXml)

    SeoGenSitemap = iCount
End Function

'======================================================================================
'  rss.xml 생성/저장. 반환: item 수
'======================================================================================
Function SeoGenRss()

    Dim sRss, sNowRfc, iTotal
    iTotal = 0
    sNowRfc = rssToRFC822(Now())

    sRss = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
    sRss = sRss & "<rss version=""2.0"" xmlns:atom=""http://www.w3.org/2005/Atom"">" & vbCrLf
    sRss = sRss & "<channel>" & vbCrLf
    sRss = sRss & "  <title>" & SEO_RSS_TITLE & "</title>" & vbCrLf
    sRss = sRss & "  <link>" & SEO_SITE & "/</link>" & vbCrLf
    sRss = sRss & "  <description>" & SEO_RSS_DESC & "</description>" & vbCrLf
    sRss = sRss & "  <language>ko</language>" & vbCrLf
    sRss = sRss & "  <atom:link href=""" & SEO_SITE & "/rss.xml"" rel=""self"" type=""application/rss+xml""/>" & vbCrLf
    sRss = sRss & "  <lastBuildDate>" & sNowRfc & "</lastBuildDate>" & vbCrLf & vbCrLf

    Dim m, mi, sInClause
    m = SeoBoardMap()
    sInClause = ""
    For mi = 0 To UBound(m, 1)
        If sInClause <> "" Then sInClause = sInClause & ","
        sInClause = sInClause & "'" & m(mi, 0) & "'"
    Next

    Dim objADO, SQL, aryRows, iRow
    Dim dtVal, sPubDate, sTitle, sDesc, sUrl, sViewPage

    Set objADO = New clsADO
    SQL = ""
    SQL = SQL & "SELECT TOP 100 CD_BOARDID, NM_TITLE, NM_CONTENTS, DT_INSYSDATE, CD_BOARDCD "
    SQL = SQL & "FROM T_BOARDDETAIL "
    SQL = SQL & "WHERE CD_BOARDCD IN (" & sInClause & ") "
    SQL = SQL & "  AND YN_USE = 'Y' AND YN_DISPLAY = 'Y' "
    SQL = SQL & "  AND NO_BOARD_DEPTH = 1 AND NO_BOARD_SEQ = 1 "
    SQL = SQL & "ORDER BY DT_INSYSDATE DESC "
    objADO.setSql(SQL)
    aryRows = objADO.getArrRs()
    Set objADO = Nothing

    If IsArray(aryRows) Then
        For iRow = 0 To UBound(aryRows, 2)
            ' 게시판코드 → view 파일 매핑 (SeoBoardMap 단일 소스)
            sViewPage = ""
            For mi = 0 To UBound(m, 1)
                If m(mi, 0) = CStr(aryRows(4, iRow)) Then
                    sViewPage = m(mi, 1)
                    Exit For
                End If
            Next

            If sViewPage <> "" Then
                sUrl   = SEO_SITE & "/html/board/" & sViewPage & "?CD_BOARDID=" & aryRows(0, iRow)
                sTitle = rssXmlEsc(CStr(aryRows(1, iRow)))
                sDesc  = rssStripHtml(CStr(aryRows(2, iRow)))
                If Len(sDesc) > 500 Then sDesc = Left(sDesc, 500) & "..."
                sDesc  = Replace(sDesc, "]]>", "]]&gt;")

                If Trim("" & aryRows(3, iRow)) <> "" Then
                    dtVal    = CDate(aryRows(3, iRow))
                    sPubDate = rssToRFC822(dtVal)
                Else
                    sPubDate = sNowRfc
                End If

                sRss = sRss & "  <item>" & vbCrLf
                sRss = sRss & "    <title>" & sTitle & "</title>" & vbCrLf
                sRss = sRss & "    <link>" & sUrl & "</link>" & vbCrLf
                sRss = sRss & "    <description><![CDATA[" & sDesc & "]]></description>" & vbCrLf
                sRss = sRss & "    <pubDate>" & sPubDate & "</pubDate>" & vbCrLf
                sRss = sRss & "    <guid isPermaLink=""true"">" & sUrl & "</guid>" & vbCrLf
                sRss = sRss & "  </item>" & vbCrLf
                iTotal = iTotal + 1
            End If
        Next
    End If

    sRss = sRss & vbCrLf & "</channel>" & vbCrLf & "</rss>" & vbCrLf

    Call seoSaveUtf8NoBom("/rss.xml", sRss)

    SeoGenRss = iTotal
End Function

' XML 1.0에서 허용되지 않는 제어문자 제거 (NUL 0x00 등) — DB 본문에 섞인 0x0 때문에 파싱 깨짐 방지
Function seoCleanXmlChars(s)
    Dim reC
    Set reC = New RegExp
    reC.Global = True
    reC.Pattern = "[\x00-\x08\x0B\x0C\x0E-\x1F]"
    seoCleanXmlChars = reC.Replace(s, "")
    Set reC = Nothing
End Function

Function rssStripHtml(s)
    Dim re
    Set re = New RegExp
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "<style[^>]*>[\s\S]*?</style>"
    s = re.Replace(s, "")
    re.Pattern = "<script[^>]*>[\s\S]*?</script>"
    s = re.Replace(s, "")
    re.Pattern = "<[^>]*>"
    s = re.Replace(s, " ")
    s = Replace(s, "&amp;",  "&")
    s = Replace(s, "&lt;",   "<")
    s = Replace(s, "&gt;",   ">")
    s = Replace(s, "&quot;", """")
    s = Replace(s, "&apos;", "'")
    s = Replace(s, "&nbsp;", " ")
    re.Pattern = "&[a-zA-Z]+;|&#[0-9]+;"
    s = re.Replace(s, "")
    Set re = Nothing
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    rssStripHtml = Trim(seoCleanXmlChars(s))
End Function

Function rssXmlEsc(s)
    s = seoCleanXmlChars(s)
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    rssXmlEsc = s
End Function

Function rssToRFC822(dt)
    Dim aDay(7), aMon(12)
    aDay(1) = "Mon" : aDay(2) = "Tue" : aDay(3) = "Wed" : aDay(4) = "Thu"
    aDay(5) = "Fri" : aDay(6) = "Sat" : aDay(7) = "Sun"
    aMon(1)  = "Jan" : aMon(2)  = "Feb" : aMon(3)  = "Mar" : aMon(4)  = "Apr"
    aMon(5)  = "May" : aMon(6)  = "Jun" : aMon(7)  = "Jul" : aMon(8)  = "Aug"
    aMon(9)  = "Sep" : aMon(10) = "Oct" : aMon(11) = "Nov" : aMon(12) = "Dec"
    rssToRFC822 = aDay(Weekday(dt, 2)) & ", " & _
                  Right("0" & Day(dt), 2) & " " & aMon(Month(dt)) & " " & Year(dt) & _
                  " 00:00:00 +0900"
End Function
%>
