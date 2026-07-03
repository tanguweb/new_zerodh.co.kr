<!-- #include virtual = "/Include/Config.asp" -->
<!-- #include virtual = "/function/fn_AdminLoginCheck.asp" -->
<!-- #include virtual = "/manager/inc/inc.seogen.asp" -->
<%
'########################################################################################
'#  File    : /manager/rss_gen.asp
'#  Info    : rss.xml 재생성 (게시글 최신 100건, 날짜순 정렬)
'#            실제 생성 로직은 /manager/inc/inc.seogen.asp (SeoGenRss) 공용 모듈
'########################################################################################

Dim bGenerate : bGenerate = (Request.Form("act") = "generate")

If bGenerate Then
    Dim iTotal : iTotal = SeoGenRss()
    Response.Write "<script>alert('rss.xml 생성 완료!\n\n게시글 " & iTotal & "건');location.href='rss_gen.asp';</script>"
    Response.End
End If
%>
<!doctype html>
<html>
<head>
    <!--#include virtual="/manager/inc/inc.ui.head.asp"-->
</head>
<body class="skin-2">
<div id="wrapper">
    <!--#include virtual="/manager/inc/inc.ui.navi.asp"-->
    <div id="page-wrapper" class="gray-bg dashbard-1">
        <!--#include virtual="/manager/inc/inc.ui.top.asp"-->

        <div class="row wrapper border-bottom white-bg page-heading">
            <div class="col-lg-10">
                <h2>RSS 재생성</h2>
                <ol class="breadcrumb">
                    <li><a href="/manager/">Home</a></li>
                    <li class="active"><strong>RSS 재생성</strong></li>
                </ol>
            </div>
        </div>

        <div class="wrapper wrapper-content animated fadeInRight">
            <div class="row">
                <div class="col-lg-6 col-lg-offset-3">
                    <div class="ibox float-e-margins">
                        <div class="ibox-title">
                            <h5><i class="fa fa-rss"></i> rss.xml 재생성</h5>
                        </div>
                        <div class="ibox-content text-center" style="padding:40px;">
                            <p class="text-muted" style="margin-bottom:20px;">
                                DB에서 최신 게시글 최대 <strong>100건</strong>을 날짜순으로 조회하여<br>
                                <strong>rss.xml</strong>을 새로 생성합니다.<br>
                                4개 게시판(1002·1009·1010·1004) 전시 게시글이 포함됩니다.
                            </p>
                            <table class="table table-bordered" style="margin-bottom:20px;text-align:left;">
                                <thead><tr><th>게시판코드</th><th>게시판명</th><th>view 페이지</th></tr></thead>
                                <tbody>
                                    <tr><td>1002</td><td>공지사항</td><td>news_view.html</td></tr>
                                    <tr><td>1009</td><td>제로치과칼럼</td><td>column_view.html</td></tr>
                                    <tr><td>1010</td><td>자주묻는질문</td><td>faq_view.html</td></tr>
                                    <tr><td>1004</td><td>치료전후</td><td>bf_view.html</td></tr>
                                </tbody>
                            </table>
                            <p class="text-muted" style="margin-bottom:20px;font-size:12px;">
                                * 글 저장·수정·삭제 시에도 rss.xml이 자동으로 재생성됩니다.
                            </p>
                            <form method="post" style="display:inline-block;margin:0;">
                                <input type="hidden" name="act" value="generate">
                                <button type="submit" class="btn btn-primary btn-lg">
                                    <i class="fa fa-refresh"></i> rss.xml 생성
                                </button>
                            </form>
                            <a href="/rss.xml" target="_blank" class="btn btn-default btn-lg" style="margin-left:8px;">
                                <i class="fa fa-external-link"></i> rss.xml 보기
                            </a>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!--#include virtual="/manager/inc/inc.ui.footer.asp"-->
    </div>
</div>
</body>
</html>
