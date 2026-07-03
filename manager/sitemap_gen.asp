<!-- #include virtual = "/Include/Config.asp" -->
<!-- #include virtual = "/function/fn_AdminLoginCheck.asp" -->
<!-- #include virtual = "/manager/inc/inc.seogen.asp" -->
<%
'########################################################################################
'#  File    : /manager/sitemap_gen.asp
'#  Info    : sitemap.xml 재생성 (정적 페이지 + 전체 게시글)
'#            실제 생성 로직은 /manager/inc/inc.seogen.asp (SeoGenSitemap) 공용 모듈
'########################################################################################

Dim bGenerate : bGenerate = (Request.Form("act") = "generate")

If bGenerate Then
    Dim iCount : iCount = SeoGenSitemap()
    Response.Write "<script>alert('sitemap.xml 생성 완료!\n\n정적 페이지 + 게시글 " & iCount & "개');location.href='sitemap_gen.asp';</script>"
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
                <h2>Sitemap 재생성</h2>
                <ol class="breadcrumb">
                    <li><a href="/manager/">Home</a></li>
                    <li class="active"><strong>Sitemap 재생성</strong></li>
                </ol>
            </div>
        </div>

        <div class="wrapper wrapper-content animated fadeInRight">
            <div class="row">
                <div class="col-lg-6 col-lg-offset-3">
                    <div class="ibox float-e-margins">
                        <div class="ibox-title">
                            <h5><i class="fa fa-sitemap"></i> sitemap.xml 재생성</h5>
                        </div>
                        <div class="ibox-content text-center" style="padding:40px;">
                            <p class="text-muted" style="margin-bottom:20px;">
                                DB에서 전체 게시글을 조회하여 <strong>sitemap.xml</strong>을 새로 생성합니다.<br>
                                정적 페이지 + 4개 게시판(1002·1009·1010·1004) 전체 게시글이 포함됩니다.
                            </p>
                            <table class="table table-bordered" style="margin-bottom:20px;text-align:left;">
                                <thead><tr><th>구분</th><th>URL</th><th>우선순위</th></tr></thead>
                                <tbody>
                                    <tr><td>메인</td><td>/html/main/</td><td>1.00</td></tr>
                                    <tr><td>의료진소개</td><td>/html/introduce/doctor.html</td><td>0.90</td></tr>
                                    <tr><td>병원소개</td><td>/html/introduce/life·space·time·contact.html (4)</td><td>0.80</td></tr>
                                    <tr><td>수면&amp;디지털센터</td><td>/html/sleep/sedation·lab·oneday.html (3)</td><td>0.80</td></tr>
                                    <tr><td>임플란트</td><td>/html/implant/computer·full·sinus·bone·denture·immediate·prf·uv.html (8)</td><td>0.80</td></tr>
                                    <tr><td>심미치료</td><td>/html/beauty/layer·resin·whitening·gum_whitening·tooth_form·gum_form.html (6)</td><td>0.80</td></tr>
                                    <tr><td>일반치료</td><td>/html/general/cavity·nerve·prosthetic·wisdom·scaling·tmj·botox·bruxism.html (8)</td><td>0.80</td></tr>
                                    <tr><td>소아치료</td><td>/html/pediatric/ortho·endo·mesiodens·prevention.html (4)</td><td>0.80</td></tr>
                                    <tr><td>커뮤니티</td><td>/html/board/bf·news·faq·column·event.html, /html/introduce/fee·price.html (7)</td><td>0.80</td></tr>
                                    <tr><td>개인정보처리방침</td><td>/html/board/personal_info.html</td><td>0.40</td></tr>
                                    <tr><td colspan="3" class="text-muted">+ 게시판 1002 / 1009 / 1010 / 1004 개별 게시글 (view)</td></tr>
                                </tbody>
                            </table>
                            <form method="post" style="display:inline-block;margin:0;">
                                <input type="hidden" name="act" value="generate">
                                <button type="submit" class="btn btn-primary btn-lg">
                                    <i class="fa fa-refresh"></i> sitemap.xml 생성
                                </button>
                            </form>
                            <a href="/sitemap.xml" target="_blank" class="btn btn-default btn-lg" style="margin-left:8px;">
                                <i class="fa fa-external-link"></i> sitemap.xml 보기
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