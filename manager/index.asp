<!-- #include virtual = "/Include/Config.asp" -->
<!-- #include virtual = "/function/fn_AdminLoginCheck.asp" -->
<!--#include virtual="/function/fn_CodeInfo.asp"-->
<%
	'########################################################################################
	'#	File		: /manager/index.asp
	'#  Create		: domoyosi / 2015.12
	'#	Info		: 어드민 index 페이지
	'#	Update		: 
	'#	Update Memo	: 
	'########################################################################################

	Dim objADO, SQL, iloop, jloop, kloop
	Dim arrRsOnlineCounsel, arrRsQuickCounsel

	Set objADO = new clsADO
		
	'온라인상담
	SQL = ""
	SQL = SQL & "SELECT TOP 10 CD_BOARDID, NM_TITLE, NM_CONTENTS, NM_FIELD_1, NM_FIELD_2, NM_FIELD_3, NM_FIELD_4, NM_FIELD_5, NM_FIELD_6, NM_FIELD_7, NM_FIELD_8, NM_FIELD_9, NM_FIELD_10, NM_FIELD_11, NM_FIELD_12, DT_INSYSDATE	"
	SQL = SQL & "FROM T_BOARDDETAIL	"
	SQL = SQL & "WHERE CD_BOARDCD = '1003' AND YN_USE = 'Y' AND YN_DISPLAY = 'Y' and NO_BOARD_DEPTH = 1	"
	SQL = SQL & "ORDER BY CD_BOARDID DESC	"
	
	objADO.setSql(SQL)
	arrRsOnlineCounsel = objADO.getArrRs()

	'빠른상담
'	SQL = ""
'	SQL = SQL & "SELECT TOP 10 CD_BOARDID, NM_TITLE, NM_CONTENTS, NM_FIELD_1, NM_FIELD_2, NM_FIELD_3, NM_FIELD_4, NM_FIELD_5, NM_FIELD_6, NM_FIELD_7, NM_FIELD_8, NM_FIELD_9, NM_FIELD_10, NM_FIELD_11, NM_FIELD_12, DT_INSYSDATE	"
'	SQL = SQL & "FROM T_BOARDDETAIL	"
'	SQL = SQL & "WHERE CD_BOARDCD = '1008' AND YN_USE = 'Y' AND YN_DISPLAY = 'Y' and NO_BOARD_DEPTH = 1	"
'	SQL = SQL & "ORDER BY CD_BOARDID DESC	"
	
'	objADO.setSql(SQL)
'	arrRsQuickCounsel = objADO.getArrRs()
	
	Set objADO = Nothing

	' 상담구분
	Dim aryCounselType : aryCounselType = fn_CodeInfo("COUNSEL_TYPE")
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
                    <h2>Dashboard</h2>
					<ol class="breadcrumb">
                        <li>
                            <a href="/manager/">Home</a>
                        </li>
                        <li class="active">
                            <strong>Dashboard</strong>
                        </li>
                    </ol>
                </div>
            </div>
			
			<div class="wrapper wrapper-content animated fadeInRight">
				
				<div class="row">
                    <div class="col-lg-4">
                        <div class="ibox float-e-margins">
                            <div class="ibox-title">
                                <h5>온라인상담</h5>
                                <div class="ibox-tools">
                                    <a class="collapse-link">
                                        <i class="fa fa-chevron-up"></i>
                                    </a>
                                </div>
                            </div>
                            <div class="ibox-content ibox-heading">
                                <h3><i class="fa fa-headphones"></i> 신규 온라인상담</h3>
                                <small><i class="fa fa-tim"></i> 최근 10개의 온라인상담을 조회합니다.</small>
                            </div>
                            <div class="ibox-content">
                                <div class="feed-activity-list">
								<%
									If IsArray(arrRsOnlineCounsel) Then 
										For iloop = 0 To UBound(arrRsOnlineCounsel,2)	
								%>
									<div class="feed-element">
                                        <div>
                                            <!--<small class="pull-right text-navy">1m ago</small>-->
                                            <strong class="text-navy">
												<%=arrRsOnlineCounsel(3,iloop)%>
											</strong>
											<div>
												<strong class="text-success"><%=CheckLength(StripHTML(arrRsOnlineCounsel(1,iloop)),35)%></strong>
											</div>
											<div>
												<i class="fa fa-mobile"></i> : <%=addHyphen(arrRsOnlineCounsel(4,iloop))%>
												&nbsp; &nbsp;/&nbsp; &nbsp;
												<i class="fa fa-at"></i> : <%=arrRsOnlineCounsel(5,iloop)%>
											</div>
                                            <div> - 상태 : <strong class="text-muted">
													<%
													For jloop = 0 To UBound(aryCounselType, 2)
														If aryCounselType(1,jloop) = arrRsOnlineCounsel(6,iloop) Then 
													%>
														<%=aryCounselType(2,jloop)%>
													<%
														End If 
													Next
													%>
												</strong>
											</div>
											<div> - 답변수신여부 : <strong class="text-muted">
													<%=getBoolean(arrRsOnlineCounsel(11,iloop)="Y","수신","미수신")%>
												</strong>
											</div>
                                            <small class="pull-right text-muted"><%=arrRsOnlineCounsel(15,iloop)%></small>
                                        </div>
                                    </div>
								<%
										Next
									Else
								%>
									  <div class="text-center">최근 온라인상담이 없습니다.</div>
								<%
									End If
								%>
                                </div>
                            </div>
                        </div>
                    </div>


				</div>

			</div>

			<!--#include virtual="/manager/inc/inc.ui.footer.asp"-->
		</div>
	</div>

	<!--#include virtual="/manager/inc/inc.ui.hidden.asp"-->
</body>
</html>		