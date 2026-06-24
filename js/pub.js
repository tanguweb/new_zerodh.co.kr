//select box
$(document).ready(function () {
	$('.sel_type01').jqTransform();//푸터 페밀리 사이트 셀렉트폼
	
	resizeFunc();
	
	$(window).resize(function(){
		resizeFunc();
	});
	
});

// 롤오버
$(function() {
	$("img").hover(function() {
		var src = $(this).attr("src");
		$(this).attr("src",src.replace("_off","_on"));
	});
	
	$("img").mouseout(function() {
		var src = $(this).attr("src");
		$(this).attr("src",src.replace("_on","_off"));
	});	
});

function resizeFunc()
{
	var windowPosH = $(window).height() - 110;
	
	var wrap1H = $('#left_menu ul').height();
	var wrap2H = $('#contents').height();
	var wrap3H = $('#batting_cart').height();

	var maxWrap = 0;
	
	if(wrap1H > maxWrap) maxWrap = wrap1H;
	if(wrap2H > maxWrap) maxWrap = wrap2H;
	if(wrap3H > maxWrap) maxWrap = wrap3H;
	
	if(maxWrap > windowPosH)
	{
		$('#left_menu').css({'height':(maxWrap)+'px'});
		$('#right_cont').css({'height':(maxWrap)+'px'});
	}
	else
	{
		$('#left_menu').css({'height':(windowPosH)+'px'});
		$('#right_cont').css({'height':(windowPosH)+'px'});
	}
}

// 팝업 레이어 띄우기 닫기
function popup_open() {
	var scrollHeight = document.body.scrollHeight
	var scrollWidth = document.body.scrollWidth
	var bodyHeight = document.body.clientHeight
	var bodyWidth = document.body.clientWidth
	var stHeight = $(window).scrollTop();
	var stWidth = $(window).scrollLeft();
	$("#header .menu_layer").hide();
	$("#layer_mask").css("height", scrollHeight);
	$("#layer_mask").css("width", scrollWidth);
	$("html").css("overflow", "hidden");
	$("#layer_mask").show();
	$("#layer_pop").css("top", stHeight + (bodyHeight / 2));
	$("#layer_pop").css("left", stWidth + (bodyWidth / 2));
	$("#layer_pop").show(200);
};

function popup_close() {
	$("html").css("overflow", "auto");
	$("#layer_mask").hide();
	$("#layer_pop").hide(200);
};

