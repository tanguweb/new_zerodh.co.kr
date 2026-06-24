//이미지 링크 점선 삭제
<!--
function bluring() {
  if (event.srcElement.tagName == "A" || event.srcElement.tagName == "IMG") document.body.focus();
}
document.onfocusin = bluring;
// -->


<!--
$(document).ready(function() {



  AOS.init({
    duration: 600,
  })

  //layer_pop
  function layer_pop() {
    $('.layer_pop').fadeIn();
  };
  $(".layer_pop > .dim_bg").click(function() {
    $('.layer_pop').fadeOut();
  });



  // meta viewport 변경
  // if (screen.width < 768) {
  //
  //   $('head').append('<meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" user-scaleable="yes">');
  // }






  // mobile side menu slide
  $(".menu_wrap").click(function() {
    // $('nav').animate({
    //   "right": '+=220'
    // });
    $('header').toggleClass("open");
    $(".sub_lnb, .lnb > li > ul").stop().slideUp(100);
    //$('.menu_wrap').removeClass("open");
    // $("nav").toggleClass("open");
   //$('header > .dim_bg').fadeToggle(200);
    //$('#wrap').css('position','fixed');
  });
  $("#close_but,  nav > div > ul > li > ul > li > a").click(function() {
   $('header').removeClass("open");
   $('nav ul li').removeClass("active");
  $('nav > div > ul > li > ul > li > ul').slideUp('normal');
    //$('header').toggleClass("open");
    //alert('dddd');
    // $('header').removeClass("open");
    //$('header > .dim_bg').fadeOut(200);
    //$('#wrap').css('position', '');
  });


  if (screen.width <= 767) {
    //$("#wrap").removeClass("pc");
    $('.lnb > li > ul ').parent().addClass('has-sub');
    // $("a[href='/category/Instagram']").attr('href', '/5')
    $('.lnb > li > a').click(function() {
      $('.lnb li').removeClass('active');
      $(this).closest('li').addClass('active');
      var checkElement = $(this).next();
      if ((checkElement.is('ul')) && (checkElement.is(':visible'))) {
        $(this).closest('li').removeClass('active');
        checkElement.slideUp('normal');
      }
      if ((checkElement.is('ul')) && (!checkElement.is(':visible'))) {
        $('.lnb ul:visible').slideUp('normal');
        checkElement.slideDown('normal');
      }
      if ($(this).closest('li').find('ul').children().length == 0) {
        return true;
      } else {
        return false;
      }
    });

    $('.lnb > li > ul > li > a').click(function() {
      // alert('More than 767');
      $('.lnb > li > ul > li').removeClass('active');
      $(this).closest('li').addClass('active');
      var checkElement = $(this).next();
      if ((checkElement.is('ul')) && (checkElement.is(':visible'))) {
        $(this).closest('li').removeClass('active');
        checkElement.slideUp('normal');
      }
      if ((checkElement.is('ul')) && (!checkElement.is(':visible'))) {
        $('.lnb > li > ul > li ul:visible').slideUp('normal');
        checkElement.slideDown('normal');
      }
      if ($(this).closest('li').find('ul').children().length == 0) {
        return true;
      } else {
        return false;
      }
    });


    $('.lnb > li > ul').css('height', 'auto');

  } else if (screen.width > 767) {
    // alert('More than 767');
    // $("#wrap").addClass("pc");
    $('.lnb > li').hover(function() {
      //$('header, .close_wrap').addClass("open");
      //$(".sub_lnb ").stop().slideDown(100);
      $(".lnb > li").children('ul').stop().slideDown(100);
     },
     function() {
      $(".lnb > li > ul").stop().slideUp(100);
    });

    // $('.lnb > li > ul > li').hover(function() {
    //   $(this).children('ul').stop().slideDown(100);
    //  },
    //  function() {
    //   $(".lnb > li > ul > li > ul").stop().slideUp(100);
    // });

    //   var marginLeft = parseInt( $("header").css('margin-left') );
    // $(window).scroll(function(e) {
    //   $("header").css("margin-left", marginLeft - $(this).scrollLeft() );
    // });
  }


  if (screen.width < 768) {
    // mobile sub_nav fixed
    //$("#wrap").removeClass("pc");
    var header = $(' .sub_visual ');

    $(window).scroll(function() {
      if ($(this).scrollTop() > 63) {
        header.addClass("f_nav");
      } else {
        header.removeClass("f_nav");
      }
    });
  } else {
    // pc sub_nav fixed
    $('.lnb').hover(function() {
        $("header").addClass('sub_nav_hover');
      }, function() {
        $("header").removeClass('sub_nav_hover');
      });
    var header = $('header, .quick_menu, .sub_nav, .sub_menu ');
    $(window).scroll(function() {
      if ($(this).scrollTop() > 350) {
        header.addClass("f_nav");
        //header.removeClass("sub_nav_hover")
        //sub_nav.addClass("f_nav");
      } else {
        //header.addClass("sub_nav_hover")
        header.removeClass("f_nav");
        //sub_nav.removeClass("f_nav");
      }
    });

  }



});
// -->


///////////////////////////////// Start : foot tab ///////////////////////////////////////
$(document).ready(function(){
    $(".tab_content").hide();
    $(".tab_content:first").show();

    $("ul.tabs li").click(function () {
      // alert('More than 767');
        $("ul.tabs li").removeClass("active");
        //$(this).addClass("active").css({"color": "darkred","font-weight": "bolder"});
        $(this).addClass("active");
        $(".tab_content").hide()
        var activeTab = $(this).attr("rel");
        $("#" + activeTab).fadeIn()
    });
});
///////////////////////////////// Start : Q&A ///////////////////////////////////////
$(document).ready(function(){
  //Add Inactive Class To All Accordion Headers
  $('.faq_q').toggleClass('a_active');

  //Set The Accordion Content Width
  var contentwidth = $('.faq_q').width();
  //$('.faq_a').css({'width' : contentwidth });

  //Open The First Accordion Section When Page Loads
  $('.faq_q').first().toggleClass('q_active').toggleClass('a_active');
  $('.faq_a').first().slideDown().toggleClass('open-content');

  // The Accordion Effect
  $('.faq_q').click(function () {
    if($(this).is('.a_active')) {
      $('.q_active').toggleClass('q_active').toggleClass('a_active').next().slideToggle().toggleClass('open-content');
      $(this).toggleClass('q_active').toggleClass('a_active');
      $(this).next().slideToggle().toggleClass('open-content');
    }

    else {
      $(this).toggleClass('q_active').toggleClass('a_active');
      $(this).next().slideToggle().toggleClass('open-content');
    }
  });

  return false;
});
///////////////////////////////// End : Q&A ///////////////////////////////////////




  // var url = window.location.pathname,
  //   urlRegExp = new RegExp(url.replace(/\/$/,'') + "$");
  //   $('.sub_nav a').each(function(){
  //       if(urlRegExp.test(this.href.replace(/\/$/,''))){
  //           $(this).parent().addClass('active');
  //       }
  //   });






    $(function() {
      var url = window.location.pathname,
        urlRegExp = new RegExp(url.replace(/\/$/,'') + "$");
        $('.sub_top_menu > ul > li a, .sub_nav li a').each(function(){
            if(urlRegExp.test(this.href.replace(/\/$/,''))){
                $(this).parent().addClass('active');
            }

        });
      });



      //input file
        var fileTarget = $('.filebox .upload-hidden');

            fileTarget.on('change', function(){
                if(window.FileReader){
                    var filename = $(this)[0].files[0].name;
                } else {
                    var filename = $(this).val().split('/').pop().split('\\').pop();
                }

                $(this).siblings('.upload-name').val(filename);
            });

            ///////////////////////////////// Start : click scroll ///////////////////////////////////////

            $(function() {

              // Cache selectors
              var lastId,
               sub_nav = $(".sub_nav");
               if (screen.width < 768) {
                   topMenuHeight = 50;
               } else {
                   topMenuHeight = 50;
               }

               // All list items
               menuItems = sub_nav.find("a"),
               // Anchors corresponding to menu items
               scrollItems = menuItems.map(function(){
                 var item = $($(this).attr("href"));
                  if (item.length) { return item; }
               });

              // Bind to scroll
              $(window).scroll(function(){
                 // Get container scroll position
                 var fromTop = $(this).scrollTop()+topMenuHeight + 50;

                 // Get id of current scroll item
                 var cur = scrollItems.map(function(){
                   if ($(this).offset().top < fromTop)
                     return this;
                 });
                 // Get the id of the current element
                 cur = cur[cur.length-1];
                 var id = cur && cur.length ? cur[0].id : "";

                 if (lastId !== id) {
                     lastId = id;
                     // Set/remove active class
                     menuItems
                       .parent().removeClass("active")
                       .end().filter("[href=#"+id+"]").parent().addClass("active");
                 }
              });

              //
              $('a[href*="#"]:not([href="#"])').click(function() {
                var target = $(this.hash);
                $('html,body').stop().animate({scrollTop: target.offset().top - topMenuHeight}, 1500);
              });


              if (location.hash){
                var id = $(location.hash);
              }
              $(window).load(function() {
                if (location.hash){
                  $('html,body').animate({scrollTop: id.offset().top -topMenuHeight}, 1500)
                };
              });
            });

            ///////////////////////////////// End : click scroll ///////////////////////////////////////
