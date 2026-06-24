$(function () {
    var eventNav = new function () {
        this.setIntervalTimer;
        this.setPausedTimeout;
        this.resumeTimer;
        this.resumeTime = 2500;
        this.interval = 2500;
        this.currentIdx = 0;

        this.start = function () {
            var _self = this;

            this.goRight();
        }
        this.direction;
        this.goRight = function () {
            var _self = this;

            this.direction = 'right';

            var $twentyContainer = $('.twentytwenty-container');
            $('.twentytwenty-handle').stop().animate({
                left: $twentyContainer.width(),
            }, {
                duration: 2000,
                easing: 'linear',
                step: function (now, fx) {
                    $('.twentytwenty-before').css({
                        clip: 'rect(0px, ' + now + 'px, ' + $twentyContainer.height() + 'px, 0px)',
                    })
                    $('.twentytwenty-after').css({
                        clip: 'rect(0px, ' + $twentyContainer.width() + 'px, ' + $twentyContainer.height() + 'px, ' + now + 'px)',
                    })
                },
                complete: function () {
                    _self.onCompleteSlide();
                },
            });

        }
        this.onCompleteSlide = function () {
            if (this.direction == 'right') {
                this.goLeft();
            } else {
                this.goRight();
            }
        }
        this.goLeft = function () {
            var _self = this;
            this.direction = 'left';

            var $twentyContainer = $('.twentytwenty-container');
            $('.twentytwenty-handle').stop().animate({
                left: 0,
            }, {
                duration: 2000,
                easing: 'linear',
                step: function (now, fx) {


                    $('.twentytwenty-before').css({
                        clip: 'rect(0px, ' + now + 'px, ' + $twentyContainer.height() + 'px, 0px)',
                    })
                    $('.twentytwenty-after').css({
                        clip: 'rect(0px, ' + $twentyContainer.width() + 'px, ' + $twentyContainer.height() + 'px, ' + now + 'px)',
                    })
                },
                complete: function () {
                    _self.onCompleteSlide();
                },
            });
        }
        this.pause = function () {
            this.stop();

            if (this.setPausedTimeout) {
                clearTimeout(this.setPausedTimeout);
            }
            this.setPausedTimeout = setTimeout(function () {
                eventNav.start();
            }, this.resumeTime);
        }
        this.stop = function () {
            $('.twentytwenty-handle').stop()
        }
    };
    eventNav.start();


    $('.twentytwenty-handle').on('mouseover touchstart touchmove', function () {
        eventNav.stop();
    }).on('mouseleave touchend', function () {
        eventNav.start();
    });


    /*var eventNav = new function(){
        this.setIntervalTimer;
        this.setPausedTimeout;
        this.resumeTimer;
        this.resumeTime = 2500;
        this.interval = 2500;
        this.currentIdx = 0;

        this.start = function(){
            var _self = this;

            this.stop();
            this.setIntervalTimer = setInterval(function(){
                _self.currentIdx = _self.currentIdx + 1;
                if(_self.currentIdx >= $(".event-nav").length){
                    _self.currentIdx = 0;
                }
                $(".event-nav").eq(_self.currentIdx).click();
            },this.interval);
        }
        this.pause = function(){
            this.stop();

            if(this.setPausedTimeout){
                clearTimeout(this.setPausedTimeout);
            }
            this.setPausedTimeout = setTimeout(function(){
                eventNav.start();
            }, this.resumeTime);
        }
        this.stop = function(){
            if(this.setIntervalTimer){
                clearInterval(this.setIntervalTimer);
            }
            if(this.setPausedTimeout){
                clearTimeout(this.setPausedTimeout);
            }
        }
    };
    // eventNav.start();


    $('.twentytwenty-handle').on('mouseover touchstart touchmove', function(){
        eventNav.stop();
    }).on('mouseleave touchend', function(){
        eventNav.start();
    })
    */

    // $(".event-nav").on("click", function () {
    //     $(".event-nav").removeClass("event-nav-active")
    //     $(this).addClass("event-nav-active")
    //
    // });
    //
    //
    // var baImage = [
    //     "//byulstar.com/front/images/20200828/1.jpg",
    //     "//byulstar.com/front/images/20200828/1_1.jpg",
    //     "//byulstar.com/front/images/20200828/2.jpg",
    //     "//byulstar.com/front/images/20200828/2_1.jpg",
    //     "//byulstar.com/front/images/20200828/3.jpg",
    //     "//byulstar.com/front/images/20200828/3_1.jpg",
    //     "//byulstar.com/front/images/20200828/4.jpg",
    //     "//byulstar.com/front/images/20200828/4_1.jpg",
    //     "//byulstar.com/front/images/20200828/5.jpg",
    //     "//byulstar.com/front/images/20200828/5_1.jpg",
    //     "//byulstar.com/front/images/20200828/6.jpg",
    //     "//byulstar.com/front/images/20200828/6_1.jpg",
    //     "//byulstar.com/front/images/20200828/7.jpg",
    //     "//byulstar.com/front/images/20200828/7_1.jpg",
    //     "//byulstar.com/front/images/20200828/8.jpg",
    //     "//byulstar.com/front/images/20200828/8_1.jpg",
    //     "//byulstar.com/front/images/20200828/9.jpg",
    //     "//byulstar.com/front/images/20200828/9_1.jpg",
    //     "//byulstar.com/front/images/20200828/10.jpg",
    //     "//byulstar.com/front/images/20200828/10_1.jpg",
    //     "//byulstar.com/front/images/20200828/11.jpg",
    //     "//byulstar.com/front/images/20200828/11_1.jpg",
    //     "//byulstar.com/front/images/20200828/12.jpg",
    //     "//byulstar.com/front/images/20200828/12_1.jpg",
    //     "//byulstar.com/front/images/20200828/13.jpg",
    //     "//byulstar.com/front/images/20200828/13_1.jpg",
    //     "//byulstar.com/front/images/20200828/14.jpg",
    //     "//byulstar.com/front/images/20200828/14_1.jpg",
    //     "//byulstar.com/front/images/20200828/15.jpg",
    //     "//byulstar.com/front/images/20200828/15_1.jpg",
    //     "//byulstar.com/front/images/20200828/16.jpg",
    //     "//byulstar.com/front/images/20200828/16_1.jpg"
    //
    // ];
    //
    // var imgTarget = $(".twentytwenty-container").find("img:first").attr("src");
    //
    //
    //
    //





    // $(".event-nav").eq(0).on("click", function() {
    //     $(".twentytwenty-container").find(".twentytwenty-before").attr("src", baImage[0]);
    //     $(".twentytwenty-container").find(".twentytwenty-after").attr("src", baImage[1]);
    // })
    //
    // $(".event-nav").eq(1).on("click", function() {
    //     $(".twentytwenty-container").find(".twentytwenty-before").attr("src", baImage[2]);
    //     $(".twentytwenty-container").find(".twentytwenty-after").attr("src", baImage[3]);
    // })
    //
    // $(".event-nav").eq(2).on("click", function() {
    //     $(".twentytwenty-container").find(".twentytwenty-before").attr("src", baImage[4]);
    //     $(".twentytwenty-container").find(".twentytwenty-after").attr("src", baImage[5]);
    // })
    //
    // $(".event-nav").eq(3).on("click", function() {
    //     $(".twentytwenty-container").find(".twentytwenty-before").attr("src", baImage[6]);
    //     $(".twentytwenty-container").find(".twentytwenty-after").attr("src", baImage[7]);
    // })
    // $(".event-nav").eq(4).on("click", function() {
    //     $(".twentytwenty-container").find(".twentytwenty-before").attr("src", baImage[8]);
    //     $(".twentytwenty-container").find(".twentytwenty-after").attr("src", baImage[9]);
    // })
    // $(".event-nav").eq(5).on("click", function() {
    //     $(".twentytwenty-container").find(".twentytwenty-before").attr("src", baImage[10]);
    //     $(".twentytwenty-container").find(".twentytwenty-after").attr("src", baImage[11]);
    // })
    // $(".event-nav").eq(6).on("click", function() {
    //     $(".twentytwenty-container").find(".twentytwenty-before").attr("src", baImage[12]);
    //     $(".twentytwenty-container").find(".twentytwenty-after").attr("src", baImage[13]);
    // })
    //
    // $(".event-nav").eq(7).on("click", function() {
    //     $(".twentytwenty-container").find(".twentytwenty-before").attr("src", baImage[14]);
    //     $(".twentytwenty-container").find(".twentytwenty-after").attr("src", baImage[15]);
    // })
    //
    // $(".event-nav").eq(8).on("click", function() {
    //     $(".twentytwenty-container").find(".twentytwenty-before").attr("src", baImage[16]);
    //     $(".twentytwenty-container").find(".twentytwenty-after").attr("src", baImage[17]);
    // })
    //
    // $(".event-nav").eq(9).on("click", function() {
    //     $(".twentytwenty-container").find(".twentytwenty-before").attr("src", baImage[18]);
    //     $(".twentytwenty-container").find(".twentytwenty-after").attr("src", baImage[19]);
    // })
    //
    // $(".event-nav").eq(10).on("click", function() {
    //     $(".twentytwenty-container").find(".twentytwenty-before").attr("src", baImage[20]);
    //     $(".twentytwenty-container").find(".twentytwenty-after").attr("src", baImage[21]);
    // })
    //
    // $(".event-nav").eq(11).on("click", function() {
    //     $(".twentytwenty-container").find(".twentytwenty-before").attr("src", baImage[22]);
    //     $(".twentytwenty-container").find(".twentytwenty-after").attr("src", baImage[23]);
    // })
    //
    // $(".event-nav").eq(12).on("click", function() {
    //     $(".twentytwenty-container").find(".twentytwenty-before").attr("src", baImage[24]);
    //     $(".twentytwenty-container").find(".twentytwenty-after").attr("src", baImage[25]);
    // })
    // $(".event-nav").eq(13).on("click", function() {
    //     $(".twentytwenty-container").find(".twentytwenty-before").attr("src", baImage[26]);
    //     $(".twentytwenty-container").find(".twentytwenty-after").attr("src", baImage[27]);
    // })
    //
    //  $(".event-nav").eq(14).on("click", function() {
    //     $(".twentytwenty-container").find(".twentytwenty-before").attr("src", baImage[28]);
    //     $(".twentytwenty-container").find(".twentytwenty-after").attr("src", baImage[29]);
    // })
    //
    //  $(".event-nav").eq(15).on("click", function() {
    //     $(".twentytwenty-container").find(".twentytwenty-before").attr("src", baImage[30]);
    //     $(".twentytwenty-container").find(".twentytwenty-after").attr("src", baImage[31]);
    //
    // })





})
