function openSocialFacebook(url, msg){
    var href = "http://www.facebook.com/sharer.php?u=" + url + "&t=" + msg;
    window.open(href, "facebook", "");
}

function openSocialTwitter(url, msg){
    var href = "http://twitter.com/home?status=" + msg + " " + encodeURIComponent(url);
    window.open(href, "twitter", "");
}

function openSocialKakaoStory(url){
    var href = "https://story.kakao.com/share?url=" + encodeURIComponent(url);
    window.open(href, "kakaostory", "");
}