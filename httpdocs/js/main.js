"use strict";
/*------------------------
   SMOOTH ANCHOR SCROLL 
------------------------*/
$('a[href*=#]:not([href=#])').click(function() {
	if (location.pathname.replace(/^\//,'') == this.pathname.replace(/^\//,'') && location.hostname == this.hostname) {
	  var target = $(this.hash);
	  target = target.length ? target : $('[name=' + this.hash.slice(1) +']');
	  if (target.length) {
	  	var extra = 0;
	  	// add nav height to scroll offset
	  	if(!$('#nav-wrap.sticky')[0]) {
	  		extra = $('#nav-wrap').outerHeight();
	  	} 
	    $('html,body').animate({
	      scrollTop: target.offset().top - extra
	    }, 1000);
	    return false;
	  }
	}
});

/*-----------------
	COUNTDOWN
-----------------*/	
$('#countdown').xmcountdown({
	width: 130,
	height: 130,
	fillWidth: 6,
	gradient: true,
	gradientColors: ['#f75e7e','#ec1a55'],
	targetDate: new Date(2016, 1, 14)
});

/*-----------------
	TWEETS
-----------------*/	
$('#news-slider').tweet({
    modpath: 'js/twitter/',
    count: 3,
    loading_text: 'Loading twitter feed...',
	username:'odindesign_tw',
	template: '<p>{text}</p><p class="timestamp">{time}</p>'
});

/*-----------------
	ACCORDION
-----------------*/
$('#accordion').xmaccordion({
	startOpen: 2,
	easing: 'swing',
	speed: 600
});

/*-----------------
	TAB
-----------------*/
$('#tab').xmtab({
	fade: true,
	fadeSpeed: 600
});

$(function() {
	/*-----------------
		NEWS SLIDER
	-----------------*/
	$('.tweet_list').bxSlider({
		pager: false,
		prevSelector: '#previous',
		prevText: '<img src="images/arrow-left.png">',
		nextSelector: '#next',
		nextText: '<img src="images/arrow-right.png">',
		easing: 'ease-in-out',
		speed: 800,
		auto: true,
		pause: 6000
	});
});

