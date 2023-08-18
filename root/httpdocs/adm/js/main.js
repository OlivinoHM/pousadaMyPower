jQuery(function ($) {
	"use strict";
	/*    ----------------------------------------------------------- */
	/*  #Fullpage
	/* ----------------------------------------------------------- */
	var currentSection;
	var currentSlide=1;
	$('.body').fullpage({
		fixedElements: '#navigation, .vmenu-wrapper, .site-overlay',
		 
		navigation: true,
		resize: false,
		scrollingSpeed: 800,
		slidesNavigation: false,
		autoScrolling:false,
		paddingTop: '0px',
		scrollBar: false,
		paddingBottom: '0px',
		afterSlideLoad: function (anchorLink, index, slideAnchor, slideIndex) {		
			currentSection = index;
			currentSlide = slideIndex;
		},
		onSlideLeave: function( anchorLink, index, slideIndex, direction){			
		},
		afterRender: function(){
			if($('#photostack').length > 0)
				new Photostack( document.getElementById( 'photostack' ));
		}
	});
	
	/* ----------------------------------------------------------- */
	/*  #Menu & Navigations
	/* ----------------------------------------------------------- */	
	$("#accordion.menu").dcAccordion({
		eventType: "click",
		hoverDelay: 600,
		menuClose: true,
		autoClose: true,
		saveState: false,
		autoExpand: true,
		classExpand: "current-menu-item",
		classDisable: "",
		showCount: false,
		disableLink: false,
		cookie: "dc_jqaccordion_widget-s1-item",
		speed: "fast"
	});
	if($('.section.active .slide').length <= 1){
		$('.navbar .next').css({'display':'none'});
		$('.navbar .prev').css({'display':'none'});
	}
	$('.navbar .next').click(function(e){
		$.fn.fullpage.moveSlideRight();
	});
	$('.navbar .prev').click(function(e){
		$.fn.fullpage.moveSlideLeft();
	});
	$('body').on('click', '.menubar', function() {
		$('html').addClass('pushed');
	});	
	$('body').on('click', '.site-overlay', function() {
		$('html').removeClass('pushed');
	});
	$('.vmenu').slimScroll({
		destroy: true,
	});
});