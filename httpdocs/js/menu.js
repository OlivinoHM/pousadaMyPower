"use strict";

/*-----------------
	RESIZE IMAGE
-----------------*/
$("figure.imgLiquidFill").imgLiquid({
	verticalAlign: 'top',
    horizontalAlign: 'center'
});

 
/*-----------------
	MAIN MENU
-----------------*/
var mainMenu = $('#main-nav > ul'),
	mainMenuItem = mainMenu.children('li'),
	pull = $('#pull'),
	windowSize = $(window).width();

if(windowSize > 1150) {
	mainMenuItem.bind('mouseenter', toggleSubMenu);
	mainMenuItem.bind('mouseleave', toggleSubMenu);
} else {
	mainMenuItem.bind('click', expandMenu);
}

/**
 * Prevent sub-item click from bubbling up the DOM tree
 */
mainMenu.find('ul > li').on('click', stopBubbling);

function toggleSubMenu() {
	$(this)
		.children('ul')
		.stop()
		.animate({opacity: 'toggle', height: 'toggle'}, 600);
	$(this)
		.children('a')
		.toggleClass('selected');
}

function expandMenu(e) {
	if($(this).children('ul')[0]) {
		e.preventDefault();
		$(this)
			.children('ul')
			.stop()
			.slideToggle(600);
		$(this)
			.children('a')
			.toggleClass('selected');
	}
}

function stopBubbling(e) {
	e.stopPropagation();
}

function toggleMenu(e) {
	e.preventDefault();
    mainMenu.slideToggle();
}

$(pull).on('click', toggleMenu);  