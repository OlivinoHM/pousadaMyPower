"use strict";

// start custom scripts
(function($) {    
jQuery(document).ready(function($){

    


    /*==========  Slideshow  ==========*/ 
    $(".rslides").responsiveSlides({
        auto: true,
        pager: true
    });

    var $window = $(window).on('resize', function(){
       var windowHeight = $(this).height();
       var windowWidth = verge.viewportW();
        if( windowWidth < 993) {
            $('.rslides li').css('height' , windowHeight);
        } else {
            $('.rslides li').css('height' , 'auto');
        }
    }).trigger('resize');


    /*==========  Contact Form Validation  ==========*/ 
    function isValidEmailAddress(emailAddress) {
        var pattern = new RegExp(/^(("[\w-\s]+")|([\w-]+(?:\.[\w-]+)*)|("[\w-\s]+")([\w-]+(?:\.[\w-]+)*))(@((?:[\w-]+\.)*\w[\w-]{0,66})\.([a-z]{2,6}(?:\.[a-z]{2})?)$)|(@\[?((25[0-5]\.|2[0-4][0-9]\.|1[0-9]{2}\.|[0-9]{1,2}\.))((25[0-5]|2[0-4][0-9]|1[0-9]{2}|[0-9]{1,2})\.){2}(25[0-5]|2[0-4][0-9]|1[0-9]{2}|[0-9]{1,2})\]?$)/i);
        return pattern.test(emailAddress);
    }    

    $('.contact-form form').submit(function() {
    
        var hasError = false;
      
        var message = $('#your_message').val();
        if ($.trim(message) == '') {
            $('#your_message').parent().addClass('has-error');
            $('#your_message').focus();
            hasError = true;
        }
        else {
            $('#message-txt').parent().removeClass('has-error');
        }

        var subject = $('#your_subject').val();
        if ($.trim(subject) == '') {
            $('#your_subject').parent().addClass('has-error');
            $('#your_subject').focus();
            hasError = true;
        }
        else {
            $('#your_subject').parent().removeClass('has-error');
        }       
        
       var phone = $('#your_phone').val();
        if ($.trim(phone) == '') {
            $('#your_phone').parent().addClass('has-error');
            $('#your_phone').focus();
            hasError = true;  
        }
        else {
            $('#your_phone').parent().removeClass('has-error');
        }

        var emailVal = $('#your_email').val();
        if ($.trim(emailVal) == '' || !isValidEmailAddress(emailVal)) {
            $('#your_email').parent().addClass('has-error');
            $('#your_email').focus();
            hasError = true;
        }
        else {
            $('#your_email').parent().removeClass('has-error');
        }                

       var fullname = $('#your_name').val();
        if ($.trim(fullname) == '') {
            $('#your_name').parent().addClass('has-error');
            $('#your_name').focus();
            hasError = true;            
        }
        else {
            $('#your_name').parent().removeClass('has-error');
        }        
        
        if (!hasError) {
            $('#submit').fadeOut('normal', function(){
                $('.loading').css({
                    display: "block"
                });
                
            });
            
            $.post($('.contact-form form').attr('action'), $('.contact-form form').serialize(), function(data){
                $('.log').html(data);
                $('.loading').remove();
                $('.contact-form form').slideUp('slow');
            });
            
        }
        
        return false;
        
    });  


}); // end jquery init
})(jQuery);
