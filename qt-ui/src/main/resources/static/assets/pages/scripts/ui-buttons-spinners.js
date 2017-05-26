var UIButtons = function () {

    var handleButtons = function () {
        Ladda.bind( '.qt-ladda-btn', { timeout: 2000 } );
        Ladda.bind( '.qt-ladda-btn.mt-progress-demo ', {
                callback: function( instance ) {
                    var progress = 0;
                    var interval = setInterval( function() {
                        progress = Math.min( progress + Math.random() * 0.1, 1 );
                        instance.setProgress( progress );

                        if( progress === 1 ) {
                            instance.stop();
                            clearInterval( interval );
                        }
                    }, 200 );
                }
            } );
    }


    return {
        //main function to initiate the module
        init: function () {
            handleButtons();
        }

    };

}();

jQuery(document).ready(function() {    
   UIButtons.init();
});