var FormValidationMd = function() {

    /**
     * 表单校验
     * @param options
     * {
     *     _form: "#form_id",   // 表单元素ID
     *     _messages: json,     // 表单校验不通过提示信息
     *     _rules: json         // 表单校验规则
     * }
     */
    var handleValidation = function(options) {
        // for more info visit the official plugin documentation: 
        // http://docs.jquery.com/Plugins/Validation
        var form = $(options._form);
        var error = $('.alert-danger', form);
        var success = $('.alert-success', form);

        form.validate({
            errorElement: 'span', //default input error message container
            errorClass: 'help-block help-block-error', // default input error message class
            focusInvalid: true, // do not focus the last invalid input
            ignore: "", // validate all fields including form hidden input
            messages: options._messages,
            rules: options._rules,

            invalidHandler: function(event, validator) { //display error alert on form submit              
                success.hide();
                error.show();
                App.scrollTo(error, -200);
            },

            errorPlacement: function(error, element) {
                if (element.is(':checkbox')) {
                    error.insertAfter(element.closest(".md-checkbox-list, .md-checkbox-inline, .checkbox-list, .checkbox-inline"));
                } else if (element.is(':radio')) {
                    error.insertAfter(element.closest(".md-radio-list, .md-radio-inline, .radio-list,.radio-inline"));
                } else {
                    error.insertAfter(element); // for other inputs, just perform default behavior
                }
            },

            highlight: function(element) { // hightlight error inputs
                $(element)
                    .closest('.form-group').addClass('has-error'); // set error class to the control group
            },

            unhighlight: function(element) { // revert the change done by hightlight
                $(element)
                    .closest('.form-group').removeClass('has-error'); // set error class to the control group
            },

            success: function(label) {
                label.closest('.form-group').removeClass('has-error'); // set success class to the control group
            },

            submitHandler: function(form) {
                success.show();
                error.hide();

                // 屏蔽提交按钮，防止页面重复提交操作
                var btn = Ladda.create( document.querySelector( 'button[type=submit]' ) );

                $(form).ajaxSubmit({
                    dataType:"json",
                    beforeSubmit : function () {
                        btn.start() ;
                    },
                    success:function( data ){

                        // 实际开发中需要删除setTimeout
                        setTimeout(function () {
                            btn.stop();

                            Common.toastr({
                                _type: 'success',
                                _title: '表单提交',
                                _msg: '数据提交成功'
                            }) ;

                        },2000);
                    },
                    error : function ( data ) {
                        btn.stop();

                        Common.toastr({
                            _type: 'error',
                            _title: '表单提交',
                            _msg: '网络超时，请重试或者联系管理员'
                        }) ;
                    }
                });
            }
        });
    }

    return {
        //main function to initiate the module
        init: function(options) {
            handleValidation(options);
        }
    };
}();

