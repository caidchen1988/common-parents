/**
 * Created by dell on 2017/5/26.
 */

var Common = {

    /**
     * 提示
     * @param options
     * {
     *  _type: success/error,
     *  _title: '标题',
     *  _msg: '内容'
     * }
     */
    toastr : function (options) {

        toastr.options = {
            "closeButton": false, //是否显示关闭按钮
            "debug": false, //是否使用debug模式
            "positionClass": "toast-top-full-width",//弹出窗的位置
            "showDuration": "300",//显示的动画时间
            "hideDuration": "1000",//消失的动画时间
            "timeOut": "1000", //展现时间
            "extendedTimeOut": "1000",//加长展示时间
            "showEasing": "swing",//显示时的动画缓冲方式
            "hideEasing": "linear",//消失时的动画缓冲方式
            "showMethod": "fadeIn",//显示时的动画方式
            "hideMethod": "fadeOut" //消失时的动画方式
        };

        toastr[options._type](options._msg, options._title)
    }
}