/***
 * @author guyl, wangxf  https://static.122.gov.cn/V1.16.1/veh1/static/js/base.js"
 * @returns
 */
function _getScriptBase() {
    var scripts = document.getElementsByTagName('script');
    for (var i = 0; i < scripts.length; i++) {
        var src = scripts[i].src, idx = src.indexOf('/static/');
        if (idx > -1)
            return src.substr(0, idx);
    }
}
function dateFormat(format, date) {
    date = date || new Date();
    var o = {
        "M+": date.getMonth() + 1, //month
        "d+": date.getDate(), //day
        "h+": date.getHours(), //hour
        "m+": date.getMinutes(), //minute
        "s+": date.getSeconds(), //second
        "q+": Math.floor((date.getMonth() + 3) / 3), //quarter
        "S": date.getMilliseconds() //millisecond
    }

    if (/(y+)/.test(format)) {
        format = format.replace(RegExp.$1, (date.getFullYear() + "").substr(4 - RegExp.$1.length));
    }

    for (var k in o) {
        if (new RegExp("(" + k + ")").test(format)) {
            format = format.replace(RegExp.$1, RegExp.$1.length == 1 ? o[k] : ("00" + o[k]).substr(("" + o[k]).length));
        }
    }
    return format;
}
var dateStr = dateFormat('yyyyMMdd', new Date());
var version = "1.5.0_" + dateStr;
require.config({
    urlArgs: version,
    baseUrl: _getScriptBase(),
    paths: {
        js: 'static/js',
        lib: 'static/jsthird',
        jquery: 'static/jsthird/jquery-1.10.2.min',
        btval: 'static/jsthird/bootstrap-validation.min',
        dtp: 'static/jsthird/bootstrap-datetimepicker.min',
        bselect: 'static/jsthird/bootstrap-select-v1.7.2/defaults-zh_CN.min',
        artDialog:'static/jsthird/artDialog/plugins/iframeTools',
        validationEngine:'veh1/static/js/jquery.validationEngine',
        vehxh: 'veh1/static/js'
    },
    shim: {
        'lib/bootstrap.min': ['jquery'],
        'validationEngine':['jquery','vehxh/jquery.validationEngine-zh_CN'],
        'btval': ['jquery'],
        'dtp': ['jquery','lib/bootstrap.min'],
        'bselect': ['jquery','lib/bootstrap.min','lib/bootstrap-select-v1.7.2/bootstrap-select'],
        'artDialog':['jquery','lib/artDialog/artDialog.source']
    }
});

require([ 'jquery','js/urls','lib/bootstrap.min','bselect','btval'], function ($,urls) {

    urls.scriptBase = _getScriptBase() + '/static';


    var pageScript = $('script[data-main*="base.js"]').data('page')
        || $('script[data-main*="base.js"]').attr('data-page');
    require(['js/funcRender'], function (funcRender) {
        funcRender.requestUsrInfo();
    })

});