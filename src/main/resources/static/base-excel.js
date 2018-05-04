$(function () {
    // nav切换事件
    $('.menu').on('click', 'li', function () {
        var _this = $(this);
        var _id = _this.attr('_id');

        _this.addClass('menu_active').siblings().removeClass('menu_active');
        $("div[id='" + _id + "']").show().siblings().hide();
    });
});




      