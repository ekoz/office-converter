This is a util to easy convert office document to *.html with customize the style.


here is some base config before to use,
 we must contain them in the default output-path 
 like OUTPUT/static/base-\*.css, OUTPUT/static/base-\*.js, OUTPUT/static/jquery-3.3.1.min.js

base-excel.css
``` css
.menu ul {list-style:none;margin: 0px;padding: 0px;width: 20000px;}
.menu ul li{float:left;cursor:pointer;}
.menu ul li {border: 1px #4e667d solid;display: block;line-height: 1.35em;padding: 4px 20px;}
.menu ul li:hover{background-color: #bfcbd6; color: #465c71;}
.menu-active {background-color: #465c71;color: #dde4ec;}
```

base-excel.js
``` javascript
$(function() {
	// nav切换事件
	$('.menu').on('click', 'li', function() {
		var _this = $(this);
		var _id = _this.attr('_id');	

		_this.addClass('menu_active').siblings().removeClass('menu_active');
		$("div[id='"+_id+"']").show().siblings().hide();
	});
});
```