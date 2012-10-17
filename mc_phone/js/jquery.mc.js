function getViewportHeight() {
    if (window.innerHeight != window.undefined) return window.innerHeight;
    if (document.compatMode == 'CSS1Compat') return document.documentElement.clientHeight;
    if (document.body) return document.body.clientHeight;

    return window.undefined;
}

function getViewportWidth() {
    var width = null;
    if (window.innerWidth != window.undefined) return window.innerWidth;
    if (document.compatMode == 'CSS1Compat') return document.documentElement.clientWidth;
    if (document.body) return document.body.clientWidth;
}

function LoadDialogs() {
    try {
        $('#mcMainWindow').dialog('destroy');
    }
    catch (e) { }

    $('#mcMainWindow').dialog({
        autoOpen: false,
        height: 200,
        width: 300,
        modal: true,
        draggable: true,
        stack: true,
        resizable: false,
        closeOnEscape: true
    });

	$(':button').button();
	$(':submit').button();
}

function CloseWindow() {
	try {
		$('#mcMainWindowFrame').attr('style', 'width:0px; height:0px;');
		$('#mcMainWindowFrame').attr('src', 'blank.htm');
	}
	catch (e) { }
	try{
		$('.ui-dialog-titlebar-close').trigger('click');
	}
	catch(e) {}
}

function closeiframe(url){
	parent.CloseWindow();
}

function OpenWindow(inTitle, inWidth, inHeight, inPage, inExtraData) {
	try{$('.ui-dialog-titlebar-close').show();} catch(e) {}
    var randid = Math.floor(Math.random() * 10001);
    if (inPage.indexOf('?') == -1) {
        inPage = inPage + '?ri=' + randid;
    }
    else {
        inPage = inPage + '&ri=' + randid;
    }

		try{
			$('.ui-widget-overlay').width(getViewportWidth());
			$('.ui-widget-overlay').height(getViewportHeight());
		}
		catch(e){}
		
        $('#mcMainWindowLoading').attr('style', 'width:300px; height:600px;');
        $('#mcMainWindowLoading').show();

        $('#mcMainWindowFrame').unbind('load');
        $('#mcMainWindowFrame').load(function () {
            $('#mcMainWindowLoading').attr('style', 'width:0px; height:0px;');
            $('#mcMainWindowLoading').hide();
        });
        $('#mcMainWindowFrame').attr('src', inPage + inExtraData + '&externalWindow=0');
        $('#mcMainWindowFrame').attr('style', 'width:300px; height:600px;');

        /*$('#mcMainWindow').dialog("option", "width", inWidth);
        if ($.browser.msie){
        	$('#mcMainWindow').dialog("option", "height", inHeight + 140);
		}
		else{
        	$('#mcMainWindow').dialog("option", "height", getViewportHeight()-40);
        	$('#mcMainWindow').dialog("option", "width", getViewportWidth()-40);
		}
		*/
		$('#mcMainWindow').dialog("option", "height", 600);
		$('#mcMainWindow').dialog("option", "width", 300);

        $("#mcMainWindow").dialog("option", "position", "center");
		$('#mcMainWindow').dialog("option", "title", 'MC Phone (' + inTitle + ')');

        $('#mcMainWindow').dialog("option", "closeOnEscape", true);
        try {
            $("#mcMainWindow").dialog("open");
        }
        catch (e) { }

}

function DoPost(iPage, iData, iReturnFunction, iErrorFunction) {
    var randid = Math.floor(Math.random() * 10001);
    try {
        $.ajax({
            type: "POST",
            url: iPage + "?ri=" + randid + iData,
            dataType: "html",
            success: function (a, b, c) {
                if (iReturnFunction) {
                    try {
                        iReturnFunction(a, b, c);
                    }
                    catch (e) { }
                }
            },
            error: function (a, b, c) {
                if (iErrorFunction) {
                    try {
                        iErrorFunction(a, b, c);
                    }
                    catch (e) { }
                }
            }
        });
    }
    catch (e) {
    }
}