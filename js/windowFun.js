window.onbeforeunload = function (e) {
    e = e || window.event;
    if (e) {
        e.returnValue = '确定离开吗？';
    }
    return '确定离开吗？';
};