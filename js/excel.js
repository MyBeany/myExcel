(function ($) {
        //单击选中的单元格
        let selectTd;
        //当前复制的单元格样式
        let selectTdStyle = {};
        $.fn.extend({
            Excel: function (options) {
                var op = $.extend({}, options);
                initFun();
                return this.each(function () {
                    var t = $(this);
                    t.addClass("excel-table");
                    if (op.data) {
                        initTable(t, {data: op.data, type: 0})
                    } else if (op.setting) {
                        op.setting.width = 0;
                        op.setting.type = 1;
                        initTable(t, op.setting)
                    } else {
                        initTable(t, {row: 11, col: 15, width: 0, type: 1})
                    }
                })
            }, getExcelHtml: function () {
                var table = $(this).find("table").first();
                if (table.length == 1) {
                    var clone = table.clone(false);
                    clone.find("tr:eq(0)").remove();
                    clone.find("tr").find("td:eq(0)").remove();
                    clone.find("td").removeClass("td-position-css").removeClass("td-chosen-css").removeClass("td-chosen-muli-css");
                    clone.find("td[class='']").removeAttr("class");
                    return clone.prop("outerHTML")
                } else {
                    return ""
                }
            }, setExcelHtml: function (html) {
                initFun();
                $(this).Excel({data: html})
            }
        });

        //初始化事件绑定
        function initFun() {
            $('body').on('input', '#selectTdValue', valueChange);
            $('body').on('change', '#fontfamily', setFontFamily);
            $('body').on('change', '#fontsize', setFontSize);
            $('body').on('click', '.btn-bold', setFontBold);
            $('body').on('click', '.btn-italic', setFontItalic);
            $('body').on('click', '.btn-underline', setUnderline);
            $('body').on('click', '.btn-strike', setFontStrike);
            $('body').on('click', '#bgColor', clickBgColor);
            $('body').on('click', '#fontColor', clickFontColor);
            $('body').on('change', '#bgColorSelect', setBgColor);
            $('body').on('change', '#fontColorSelect', setFontColor);
            $('body').on('click', '.btn-htTop', 'top', setValign);
            $('body').on('click', '.btn-htMiddle', 'middle', setValign);
            $('body').on('click', '.btn-htBottom', 'bottom', setValign);
            $('body').on('click', '.btn-htLeft', 'left', setAlign);
            $('body').on('click', '.btn-htCenter', 'center', setAlign);
            $('body').on('click', '.btn-htRight', 'right', setAlign);
            $('body').on('click', '.merge-btn', mergeBtn);
            $('body').on('click', '.split-btn', splitBtn);
            $('body').on('click', '.whiteSpace', whiteSpace);
            $('body').on('click', '.borderLeft', setBorderLeft);
            $('body').on('click', '.borderRight', setBorderRight);
            $('body').on('click', '.borderTop', setBorderTop);
            $('body').on('click', '.borderBottom', setBorderBottom);
            $('body').on('click', '.borderColor', clickBorderColor);
            $('body').on('click', '.borderStyle', showBorderStyleDiv);
            // $('body').on('change', '#borderColor', setBorderColor);
            $('body').on('change', '.cell-width', setCellWidth);
            $('body').on('change', '.cell-height', setCellHeight);
            $('body').on('click', '.borderAll', setBorderAll);
            $('body').on('click', '.borderSolid', 'solid', setBorderStyleOption);
            $('body').on('click', '.borderDashed', 'dashed', setBorderStyleOption);
            $('body').on('click', '.borderDouble', 'double', setBorderStyleOption);
            $('body').on('click', '.borderNone', 'none', setBorderStyleOption);
            tableScroll();
        }

        function tableScroll() {
            var tableCont = document.querySelector('.excel');

            function scrollHandle(e) {
                var scrollLeft = this.scrollLeft;
                var scrollTop = this.scrollTop;
                var d = $(this).data('slt');
                if (scrollLeft != (d == undefined ? 0 : d.sl)) {
                    $('.drug-ele-td-vertical').css('transform', 'translateX(' + scrollLeft + 'px)');
                    $('.row-height-panel-item').css('transform', 'translateX(' + scrollLeft + 'px)');
                }
                if (scrollTop != (d == undefined ? 0 : d.st)) {
                    $('.drug-ele-td-horizontal').css('transform', 'translateY(' + scrollTop + 'px)');
                    $('.col-width-panel-item').css('transform', 'translateY(' + scrollTop + 'px)');
                }
                $(this).data('slt', {sl: scrollLeft, st: scrollTop});
            }

            tableCont.addEventListener('scroll', scrollHandle)
        }

        //初始化table
        function initTable(t, setting) {
            t.empty();
            var table;
            //回显用的
            if (setting.type == 0) {
                t.html(setting.data);
                table = t.find("table").first();
                var fir = table.find("tr:eq(0)");
                var clone = fir.clone(false).height(25).insertBefore(fir);
                clone.find("td").css("display", "").removeAttr("rowspan").removeAttr("colspan").html("").removeClass("td-chosen-css");
                $("<td></td>").insertBefore(table.find("tr").find("td:eq(0)"))
            }
            //生成table
            else if (setting.type == 1) {
                table = $("<table></table>").appendTo(t);
                for (var i = 0; i < setting.row; i++) {
                    var tr = $("<tr></tr>").height(25).appendTo(table);
                    for (var j = 0; j < setting.col; j++) {
                        $("<td></td>").appendTo(tr)
                    }
                }
                if (setting.width && setting.width > 0) {
                    $('td').css('width', setting.width);
                }
            }
            drawDrugArea(table);
            eventBind(table, t);
            drugCell(table, t);
            //设置鼠标右键菜单的
            t.unbind("contextmenu");
            t.on('contextmenu', function () {
                return false
            })
        }

        function selectTable(table, t, e) {
            if (e.button == 2 && !$(e.target).hasClass("drug-ele-td")) {
                if (table.find(".td-chosen-css").length == 0) {
                    $(e.target).addClass("td-chosen-css")
                }
                showRightPanel(table, t, e)
            } else {
                closeRightPanel(t);
                var ele = $(e.target);
                if (!ele.hasClass("drug-ele-td")) {
                    clearPositionCss(table);
                    if (!ele.is("table") && table.data("beg-td-ele") && table.data("beg-td-ele").is(ele)) {
                        ele.addClass("td-chosen-css");
                        var posi = getTdPosition(ele);
                        table.find("tr").find("td:eq(" + posi.col + ")").addClass("td-position-css");
                        table.find("tr:eq(" + posi.row + ")").find("td").addClass("td-position-css")
                    } else {
                        getChosenList(table, getTdPosition(table.data("beg-td-ele")), getTdPosition(ele))
                    }
                    drawChosenArea(table, t)
                }
            }
        }

        function mouseMove(table, t) {
            table.mouseover(function (e) {
                table.find("td").removeClass("td-chosen-muli-css");
                table.find("td").removeClass("td-chosen-css");
                selectTable(table, t, e)
            });
        }

        //赋值文本框   改变之后赋值到td内
        function setSelectTdValue(ele) {
            let val = $(ele).html();
            let $input = $('#selectTdValue');
            $input.val(val);
            setTimeout(function () {
                $input.select();
            }, 10);
        }

        //赋值文本框  change事件
        function valueChange() {
            let val = $('#selectTdValue').val();
            if (selectTd) {
                selectTd.html(val)
            }
        }

        //设置点击td时   赋值文本框的事件
        function settingInput(e) {
            setTimeout(function () {
                $('#selectTdValue').focus();
            }, 100);
            let pos = getTdPosition($(e));
            $('#coordinate').html('<span>' + getChar(pos.col - 1) + pos.row + "</span>")
            setSelectTdValue(e);
        }

        //判断元素是否有某属性
        function hasAttr(e, attr) {
            let Attr = e.attr(attr);
            if (typeof Attr !== typeof undefined && Attr !== false) {
                return true;
            } else {
                return false;
            }
        }

        //选择一行或一列时，设置选择框样式
        function selectWhole(table, addWidth, addHeight) {
            var coll = table.find(".td-chosen-css");
            var first = coll.first();
            var posi = getTdPosition(first);
            var width = 0, height = 0;
            coll.each(function () {
                var p = getTdPosition($(this));
                if (p.row == posi.row) {
                    width += this.offsetWidth
                }
                if (p.col == posi.col) {
                    height += this.offsetHeight
                }
            });
            if (addWidth === 0) {
                addWidth = width;
            }
            if (addHeight === 0) {
                addHeight = height;
            }
            setSelectBorder(table, addWidth, addHeight, first[0].offsetTop, first[0].offsetLeft);
        }

        //点击td 设置样式栏中各项的值
        function triggerStyle(e, table) {
            $('.sub-bottom').children().removeClass('buttonBgColor');
            let ele = $(e);
            let fontFamily = ele.css('font-family');
            $('#fontfamily').val(fontFamily);

            let fontSize = ele.css('font-size');
            $('#fontsize').val(fontSize);

            let fontWeight = ele.css('font-weight');
            if (fontWeight !== '400') {
                $('.btn-bold').addClass('buttonBgColor');
            }

            let fontItalic = ele.css('font-style');
            if (fontItalic !== 'normal') {
                $('.btn-italic').addClass('buttonBgColor');
            }

            let underline = ele.css('text-decoration-line');
            if (underline === 'underline') {
                $('.btn-underline').addClass('buttonBgColor');
            }

            let fontStrike = ele.css('text-decoration-line');
            if (fontStrike === 'line-through') {
                $('.btn-strike').addClass('buttonBgColor');
            }

            let valign = ele.css('vertical-align');
            $('.btn-av').removeClass('buttonBgColor');
            if (valign === 'top') {
                $('.btn-htTop').addClass('buttonBgColor');
            } else if (valign === 'middle') {
                $('.btn-htMiddle').addClass('buttonBgColor');
            } else if (valign === 'bottom') {
                $('.btn-htBottom').addClass('buttonBgColor');
            }

            let textAlign = ele.css('text-align');
            $('.btn-ah').removeClass('buttonBgColor');
            if (textAlign === 'left') {
                $('.btn-htLeft').addClass('buttonBgColor');
            } else if (textAlign === 'center') {
                $('.btn-htCenter').addClass('buttonBgColor');
            } else if (textAlign === 'right') {
                $('.btn-htRight').addClass('buttonBgColor');
            }

            let whiteSpace = ele.css('whiteSpace');
            if (whiteSpace !== 'nowrap') {
                $('.whiteSpace').addClass('buttonBgColor');
            } else {
                $('.whiteSpace').removeClass('buttonBgColor');
            }
        }

        function selectMoreCell(e, table, t) {
            let addWidth = 0;
            let addHeight = 0;
            table.find("td").removeClass("td-chosen-css").removeClass('td-chosen-muli-css');
            if ($(e.target).index() === 0 && $(e.target).html() === '') {
                return
            }
            if ($(e.target).hasClass('drug-ele-td-vertical')) {
                selectTd = $(e.target).next();
                $(e.target).nextAll().each(function (index, ele) {
                    if ((!hasAttr($(this), 'colspan')) || index == 0) {
                        $(this).addClass('td-chosen-css').addClass('td-chosen-muli-css');
                    }
                })
                $(e.target).next().addClass('selectTd');
                addWidth = $(table)[0].offsetWidth - 63;
            } else {
                let index = $(e.target).index();
                table.find("tr").each(function (i, ele) {
                    $td = $(this).children().eq(index);
                    if ((!hasAttr($td, 'rowspan')) || i === 1) {
                        $td.addClass('td-chosen-css').addClass('td-chosen-muli-css');
                    }
                    if (i === 1) {
                        $td.addClass('selectTd');
                        selectTd = $td;
                    }
                });

                $(e.target).removeClass("td-chosen-css").removeClass('td-chosen-muli-css');
                addHeight = $(table)[0].offsetHeight - 25;
            }
            let pos = getTdPosition(selectTd);
            $('#coordinate').html('<span>' + getChar(pos.col - 1) + pos.row + "</span>")
            selectWhole(table, addWidth, addHeight);
        }

        function tdMousedown(e, table) {
            selectTd = $(e);
            table.find("td").removeClass("td-chosen-css");
            table.removeData("beg-td-ele");
            table.data("beg-td-ele", $(e));
            $(e).addClass('selectTd');
        }

        function clickTd(e, table, t) {
            table.find("td").removeClass('selectTd');
            tdMousedown(e, table);
            closeRightPanel(t);
            clearPositionCss(table);
            e.addClass("td-chosen-css");
            var posi = getTdPosition(e);
            table.find("tr").find("td:eq(" + posi.col + ")").addClass("td-position-css");
            table.find("tr:eq(" + posi.row + ")").find("td").addClass("td-position-css")
            drawChosenArea(table, t);
            settingInput(e);
            triggerStyle(e, table);
        }

        function selectTdScroll() {
            let $node = $('.chosen-area-p-drug');
            let windowH = $('.excel').height(),
                windowW = $('.excel').width(),
                $nodeOffsetH = parseInt($node.css('margin-top')),
                $nodeOffsetW = parseInt($node.css('margin-left')),
                $nodeInitLeft = selectTd.innerWidth() + selectTd.prevAll().last().innerWidth(),
                $nodeInitTop = selectTd.innerHeight() + selectTd.parent().prevAll().last().innerHeight() - 2;
            //备注  19为滚动条宽度
            if (($nodeOffsetW + 19 >= windowW) && ($nodeOffsetW - $('.excel').scrollLeft() + 19 >= windowW)) {
                $('.excel').scrollLeft(selectTd.width() + $('.excel').scrollLeft() + 4);
            } else if ($nodeInitLeft + $('.excel').scrollLeft() > $nodeOffsetW) {
                $('.excel').scrollLeft($('.excel').scrollLeft() - selectTd.width() - 4);
            } else if (($nodeOffsetH + 19 >= windowH) && ($nodeOffsetH - $('.excel').scrollTop() + 19 >= windowH)) {
                $('.excel').scrollTop(selectTd.height() + $('.excel').scrollTop() + 4);
            } else if ($nodeInitTop + $('.excel').scrollTop() > $nodeOffsetH) {
                $('.excel').scrollTop($('.excel').scrollTop() - selectTd.height() - 4);
            }
        }

        function tableKeyDown(e, table, t) {
            if (selectTd == undefined || $('.rightmouse-panel-div').length != 0) {
                return;
            }
            let eCode = e.keyCode ? e.keyCode : e.which ? e.which : e.charCode;
            if (e.ctrlKey && eCode === 90) {
                chexiaoFunc(t)
            } else if (eCode === 13 || eCode === 39) {
                let $nextTd = selectTd.nextAll(':visible').first();
                if ($nextTd.length > 0) {
                    clickTd($nextTd, table, t);
                }
            } else if (eCode === 37) {
                let $prevTd = selectTd.prevAll(':visible').first();
                if ($prevTd.prev().length > 0) {
                    clickTd($prevTd, table, t);
                }
            } else if (eCode === 40) {
                let index = selectTd.index();
                let $nextTd = {};
                selectTd.parent().nextAll().each(function () {
                    $nextTd = $(this).children().eq(index);
                    if (!$nextTd.is(":hidden")) {
                        return false;
                    }
                });
                if ($nextTd.length > 0) {
                    clickTd($nextTd, table, t);
                }
            } else if (eCode === 38) {
                let index = selectTd.index();
                let $prevTd = {};
                selectTd.parent().prevAll().each(function () {
                    $prevTd = $(this).children().eq(index);
                    if (!$prevTd.is(":hidden")) {
                        return false;
                    }
                });

                let $prevTdPrev = $prevTd.parent().prev();
                if ($prevTdPrev.length > 0) {
                    clickTd($prevTd, table, t);
                }
            }
            selectTdScroll();
        }

        //为table绑定事件
        function eventBind(table, t) {
            table.mousedown(function (e) {
                if (e.button == 0) {
                    table.find("td").removeClass('selectTd');
                    if (!$(e.target).hasClass("drug-ele-td")) {
                        tdMousedown(e.target, table);
                        settingInput(e.target);
                        triggerStyle(e.target, table);
                    } else {
                        selectMoreCell(e, table, t)
                    }
                    mouseMove(table, t);
                }
            }).mouseup(function (e) {
                table.unbind('mouseover');
                selectTable(table, t, e);
            });
            $(document).unbind("keydown");
            $(document).keydown(function (e) {
                tableKeyDown(e, table, t);
            });
        }

        function getChosenList(table, begPosi, endPosi) {
            if (begPosi != undefined && endPosi != undefined) {
                for (var i = (begPosi.row > endPosi.row ? endPosi.row : begPosi.row); i <= (begPosi.row > endPosi.row ? begPosi.row : endPosi.row); i++) {
                    var tr = table.find("tr:eq(" + i + ")");
                    for (var j = (begPosi.col > endPosi.col ? endPosi.col : begPosi.col); j <= (begPosi.col > endPosi.col ? begPosi.col : endPosi.col); j++) {
                        var td = tr.find("td:eq(" + j + ")");
                        td.addClass("td-chosen-css");
                    }
                }
                var coll = table.find(".td-chosen-css");
                var firstPosi = getTdPosition($(coll.get(0)));
                var beg_row = firstPosi.row;
                var beg_col = firstPosi.col;
                table.find("td").removeData("add-chosen-state").removeData("get-father-state");
                while (true) {
                    var end_row = 0;
                    var end_col = 0;
                    var con = false;
                    coll.each(function () {
                        var p = getTdPosition($(this));
                        var r = p.row + ($(this).attr("rowspan") == undefined ? 0 : (Number($(this).attr("rowspan")) - 1));
                        var c = p.col + ($(this).attr("colspan") == undefined ? 0 : (Number($(this).attr("colspan")) - 1));
                        end_row = end_row < r ? r : end_row;
                        end_col = end_col < c ? c : end_col;
                        beg_row = beg_row > p.row ? p.row : beg_row;
                        beg_col = beg_col > p.col ? p.col : beg_col
                    });
                    for (var i = beg_row; i <= end_row; i++) {
                        var tr = table.find("tr:eq(" + i + ")");
                        for (var j = beg_col; j <= end_col; j++) {
                            var dt = tr.find("td:eq(" + j + ")");
                            if (dt.is(":hidden") && dt.data("get-father-state") == undefined) {
                                var p = getFatherCell(dt);
                                dt.data("get-father-state", 0);
                                if (p != null && p.length == 1) {
                                    p.data("add-chosen-state", 0);
                                    if (p != null && coll.index(p) == -1) {
                                        p.addClass("td-chosen-css");
                                        coll = table.find(".td-chosen-css");
                                        con = true
                                    }
                                }
                            } else {
                                if (!dt.hasClass("td-chosen-css")) {
                                    dt.addClass("td-chosen-css");
                                    coll = table.find(".td-chosen-css");
                                    con = true
                                }
                            }
                        }
                    }
                    if (!con) {
                        break
                    }
                }
                return coll
            }
        }

        function getTdPosition(td) {
            if (td != undefined && td.length == 1) {
                var table = td.closest("table");
                var pos = {};
                var tr = td.closest("tr");
                pos.row = table.find("tr").index(tr);
                pos.col = tr.find("td").index(td);
                return pos
            }
        }

        function mergeCell(table) {
            if (table.length == 1) {
                var coll = table.find(".td-chosen-css");
                if (coll.length > 1) {
                    var fir = $(coll.get(0));
                    var posi = getTdPosition(fir);
                    var r = 0, c = 0;
                    if (fir.attr("rowspan") != undefined && fir.attr("colspan") != undefined) {
                        r = Number(fir.attr("rowspan")) - 1;
                        c = Number(fir.attr("colspan")) - 1
                    }
                    coll.each(function () {
                        var p = getTdPosition($(this));
                        r = (p.row - posi.row) > r ? p.row - posi.row : r;
                        c = (p.col - posi.col) > c ? (p.col - posi.col) : c;
                        if (!$(this).is(fir)) {
                            $(this).removeClass("td-chosen-css").css("display", "none");
                            if ($(this).attr("rowspan") != undefined && $(this).attr("colspan") != undefined) {
                                r = (p.row + (Number($(this).attr("rowspan")) - 1) - posi.row) > r ? (p.row + (Number($(this).attr("rowspan")) - 1) - posi.row) : r;
                                c = (p.col + (Number($(this).attr("colspan")) - 1) - posi.col) > c ? (p.col + (Number($(this).attr("colspan")) - 1) - posi.col) : c
                            }
                        }
                    });
                    $(coll.get(0)).attr("rowspan", r + 1).attr("colspan", c + 1).css("display", "")
                } else if (coll.length == 1) {
                    var fir = $(coll.get(0));
                    if (fir.attr("rowspan") != undefined && fir.attr("colspan") != undefined) {
                        var posi = getTdPosition(fir);
                        for (var i = posi.row; i <= (posi.row + (Number($(fir).attr("rowspan")) - 1)); i++) {
                            var tr = table.find("tr:eq(" + i + ")");
                            for (var j = posi.col; j <= (posi.col + (Number($(fir).attr("colspan")) - 1)); j++) {
                                var td = tr.find("td:eq(" + j + ")").css("display", "").addClass("td-chosen-css");
                                if (!td.is(fir)) {
                                    td.removeAttr("rowspan").removeAttr("colspan")
                                }
                            }
                        }
                        fir.removeAttr("rowspan").removeAttr("colspan")
                    }
                }
            }
        }

        function getFatherCell(noneTd) {
            var table = noneTd.closest("table");
            var fatherCell = [];
            table.find("td[rowspan][colspan]").each(function () {
                var posi = getTdPosition($(this));
                var cell = $(this);
                var con = false;
                for (var i = posi.row; i <= (posi.row + (Number($(this).attr("rowspan")) - 1)); i++) {
                    var tr = table.find("tr:eq(" + i + ")");
                    for (var j = posi.col; j <= (posi.col + (Number($(this).attr("colspan")) - 1)); j++) {
                        var dt = tr.find("td:eq(" + j + ")");
                        if (noneTd.is(dt)) {
                            fatherCell[fatherCell.length] = cell;
                            con = true;
                            break
                        }
                    }
                    if (con) {
                        break
                    }
                }
            });
            if (fatherCell.length == 1) {
                return fatherCell[0]
            } else {
                return null
            }
        }

        function panelItemMouseleave(ele, table, t) {
            ele.mouseleave(function (e) {
                clearDurgEle(table, t)
            });
        }

        function drugCell(table, t) {
            var colTransform = $('.col-width-panel-item').eq(1).css('transform');
            t.find(".col-width-panel,.row-height-panel").remove();
            t.find(".chosen-area-p").remove();
            var colWidthPanel = $("<div class='col-width-panel'></div>");
            var rowHeightPanel = $("<div class='row-height-panel'></div>");
            var left = 0, top = 0;
            var firstTr = table.find("tr").first();
            colWidthPanel.insertBefore(table);
            rowHeightPanel.insertBefore(table);
            table.find("tr").first().find("td").each(function () {
                left = this.offsetLeft;
                let colWidthPanelItem = $("<div class='col-width-panel-item'></div>");
                colWidthPanelItem.attr("draggable", true).mousedown(function (e) {
                    e.preventDefault && e.preventDefault();
                    var ele = $(e.target);
                    if (ele.data("left") == undefined) {
                        recordData(t);
                        ele.data("left", ele.css("left"));
                        ele.data("e-left", e.clientX);
                        t.data("drug-ele", ele);
                    }
                }).mouseup(function () {
                    clearDurgEle(table, t)
                }).css("transform",colTransform).css("left", left +this.offsetWidth - 4).css("height", firstTr[0].offsetHeight).appendTo(colWidthPanel)
            });
            table.find("tr").each(function () {
                top = this.offsetTop;
                $(this).height($(this).height());
                let rowHeightPanelItem = $("<div class='row-height-panel-item'></div>");
                rowHeightPanelItem.attr("draggable", true).mousedown(function (e) {
                    e.preventDefault && e.preventDefault();
                    var ele = $(e.target);
                    if (ele.data("top") == undefined) {
                        recordData(t);
                        ele.data("top", ele.css("top"));
                        ele.data("e-top", e.clientY);
                        t.data("drug-ele", ele);
                    }
                }).mouseup(function () {
                    clearDurgEle(table, t);
                }).css("top", top + this.offsetHeight - 4).css("width", firstTr.find("td")[0].offsetWidth).appendTo(rowHeightPanel)
            });
            colWidthPanel.find(".col-width-panel-item:first,.col-width-panel-item:last").css("display", "none");
            rowHeightPanel.find(".row-height-panel-item:first,.row-height-panel-item:last").css("display", "none");
            t.unbind("mouseup").unbind("mousemove").unbind("mousedown").unbind("mouseleave");
            t.mousedown(function (e) {
                var ele = t.data("drug-ele");
                if (ele !== undefined) {
                    if (ele.hasClass("col-width-panel-item")) {
                        panelItemMouseleave(colWidthPanel, table, t);
                    }
                    if (ele.hasClass("row-height-panel-item")) {
                        panelItemMouseleave(rowHeightPanel, table, t);
                    }
                }
            }).mouseup(function (e) {
                clearDurgEle(table, t);
            }).mousemove(function (e) {
                if (t.data("drug-ele") != undefined) {
                    closeRightPanel(t);
                    var ele = t.data("drug-ele");
                    if (ele.hasClass("col-width-panel-item") && ele.data("left") != undefined) {
                        var left = parseInt(ele.data("left")) + (e.clientX - ele.data("e-left"));
                        var ind = colWidthPanel.find(".col-width-panel-item").index(ele);
                        var upLeft = 0;
                        if (ind > 0) {
                            upLeft = parseInt(ele.prev(".col-width-panel-item").css("left")) + 4
                        }
                        var now = table.find("tr").find("td:eq(" + ind + ")");
                        now.width(left - upLeft);
                        //将负责调整宽度的元素加宽，以免出现鼠标滑动过快而导致调整失败
                        ele.css("left", left-250).css("width",500);
                    }
                    if (ele.hasClass("row-height-panel-item") && ele.data("top") != undefined) {
                        var top = parseInt(ele.data("top")) + (e.clientY - ele.data("e-top"));
                        var ind = rowHeightPanel.find(".row-height-panel-item").index(ele);
                        var upTop = 0;
                        if (ind > 0) {
                            upTop = parseInt(ele.prev(".row-height-panel-item").css("top")) + 4
                        }
                        if (top - upTop > 5) {
                            var now = table.find("tr:eq(" + ind + ")");
                            now.height(top - upTop);
                            ele.css("top", top-250).css("height",500);
                        }
                    }
                }
            })
        }

        function clearDurgEle(table, t) {
            if (t.data("drug-ele") != undefined) {
                t.data("drug-ele").removeData("left").removeData("e-left").removeData("top").removeData("e-top");
                t.removeData("drug-ele");
                drugCell(table, t)
            }
        }

        function addRowCol(table, type, t) {
            var chosenColl = table.find(".td-chosen-css");
            if (chosenColl.length == 1) {
                var chosen = chosenColl.first();
                var tr = chosen.closest("tr");
                var col = table.find("tr").find("td:eq(" + (tr.find("td").index(chosen)) + ")");
                if (type == 0) {
                    addRowColSpan(tr, type).insertBefore(tr)
                } else if (type == 1) {
                    addRowColSpan(tr, type).insertAfter(tr)
                } else if (type == 4) {
                    addRowColSpan(tr, type);
                    tr.remove()
                } else if (type == 2) {
                    addRowColSpan(col, type)
                } else if (type == 3) {
                    addRowColSpan(col, type)
                } else if (type == 5) {
                    addRowColSpan(col, type);
                    col.remove()
                }
            }
            table.find("td[rowspan=1][colspan=1]").removeAttr("rowspan").removeAttr("colspan");
            t.find(".chosen-area-p").remove();
            clearDurgEle(table, t);
            drawDrugArea(table)
        }

        function addRowColSpan(list, ty) {
            var coll = [];
            if (ty == 0 || ty == 1 || ty == 4) {
                var tr = list;
                tr.find("td").each(function () {
                    if ($(this).is(":hidden")) {
                        var p = getFatherCell($(this));
                        var con = true;
                        for (var i = 0; i < coll.length; i++) {
                            if (coll[i].is(p)) {
                                con = false;
                                break
                            }
                        }
                        if (con && p != null) {
                            coll[coll.length] = p;
                            p.attr("rowspan", spanNum(p.attr("rowspan"), ty == 4 ? -1 : 1))
                        }
                    } else {
                        if ($(this).attr("rowspan") && $(this).attr("colspan")) {
                            coll[coll.length] = $(this);
                            if (ty == 4) {
                                var nextTr = tr.next("tr");
                                if (nextTr.length == 1 && Number($(this).attr("rowspan")) > 1) {
                                    var ind = tr.find("td").index($(this));
                                    nextTr.find("td:eq(" + ind + ")").attr("rowspan", spanNum($(this).attr("rowspan"), -1)).attr("colspan", $(this).attr("colspan")).css("display", "")
                                }
                            } else {
                                $(this).attr("rowspan", Number($(this).attr("rowspan")) + 1)
                            }
                        }
                    }
                });
                var clone = tr.clone(true);
                if (ty == 0) {
                    tr.find("td[rowspan][colspan]").each(function () {
                        $(this).removeAttr("rowspan").removeAttr("colspan").css("display", "none")
                    })
                }
                if (ty == 1) {
                    clone.find("td[rowspan][colspan]").each(function () {
                        $(this).removeAttr("rowspan").removeAttr("colspan").css("display", "none")
                    })
                }
                clone.height(25);
                clone.find("td").removeClass("td-chosen-css").html("");
                return clone
            } else {
                var cloneLs = [];
                list.each(function () {
                    if ($(this).is(":hidden")) {
                        var p = getFatherCell($(this));
                        var con = true;
                        for (var i = 0; i < coll.length; i++) {
                            if (coll[i].is(p)) {
                                con = false;
                                break
                            }
                        }
                        if (con && p != null) {
                            coll[coll.length] = p;
                            p.attr("colspan", spanNum(p.attr("colspan"), ty == 5 ? -1 : 1))
                        }
                    } else {
                        if ($(this).attr("rowspan") && $(this).attr("colspan")) {
                            coll[coll.length] = $(this);
                            if (ty == 5) {
                                var nextTd = $(this).next("td");
                                if (nextTd.length == 1 && Number($(this).attr("colspan")) > 1) {
                                    nextTd.width($(this).width()).attr("rowspan", $(this).attr("rowspan")).attr("colspan", spanNum($(this).attr("colspan"), -1)).css("display", "")
                                }
                            } else {
                                $(this).attr("colspan", Number($(this).attr("colspan")) + 1)
                            }
                        }
                    }
                    var clone = $(this).clone(true);
                    clone.width($(this).width());
                    clone.removeClass("td-chosen-css").html("");
                    cloneLs[cloneLs.length] = clone
                });
                for (var i = 0; i < cloneLs.length; i++) {
                    if (ty == 2) {
                        cloneLs[i].insertBefore($(list.get(i)));
                        var t = $(list.get(i));
                        if (t.attr("rowspan") && t.attr("colspan")) {
                            t.removeAttr("rowspan").removeAttr("colspan").css("display", "none")
                        }
                    }
                    if (ty == 3) {
                        cloneLs[i].insertAfter($(list.get(i)));
                        var t = cloneLs[i];
                        if (t.attr("rowspan") && t.attr("colspan")) {
                            t.removeAttr("rowspan").removeAttr("colspan").css("display", "none")
                        }
                    }
                }
            }
        }

        function spanNum(spanNum, n) {
            var num = Number(spanNum) + n;
            num = num < 1 ? 1 : num;
            return num
        }

        function drawChosenArea(table, t) {
            var coll = table.find(".td-chosen-css");
            table.find("td").removeClass("td-chosen-muli-css");
            if (coll.length > 0) {
                var first = coll.first();
                var posi = getTdPosition(first);
                var width = 0, height = 0;
                var p = table.parent();
                coll.each(function () {
                    var p = getTdPosition($(this));
                    if (p.row == posi.row) {
                        width += this.offsetWidth
                    }
                    if (p.col == posi.col) {
                        height += this.offsetHeight
                    }
                });
                if (coll.length > 1) {
                    coll.addClass("td-chosen-muli-css");
                    //复制值
                    if (p.find(".chosen-area-p-drug").length === 1) {
                        var con = false;
                        var fir = coll.first();
                        if (p.find(".chosen-area-p-drug").data("text") !== undefined) {
                            recordData(t);
                            coll.html(p.find(".chosen-area-p-drug").data("text"));
                            p.find(".chosen-area-p-drug").removeData("text");
                            con = true
                        }
                        if (p.find(".chosen-area-p-drug").data("textNum") !== undefined) {
                            recordData(t);
                            var n = p.find(".chosen-area-p-drug").data("textNum");
                            var v = Number($.trim(fir.text()));
                            coll.each(function () {
                                $(this).html(v);
                                v += n
                            });
                            p.find(".chosen-area-p-drug").removeData("textNum");
                            con = true
                        }
                        if (con) {
                            if (fir.css("vertical-align") && fir.css("vertical-align") !== "") {
                                coll.css("vertical-align", fir.css("vertical-align"))
                            }
                            if (fir.css("text-align") && fir.css("text-align") !== "") {
                                coll.css("text-align", fir.css("text-align"))
                            }
                        }
                    }
                }
                setSelectBorder(table, width, height, first[0].offsetTop, first[0].offsetLeft);
            }
        }

        function setSelectBorder(table, width, height, top, left) {
            var p = table.parent();
            var coll = table.find(".td-chosen-css");
            p.find(".chosen-area-p").remove();
            $("<div class='chosen-area-p'></div>").width(1).height(height + 1).css("margin-top", top - 1).css("margin-left", left - 1).insertBefore(table);
            $("<div class='chosen-area-p'></div>").width(width + 1).height(1).css("margin-top", top - 1).css("margin-left", left - 1).insertBefore(table);
            $("<div class='chosen-area-p'></div>").width(1).height(height).css("margin-top", top - 1).css("margin-left", left + width - 1).insertBefore(table);
            $("<div class='chosen-area-p'></div>").width(width).height(1).css("margin-top", top + height - 1).css("margin-left", left - 1).insertBefore(table);
            $("<div class='chosen-area-p chosen-area-p-drug'></div>").mousedown(function () {
                //控制只有当选择一个的时候才能复制
                // if (coll.length === 1) {
                $(this).data("text", $.trim(coll.first().text()))
                // }
                if (coll.length === 2) {
                    var reg = /^\d{1,9}$/;
                    if (reg.test($.trim(coll.first().text())) && reg.test($.trim($(coll.get(1)).text()))) {
                        $(this).data("textNum", Number($.trim($(coll.get(1)).text())) - Number($.trim(coll.first().text())))
                    }
                }
            }).width(3).height(3).css("padding", "2px").css("margin-top", top + height - 4).css("margin-left", left + width - 4).insertBefore(table)
        }

        function getChar(ind) {
            var char = String.fromCharCode(65 + ind);
            if (ind >= 26) {
                char = String.fromCharCode(65 + (parseInt(ind / 26) - 1)) + String.fromCharCode(65 + ind % 26)
            }
            return char;
        }

        //设置表头
        function drawDrugArea(table) {
            table.parent().append('<div class="tableLeftTop" id="coordinate"></div>')
            var ind = 0;
            table.find("tr").first().addClass('thead')
            table.find("tr").first().find("td:gt(0)").unbind("click");
            table.find("tr").find("td:eq(0)").unbind("click");
            table.find("tr").first().find("td:gt(0)").each(function () {
                var char = getChar(ind);
                $(this).addClass("drug-ele-td drug-ele-td-horizontal ").css("text-align", "center").html(char);
                ind++
            });
            ind = 0;
            table.find("tr").find("td:eq(0)").each(function () {
                $(this).width(60).addClass("drug-ele-td drug-ele-td-vertical").css("text-align", "center").html(ind === 0 ? "" : ind);
                ind++
            });
        }

        function clearPositionCss(table) {
            table.find("td").removeClass("td-position-css")
        }

        function showRightPanel(table, t, e) {
            var coll = table.find(".td-chosen-css");
            closeRightPanel(t);
            var rightMousePanel = $("<div class='rightmouse-panel-div'></div>").css("left", e.clientX).css("top", e.clientY).insertBefore(table);
            var leftPanel = $("<div class='panel-div-left'></div>").width(200).appendTo(rightMousePanel);
            var rightPanel = $("<div class='panel-div-right'></div>").width(130).appendTo(rightMousePanel);
            $("<div class='wb duiqifangsi'></div>").html("<span class='excel-rightmomuse-icon-css'><i class='fa fa-th'></i></span><span class='excel-rightmomuse-text-css'>对齐方式</span><span class='excel-rightmomuse-icon-css excel-rightmomuse-icon-next-css'><i class='fa fa-caret-right'></i></span>").appendTo(leftPanel);
            $("<div class='wb hebingdanyuange'></div>").html("<span class='excel-rightmomuse-icon-css'><i class='fa fa-columns'></i></span><span class='excel-rightmomuse-text-css'>合并单元格</span>").appendTo(leftPanel);
            $("<div class='wb fuzhidanyuange'></div>").html("<span class='excel-rightmomuse-icon-css'><i class='fa fa-columns'></i></span><span class='excel-rightmomuse-text-css'>复制单元格样式</span>").appendTo(leftPanel);
            $("<div class='wb zhantiedanyuange'></div>").html("<span class='excel-rightmomuse-icon-css'><i class='fa fa-columns'></i></span><span class='excel-rightmomuse-text-css'>粘贴单元格样式</span>").appendTo(leftPanel);
            $("<div class='hr'></div>").appendTo(leftPanel);
            $("<div class='wb shangchayihang'></div>").html("<span class='excel-rightmomuse-icon-css'><i class='fa fa-angle-up'></i></span><span class='excel-rightmomuse-text-css'>上方插入一行</span>").appendTo(leftPanel);
            $("<div class='wb xiachayihang'></div>").html("<span class='excel-rightmomuse-icon-css'><i class='fa fa-angle-down'></i></span><span class='excel-rightmomuse-text-css'>下方插入一行</span>").appendTo(leftPanel);
            $("<div class='wb zuochayilie'></div>").html("<span class='excel-rightmomuse-icon-css'><i class='fa fa-angle-left'></i></span><span class='excel-rightmomuse-text-css'>左边插入一列</span>").appendTo(leftPanel);
            $("<div class='wb youchayilie'></div>").html("<span class='excel-rightmomuse-icon-css'><i class='fa fa-angle-right'></i></span><span class='excel-rightmomuse-text-css'>右边插入一列</span>").appendTo(leftPanel);
            $("<div class='hr'></div>").appendTo(leftPanel);
            $("<div class='wb shanchuhang'></div>").html("<span class='excel-rightmomuse-icon-css'><i class='fa fa-minus-square-o'></i></span><span class='excel-rightmomuse-text-css'>删除行</span>").appendTo(leftPanel);
            $("<div class='wb shanchulie'></div>").html("<span class='excel-rightmomuse-icon-css'><i class='fa fa-minus-square'></i></span><span class='excel-rightmomuse-text-css'>删除列</span>").appendTo(leftPanel);
            $("<div class='wb chexiao' title='只能结构改变撤销'></div>").html("<span class='excel-rightmomuse-icon-css'><i class='fa fa-reply-all'></i></span><span class='excel-rightmomuse-text-css'>撤销</span>").appendTo(leftPanel);
            $("<div class='wb juzhong'></div>").html("<span class='excel-rightmomuse-icon-css'><i class='fa fa-align-justify'></i></span><span class='excel-rightmomuse-text-css'>居中</span>").appendTo(rightPanel);
            $("<div class='wb zuoduiqi'></div>").html("<span class='excel-rightmomuse-icon-css'><i class='fa fa-align-left'></i></span><span class='excel-rightmomuse-text-css'>左对齐</span>").appendTo(rightPanel);
            $("<div class='wb youduiqi'></div>").html("<span class='excel-rightmomuse-icon-css'><i class='fa fa-align-right'></i></span><span class='excel-rightmomuse-text-css'>右对齐</span>").appendTo(rightPanel);
            $("<div class='hr'></div>").appendTo(rightPanel);
            $("<div class='wb chuizhijuzhong'></div>").html("<span class='excel-rightmomuse-icon-css'><i class='fa fa-navicon'></i></span><span class='excel-rightmomuse-text-css'>垂直居中</span>").appendTo(rightPanel);
            $("<div class='wb dingduanduiqi'></div>").html("<span class='excel-rightmomuse-icon-css'><i class='fa fa-angle-double-up'></i></span><span class='excel-rightmomuse-text-css'>顶端对齐</span>").appendTo(rightPanel);
            $("<div class='wb dibuduiqi'></div>").html("<span class='excel-rightmomuse-icon-css'><i class='fa fa-angle-double-down'></i></span><span class='excel-rightmomuse-text-css'>底部对齐</span>").appendTo(rightPanel);
            var setting = $("<div class='wb setting'></div>").html("<span class='setting-item'><span class='setting-text'>宽</span><span class='setting-input'><input type='text' name='width' title='单元格宽度'/></span></span><span class='setting-item'><span class='setting-text'>行</span><span class='setting-input'><input type='text' name='row'/></span></span><span class='setting-item'><span class='setting-text'>列</span><span class='setting-input'><input type='text' name='col'/></span></span>").appendTo(leftPanel);
            leftPanel.mousemove(function (e) {
                var ele = $(e.target);
                if (ele.hasClass("duiqifangsi") || ele.closest(".duiqifangsi").length == 1) {
                    rightPanel.css("display", "")
                } else {
                    rightPanel.css("display", "none")
                }
            });
            setting.find("input").keyup(function (e) {
                if (e.keyCode == 13) {
                    var width = $.trim(setting.find("input[name='width']").val());
                    var row = $.trim(setting.find("input[name='row']").val());
                    var col = $.trim(setting.find("input[name='col']").val());
                    var reg = /^\d{1,4}$/;
                    if (reg.test(row) && reg.test(col)) {
                        width = reg.test(width) ? width : 0;
                        initTable(t, {row: Number(row) + 1, col: Number(col) + 1, width: width, type: 1})
                    }
                }
            });
            rightMousePanel.find(".wb").click(function () {
                var obj = $(this);
                if (!obj.hasClass("duiqifangsi") && !obj.hasClass("setting")) {
                    if (!obj.hasClass("chexiao")) {
                        recordData(t)
                    }
                    if (obj.hasClass("hebingdanyuange")) {
                        mergeCell(table)
                    }
                    if (obj.hasClass("fuzhidanyuange")) {
                        copydanyuange(table)
                    }
                    if (obj.hasClass("zhantiedanyuange")) {
                        pastedanyuange(table)
                    }
                    if (obj.hasClass("shangchayihang")) {
                        addRowCol(table, 0, t)
                    }
                    if (obj.hasClass("xiachayihang")) {
                        addRowCol(table, 1, t)
                    }
                    if (obj.hasClass("zuochayilie")) {
                        addRowCol(table, 2, t)
                    }
                    if (obj.hasClass("youchayilie")) {
                        addRowCol(table, 3, t)
                    }
                    if (obj.hasClass("shanchuhang")) {
                        addRowCol(table, 4, t)
                    }
                    if (obj.hasClass("shanchulie")) {
                        addRowCol(table, 5, t)
                    }
                    if (obj.hasClass("chexiao")) {
                        chexiaoFunc(t)
                    }
                    if (obj.hasClass("juzhong")) {
                        coll.css("text-align", "center")
                    }
                    if (obj.hasClass("zuoduiqi")) {
                        coll.css("text-align", "left")
                    }
                    if (obj.hasClass("youduiqi")) {
                        coll.css("text-align", "right")
                    }
                    if (obj.hasClass("chuizhijuzhong")) {
                        coll.css("vertical-align", "middle")
                    }
                    if (obj.hasClass("dingduanduiqi")) {
                        coll.css("vertical-align", "top")
                    }
                    if (obj.hasClass("dibuduiqi")) {
                        coll.css("vertical-align", "bottom")
                    }
                    if (obj.hasClass("shangchayihang") || obj.hasClass("xiachayihang") || obj.hasClass("zuochayilie") || obj.hasClass("youchayilie") || obj.hasClass("shanchuhang") || obj.hasClass("shanchulie")) {
                        drugCell(table, t)
                    }
                    rightMousePanel.remove()
                }
            });
            if (!(t.data("record") != undefined && t.data("record").length > 0)) {
                leftPanel.find(".chexiao").remove()
            }
        }

        function copydanyuange(table) {
            if (hasAttr(selectTd, 'rowspan') || hasAttr(selectTd, 'colspan')) {
                alert("合并后的单元格不允许复制样式");
                return;
            }
            selectTdStyle = selectTd.prop("outerHTML");
        }

        function pastedanyuange() {
            selectTd.replaceWith(selectTdStyle);
            selectTd = $(selectTdStyle)
            setSelectTdValue(selectTd);
        }

        function chexiaoFunc(t) {
            if (t.data("record") != undefined) {
                var record = t.data("record");
                if (record.length > 0) {
                    initTable(t, {data: record[record.length - 1], type: 0});
                    record.splice(record.length - 1, 1)
                }
            }
        }

        function closeRightPanel(t) {
            t.find(".rightmouse-panel-div").remove()
        }

        function recordData(t) {
            var record = [];
            if (t.data("record") != undefined) {
                record = t.data("record")
            }
            record[record.length] = t.getExcelHtml();
            t.data("record", record)
        }

        function setFontFamily() {
            let fontFamily = $(this).val();
            let coll = $('table').first().find('.td-chosen-css');
            coll.css('font-family', fontFamily);
        }

        function setFontSize() {
            let fontSize = $(this).val();
            let coll = $('table').first().find('.td-chosen-css');
            coll.css('font-size', fontSize);
        }

        function setFontBold() {
            $(this).toggleClass("buttonBgColor");
            let coll = $('table').first().find('.td-chosen-css');
            let fontWeight = selectTd.css('font-weight');
            if (fontWeight === '400') {
                coll.css('font-weight', '700');
            } else {
                coll.css('font-weight', '');
            }
        }

        function setFontItalic() {
            $(this).toggleClass("buttonBgColor");
            let coll = $('table').first().find('.td-chosen-css');
            let fontItalic = selectTd.css('font-style');
            if (fontItalic === 'normal') {
                coll.css('font-style', 'italic');
            } else {
                coll.css('font-style', '');
            }
        }

        function setUnderline() {
            $('.btn-strike').removeClass('buttonBgColor');
            let coll = $('table').first().find('.td-chosen-css');
            let underline = selectTd.css('text-decoration-line');
            if (underline === 'none' || underline === 'line-through') {
                coll.css('text-decoration-line', 'underline');
                $(this).addClass("buttonBgColor");
            } else {
                coll.css('text-decoration-line', '');
                $(this).removeClass("buttonBgColor");
            }

        }

        function setFontStrike() {
            $('.btn-underline').removeClass('buttonBgColor');
            let coll = $('table').first().find('.td-chosen-css');
            let fontStrike = selectTd.css('text-decoration-line');
            if (fontStrike === 'none' || fontStrike === 'underline') {
                coll.css('text-decoration-line', 'line-through');
                $(this).addClass("buttonBgColor");
            } else {
                coll.css('text-decoration-line', '');
                $(this).removeClass("buttonBgColor");
            }
        }

        function clickBgColor() {
            if (selectTd === undefined || selectTd === {}) {
                alert("请选择需要设置的单元格");
                return;
            }
            $('#bgColorSelect').click();
        }

        function clickFontColor() {
            if (selectTd === undefined || selectTd === {}) {
                alert("请选择需要设置的单元格");
                return;
            }
            $('#fontColorSelect').click();
        }

        function setBgColor() {
            let color = $("#bgColorSelect").val();
            let coll = $('table').first().find('.td-chosen-css');
            coll.css('background-color', color);
        }

        function setFontColor() {
            let color = $("#fontColorSelect").val();
            let coll = $('table').first().find('.td-chosen-css');
            coll.css('color', color);
        }

        function setValign(event) {
            $('.btn-av').removeClass('buttonBgColor');
            $(this).addClass("buttonBgColor");
            let coll = $('table').first().find('.td-chosen-css');
            coll.css("vertical-align", event.data);
        }

        function setAlign(event) {
            $('.btn-ah').removeClass('buttonBgColor');
            $(this).addClass("buttonBgColor");
            let coll = $('table').first().find('.td-chosen-css');
            coll.css("text-align", event.data);
        }

        function mergeBtn() {
            let table = $('.excel').find("table").first();
            var coll = table.find(".td-chosen-css");
            if (coll.length !== 1) {
                mergeCell(table);
            }
        }

        function splitBtn() {
            let table = $('.excel').find("table").first();
            var coll = table.find(".td-chosen-css");
            if (coll.length === 1) {
                mergeCell(table);
            }
        }

        function whiteSpace() {
            $(this).toggleClass('buttonBgColor');
            if (selectTd != null) {
                if (selectTd.css('white-space') === 'nowrap') {
                    selectTd.css('white-space', 'normal')
                } else {
                    selectTd.css('white-space', 'nowrap')
                }
            }
        }

        function getBorderCssStr() {
            let color = $("#borderColor").val();
            let style = $('.borderStyleOption option:selected').val();
            return getBorderWidthByStyle(style) + style + ' ' + color;
        }

        function setBorderLeft() {
            let coll = $('table').first().find('.td-chosen-css');
            let rowSpanNum = 0;
            coll.parent().each(function (index) {
                let $td = $(this).children('.td-chosen-css:visible').first();
                if (index > 0) {
                    let $tdrowspan = $td.attr('rowspan');
                    let $parentTd = $td.parent().prev().children('.td-chosen-css[rowspan]');
                    let rowspan = $parentTd.attr('rowspan');
                    if (rowspan!= null && rowspan!== '' && rowspan!== 0) {
                        rowSpanNum = rowspan;
                    }
                    if (rowSpanNum !== 0) {
                        rowSpanNum--;
                    }
                    if (rowSpanNum === 0 ||($tdrowspan===undefined && $td.prev().css('display')!=='none')) {
                        $td.css('border-left', getBorderCssStr());
                    }
                } else {
                    $td.css('border-left', getBorderCssStr());
                }
            });
        }

        function setBorderTop() {
            let coll = $('table').first().find('.td-chosen-css');
            coll.parent().first().children('.td-chosen-css:visible').css('border-top', getBorderCssStr())
        }

        function setBorderRight() {
            let coll = $('table').first().find('.td-chosen-css');
            let rowSpanNum = 0;
            coll.parent().each(function (index) {
                let $td = $(this).children('.td-chosen-css:visible').last();
                if (index > 0) {
                    let $tdrowspan = $td.attr('rowspan');
                    let $parentTd = $td.parent().prev().children('.td-chosen-css[rowspan]');
                    let rowspan = $parentTd.attr('rowspan');
                    if (rowspan!= null && rowspan!== '' && rowspan!== 0) {
                        rowSpanNum = rowspan;
                    }
                    if (rowSpanNum !== 0) {
                        rowSpanNum--;
                    }
                    if (rowSpanNum === 0 ||($tdrowspan===null && $td.next().css('display')!=='none')) {
                        $td.css('border-right', getBorderCssStr());
                    }
                } else {
                    $td.css('border-right', getBorderCssStr());
                }
            });
        }

        //寻找td上面的第一个td
        function findPrevShowTd($td) {
            let $prevTd = $td.parent().prev().children().eq($td.index());
            if($prevTd.css('display') === 'none'){
                return findPrevShowTd($prevTd);
            }else{
                return $prevTd;
            }
        }

        function setBorderBottom() {
            let coll = $('table').first().find('.td-chosen-css');
            let i = 0;
            coll.parent().last().children('.td-chosen-css').each(function () {
                if($(this).css('display') === 'none' && i === 0){
                    i = 1;
                    findPrevShowTd($(this)).css('border-bottom', getBorderCssStr());
                }else{
                    $(this).css('border-bottom', getBorderCssStr())
                }
            })
        }

        function setBorderAll() {
            let coll = $('table').first().find('.td-chosen-css');
            let style = $('.borderStyleOption option:selected').val();
            if (style === 'none') {
                coll.css('border', '#ccc 1px solid')
            } else {
                coll.css('border', getBorderCssStr())
            }
        }

        function clickBorderColor() {
            $('#borderColor').click();
        }

        function setBorderColor() {
            let color = $("#borderColor").val();
            if (selectTd !== undefined && selectTd !== {}) {
                if (selectTd.css('border-right-color') !== 'rgb(204, 204, 204)') {
                    selectTd.css('border-right-color', color);
                }
                if (selectTd.css('border-top-color') !== 'rgb(204, 204, 204)') {
                    selectTd.css('border-top-color', color);
                }
                if (selectTd.css('border-bottom-color') !== 'rgb(204, 204, 204)') {
                    selectTd.css('border-bottom-color', color);
                }
                if (selectTd.css('border-left-color') !== 'rgb(204, 204, 204)') {
                    selectTd.css('border-left-color', color);
                }
            }
        }

        function showBorderStyleDiv() {
            $('.selectBorderStyle').toggleClass('show');
        }

        function setBorderStyleOption(event) {
            let style = event.data;
            $('.borderStyleOption').val(style);
        }

        function setCellWidth() {
            if (selectTd !== undefined && selectTd !== {}) {
                let width = $('#cell-width').val();
                let index = selectTd.index();
                let oldWidth = selectTd.width();
                selectTd.parent().siblings().each(function (i, ele) {
                    $(this).children('td').eq(index).width(width);
                })
                selectTd.width(width);
                let left = $('.col-width-panel-item').eq(index).position().left;
                let nowPanelItem = $('.col-width-panel-item').eq(index);
                nowPanelItem.css('left', left + selectTd.width() - oldWidth);
                nowPanelItem.nextAll().each(function (i, ele) {
                    let left = $(this).position().left
                    $(this).css('left', left + selectTd.width() - oldWidth);
                })
                updateBorderLeft(oldWidth, selectTd.innerWidth);
            }
        }

        function updateBorderLeft(oldWidth, width) {
            $('.chosen-area-p').eq(1).width($('.chosen-area-p').eq(1).width() - oldWidth + parseInt(width))
            let oldMarginLeft = parseInt($('.chosen-area-p').eq(2).css('margin-left'));
            $('.chosen-area-p').eq(2).css('margin-left', oldMarginLeft - oldWidth + parseInt(width) - 1)
            $('.chosen-area-p').eq(3).width($('.chosen-area-p').eq(3).width() - oldWidth + parseInt(width) - 1)
            $('.chosen-area-p').eq(4).css('margin-left', oldMarginLeft - oldWidth + parseInt(width) - 3)
        }

        function setCellHeight() {
            if (selectTd !== undefined && selectTd !== {}) {
                let index = selectTd.prevAll().last().html();
                let height = $('#cell-height').val();
                let oldHeight = selectTd.height();
                selectTd.parent().height(height);
                let top = $('.row-height-panel-item').eq(index).position().top;
                let nowPanelItem = $('.row-height-panel-item').eq(index);
                nowPanelItem.css('top', top + selectTd.height() - oldHeight - 4);
                nowPanelItem.nextAll().each(function (i, ele) {
                    let top = $(this).position().top
                    $(this).css('top', top + selectTd.height() - oldHeight - 4);
                })
                updateBorderTop(selectTd.outerHeight(), oldHeight)
            }
        }

        function updateBorderTop(height, oldHeight) {
            $('.chosen-area-p').eq(0).height($('.chosen-area-p').eq(0).height() - oldHeight + parseInt(height) - 4)
            $('.chosen-area-p').eq(2).height($('.chosen-area-p').eq(2).height() - oldHeight + parseInt(height) - 4)
            let oldMarginTop = parseInt($('.chosen-area-p').eq(3).css('margin-top'));
            $('.chosen-area-p').eq(3).css('margin-top', oldMarginTop - oldHeight + parseInt(height) - 4)
            $('.chosen-area-p').eq(4).css('margin-top', oldMarginTop - oldHeight + parseInt(height) - 7)
        }

        function getBorderWidthByStyle(style) {
            if (style === 'double') {
                return '3px '
            } else {
                return '2px ';
            }
        }

    }

)(jQuery);
