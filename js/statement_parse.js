/**
 * Created by a7268 on 2016/12/1.
 */

+function ($) {

    var $uploadExcel = $("#upload_excel");
    var $filenameSpan = $("#filename");
    var $uploadForm = $("#upload_form");
    var $selectStatement = $("#select_statement");
    var $selectQN = $("#selectQN");
    var $selectDate = $("#statement_date");
    var $textArea = $("#view_code");
    $uploadExcel.on("change", uploadHandler);

    function uploadHandler() {
        var path = $uploadExcel.val();
        $filenameSpan.html(path.substring(path.lastIndexOf("\\") + 1));
        $.ajax({
            url: "/statements_parse_action",
            type: "post",
            dataType: "json",
            data: new FormData($uploadForm[0]),
            async: false,
            cache: false,
            contentType: false,
            processData: false,
            success: function (data) {

                if (data.length == 0) {
                    alert("报表格式错误或无内容");
                    $selectStatement.html("");
                    $selectQN.html("");
                    $selectDate.html("");
                    $textArea.html("");
                }
                else {
                    viewCode(data);
                }
            },
            error: function (data, status, e) {
                console.error(e);
            }
        });
        return false;
    }

    /*
     * 创建option元素并设置值为str，然后添加到$select的最后
     */
    function createOption(str, value, $select) {
        var $option = $("<option>");
        $option.prop({
            value: value
        });
        $option.html(str);
        $select.append($option);
    }

    function viewCode(data) {
        var codes = {};
        $selectStatement.html("");
        for (var name in data) {
            if (data.hasOwnProperty(name) && name != "length") {
                createOption(name, name, $selectStatement);
                codes[name] = parseStatement(data[name]);
            }
        }

        //select改变事件
        $selectStatement.on("change", function () {
            $selectQN.html("");
            if ($selectStatement.val().indexOf("资产负债表") != -1) {
                createOption("期末余额", "qms", $selectQN);
                createOption("年初余额", "ncs", $selectQN);
            }
            else {
                createOption("本月金额", "qms", $selectQN);
                createOption("本年累计", "ncs", $selectQN);
            }
            $selectQN.trigger("change");
        });
        $selectQN.on("change", function () {
            var date = codes[$selectStatement.val()].date;
            $selectDate.html("");
            createOption(date.getFullYear() + "年" + (date.getMonth() + 1) + "月" + date.getDate() + "日", "statement_date", $selectDate);
            $textArea.html(codes[$selectStatement.val()][$selectQN.val()]);
        });


        $selectStatement.trigger("change");

        /*
         * 解析sheet对象，并返回解析后的对象
         * 该对象有date、qms、ncs成员对象
         */
        function parseStatement(sheet) {
            var data = sheet.data;
            var obj = {};
            var no,         //编号
                qms,        //期末数或本月数所在列
                ncs,        //年初数或本年累计所在列
                rows;        //第几行至第几行
            var QN = {};    //资产负债表、利润表input name值

            obj.qms = {};
            obj.ncs = {};

            if (sheet.name.indexOf("资产负债表") != -1) {
                obj.date = xlsxDateParse(data[2][8]);
                no = [2, 6];
                qms = [3, 7];
                ncs = [4, 8];
                rows = [5, 36];
                QN = {
                    qms: "qms",
                    ncs: "ncs"
                };
            }
            else if (sheet.name.indexOf("利润表") != -1) {
                obj.date = xlsxDateParse(data[2][5]);
                no = [2];
                qms = [5];
                ncs = [3];
                rows = [4, 36];
                QN = {
                    qms: "sns",
                    ncs: "bns"
                };
            }
            parse(no, qms, ncs, rows);

            for (var name in obj) {
                if (name != "date") {
                    obj[name] = dataToCode(obj[name], QN[name]);
                }
            }

            return obj;

            /*
             * 将数据转换成js代码
             */
            function dataToCode(data, QN) {
                var code = "";
                code = '+function () {var inputs = document.getElementsByTagName("input");var ' + QN + ' = {};for (var i = 0; i < inputs.length; i++) {var name = inputs[i].name;if (name.substr(0, 3) == "' + QN + '") {if(inputs[i].disabled==false){' + QN + '[name.substr(4, 5)] = inputs[i];}}}var obj=' + JSON.stringify(data) + ';for(var q in ' + QN + ') {' + QN + '[q].value = obj[q];' + QN + '[q].onblur();}}()';
                return code;
            }

            /*
             * 按指定参数解析报表
             * no：编号所在列
             * qms：期末数或本月数所在列
             * ncs：年初数或本年累计所在列
             * 以上参数为数组，no、qms、ncs的长度需相同
             * rows：每列扫描行数  i=rows[0];i<rows[1]
             */
            function parse(no, qms, ncs, rows) {

                for (var i = 0; i < no.length; i++) {
                    for (var row = rows[0]; row < rows[1]; row++) {
                        var noName = numberToString(data[row][no[i]]);      //获取编号
                        if (noName) {
                            obj.qms[noName] = Math.round(data[row][qms[i]] * 100) / 100 || 0;
                            obj.ncs[noName] = Math.round(data[row][ncs[i]] * 100) / 100 || 0;
                        }
                    }
                }

                //数字转成字符串并不成两位数
                function numberToString(n) {
                    if (n == null) {
                        return;
                    }
                    if (n < 10) {
                        return "0" + n;
                    }
                    else {
                        return "" + n;
                    }
                }
            }

            //xlsx的日期单元格转换为日期格式
            function xlsxDateParse(time) {
                return new Date(new Date(1900, 0, 0).getTime() - 1 + time * 24 * 60 * 60 * 1000);
            }
        }

    }

}(jQuery);

