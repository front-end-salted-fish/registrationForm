//打印表格
var idTmr;

function getExplorer() {
    var explorer = window.navigator.userAgent;
    //ie  
    if (explorer.indexOf("MSIE") >= 0) {
        return 'ie';
    }
    //firefox  
    else if (explorer.indexOf("Firefox") >= 0) {
        return 'Firefox';
    }
    //Chrome  
    else if (explorer.indexOf("Chrome") >= 0) {
        return 'Chrome';
    }
    //Opera  
    else if (explorer.indexOf("Opera") >= 0) {
        return 'Opera';
    }
    //Safari  
    else if (explorer.indexOf("Safari") >= 0) {
        return 'Safari';
    }
}

function method5(tableid) {
    if (getExplorer() == 'ie') {
        var curTbl = document.getElementById(tableid);
        var oXL = new ActiveXObject("Excel.Application");
        var oWB = oXL.Workbooks.Add();
        var xlsheet = oWB.Worksheets(1);
        var sel = document.body.createTextRange();
        sel.moveToElementText(curTbl);
        sel.select();
        sel.execCommand("Copy");
        xlsheet.Paste();
        oXL.Visible = true;

        try {
            var fname = oXL.Application.GetSaveAsFilename("Excel.xls",
                "Excel Spreadsheets (*.xls), *.xls");
        } catch (e) {
            print("Nested catch caught " + e);
        } finally {
            oWB.SaveAs(fname);
            oWB.Close(savechanges = false);
            oXL.Quit();
            oXL = null;
            idTmr = window.setInterval("Cleanup();", 1);
        }

    } else {
        tableToExcel(tableid)
    }
}

function Cleanup() {
    window.clearInterval(idTmr);
    CollectGarbage();
}
var tableToExcel = (function () {
    var uri = 'data:application/vnd.ms-excel;base64,',
        template = '<html><head><meta charset="UTF-8"></head><body><table  border="1">{table}</table></body></html>',
        base64 = function (
            s) {
            return window.btoa(unescape(encodeURIComponent(s)))
        },
        format = function (s, c) {
            return s.replace(/{(\w+)}/g, function (m, p) {
                return c[p];
            })
        }
    return function (table, name) {
        if (!table.nodeType)
            table = document.getElementById(table)
        var ctx = {
            worksheet: name || 'Worksheet',
            table: table.innerHTML
        }
        window.location.href = uri + base64(format(template, ctx))
    }
})();
/* 提取公共的文本类表单项验证 */
function commonInfo(reg, id) {
    var inputText = document.getElementById(id);
    var inputValue = inputText.value;
    var inputSpan = document.getElementById(id + "Span");

    if (inputValue == null || inputValue.length == 0) {
        inputSpan.innerHTML = "不能为空！";
        inputSpan.style.color = "red";
        inputSpan.style.background = "none";
        return false;
    } else {
        if (!reg.test(inputValue)) {
            inputSpan.innerHTML = "格式有误！";
            inputSpan.style.color = "red";
            inputSpan.style.background = "none";
            return false;
        } else {
            inputSpan.innerHTML = "√";
            inputSpan.style.color = "white";
            inputSpan.style.background = "rgb(50,200,100)";
            inputSpan.style.padding = "2px 4px";
            inputSpan.style.borderRadius = "15px";
            return true;
        }
    }

}
/* 验证用户名 ，学院，专业班级和宿舍*/
function checkUsername(ID) {
    var reg = /^[a-zA-Z\u4e00-\u9fa5]{1}[a-zA-Z0-9_\u4e00-\u9fa5]{1,19}$/ig; /* 既支持中文，又支持英文字符，不能以数字开头。限制为2-20个字符 */
    var id = ID;
    return commonInfo(reg, id);
}
var userName = document.getElementById("userName");
var ademicInstitution = document.getElementById("ademicInstitution");
var Class = document.getElementById('class');
var dormitory = document.getElementById("dormitory");
userName.addEventListener("blur", function () {
    checkUsername('userName');
})
ademicInstitution.addEventListener("blur", function () {
    checkUsername('ademicInstitution');
})
Class.addEventListener("blur", function () {
    checkUsername('class');
})
dormitory.addEventListener("blur", function () {
    checkUsername("dormitory");
})
/* 验证手机号 */
function checkPhone() {
    var reg = /^[1][0-9]{10}$/ig; /* 验证手机号 */
    var id = "phone";
    return commonInfo(reg, id);
}
let $phone = $("#phone");
$phone.on("blur", () => {
    checkPhone();
})
/*验证学号 */
function checkstudentNumber() {
    var reg = /^\d*$/ig; /* 验证学号 */
    var id = 'studentNumber';
    return commonInfo(reg, id);
}
let $studentNumber = $("#studentNumber");
$studentNumber.on("blur", () => {
    checkstudentNumber();
});

/*点击关闭按钮会关闭页面 */
(() => {
    let $close = $("#close");
    let $outerBox = $("#tableid");
    $close.on("click", () => {
        $outerBox.hide();
    })
})();

let $departmentName = $("#departmentName");
$.ajax({
    method: "GET", // 一般用 POST 或 GET 方法
    url: "http://10.21.23.177:8080/apply/getDepartmentName", // 要请求的地址
    dataType: "json", // 服务器返回的数据类型，可能是文本 ，音频 视频 script 等浏览 （MIME类型）器会采用不同的方法来解析。
    data: {
        //communityId
        //发送社团id，待定怎么确认，
        //发送到服务器的数据。将自动转换为请求字符串格式。GET 请求中将附加在 URL 后。查看 processData 选项说明以禁止此自动转换。必须为 Key/Value 格式。如果为数组，jQuery 将自动为不同值对应同一个名称。如 {foo: ["bar1", "bar2"]} 转换为 "&foo=bar1&foo=bar2"。
        communityId: 2
    },
    // beforeSend: function (xhr) {
    //     xhr.setRequestHeader('Authorization', 'Barber Token');
    // }, //这里设置header
    // headers: {
    //     'Content-Type': 'application/json;charset=utf8',
    //     'Authorization': 'Barber Token'
    // },
    success(data) {

        datas = data;
        console.log("成功了"); // 成功之后执行这里面的代码

        let length = datas.object.length;
        console.log(datas.object);

        // let name = datas.object.name;
        $.each(datas.object, (length, e) => {
            let innerName = '<option>' + datas.object[length].name + '</option>';
            $departmentName.append(innerName);
            return true;
        })

    },
    error(e) {
        console.log(e) //请求失败是执行这里的函数
    }
});
/*点提交按钮会先检查是否都已填写*/
(() => {
    let $btn = $("#btn");
    let $form = $(".form");
    $btn.on("click", () => {

        let $formSpan = $(".formSpan");

        $form.each((i) => {
            if ($form.eq(i).val() == "") {
                $formSpan.eq(i).text('不能为空！');
                $formSpan.eq(i).css({
                    "color": "red",
                    "background": "none"
                });

                i++;
            }
        })
        //点击发送Ajax请求
        let $btn = $("#btn");
        let $userNameval = $("#userName").val();
        let $sexval = $("#sex").val();
        let $ademicInstitutionval = $("#ademicInstitution").val();
        let $classval = $("#class").val();
        let $dormitoryval = $("#dormitory").val();
        let $phoneval = $("#phone").val();
        let $studentNumberval = $("#studentNumber").val();
        let $departmentNameval = $("#departmentName").val();
        let $textAreaval = $("#textArea").val();
        // let $departmentNameID = $("#departmentName").selectedindex;
        let $departmentNameID = $("#departmentName").get(0).selectedIndex + 1;

        let usersData = {
            name: $userNameval,
            sex: $sexval,
            majorClass: $classval,
            academy: $ademicInstitutionval,
            phone: $phoneval,
            studentNumber: $studentNumberval,
            introduce: $textAreaval,
            dormitory: $dormitoryval,
            communityId: 1, //社团id
            departmentId: $departmentNameID //部门id
        };
        console.log(usersData);
        if ($userNameval == '' ||
            $sexval == '' ||
            $ademicInstitutionval == '' ||
            $classval == '' ||
            $dormitoryval == '' ||
            $phoneval == '' ||
            $studentNumberval == '' ||
            $departmentNameval == '' ||
            $textAreaval == '') {
            alert("还有没填的项！");

        } else if ($(".formSpan").eq(0).html() == "格式有误！" ||
            $(".formSpan").eq(1).html() == "格式有误！" ||
            $(".formSpan").eq(2).html() == "格式有误！" ||
            $(".formSpan").eq(3).html() == "格式有误！" ||
            $(".formSpan").eq(4).html() == "格式有误！" ||
            $(".formSpan").eq(5).html() == "格式有误！" ||
            $(".formSpan").eq(6).html() == "格式有误！") {
            alert("存在格式有误的项！");
        } else {
            $.ajax({

                type: 'POST',

                data: JSON.stringify(usersData),

                contentType: 'application/json',

                dataType: 'json',

                url: 'http://10.21.23.177:8080/apply/studentApplicationForm',

                success: function (data) {

                    var datas = data;
                    console.log("成功了");


                    switch (datas.code) {
                        case 0:
                            alert(datas.msg + "恭喜你报名成功!");
                            break;
                        case 2:
                            alert("还有没填的项！");
                            break;
                        case 4:
                            alert(datas.msg);
                            break;
                        default:
                            alert(datas.msg);
                            break;
                    }




                },

                error: function (e) {

                    alert("操作失败请重试");

                }

            });
        }


    })
})();