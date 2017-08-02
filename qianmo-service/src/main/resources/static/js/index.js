function submitExcel() {
    var excelFile = $("#excelFile").val();
    if (excelFile == '') {
        alert("请选择需上传的文件!");
        return false;
    }
    if (excelFile.indexOf('.xls') == -1) {
        alert("文件格式不正确，请选择正确的Excel文件(后缀名.xls)！");
        return false;
    }
    $("#fileUpload").submit();
}

//文件下载
function downloadExcel() {
    //var url="http://localhost:8088/excelTest/excelDownload";
    //$.fileDownload("http://localhost:8088/excelTest/111.json");
    /*
     $.fileDownload('http://localhost:8088/excelTest/A1411_1410.xls', {
     successCallback: function (url) {
     alert("下载成功");
     },
     failCallback: function (html, url) {
     alert("下载失败");
     }
     });
     */
    var json1 = '';

    var ajax
    if (window.XMLHttpRequest) {
        ajax = new XMLHttpRequest()
    } else if (window.ActiveXObject) {
        ajax = new window.ActiveXObject()
    } else {
        alert("请升级至最新版本的浏览器")
    }
    if (ajax != null) {
        //ajax.open("POST", "http://123.56.22.114:8080/qianmo-service/excelDownload", true)
        ajax.open("POST", "http://localhost:8088/qianmo-service/excelDownload", true)
        ajax.onload = function () {
            if (ajax.status == 200) {
                var filename = "";
                console.log(ajax.getResponseHeader)
                var disposition = ajax.getResponseHeader('Content-Disposition');
                if (disposition && disposition.indexOf('attachment') !== -1) {
                    var filenameRegex = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/;
                    var matches = filenameRegex.exec(disposition);
                    if (matches != null && matches[1]) filename = matches[1].replace(/['"]/g, '');
                }
                var type = ajax.getResponseHeader('Content-Type');
                console.log(type)
                var blob = new Blob([this.response], {type: type});
                console.log(blob)
                console.log(window.navigator.msSaveBlob);
                if (typeof window.navigator.msSaveBlob !== 'undefined') {
                    // IE workaround for "HTML7007: One or more blob URLs were revoked by closing the blob for which they were created. These URLs will no longer resolve as the data backing the URL has been freed."
                    window.navigator.msSaveBlob(blob, filename);
                } else {
                    var URL = window.URL || window.webkitURL;
                    var downloadUrl = URL.createObjectURL(blob);
                    console.log(downloadUrl)
                    if (filename) {
                        console.log(filename)
                        // use HTML5 a[download] attribute to specify filename
                        var a = document.createElement("a");
                        // safari doesn't support this yet
                        if (typeof a.download === 'undefined') {
                            window.location = downloadUrl;
                        } else {
                            a.href = downloadUrl;
                            a.download = filename;
                            console.log(a)
                            document.body.appendChild(a);
                            a.click();
                        }
                    } else {
                        window.location = downloadUrl;
                    }

                    setTimeout(function () {
                        URL.revokeObjectURL(downloadUrl);
                    }, 100); // cleanup
                }
            }
        }
        ajax.responseType = 'blob';
        ajax.setRequestHeader('Content-type', 'application/json;charset=utf-8');
        ajax.send(json1)
    }
}

//获取json
function getContentJson() {
    var url = "G2501-141-1-2017072116243504.txt"
    $.ajax({
        type: "post",
        url: "http://localhost:8088/qianmo-service/getContentJson",
        data: 'url=' + url,
        contentType: "application/x-www-form-urlencoded; charset=utf-8",
        success: function (data) {
            alert(data);
        }
    });
}

function changeContent() {
    $(function () {
        $.ajax({
            type: "post",
            url: "https://localhost:8088/qianmo-service/changeContent",
            contentType: "application/json;charset=utf-8",
            data: '[{"coord":"测试A_1","value":"=i23+1"},{"coord":"测试B_1","value":"=o23+1"}]',
            success: function (data) {
                alert(data);
            }
        });
    });
}