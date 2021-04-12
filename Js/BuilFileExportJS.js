 function DownLoad() {
        if ($('#txtFromDate').val() == '' || $('#txtToDate').val() == '') {
            alert("bạn phải chọn ngày");
            return;
        }
        $.ajax({
            url: _url + '/api/Report/spRptExcelCustomerSummary',
            type: 'POST',
            crossDomain: true,
            contentType: 'application/json; charset=utf-8',
            data: JSON.stringify({
                fromdate: $('#txtFromDate').val(),
                todate: $('#txtToDate').val()
            }),
            success: function (rs) {
                var bytes = Base64ToBytes(rs);

                //Convert Byte Array to BLOB.
                var blob = new Blob([bytes], { type: "application/octetstream" });

                //Check the Browser type and download the File.
                var isIE = false || !!document.documentMode;
                if (isIE) {
                    window.navigator.msSaveBlob(blob, "CustomerSummary.xlsx");
                } else {
                    var url = window.URL || window.webkitURL;
                    link = url.createObjectURL(blob);
                    var a = $("<a/>");
                    a.attr("download", "CustomerSummary.xlsx");
                    a.attr("href", link);
                    $("body").append(a);
                    a[0].click();
                    $("body").remove(a);
                }
            },
            error: function () {
                bootbox.alert("Get data error");
            }
        });
    }
    function Base64ToBytes(base64) {
        var s = window.atob(base64);
        var bytes = new Uint8Array(s.length);
        for (var i = 0; i < s.length; i++) {
            bytes[i] = s.charCodeAt(i);
        }
        return bytes;
    };