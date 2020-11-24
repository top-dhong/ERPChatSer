$(document).on("pageinit", "#jqUserInfo", function () {

    $("#btnSubmit").bind("click", function (e) {

        var pram = "{ManName:'" + "张三" + "', ManTel:'" + "13540486632" + "', LastCompany:'" + ""
            + "', LastStation:'" + "" + "', LastSalary:'" + "" + "', WantSalary:'" + ""
            + "', HighestDegree:'" + "" + "', School:'" + "" + "', Speciality:'" + "" + "'}";
        alert(pram);

        $.ajax({
            type: "POST",
            contentType: "application/json",
            url: "SHWXSvr.asmx/PutCandidateInfo",
            data: pram,
            dataType: 'json',
            success: function (result) {
                alert('success');
            },
            error: function (jqXHR, textStatus, errorThrown) {
                alert('fail');
            }
        });

        return false; 
    });

});
