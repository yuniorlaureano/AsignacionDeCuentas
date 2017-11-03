$(function () {

    $("#command-button-upload-file").on("click", function () {
        $("#excel").click();
    });

    showUploadedFileProperty();
    showErrorMessage();
    showMessage();
    sendSheet();

});

function showUploadedFileProperty() {

    $("#excel").on("change", function () {

        var files = $(this)[0].files;

        if (files.length > 0) {
            showColorLoading();
            var file = files[0];
            var fileParts = file.name.split(".");
            $(".file-size").text(file.size);
            $(".file-name").text(fileParts[0]);
            $(".file-type").text(fileParts[1]);

            $("#form-upload-excel").submit();
        }
    });
}

function showErrorMessage() {

    var errorc = $("#err-container");

    if (errorc.length != 0) {

        var html = "<p>" + errorc.attr("data-error") + "</p>";
        getTimerDialogMsg("danger", html, "panel-upload-file", 50000);
    }

}

function showMessage() {

    var messagec = $("#msg-container");

    if (messagec.length != 0) {

        var html = "<p>" + messagec.attr("data-message") + "</p>";

        getTimerDialogMsg("info", html, "panel-upload-file", 50000);
    }

}

function sendSheet() {

    $("#command-button-sheet").on("click", function () {
        
        var form = $($("#form-send-excel-sheets")[0]);
        var inputElement = form.find("input#sheet");

        console.log(inputElement.val());
        if (inputElement.val() != "invalid") {
            showColorLoading();
            form.submit();            
        }
        else
        {
            getTimerDialogMsg("info", "La información no pudo ser cargada. Favor intentar nuevamente.", "panel-upload-file", 50000);
        }


    });

}