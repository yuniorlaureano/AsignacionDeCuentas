$(function () {

    showErrorMessage();
    showMessage();

    $("#command-button-assignar").on("click", function () {
        showColorLoading();
    });
});

function showErrorMessage() {

    var errorc = $("#err-container");

    if (errorc.length != 0) {

        var html = "<p>" + errorc.attr("data-error") + "</p>";

        getTimerDialogMsg("danger", html, "panel-result-account", 7000);
    }

}

function showMessage() {

    var messagec = $("#msg-container");

    if (messagec.length != 0) {

        var html = "<p>" + messagec.attr("data-message") + "</p>";

        getTimerDialogMsg("info", html, "panel-result-account", 7000);
    }

}