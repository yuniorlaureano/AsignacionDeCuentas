/*
Utils library
Contains js functions for reuse in the project

Mirla Del Rio 16/04/2015
*/


//create an editable popup celd
function getEditable() {

    $.fn.editable.defaults.mode = 'popup';

    $('.editable').editable({
        validate: function (value) {
            if (value === null || value === '') {
                return 'El campo no puede estar vacío';
            }
        }

    });

    $('.dateEditable').editable({
        format: 'YYYY-MM-DD',
        viewformat: 'YYYY-MM-DD',
        template: 'DD/ MM / YYYY',
        combodate: {
            minYear: 2015,
            maxYear: 2030
        },
        validate: function (value) {
            if (value === null || value === '') {
                return 'Debe ingresar una fecha';
            }
        }
    });

}

function getEditableAllRows(nrow) {

    $.fn.editable.defaults.mode = 'popup';

    $('.editable', nrow).editable({
        validate: function (value) {
            if (value === null || value === '') {
                return 'El campo no puede estar vacío';
            }
        }

    });

    $('.dateEditable', nrow).editable({
        format: 'YYYY-MM-DD',
        viewformat: 'YYYY-MM-DD',
        template: 'DD/ MM / YYYY',
        combodate: {
            minYear: 2015,
            maxYear: 2030
        },
        validate: function (value) {
            if (value === null || value === '') {
                return 'Debe ingresar una fecha';
            }
        }
    });

}

function getEditableInline() {

    $.fn.editable.defaults.mode = 'inline';

    $('.editable').editable();
}

//Change html control text
function changeText(ID, msg) {
    $("#" + ID).html(msg);
}

function activeLoading(id) {
    $("#" + id).addClass("fa fa-spinner fa-pulse fa-lg");
}

function disactiveLoading(id) {
    $("#" + id).removeClass("fa fa-spinner fa-pulse fa-lg");
}

//Format plain number to number with comma exp: 9999 -> 9,999
function formatNumber(money) {
    return accounting.formatNumber(money);
}

//Toggle bill projection report between detail table and pivot table
function toggleTable(idTableResume, idTableDetail, idToggleBtn, tableDetail, tablePivot, idAlertInfo) {

        $("#" + idToggleBtn).click(function () {

            if ($("#" + idTableResume).css("display") == "block") {

                $("#" + idTableResume).hide();
                $("#" + idTableDetail).show();
                tableDetail;
                $("#" + idToggleBtn).html('Ver Resumen');

            } else {

                $("#" + idTableDetail).hide();
                $("#" + idTableResume).show();
                tablePivot;
                $("#" + idToggleBtn).html('Ver Detalles');

            }

        });


}
//Export to Excel Novedades..............................
function generateReportsToExcel(url, idBtnExportExcel, idAlertInfo) {
    $('#' + idBtnExportExcel).click(function () {
        changeText(idBtnExportExcel, "Exportando..");
        $.ajax({
            url: url,
            type: 'post',
            dataType: 'json',
            contentType: "application/json",
            success: function (data) {
                window.location = data.file;
                changeText(idBtnExportExcel, "Exportar Excel");
            },
            error: function (ex) {
                $("#" + idBtnExportExcel).html("Exportar Excel");
                getDialogMsg("error", "Error", "No se pudo exportar a Excel, detalle: " + ex.message, idAlertInfo);
            },
            data: JSON.stringify({
                "data": data
            })
        });

    });
}

//Export report in excel file
function generateReportsExcel(url, idBtnExportExcel, report, idAlertInfo) {
    $('#' + idBtnExportExcel).click(function () {
        changeText(idBtnExportExcel, "Exportando..");
        $.ajax({
            url: url,
            type: 'post',
            dataType: 'json',
            contentType: "application/json",
            success: function (data) {
                window.location = data.file;
                changeText(idBtnExportExcel, "Exportar Excel");
            },
            error: function (ex) {
                $("#" + idBtnExportExcel).html("Exportar Excel");
                getDialogMsg("error", "Error", "No se pudo exportar a Excel, detalle: " + ex.message, idAlertInfo);
            },
            data: JSON.stringify({
                "data": report
            })
        });

    });

}

function ExportToExcel(url, idBtnExportExcel, data) {
    changeText(idBtnExportExcel, "Exportando..");
        $.ajax({
            url: url,
            type: 'post',
            dataType: 'json',
            contentType: "application/json",
            success: function (data) {
                window.location = data.file;
                changeText(idBtnExportExcel, "Exportar Excel");
            },
            error: function (ex) {
                $("#" + idBtnExportExcel).html("Exportar Excel");
            },
            data: JSON.stringify({
                "data": data
            })
        });

}

//show a dialog with a msg
function getDialogMsg(action, subject, msg, idDialogMsg) {

    $("#" + idDialogMsg).empty();
    var alert = "alert-purple";

    if (action == "success") {
        alert = "alert-success"
    } else if (action == "error") {
        alert = "alert-danger"
    } else if (action == "warning") {
        alert = "alert-warning"
    }

    $("#" + idDialogMsg).addClass(alert).append("<button class='close' data-dismiss='alert'><span>&times;</span></button><strong>" + subject + "</strong> " + msg).css("display", "block");

}

function getTimerDialogMsg(action, contentHTML, idPanel, timer) {
   
    $.niftyNoty({
        type: action,
        container: '#' + idPanel,
        html: contentHTML,
        focus: false,
        timer: timer ? timer : 0
    });

}

function showLoading(obj) {
    var $el = obj.niftyOverlay(), relTime;
    $el.niftyOverlay('hide');
    $el.niftyOverlay('show')
    return $el;
}

function hideLoading($el) {
    $el.niftyOverlay('hide');
}

function showColorLoading() {
    $('.loader-wrapper').show();
}

function hideColorLoading() {
    $('.loader-wrapper').hide();
}

function getBtnExcelGenerator(idBtn, idTable) {

    $("#" + idBtn).remove();
    $("#" + idTable + "_filter").append('<button id="' + idBtn + '" class="btn btn-mint btn-labeled fa fa-table add-tooltip mar-lft' +
        ' data-placement="top" data-toggle="tooltip" data-original-title="Exporta tabla en Excel">Exportar Excel</button>');

}

function relocatePag_Filt (tabID) {
    $(tabID + "_filter").css("float", "right");
    $(tabID + "_paginate").css("float", "right");
}

function formatDate (dateUnformated) {

    var dateFormated = '';

    if (dateUnformated != null || dateUnformated != undefined) {

        dateFormated = dateUnformated.trim().split('T')[0];
    }
    return dateFormated;
}

 function checkAll(checkAllID) {

    $('#' + checkAllID).change(function () {
        var checkboxes = $(this).closest('form').find(':checkbox');
        if ($(this).is(':checked')) {
            checkboxes.prop('checked', true);
        } else {
            checkboxes.prop('checked', false);
        }
    });
 }


 /*DESPLEGA UNA ALERTA, */
 function alertNoty(msg, title, type) {

     $.niftyNoty({
         type: type,
         title: title,
         message: msg,
         container: 'floating',
         timer: 5000
     });    
 }