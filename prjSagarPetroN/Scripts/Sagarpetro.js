var baseUrl = '/prjSagarPetroN';
var requestsProcessed = [];
var iRequestId = 1;
var docNo = "0";

function isRequestCompleted(iRequestId, processedRequestsArray) {
    return processedRequestsArray.indexOf(iRequestId) === -1 ? false : true;
}

function isRequestProcessed(iRequestId) {
    for (let i = 0; i < requestsProcessed.length; i++) {
        if (requestsProcessed[i] == iRequestId) {
            return true;
        }
    } return false;
}
//on authorize     /prjSagarPetroN/Scripts/Sagarpetro.js     SagarpetroOnauth
function SagarpetroOnauth() {
    debugger
    requestsProcessed = [];
    Focus8WAPI.getFieldValue("setOnAuthPPl", ["", "DocNo"], Focus8WAPI.ENUMS.MODULE_TYPE.TRANSACTION, false, iRequestId);
}

function setOnAuthPPl(response) {
    debugger
    if (isRequestCompleted(response.iRequestId, requestsProcessed)) {
        return;
    }
    requestsProcessed.push(response.iRequestId);
    iRequestId++;
    logDetails = response.data[0];
    docNo = response.data[1].FieldValue;
    console.log(response);
    // Focus8WAPI.getBodyFieldValue("setOnAuthPPl2", ["Item", "Unit", "QuantitytobeProduced", "DueDate"], Focus8WAPI.ENUMS.MODULE_TYPE.TRANSACTION, false, body_Index, iRequestId);
    $.ajax({
        url: baseUrl + "/SagarPetro/GetPpl",
        type: "POST",
        datatype: 'JSON',
        data: { "CompanyId": logDetails.CompanyId, "vtype": logDetails.iVoucherType, "DocNo": docNo, "UserId": logDetails.LoginId },
        //data,//
        success: function (response) {
            debugger
            if (response.status == true) {
                alert(response.Message);
                Focus8WAPI.continueModule(Focus8WAPI.ENUMS.MODULE_TYPE.TRANSACTION, true)
            }
            if (response.status == false) {
                alert(response.Message);
                Focus8WAPI.continueModule(Focus8WAPI.ENUMS.MODULE_TYPE.TRANSACTION, false)//true
            }

        },
        error: function (error) {
            Focus8WAPI.continueModule(Focus8WAPI.ENUMS.MODULE_TYPE.TRANSACTION, false)
            console.log('Error :: ', error);
        }
    });
}
//function setOnAuthPPl2(response) {
//    debugger;
//    if (isRequestProcessed(response.iRequestId)) {
//        return;
//    }
    
//    requestsProcessed.push(response.iRequestId);
//    console.log(response);
//    iRequestId++;
//    var itemid = response.data[0].FieldValue;
//    var unitid = response.data[1].FieldValue;
//    var qty = response.data[2].FieldValue;
//    var wh = response.data[3].FieldValue;
//}