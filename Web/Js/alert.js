if(typeof(PAGETYPE) == "undefined")
    var PAGETYPE = "";

$(document).ready(function() {
    var getTimerID = setInterval(function() {
        var getTimerUrl = "/getNewInfo.aspx";
        $.ajax({
            url: getTimerUrl,
            dataType: 'json',
            type: 'POST',
            success: function(data) {
                var iChReq = parseInt(data.chreq);
                var iChStandby = parseInt(data.chstand);
                var iChComplete = parseInt(data.chcom);

                var iDischReq = parseInt(data.dischreq);
                var iDischStandby = parseInt(data.dischstand);
                var iDischComplete = parseInt(data.dischcom);
                
                var iMemTotal = parseInt(data.memtotal);
                var iMemNew = parseInt(data.memnew);
                var iMemNew_0 = parseInt(data.memnew_0);

                if ((isNaN(iChReq) || iChReq < 1) &&
                   (isNaN(iDischReq) || iDischReq < 1) &&
                   (isNaN(iMemNew_0) || iMemNew_0 < 1)) {
                    getFlashElement("objPlaySound", "objPlaySoundName").stopMP3();
                }
                else if (iChReq > 0)
                    getFlashElement("objPlaySound", "objPlaySoundName").playMP3("charge");
                else if (iDischReq > 0)
                    getFlashElement("objPlaySound", "objPlaySoundName").playMP3("discharge");
                
                $("#spChargeRequest_pip").html(isNaN(iChReq) ? 0 : iChReq);
                $("#spChargeStandby_pip").html(isNaN(iChStandby) ? 0 : iChStandby);
                $("#spChargeComplete_pip").html(isNaN(iChComplete) ? 0 : iChComplete);

                $("#spDischargeRequest_pip").html(isNaN(iDischReq) ? 0 : iDischReq);
                $("#spDischargeStandby_pip").html(isNaN(iDischStandby) ? 0 : iDischStandby);
                $("#spDischargeComplete_pip").html(isNaN(iDischComplete) ? 0 : iDischComplete);

                $("#spMember_pip").html(isNaN(iMemTotal) ? 0 : iMemTotal);
                $("#spMemberNew_pip").html(isNaN(iMemNew) ? 0 : iMemNew);

                if (PAGETYPE == "CHARGE") {
                    setCookie("saveChargeCount", iChReq);
                } else if (PAGETYPE == "DISCHARGE") {
                    setCookie("saveDischargeCount", iDischReq);
                }
            },
            error: function(XMLHttpRequest, textStatus, errorThrown) {
            }
        });
    }, 10000);
});