// ==UserScript==
// @name        wechat_contact_expert
// @name:zh-CN        wechat_contact_expert
// @name:zh-TW        wechat_contact_expert
// @name:ja        wechat_contact_expert
// @name:ru        wechat_contact_expert
// @name:kr        wechat_contact_expert
// @namespace   wechat_contact_expert
// @supportURL  https://github.com/zhuzemin
// @description expert wechat contact NickName, Remark, Gender, City, Avatar
// @description:zh-CN expert wechat contact NickName, Remark, Gender, City, Avatar
// @description:zh-TW expert wechat contact NickName, Remark, Gender, City, Avatar
// @description:ja expert wechat contact NickName, Remark, Gender, City, Avatar
// @description:ru expert wechat contact NickName, Remark, Gender, City, Avatar
// @description:kr expert wechat contact NickName, Remark, Gender, City, Avatar
// @include     https://wx.qq.com/
// @version     1.0
// @run-at      document-end
// @author      zhuzemin
// @license     Mozilla Public License 2.0; http://www.mozilla.org/MPL/2.0/
// @license     CC Attribution-ShareAlike 4.0 International; http://creativecommons.org/licenses/by-sa/4.0/
// @grant       GM_xmlhttpRequest
// @grant         GM_registerMenuCommand
// @grant         GM_setValue
// @grant         GM_getValue
// @require      https://github.com/gyeongseokKang/exportExcel_javascript/raw/master/lib/exceljs.js
// ==/UserScript==

//config
let config = {
    'debug': true,
    'remark_keyword': [],
    'excelName': 'wechat_contact.xlsx',
    'table': [['微信头像', '备注名', '微信昵称', '性别', '省份', '城市', '微信签名']]
}
let debug = config.debug ? console.log.bind(console) : function () {
};


// prepare UserPrefs
setUserPref(
    'remark_keyword',
    config.remark_keyword,
    '备注名_关键字',
    ``,
);


//userscript entry
let init = function () {
    //create button
    if (window.self === window.top) {
        debug("init");
        let interval = setInterval(function () {
            let contactList = document.querySelector('div.tab_item.no_extra');
            if (contactList != null) {
                clearInterval(interval);
                debug("get contactList");
                for (let contact of Object.values(unsafeWindow._contacts)) {
                    if (contact.isBrandContact() == 0) {
                        if (contact.isConversationContact()) {
                            if (!contact.isFileHelper()) {
                                if (!contact.isRecommendHelper()) {
                                    if (!contact.isRoomContact()) {
                                        if (!contact.isSpContact()) {
                                            debug(contact.RemarkName);
                                            config.table.push([
                                                document.location.href.replace(/\/$/, '') + contact.HeadImgUrl,
                                                contact.RemarkName,
                                                contact.NickName,
                                                gender = contact.Sex == 1 ? '男' : '女',
                                                contact.Province,
                                                contact.City,
                                                contact.Signature
                                            ]);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                let promise = [];
                let workbook = new exceljs.Workbook();


                const sheet = workbook.addWorksheet("sheet1");
                config.table.forEach((contact, rowID) => {
                    contact.forEach((property, colID) => {
                        let row = sheet.getRow(rowID + 1);
                        row.height = 50;
                        let cell = row.getCell(colID + 1);
                        if (property.includes('wx.qq.com')) {
                            promise.push(
                                getBase64FromUrl(property).then((base64Code) => {
                                    const image = workbook.addImage({
                                        base64: base64Code,
                                        extension: "jpg",
                                    });

                                    sheet.addImage(image, {
                                        tl: { col: colID, row: rowID },
                                        ext: { width: 35, height: 35 }
                                    });
                                })
                            );
                        }
                        else {

                            cell.value = property
                        }
                    });
                });
                Promise.all(promise).then(() => {
                    workbook.xlsx.writeBuffer().then((b) => {
                        let a = new Blob([b]);
                        let url = window.URL.createObjectURL(a);

                        let elem = document.createElement("a");
                        elem.href = url;
                        elem.download = `wechat_contact.xlsx`;
                        document.body.appendChild(elem);
                        elem.style = "display: none";
                        elem.click();
                        elem.remove();
                    });
                });
            }
        }, 10000);
    }

}
window.addEventListener('DOMContentLoaded', init);


function request(object, func, timeout = 60000) {
    GM_xmlhttpRequest({
        method: object.method,
        url: object.url,
        headers: object.headers,
        responseType: object.respType,
        data: object.body,
        timeout: timeout,
        onload: function (responseDetails) {
            debug(responseDetails);
            //Dowork
            func(responseDetails, object);
        },
        ontimeout: function (responseDetails) {
            debug(responseDetails);
            //Dowork
            func(responseDetails);

        },
        ononerror: function (responseDetails) {
            debug(responseDetails);
            //Dowork
            func(responseDetails);

        }
    });
}


/**
 * Create a user setting prompt
 * @param {string} varName
 * @param {any} defaultVal
 * @param {string} menuText
 * @param {string} promtText
 * @param {function} func
 */
function setUserPref(varName, defaultVal, menuText, promtText, func = null) {
    GM_registerMenuCommand(menuText, function () {
        let val = prompt(promtText, GM_getValue(varName, defaultVal));
        if (val === null) { return; }  // end execution if clicked CANCEL
        GM_setValue(varName, val);
        if (func != null) {
            func(val);
        }
    });
}

const getBase64FromUrl = async (url) => {
    const data = await fetch(url);
    const blob = await data.blob();
    return new Promise((resolve) => {
        const reader = new FileReader();
        reader.readAsDataURL(blob);
        reader.onloadend = () => {
            const base64data = reader.result;
            resolve(base64data);
        }
    });
}