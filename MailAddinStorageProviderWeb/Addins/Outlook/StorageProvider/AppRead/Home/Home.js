/// <reference path="../App.js" />
var item;
var mailbox;
var context;
function FavoriteBands() { }
function Band(bandname, musicalgenre) {
    this.Name = bandname;
    this.Genre = musicalgenre;
}

(function () {
    "use strict";
    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {

            //NOTE Initialize add-in

            app.initialize();
            
            item = Office.cast.item.toItemRead(Office.context.mailbox.item);
            mailbox = Office.context.mailbox;
            context = Office.context;

            try {
                $(".ms-NavBar").NavBar();
                $(".ms-Dropdown").Dropdown();
                
                $("#navBarCreateFolder").click(createFolder);
                $("#navBarCheckStorage").click(checkStorage);
                $("#navBarRetrieveStorage").click(retrieveStorage);
                $("#navBarUpdateStorage").click(updateStorage);
                $("#navBarClearSettings").click(clearSettings);
                $("#addArtist").click(addArtist);
                $("#clearArtists").click(clearArtists);
                $("#retrieveStorage").click(retrieveStorage);
                $("#updateStorage").click(updateStorage);
            } catch (e) {
                app.showNotification('Uh-oh!', 'Error initializing components');
            } 
            
            solutionStorage.initialize(context, app);

            $('#dialog-confirm-clearsettings').hide();
        });
    };
})();
var ewsCallbacks = (function () {
    "use strict";

    var createSolutionStorageFolderCallback = function (asyncResult){
        //Use arguments[0].asyncContext to get at userContext parameter value (if used in caller)
        
        var result = null;

        if (asyncResult === null) {
            app.showNotification('Error!', '[in createSolutionStorageFolderCallback]: null result');
            return 'error';
        }

        if (asyncResult.error !== null) {
            app.showNotification('Error!', '[in createSolutionStorageFolderCallback]: ' + asyncResult.error.message);
            return 'error';
        } else {

            try {
                var response = $.parseXML(asyncResult.value);
                var responseDOM = $(response);
                var prop;

                if (responseDOM) {

                    if (responseDOM) {
                        prop = responseDOM.filterNode("m:ResponseCode")[0];
                    }

                    if (!prop) {
                        app.showNotification('Error!', '[in createSolutionStorageFolderCallback]: Failed to parse response');
                        return 'error';
                    } else {
                        var errorType = prop.textContent;
                        if (prop.textContent === "ErrorFolderExists") {
                            //HIGH If folder exists, do what?
                            return 'error';
                        }

                        if (prop.textContent === "NoError") {
                            var foldersNode = null;
                            
                            foldersNode = responseDOM.filterNode("m:Folders")[0];

                            if (!foldersNode) {
                                app.showNotification('Error!', '[in createSolutionStorageFolderCallback]: Failed to retrieve folder data');
                                return 'error';
                            }
                            
                            var folderchildNodes;
                            try {

                                //NOTE Get ID for new solution storage folder
                                folderchildNodes = foldersNode.childNodes[0];
                                solutionStorage.solutionFolderID = folderchildNodes.childNodes.item("Folder").getAttribute("Id");
                                result = solutionStorage.solutionFolderID;

                            } catch (e) {
                                return 'error';
                            }
                        } else {
                            prop = responseDOM.filterNode("m:MessageText")[0];
                            app.showNotification('Error!', '[in createSolutionStorageFolderCallback]: ' + errorType + ": " + prop.textContent);
                            return 'error';
                        }
                    }
                }

            } catch (errorMsg) {                
                app.showNotification('Error!', '[in createSolutionStorageFolderCallback]: Failed to parse response (' + errorMsg + ')');
                return 'error';
            }
        }
        return result;
    };
    var createStorageItemCallback = function (asyncResult) {
        //Use arguments[0].asyncContext to get at userContext parameter value (if used in caller)
        
        var result = null;

        if (asyncResult === null) {
            app.showNotification('Error!', '[in createStorageItemCallback]: null result');
            result = 'error';
        }

        if (asyncResult.error !== null) {
            app.showNotification('Error!', '[in createStorageItemCallback]: ' + asyncResult.error.message);
            result = 'error';
        } else {

            var errorMsg;
            try {
                var response = $.parseXML(asyncResult.value);
                var responseDOM = $(response);
                var prop;

                if (responseDOM) {
                    if (responseDOM) {
                        prop = responseDOM.filterNode("m:ResponseCode")[0];
                    }

                    if (!prop) {
                        app.showNotification('Error!', '[in createStorageItemCallback]: Failed to parse response');
                        result = 'error';
                    } else {
                        if (prop.textContent === "NoError") {
                            var itemsNode = null;
                            var childNodesCnt;

                            itemsNode = responseDOM.filterNode("m:Items")[0];

                            if (!itemsNode) {
                                app.showNotification('Error!', '[in createStorageItemCallback]: Failed to retrieve item data');
                                result = 'error';
                            }
                            else
                            {
                                childNodesCnt = itemsNode.childElementCount;
                                for (var i = 0; i < childNodesCnt; i++) {
                                    var itemchildNodes;
                                    try {
                                        //Get ID for new solution storage message
                                        itemchildNodes = itemsNode.childNodes[i];
                                        solutionStorage.solutionStorageMessageID = itemchildNodes.childNodes.item("Item").getAttribute("Id");
                                        solutionStorage.saveSettings();
                                        result = solutionStorage.solutionStorageMessageID;
                                        break;
                                    } catch (e) {
                                        result = 'error';
                                    }
                                }    
                            }
                            
                        } else {
                            app.showNotification('Error!', '[in createStorageItemCallback]:' + prop.textContent);
                            result = 'error';
                        }
                    }
                }

            } catch (errorMsg) {                
                app.showNotification('Error!', '[in createStorageItemCallback]: Failed to parse response (' + errorMsg + ')');
                result = 'error';
            }
        }

        return result;
    };
    var getStorageItemCallback = function(asyncResult) {
        //Use arguments[0].asyncContext to get at userContext parameter value (if used in caller)

        var result = null;
        if (asyncResult === null) {
            app.showNotification('Error!', '[in getStorageItemCallback]: null result');
            return 'error';
        }

        if (asyncResult.error !== null) {
            app.showNotification('Error!', '[in getStorageItemCallback]: ' + asyncResult.error.message);
            return 'error';
        } else {

            var errorMsg;
            try {
                var response = $.parseXML(asyncResult.value);
                var responseDOM = $(response);
                var prop;

                if (responseDOM) {
                    if (responseDOM) {
                        prop = responseDOM.filterNode("m:ResponseCode")[0];
                    }

                    if (!prop) {
                        app.showNotification('Error!', '[in getStorageItemCallback]: Failed to parse response');
                        return 'error';

                    } else {
                        if (prop.textContent === "NoError") {
                            var bodyProp;
                            app.showNotification("Please wait...", "Retrieving XML from message body...");
                            try {
                                bodyProp = responseDOM.filterNode("t:Body")[0];

                                var xmlData = bodyProp.textContent;
                                var x2js = new X2JS();

                                solutionStorage.applicationData = new Object(); //ApplicationData()??                                
                                solutionStorage.applicationData = x2js.xml_str2json(xmlData);
                                $("#xmlText").prop('value', xmlData);
                                $("#numberOfBusinessObjects").prop('innerText', "#Business Objects in memory: " + solutionStorage.applicationData.FavoriteBands.Band.length);
                                result = 'success';

                            } catch (e) {
                                app.showNotification('Error!', '[in getStorageItemCallback(B)]:' + prop.textContent);
                                result = 'error';
                            }
                        } else {
                            app.showNotification('Error!', '[in getStorageItemCallback]:' + prop.textContent);
                            return 'error';
                        }
                    }
                }

            } catch (e) {
                errorMsg = e;
                app.showNotification('Error!', '[in getStorageItemCallback]: Failed to parse response (' + errorMsg + ')');
                return 'error';
            }
            app.showNotification("Ready to rock!", "XML data from solution storage has been loaded into memory as business objects.");
        }
        return result;
    };
    var updateFolderCallback = function(asyncResult) {
        //Use arguments[0].asyncContext to get at userContext parameter value (if used in caller)

        if (asyncResult === null) {
            app.showNotification('Error!', '[in updateFolderCallback]: null result');
            return 'error';
        }

        if (asyncResult.status === 'succeeded') {
            //NOTE Save ID of solution storage folder to add-in settings
            solutionStorage.saveSettings();
            app.showNotification("Done!", "Hidden folder '" + solutionStorage.solutionFolderName + "' created.");
            return 'success';        
        }

        if (asyncResult.error !== null) {
            app.showNotification('Error!', '[in updateFolderCallback]: ' + asyncResult.error.message);
            return 'error';
        }
        else {
            var errorMsg;
            try {
                var response = $.parseXML(asyncResult.value);
                var responseDOM = $(response);

                if (responseDOM) {
                    if (responseDOM) {
                        prop = responseDOM.filterNode("m:ResponseCode")[0];
                    }

                    if (!prop) {                        
                        app.showNotification('Error!', '[in updateFolderCallback]: Failed to parse response');            
                    }
                }
            } catch (e) {
                errorMsg = e;
                app.showNotification('Error!', '[in updateFolderCallback]: Failed to parse response (' + errorMsg + ')');
            }
        }

        return 'error';
    };
    var updateItemCallback = function(asyncResult) {
        //Use arguments[0].asyncContext to get at userContext parameter value (if used in caller)

        var result = null;

        if (asyncResult === null) {
            app.showNotification('Error!', '[in updateItemCallback]: null result');
            return 'error';
        }

        if (asyncResult.error !== null) {
            app.showNotification('Error!', '[in updateItemCallback]: ' + asyncResult.error.message);
            return 'error';
        }
        else {
            var prop = null;

            try {
                var response = $.parseXML(asyncResult.value);
                var responseDOM = $(response);

                if (responseDOM) {
                    if (responseDOM) {
                        prop = responseDOM.filterNode("m:ResponseCode")[0];
                    }

                    if (asyncResult.status === "succeeded") {                 
                        result = 'success';
                    }

                    //REVIEW Is prop eval needed?

                    if (!prop) {                        
                        app.showNotification('Error!', '[in updateItemCallback]: Failed to parse response');
                        return 'error';
                    } else {
                        if (prop.textContent === "NoError") {                            
                            result = 'success';
                        }
                        else {
                            app.showNotification('Error!', '[in updateItemCallback]:' + prop.textContent);
                            return 'error';
                        }
                    }
                }

            } catch (e) {
                app.showNotification('Error!', '[in updateItemCallback]: Failed to parse response (' + e + ')');
                return 'error';
            }
        }
        return result;
    };

    return {
  
        createSolutionStorageFolderCallback: createSolutionStorageFolderCallback,
        createStorageItemCallback: createStorageItemCallback,      
        getStorageItemCallback: getStorageItemCallback,
        updateFolderCallback: updateFolderCallback,
        updateItemCallback: updateItemCallback
    };
})();
var ewsRequests = (function () {

    "use strict";

    var getCreateSolutionStorageFolderRequest = function(folderName, isHidden) {
        var request;
        var distinguishedFolderId;

        //DistinguishedFolderId values: https://msdn.microsoft.com/en-us/library/office/aa580808(v=exchg.150).aspx
        //NOTE: Use root to create in visible folder at Mailbox root instead, msgfolderrot to create at Top of Information Store folder (visible folders with default folders at root)

        if (isHidden) {
            distinguishedFolderId = "root";
        } else {
            distinguishedFolderId = "msgfolderroot";
        }
        
        request = '<?xml version="1.0" encoding="utf-8"?> ' +
            '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" ' +
            '        xmlns:xsd="http://www.w3.org/2001/XMLSchema" ' +
            '        xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" ' +
            '        xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"> ' +
            '   <soap:Header>' +
            '       <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
            '   </soap:Header>' +
            '       <soap:Body>' +
            '          <CreateFolder xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
            '            <ParentFolderId>' +
            '              <t:DistinguishedFolderId Id="' + distinguishedFolderId + '"/>' +
            '            </ParentFolderId>' +
            '            <Folders>' +
            '               <t:Folder>' +
            '                   <t:DisplayName>' + folderName + '</t:DisplayName>' +
            '               </t:Folder>' +
            '            </Folders>' +
            '          </CreateFolder>' +
            '        </soap:Body>' +
            '</soap:Envelope>';

        return request;
    };
    var getUpdateItemRequest = function(solutionStorageMessageID, applicationData) {
        var request;
        var body;
        var xmlData;
        var x2js = new X2JS();

        xmlData = x2js.json2xml_str(applicationData);
        body = xmlData;
        body = htmlEncode(body);

        request = '<?xml version="1.0" encoding="utf-8"?>' +
            '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
            '               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
            '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
            '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
            '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
            '  <soap:Header>' +            
            '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types"/>' +
            '  </soap:Header>' +
            '  <soap:Body>' +
            '    <m:UpdateItem MessageDisposition="SaveOnly" ConflictResolution="AlwaysOverwrite">' +
            '      <m:ItemChanges>' +
            '       <t:ItemChange>' +            
            '       <t:ItemId Id="' + solutionStorageMessageID + '" />' +
            '       <t:Updates>' +
            '           <t:SetItemField>' +
            '               <t:FieldURI FieldURI="item:Body" />' +
            '               <t:Message>' +
            '                   <t:Body BodyType="Text">' + body + '</t:Body>' +
            '               </t:Message>' +
            '           </t:SetItemField>' +
            '       </t:Updates>' +
            '       </t:ItemChange>' +
            '      </m:ItemChanges>' +
            '    </m:UpdateItem>' +
            '  </soap:Body>' +
            '</soap:Envelope>';

        return request;
    };
    var getStorageItemDataRequest = function(itemid) {
        var request;

        request = '<?xml version="1.0" encoding="utf-8"?>' +
            '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
            '               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
            '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
            '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
            '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
            '  <soap:Header>' +
            '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
            '  </soap:Header>' +
            '  <soap:Body>' +
            '    <GetItem' +
            '                xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"' +
            '                xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
            '      <ItemShape>' +
            '        <t:BaseShape>Default</t:BaseShape>' +
            '        <t:BodyType>Text</t:BodyType>' +
            '           <t:AdditionalProperties>' +
            '               <t:FieldURI FieldURI="item:Body" />' +
            '           </t:AdditionalProperties>' +
            '      </ItemShape>' +
            '      <ItemIds>' +
            '        <t:ItemId Id="' + itemid + '"/>' +
            '      </ItemIds>' +
            '    </GetItem>' +
            '  </soap:Body>' +
            '</soap:Envelope>';

        return request;
    };
    var getCreateStorageItemRequest = function(solutionFolderID, applicationData) {
        var request;
        var body;
        var xmlData;
        var x2js = new X2JS();

        xmlData = x2js.json2xml_str(applicationData);
        body = xmlData;
        body = htmlEncode(body);

        request = '<?xml version="1.0" encoding="utf-8"?>' +
            '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
            '               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
            '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
            '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
            '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
            '  <soap:Header>' +
            '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
            '  </soap:Header>' +
            '  <soap:Body>' +
            '    <m:CreateItem MessageDisposition="SaveOnly">' +
            '      <m:SavedItemFolderId>' +
            '       <t:FolderId Id="' + solutionFolderID + '"/>' +
            '      </m:SavedItemFolderId>' +
            '      <m:Items>' +
            '        <t:Message>' +
            '          <t:ItemClass>IPM.Post</t:ItemClass>' +
            '          <t:Subject>MessageFiler Solution Storage</t:Subject>' +
            '          <t:Body BodyType="Text">' + body + '</t:Body>' +
            '        </t:Message>' +
            '      </m:Items>' +
            '    </m:CreateItem>' +
            '  </soap:Body>' +
            '</soap:Envelope>';

        console.log("getCreateStorageItemRequest leave");
        return request;
    };
    var getUpdateFolderRequest = function(id) {
        var request;

        request = '<?xml version="1.0" encoding="utf-8"?>' +
            '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
            '               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
            '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
            '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
            '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
            '  <soap:Header>' +
            '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
            '  </soap:Header>' +
            '       <soap:Body>' +
            '          <m:UpdateFolder>' +
            '            <m:FolderChanges>' +            
            '               <t:FolderChange>' +
            '                   <t:FolderId Id="' + id + '" />' +
            '                   <t:Updates>' +
            '                       <t:SetFolderField>' +
            '                           <t:ExtendedFieldURI PropertyTag="4340"' + ' PropertyType="Boolean" />' +
            '                           <t:Folder>' +
            '                               <t:ExtendedProperty>' +
            '                                   <t:ExtendedFieldURI PropertyTag="4340" PropertyType="Boolean" />' + '<t:Value>True</t:Value>' +
            '                               </t:ExtendedProperty>' +
            '                           </t:Folder>' +
            '                       </t:SetFolderField>' +
            '                   </t:Updates>' +
            '               </t:FolderChange>' +
            '            </m:FolderChanges>' +
            '          </m:UpdateFolder>' +
            '        </soap:Body>' +
            '</soap:Envelope>';

        return request;
    };

    return {
        getCreateSolutionStorageFolderRequest: getCreateSolutionStorageFolderRequest,
        getUpdateFolderRequest: getUpdateFolderRequest,
        getStorageItemDataRequest: getStorageItemDataRequest,
        getCreateStorageItemRequest: getCreateStorageItemRequest,
        getUpdateItemRequest: getUpdateItemRequest

    };
    
})();
var solutionStorage = (function () {
    "use strict";

    var applicationData;
    var mailbox;
    var roamingSettings;
    var solutionFolderID;
    var solutionFolderName;
    var solutionStorage = {};
    var solutionStorageMessageID;
    var usingHiddenFolder;

    applicationData = new ApplicationData();

    function ApplicationData(favoriteBands) {

        if (favoriteBands === undefined) {
            this.FavoriteBands = new FavoriteBands();
            this.FavoriteBands.Band = [];
        } else {
            this.FavoriteBands = favoriteBands;
        }
    }
    function Band(bandname, musicalgenre) {
        this.Name = bandname;
        this.Genre = musicalgenre;
    }
    function FavoriteBands() {
        
    }
    
    var clearSettings = function() {
        try {
            roamingSettings.remove("solutionFolderID");
            roamingSettings.remove("solutionStorageMessageID");
            roamingSettings.remove("solutionFolderName");
            roamingSettings.saveAsync(saveMyAddInSettingsCallback);
        } catch (e) {

        }
        $("#folderID").prop('value', "");
        $("#folderName").prop('value', "");
        $("#messageID").prop('value', "");
    }
    var createSolutionStorage = function (folderName, isHidden) {
        app.showNotification('Please wait...', 'Creating solution storage folder...');
        solutionStorage.usingHiddenFolder = isHidden;
        solutionStorage.solutionFolderName = folderName;

        //NOTE Resource for blog: http://learn.jquery.com/code-organization/deferreds/jquery-deferreds/
        $.when(
            mailbox.makeEwsRequestAsync(ewsRequests.getCreateSolutionStorageFolderRequest(folderName, isHidden), ewsCallbacks.createSolutionStorageFolderCallback)
            ).then(function () {
            if (solutionStorage.solutionFolderID) {
                return 'success';
            };
        }).fail(function() {
            console.log("oops");
        });  
    };
    var createStorageItem = function () {

        if (!solutionStorage.settingsLoaded === true) {
            getStorageIds();
        }

        if (solutionStorage.solutionFolderID === undefined) {
            app.showNotification("Uh-oh!", "The storage item folder hasn't been created yet. Click the 'Create Folder'  button.");
            return 'error';
        }

        mailbox.makeEwsRequestAsync(ewsRequests.getCreateStorageItemRequest(solutionStorage.solutionFolderID, solutionStorage.applicationData), ewsCallbacks.createStorageItemCallback);

        return solutionStorage.solutionStorageMessageID;

    };
    var getStorageIds = function() {
        try {
            solutionStorage.solutionFolderID = roamingSettings.get("solutionFolderID");
            solutionStorage.solutionStorageMessageID = roamingSettings.get("solutionStorageMessageID");
            solutionStorage.solutionFolderName = roamingSettings.get("solutionFolderName");
            solutionStorage.settingsLoaded = true;
        } catch (e) {

        }                
    }
    var getStorageItem = function () {
        if (!solutionStorage.settingsLoaded === true) {
            getStorageIds();
        }

        if (solutionStorage.solutionStorageMessageID === undefined) {
            app.showNotification("Uh-oh!", "The storage item hasn't been created yet. First create some Business Objects below, then click the 'Update Storage' button.");
            return 'error';
        }

        //REVIEW Use deferreds and return 'success' on getStorageItemCallback completion?

        mailbox.makeEwsRequestAsync(ewsRequests.getStorageItemDataRequest(solutionStorage.solutionStorageMessageID), ewsCallbacks.getStorageItemCallback);
        return 'success';
    };
    var saveMyAddInSettingsCallback = function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            // Handle the failure in asyncResult.error
            app.showNotification('Add-in Error:', asyncResult.error);
        }
    };
    var saveSettings = function () {
        try {
            roamingSettings.set("solutionFolderID", solutionStorage.solutionFolderID);
            roamingSettings.set("solutionStorageMessageID", solutionStorage.solutionStorageMessageID);
            roamingSettings.set("solutionFolderName", solutionStorage.solutionFolderName);
            roamingSettings.saveAsync(saveMyAddInSettingsCallback);
        } catch (e) {

        }

        if (solutionStorage.solutionFolderID !== null && solutionStorage.solutionFolderID !== undefined) 
            $("#folderID").prop('value', solutionStorage.solutionFolderID);
        if (solutionStorage.solutionFolderName !== null && solutionStorage.solutionFolderName !== undefined)
            $("#folderName").prop('value', solutionStorage.solutionFolderName);
        if (solutionStorage.solutionStorageMessageID !== null && solutionStorage.solutionStorageMessageID !== undefined)
            $("#messageID").prop('value', solutionStorage.solutionStorageMessageID);
    }
    var updateStorageItem = function () {
        
        if (solutionStorage.solutionStorageMessageID === undefined || solutionStorage.applicationData === undefined) {
            return 'error';
        }
        {
            mailbox.makeEwsRequestAsync(ewsRequests.getUpdateItemRequest(solutionStorage.solutionStorageMessageID, solutionStorage.applicationData), ewsCallbacks.updateItemCallback);
            return 'success';
        }        
    };

    solutionStorage.initialize = function(officeContext, appObject) {
        roamingSettings = officeContext.roamingSettings;
        mailbox = officeContext.mailbox;
        app = appObject;
    }

    solutionStorage.applicationData = applicationData;
    solutionStorage.clearSettings = clearSettings;
    solutionStorage.createSolutionStorage = createSolutionStorage;
    solutionStorage.createStorageItem = createStorageItem;
    solutionStorage.getStorageIds = getStorageIds;
    solutionStorage.getStorageItem = getStorageItem;
    solutionStorage.saveSettings = saveSettings;
    solutionStorage.solutionFolderID = solutionFolderID;
    solutionStorage.solutionFolderName = solutionFolderName;
    solutionStorage.solutionStorageMessageID = solutionStorageMessageID;
    solutionStorage.updateFolder = updateFolder;
    solutionStorage.updateStorageItem = updateStorageItem;
    solutionStorage.usingHiddenFolder = usingHiddenFolder;

    return solutionStorage;

})();

function addArtist() {
    var artistName = $("#artistName").prop('value');
    var genre = $("#genres").prop('value');
    
    if (artistName === undefined) {
        app.showNotification("Wait!", "You must enter an artist name.");        
        return;
    }
    if (genre === undefined) {
        app.showNotification("Wait!", "You must select a genre.");        
        return;
    }

    var artist = new Band(artistName, genre);
    var xmlData;
    var x2js = new X2JS();

    if (Array.isArray(solutionStorage.applicationData.FavoriteBands.Band)) {
        //Add band to bands
        solutionStorage.applicationData.FavoriteBands.Band.push(artist);
    } else {
        if (solutionStorage.applicationData.FavoriteBands.Band === "") {
            //For some reason the x2js.xml_str2json will set the root object as an empty string if there are no XML data nodes, so we need to initialize it as an array
            solutionStorage.applicationData.FavoriteBands.Band = [];
            solutionStorage.applicationData.FavoriteBands.Band.push(artist);
        } else {
            //There is one existing Band, but it is an Object and not an array; we need to change the Object to an Array, re-add the existing band and then add the second one. There's probably something I don't understand here...

            var firstFave = solutionStorage.applicationData.FavoriteBands.Band;
            solutionStorage.applicationData.FavoriteBands.Band = [];
            solutionStorage.applicationData.FavoriteBands.Band.push(firstFave);
            solutionStorage.applicationData.FavoriteBands.Band.push(artist);
        }        
    }   

    $("#numberOfBusinessObjects").prop('innerText', "#Business Objects in memory: " + solutionStorage.applicationData.FavoriteBands.Band.length);
    xmlData = x2js.json2xml_str(solutionStorage.applicationData);
    $("#xmlText").prop('value', xmlData); //NOTE Use .prop for dynamic attributes like checked, selected and value (http://api.jquery.com/prop/)    
}
function checkStorage() {
    solutionStorage.getStorageIds();
    try {

        if (solutionStorage.solutionFolderID === null || solutionStorage.solutionFolderID === undefined) {
            $("#folderID").prop('value', "not set");            
        } else {
            $("#folderID").prop('value', solutionStorage.solutionFolderID);
        }
        if (solutionStorage.solutionStorageMessageID === null || solutionStorage.solutionStorageMessageID === undefined) {
            $("#messageID").prop('value', "not set");
        } else {
            $("#messageID").prop('value', solutionStorage.solutionStorageMessageID);
        }
        
        if (solutionStorage.solutionFolderName === null || solutionStorage.solutionFolderName === undefined) {
            $("#folderName").prop('value', "not set");
        } else {
            $("#folderName").prop('value', solutionStorage.solutionFolderName);
        }
    } catch (e) {

    }

    app.showNotification("Done!", "Settings retrieved.");
}
function clearArtists() {
    if (solutionStorage.applicationData !== undefined) {        
        solutionStorage.applicationData.FavoriteBands.Band = []; //Reset the array
        $("#numberOfBusinessObjects").prop('innerText', "#Business Objects in memory: 0");
        $("#xmlText").prop('value', '');
    }
}
function clearSettings() {
    $(function () {
        $('#dialog-confirm-clearsettings').show();

        $("#dialog-confirm-clearsettings").dialog({
            resizable: false,
            height: "auto",
            modal: true,
            buttons: {
                "Yes": function () {
                    $(this).dialog("close");
                    solutionStorage.clearSettings();
                    app.showNotification("Done!", "Settings cleared.");
                },
                "No": function() {
                    $(this).dialog("close");
                }
            }
        });
    });    
}
function createFolder() {
    
    var folderName = $("#folderName").prop('value');
    var isHidden = $("#hiddenFolder").prop('checked');
    
    if (folderName === "") {
        app.showNotification("Wait!", "You must enter a folder name.");
        return;
    }        

    solutionStorage.clearSettings();
    var result;    
    
    //NOTE Using deferreds: http://learn.jquery.com/code-organization/deferreds/jquery-deferreds/
    $.when(
        result = solutionStorage.createSolutionStorage(folderName, isHidden)
    ).then(function () {
        solutionStorage.saveSettings();
        if (isHidden) {
            app.showNotification("Folder created!", "Now one more call to make it hidden...");
            result = solutionStorage.updateFolder();
        }        
    }).fail(function() {
        app.showNotification("Zoot alors!", "!!!");
        return;
    });

    //result = solutionStorage.createSolutionStorage(folderName, isHidden);
    if (result = 'success') {
        app.showNotification("Woo-hoo!", "Folder '" + solutionStorage.solutionFolderName + "' created. Now go ahead and start creating some business objects and then click the 'Update Storage' button to save that data to xml in a message within that folder.");
    }
}
function htmlEncode(value)
{
        //create a in-memory div, set it's inner text(which jQuery automatically encodes)
        //then grab the encoded contents back out.  The div never exists on the page.
        return $('<div/>').text(value).html();
    }
function htmlDecode(value)
{
    return $('<div/>').html(value).text();
}
function retrieveStorage() {
    app.showNotification("Please wait...", "Getting storage...");

    var result = solutionStorage.getStorageItem();
    if (result === 'success') {
        app.showNotification("Storage retrieved!", "BTW: how long have you been sitting at your computer? Get up and stretch!");
        window.location = "#xmlText";
    } else {
        app.showNotification("What the...", "ERROR: " + result);
    }
}
function updateFolder() {
    mailbox.makeEwsRequestAsync(ewsRequests.getUpdateFolderRequest(solutionStorage.solutionFolderID), ewsCallbacks.updateFolderCallback);
}
function updateStorage() {

    if (solutionStorage.applicationData.FavoriteBands === undefined) {
        app.showNotification("Uh-oh!", "There's nothing to store! Try creating some business objects first.");
        window.location("#artistName");
        return;
    }

    //NOTE First see if the storage item has been created yet; the id for the message should be stored in RoamingSettings and passed to solutionStorage.solutionStorageMessageID
    if (solutionStorage.solutionStorageMessageID === undefined) {
        app.showNotification("Please wait...", "We have to create the storage message first...");

        if (!solutionStorage.settingsLoaded === true) {
    getStorageIds();
        }
        if (solutionStorage.solutionFolderID === undefined) {
    app.showNotification("Uh-oh!", "The storage item folder hasn't been created yet. Click the 'Create Folder'  button.");
            return;
        }

        //NOTE Using deferreds: http://learn.jquery.com/code-organization/deferreds/jquery-deferreds/
        $.when(
            solutionStorage.createStorageItem()
        ).then(function() {    
            app.showNotification("Storage Item created!", "Now one more call to update the storage...");
            solutionStorage.updateStorageItem();
        }).fail(function() {
            app.showNotification("Zoot alors!", "!!!");
            return;
        });
    } else {
        app.showNotification("Please wait...", "Updating storage, feeding llamas, etc...");
        solutionStorage.updateStorageItem();
    }
    app.showNotification("Storage updated!", "P.S.: You look marvelous!");
}

// This function plug in filters nodes for the one that matches the given name.
// This sidesteps the issues in jquerys selector logic.
(function ($) {
    $.fn.filterNode = function (node) {
        return this.find("*").filter(function () {
            return this.nodeName === node;
        });
    };
})(jQuery);