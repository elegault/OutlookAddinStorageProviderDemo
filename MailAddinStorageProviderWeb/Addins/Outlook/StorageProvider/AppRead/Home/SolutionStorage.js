

function createSolutionStorageFolderCallback(asyncResult) {
    //Use ? arguments[0].asyncContext to get at userContext parameter value

    if (asyncResult == null) {
        app.showNotification('Error!', '[in createSolutionStorageFolderCallback]: null result');
        return;
    }

    if (asyncResult.error != null) {
        app.showNotification('Error!', '[in createSolutionStorageFolderCallback]: ' + asyncResult.error.message);
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
                    if (errorMsg) {
                        app.showNotification('Error!', '[in createSolutionStorageFolderCallback]: ' + errorMsg);
                    } else {
                        app.showNotification('Error!', '[in createSolutionStorageFolderCallback]: Failed to parse response');
                    }

                    return;
                } else {

                    if (prop.textContent == "ErrorFolderExists") {
                        //BUG IF folder exists, do what? Should only occur during testing??
                    }

                    if (prop.textContent == "NoError") {
                        var foldersNode = null;
                        var childNodesCnt;

                        foldersNode = responseDOM.filterNode("m:Folders")[0];

                        if (!foldersNode) {
                            app.showNotification('Error!', '[in createSolutionStorageFolderCallback]: Failed to retrieve folder data');
                            return;
                        }

                        childNodesCnt = foldersNode.childElementCount;
                        var folderchildNodes;
                        try {

                            //NOTE Load Step 3B: Get ID and ChangeKey for new solution storage folder
                            folderchildNodes = foldersNode.childNodes[0];
                            solutionFolderID = folderchildNodes.childNodes.item("Folder").getAttribute("Id");
                            solutionFolderChangeKey = folderchildNodes.childNodes.item("Folder").getAttribute("ChangeKey");

                            //NOTE TEST Load Step 3C: Make new solution storage folder hidden
                            //Do not update folder now if we are creating it at the Mailbox root
                            mailbox.makeEwsRequestAsync(getUpdateFolderRequest(solutionFolderID, solutionFolderChangeKey), updateFolderCallback);

                        } catch (e) {

                        }
                    }
                    else {
                        app.showNotification('Error!', '[in createSolutionStorageFolderCallback]:' + prop.textContent);
                    }
                }                
            }

        } catch (e) {
            errorMsg = e;
            app.showNotification('Error!', '[in createSolutionStorageFolderCallback]: Failed to parse response (' + errorMsg + ')');
        }
    }
}
function createStorageItemCallback(asyncResult) {
    //Use ? arguments[0].asyncContext to get at userContext parameter value

    app.hideNotification();
    var result = null;

    if (asyncResult == null) {
        app.showNotification('Error!', '[in createStorageItemCallback]: null result');
        return 'error';
    }

    if (asyncResult.error != null) {
        app.showNotification('Error!', '[in createStorageItemCallback]: ' + asyncResult.error.message);
        return 'error';
    }
    else {
        
        try {
            var response = $.parseXML(asyncResult.value);
            var responseDOM = $(response);
            var prop =  null;

            if (responseDOM) {
                if (responseDOM) {
                    prop = responseDOM.filterNode("m:ResponseCode")[0];
                }

                if (!prop) {
                    app.showNotification('Error!', '[in createStorageItemCallback]: Failed to parse response');
                    return 'error';
                    
                } else {
                    if (prop.textContent == "NoError") {
                        var itemsNode = null;
                        var childNodesCnt;

                        itemsNode = responseDOM.filterNode("m:Items")[0];

                        if (!itemsNode) {
                            app.showNotification('Error!', '[in createStorageItemCallback]: Failed to retrieve item data');
                            return 'error';
                        }

                        childNodesCnt = itemsNode.childElementCount;
                        for (var i = 0; i < childNodesCnt; i++) {
                            var itemchildNodes;
                            try {
                                //NOTE Get ID for new solution storage message
                                itemchildNodes = itemsNode.childNodes[i];
                                solutionStorageMessageID = itemchildNodes.childNodes.item("Item").getAttribute("Id");
                                solutionStorageMessageChangeKey = itemchildNodes.childNodes.item("Item").getAttribute("ChangeKey");
                                _settings.set("solutionStorageMessageID", solutionStorageMessageID);
                                _settings.set("solutionStorageMessageChangeKey", solutionStorageMessageChangeKey);
                                initialSetup = false;
                                _settings.saveAsync(saveMyAddInSettingsCallback);
                                result = 'success';                                
                                //HIGH Display CallOut to show how to use the search to get started
                                break;
                            }
                            catch (e) {

                            }
                        }
                    }
                    else {
                        app.showNotification('Error!', '[in createStorageItemCallback]:' + prop.textContent);
                        return 'error';
                    }
                }
            }

        } catch (e) {            
            app.showNotification('Error!', '[in createStorageItemCallback]: Failed to parse response (' + e + ')');
            return 'error';
        }
    }
    return result;
}
function findFolderCallback(asyncResult) {
    //Use ? arguments[0].asyncContext to get at userContext parameter value

    if (asyncResult == null) {
        app.showNotification('Error!', '[in findFolderCallback]: null result');
        return;
    }

    if (asyncResult.error != null) {
        app.showNotification('Error!', '[in findFolderCallback]: ' + asyncResult.error.message);
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
                    if (errorMsg) {
                        app.showNotification('Error!', '[in findFolderCallback]: ' + errorMsg);
                    } else {
                        app.showNotification('Error!', '[in findFolderCallback]: Failed to parse response');
                    }

                    return;
                } else {
                    if (prop.textContent == "NoError") {
                        var foldersNode = null;

                        foldersNode = responseDOM.filterNode("m:Folders")[0];

                        if (!foldersNode) {
                            app.showNotification('Error!', '[in findFolderCallback]: Failed to retrieve folder data');
                            return;
                        }

                        var folderchildNodes;
                        try {

                            folderchildNodes = foldersNode.childNodes[0];
                            var folderChangeKey = folderchildNodes.childNodes.item("Folder").getAttribute("ChangeKey");
                            var folder = new Folder();
                            folder = arguments[0].asyncContext;
                            moveEmail(folder);
                        } catch (e) {

                        }
                    }
                    else {
                        //TEST Prompt to reset storage
                        //HIGH Create dialog - canot use alert. message function in Mailbox??
                        //alert("The specific folder cannot be found. Please wait while we refresh our cache of the email folders in your Mailbox...");
                        loadFolders();
                        //app.showNotification('Error!', '[in findFolderCallback]:' + prop.textContent);
                    }
                }
            }

        } catch (e) {
            errorMsg = e;
            app.showNotification('Error!', '[in findFolderCallback]: Failed to parse response (' + errorMsg + ')');
        }
    }
}
function getCreateSolutionStorageFolderRequest() {
    var request;

    //DistinguishedFolderId values: https://msdn.microsoft.com/en-us/library/office/aa580808(v=exchg.150).aspx
    //NOTE: Folder isn't hidden when created in msgfolderroot and updated with hidden property. Use root instead
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
        '              <t:DistinguishedFolderId Id="msgfolderroot"/>' +
        '            </ParentFolderId>' +
        '            <Folders>' +
        '               <t:Folder>' +
        '                   <t:DisplayName>MessageFiler</t:DisplayName>' +
        '               </t:Folder>' +
        '            </Folders>' +                    
        '          </CreateFolder>' +
        '        </soap:Body>' +
        '</soap:Envelope>';

    return request;
}
function getCreateStorageItemRequest() {
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
        //'       <t:DistinguishedFolderId Id="drafts" />' +
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

    return request;
}
function getFindFolderRequest(folderID) {
    var request;

    //DistinguishedFolderId values: https://msdn.microsoft.com/en-us/library/office/aa580808(v=exchg.150).aspx

    request = '<?xml version="1.0" encoding="utf-8"?> ' +
        '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" ' +
        '        xmlns:xsd="http://www.w3.org/2001/XMLSchema" ' +
        '        xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" ' +
        '        xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"> ' +
        '   <soap:Header>' +
        '       <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
        '   </soap:Header>' +
        '       <soap:Body>' +
        '          <GetFolder xmlns="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
        '           <FolderShape>' +
        '               <t:BaseShape>Default</t:BaseShape>' +
        '           </FolderShape>' +
        '            <FolderIds>' +
        '               <t:FolderId Id="' + folderID + '"/> ' +
        '            </FolderIds>' +
        '          </GetFolder>' +
        '        </soap:Body>' +
        '</soap:Envelope>';

    return request;
}
function getStorageItemCallback(asyncResult) {
    //Use ? arguments[0].asyncContext to get at userContext parameter value

    disableSpinner();
    //app.hideNotification();

    if (asyncResult == null) {
        app.showNotification('Error!', '[in getStorageItemCallback]: null result');
        return;
    }

    if (asyncResult.error != null) {
        app.showNotification('Error!', '[in getStorageItemCallback]: ' + asyncResult.error.message);
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
                    if (errorMsg) {
                        app.showNotification('Error!', '[in getStorageItemCallback]: ' + errorMsg);
                    } else {
                        app.showNotification('Error!', '[in getStorageItemCallback]: Failed to parse response');
                    }

                    return;
                } else {
                    if (prop.textContent == "NoError") {
                        var bodyProp;
                        var bodyNode;

                        app.showNotification("Please wait...", "Loading the search box thingy with folders...");
                        try {
                            bodyProp = responseDOM.filterNode("t:Body")[0];
                            //bodyNode = bodyProp.childNodes[i];
                            //bodyContent = body.textContent;
                            //folderData = bodyNode.textContent;

                            //Switch to XML storage

                            folderData = bodyProp.textContent;
                            //folderData = htmlDecode(folderData); //TESTED Decoding folder data

                            var x2js = new X2JS();
                            applicationData = new Object();
                            applicationData = x2js.xml_str2json(folderData);

                            //var foldersRaw = folderData.split(";");
                            //folders = new Array(foldersRaw.length);

                            //for (var j = 0; j < foldersRaw.length; j++) {
                            //    var myFolderData = foldersRaw[j].split("|");
                            //    var myFolder = new FolderData(myFolderData[0], myFolderData[1]);
                            //    folders[j] = myFolder;
                            //}
                            populateFolderUI();
                            populateSuggestionsUI();
                            populateFavoritesUI();
                        } catch (e) {
                            app.showNotification('Error!', '[in getStorageItemCallback(B)]:' + prop.textContent);
                        }
                    }
                    else {
                        app.showNotification('Error!', '[in getStorageItemCallback]:' + prop.textContent);
                    }
                }
            }

        } catch (e) {
            errorMsg = e;
            app.showNotification('Error!', '[in getStorageItemCallback]: Failed to parse response (' + errorMsg + ')');
        }
        app.showNotification("Ready to rock your Inbox!", "Start typing in the search box to find a folder to file the email to");
    }
}
function getStorageItemDataRequest(item_id) {
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
        //'               <t:FieldURI FieldURI="item:MimeContent" />' +
        '               <t:FieldURI FieldURI="item:Body" />' +
        //'               <t:FieldURI FieldURI="item:Subject" />' +
        '           </t:AdditionalProperties>' +
        '      </ItemShape>' +
        '      <ItemIds>' +
        '        <t:ItemId Id="' + item_id + '"/>' +
        '      </ItemIds>' +
        '    </GetItem>' +
        '  </soap:Body>' +
        '</soap:Envelope>';


    return request;
}
function getUpdateFolderRequest(id, changeKey) {
    var request;

    //BUG Folder isn't hidden! [Tag:] 0x10f4, 0n4340. Microsoft.Exchange.WebServices.Data.MapiPropertyType.Boolean. PidTagAttributeHidden, PR_ATTR_HIDDEN, ptagAttrHidden, DAV:ishidden

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
        //'              <t:FolderChange><t:FolderId Id=' + id + ' ChangeKey=' + changeKey + '/>' +
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
}
function getUpdateItemRequest() {
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
        //'    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
        '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types"/>' +
        '  </soap:Header>' +
        '  <soap:Body>' +
        '    <m:UpdateItem MessageDisposition="SaveOnly" ConflictResolution="AlwaysOverwrite">' +
        '      <m:ItemChanges>' +
        '       <t:ItemChange>' +
        //'       <t:ItemId Id="' + solutionStorageMessageID + '" ChangeKey="' + solutionStorageMessageChangeKey + '"/>' +
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
}
function resetSettings() {

    //HIGH Add yes/no dialog to confirm resetting storage
    _settings.remove("solutionStorageMessageID");
    _settings.remove("solutionStorageMessageChangeKey");
    _settings.remove("solutionFolderID");
    //_settings.remove("folderData");
    _settings.saveAsync(saveMyAddInSettingsCallback);
}
function updateFolderCallback(asyncResult) {
    //Use ? arguments[0].asyncContext to get at userContext parameter value
    if (asyncResult == null) {
        app.showNotification('Error!', '[in updateFolderCallback]: null result');
        return;
    }

    if (asyncResult.status === 'succeeded') {
        //NOTE Load Step 3D: Save ID of solution storage folder to add-in settings
        _settings.set("solutionFolderID", solutionFolderID);
        _settings.saveAsync(saveMyAddInSettingsCallback);

        //Long request to load folders - set spinner in UI
        app.showNotification('Please wait...', 'Loading Mailbox folders...');
        mailbox.makeEwsRequestAsync(getAllFoldersRequest(), getFoldersCallback);
        return;
    }
    else {
        app.showNotification('Error!', '[in updateFolderCallback]:' + prop.textContent);
    }

    if (asyncResult.error != null) {
        app.showNotification('Error!', '[in updateFolderCallback]: ' + asyncResult.error.message);
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
                    if (errorMsg) {
                        app.showNotification('Error!', '[in updateFolderCallback]: ' + errorMsg);
                    } else {
                        app.showNotification('Error!', '[in updateFolderCallback]: Failed to parse response');
                    }
                    return;
                }
            }
        } catch (e) {
            errorMsg = e;
            app.showNotification('Error!', '[in updateFolderCallback]: Failed to parse response (' + errorMsg + ')');
        }
    }
}
function updateItemCallback(asyncResult) {
    //Use ? arguments[0].asyncContext to get at userContext parameter value

    //app.hideNotification();

    if (asyncResult == null) {
        app.showNotification('Error!', '[in updateItemCallback]: null result');
        return;
    }

    if (asyncResult.error != null) {
        app.showNotification('Error!', '[in updateItemCallback]: ' + asyncResult.error.message);
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

                if (asyncResult.status == "succeeded") {
                    if (arguments[0].asyncContext) {
                        mailbox.makeEwsRequestAsync(getItemDataRequest(item.itemId), getItemDataCallback, arguments[0].asyncContext);
                    } else {
                        //Reset folders was called if no param - we don't move the email in this situation
                    }
                    return;
                }

                if (!prop) {
                    if (errorMsg) {
                        app.showNotification('Error!', '[in updateItemCallback]: ' + errorMsg);
                    } else {
                        app.showNotification('Error!', '[in updateItemCallback]: Failed to parse response');
                    }

                    return;
                } else {
                    if (prop.textContent == "NoError") {
                        //HIGH Check to see if solution storage data was updated prior to a moveItem call, and call getItemDataCallBack -> moveItemCallBack
                        if (arguments[0].asyncContext) {
                            mailbox.makeEwsRequestAsync(getItemDataRequest(item.itemId), getItemDataCallback, arguments[0].asyncContext);
                        }
                    }
                    else {
                        app.showNotification('Error!', '[in updateFolderCallback]:' + prop.textContent);
                    }
                }
            }

        } catch (e) {
            errorMsg = e;
            app.showNotification('Error!', '[in updateItemCallback]: Failed to parse response (' + errorMsg + ')');
        }
    }
}