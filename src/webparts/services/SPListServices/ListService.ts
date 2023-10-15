import "@pnp/polyfill-ie11";
//import "react-app-polyfill/ie11";

import { sp } from "@pnp/sp";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/folders";
import "@pnp/sp/attachments";
import "@pnp/sp/site-users/web";
import "@pnp/sp/files";

import { IList } from '@pnp/sp/lists/types';
import { IListService } from '../SPListServices';

export class ListService implements IListService {
    constructor(context: any) {
        sp.setup({
            ie11: true,
            spfxContext: context
        });
    }
    // Get List items by passing List title
    public GetItemsByListTitle = async (listTitle: string, selectFields: string, filter: string, expand: string, orderBy: string, orderByDir: boolean): Promise<any> => {
        return await this._getItemsByListTitle(listTitle, selectFields, filter, expand, orderBy, orderByDir);
    }
    // Get List items by passing List Id
    public GetItemsByListId = async (listId: string, selectFields: string, categoryFilter: string, expand: string, orderBy: string, orderByDir: boolean, top?: number): Promise<any> => {
        return await this._getItemsByListId(listId, selectFields, categoryFilter, expand, orderBy, orderByDir, top);
    }
    public GetListItemsByListId = async (listId: string): Promise<any> => {
        return await this._getListItemsByListId(listId);
    }

    public GetItemsByListIdAndFolder = async (listId: string, folderName: string): Promise<any> => {
        return await this._getItemsByListIdAndFolder(listId, folderName);
    }
    // Delete List item by Item Id
    public DeleteItemById = async (listTitle: string, itemId: number): Promise<any> => {
        return await this._deleteItemById(listTitle, itemId);
    }
    // Get List Fields by passing List Title
    public GetFieldsByListTitle = async (listTitle: string, selectFields: string, filter: string, expand: string, orderBy: string, orderByDir: boolean): Promise<any> => {
        return await this._getFieldsByListTitle(listTitle, selectFields, filter, expand, orderBy, orderByDir);
    }
    // Get List Fields by passing List Id
    public GetFieldsByListId = async (listId: string, selectFields: string, filter: string, expand: string, orderBy: string, orderByDir: boolean): Promise<any> => {
        return await this._getFieldsByListId(listId, selectFields, filter, expand, orderBy, orderByDir);
    }
    // Get all List details
    public GetAllLists = async (selectFields: string, filter: string, expand: string, orderBy: string, orderByDir: boolean): Promise<any> => {
        return await this._getAllLists(selectFields, filter, expand, orderBy, orderByDir);
    }
    // Get Latest item Id by passing List Title
    public GetLatestItemId = async (listTitle: string): Promise<any> => {
        return await this._getLatestItemId(listTitle);
    }
    // Create List Item
    public CreateListItem = async (listName: string, items: any): Promise<any> => {
        return await this._createListItem(listName, items);
    }

    //Attach files to list tiem

    public AddListAttachments = async (listId: string, itemId: number, attachmentItems: any): Promise<any> => {
        return await this._addListAttachments(listId, itemId, attachmentItems);
    }

    //Delete multiple file attachements 

    public DeleteListAttachments = async (listId: string, itemId: number, fileList: any): Promise<any> => {
        return await this._deleteListAttachments(listId, itemId, fileList);
    }
    // add batch items



    public AddMultipleItems = async (listId: string, items: any): Promise<any> => {
        return await this._addMultipleItems(listId, items);
    }


    public AddMultipleTransactionItems = async (listId: string, items: any, jvRequestId: string, ID: number, Year: string): Promise<any> => {
        return await this._addMultipleTransactionItems(listId, items, jvRequestId, ID, Year);
    }

    public AddMultipleTransactionItemsForEdit = async (listId: string, items: any): Promise<any> => {
        return await this._addMultipleTransactionItemsForEdit(listId, items);
    }

    //update Multiple Items 
    public UpdateMultipleTransactionItems = async (listId: string, items: any, IsDelete: boolean): Promise<any> => {
        return await this._updateMultipleTransactionItems(listId, items, IsDelete);
    }

    // Update List item
    public UpdateListItem = async (listName: string, items: any, listId: number): Promise<any> => {
        return await this._updateListItem(listName, items, listId);
    }

    public UpdateListItemByTitle = async (listName: string, items: any, listId: number): Promise<any> => {
        return await this._updateListItemByTitle(listName, items, listId);
    }

    // Get List item by passing List item id
    public GetListItembyItemId = async (listName: string, itemId: number, selectFields: string, expand: string): Promise<any> => {
        return await this._getListItemByItemId(listName, itemId, selectFields, expand);
    }

    public GetUserID = async (email: string): Promise<any> => {
        return await this._getUserId(email);
    }

    public FileUploadtoLibrary = async (serverRelativeURL: string, file: any): Promise<any> => {
        return await this._fileUploda(serverRelativeURL, file);
    }

    public GetListDetailsbyID = async (listID: string, select: string, expand: string): Promise<any> => {
        return await this._getListDetailsbyID(listID, select, expand);
    }

    public GetFileBuffer = async (fileURL): Promise<any> => {
        return await this._getFileBuffer(fileURL);
    }

    public CreateFolder = async (libName: string, folderName: string): Promise<any> => {
        return await this._createFolder(libName, folderName);
    }
    public UploadFiletoLibrary = async (serverrelativeURL: string, fileData: any): Promise<any> => {
        return await this._uploadFiletoLibrary(serverrelativeURL, fileData);
    }

    public CopyFiletoLibrary = async (srcPath: string, destPath: string, name: string): Promise<any> => {
        return await this._copyFiletoLibrary(srcPath, destPath, name)
    }


    public DeleteFile = async (serverrelativeURL: string): Promise<any> => {
        return await this._deleteFile(serverrelativeURL);
    }

    public GetFilesfromFolder = async (serverrelativeURL: string): Promise<any> => {
        return await this._getFilesfromFolder(serverrelativeURL);
    }

    public GetFileMetadata = async (Name: string, serverrelativeURL: string): Promise<any> => {
        return await this._getFileMetadata(Name, serverrelativeURL);
    }


    public GetfileFromLibrary = async (folderPath: string, select: string, filter: string): Promise<any> => {
        return await this._getfileFromLibrary(folderPath, select, filter);
    }

    public GetOtherSiteItemsByListId = async (siteURL: string, listId: string, selectFields: string, categoryFilter: string, expand: string, orderBy: string, orderByDir: boolean, top?: number): Promise<any> => {
        return await this._getOtherSiteItemsByListId(siteURL, listId, selectFields, categoryFilter, expand, orderBy, orderByDir, top);
    }
    //_getOtherSiteListItemsByListId

    private _getItemsByListTitle = async (listTitle: string, selectFields: string, filter: string, expand: string, orderBy: string, orderByDir: boolean): Promise<any> => {
        let _listItems: any = null;
        let filterCatVal: string = filter.length == 0 ? "" : filter;
        let expandVal: string = expand.length == 0 ? "" : expand;
        let orderByVal: string = orderBy.length == 0 ? "Created" : orderBy;
        let selectFieldsVal: string = selectFields.length == 0 ? "" : selectFields;
        try {
            _listItems = await sp.web.lists.getByTitle(listTitle).items
                .select(selectFieldsVal)
                .expand(expandVal)
                .filter(filterCatVal)
                .top(5000)
                .orderBy(orderByVal, orderByDir)
                .get();
        }
        catch (err) {
            console.log(err);
            return null;
        }
        return _listItems;

    }

    private _getListItemsByListId = async (listId: string): Promise<any> => {
        let _listItems: any = null;
        try {
            _listItems = await sp.web.lists.getById(listId).items
                .get();
        }
        catch (err) {
            console.log(err);
            return null;
        }
        return _listItems;
    }




    private _getItemsByListId = async (listId: string, selectFields: string, categoryFilter: string, expand: string, orderBy: string, orderByDir: boolean, top?: number): Promise<any> => {
        let _listItems: any = null;
        let filterCatVal: string = categoryFilter.length == 0 ? "" : categoryFilter;
        let expandVal: string = expand.length == 0 ? "" : expand;
        let orderByVal: string = orderBy.length == 0 ? "Created" : orderBy;
        let selectFieldsVal: string = selectFields.length == 0 ? "" : selectFields;
        let topRecords = top ? top : 5000;
        try {
            _listItems = await sp.web.lists.getById(listId).items
                .select(selectFieldsVal)
                .expand(expandVal)
                .filter(filterCatVal)
                .top(topRecords)
                .orderBy(orderByVal, orderByDir)
                .get();
        }
        catch (err) {
            console.log(err);
            return null;
        }
        return _listItems;
    }
    private _getItemsByListIdAndFolder = async (listId: string, folderName: string): Promise<any> => {
        let _listItems: any = null;
        try {
            _listItems = await sp.web.getFolderByServerRelativeUrl(folderName).getItem();
        }
        catch (err) {
            console.log(err);
            return null;
        }
        return _listItems;
    }

    private _deleteItemById = async (listTitle: string, itemId: number): Promise<any> => {

        let deletedItems: any = null;
        try {
            deletedItems = await sp.web.lists.getByTitle(listTitle).items
                .getById(itemId)
                .delete();
        }
        catch (err) {
            console.log(err);
            return null;
        }
        return deletedItems;
    }

    private _getFieldsByListTitle = async (listTitle: string, selectFields: string, filter: string, expand: string, orderBy: string, orderByDir: boolean): Promise<any> => {
        let listFields: any = null;
        let filterCatVal: string = filter.length == 0 ? "Hidden eq false" : filter;
        let expandVal: string = expand.length == 0 ? "" : expand;
        let orderByVal: string = orderBy.length == 0 ? "Title" : orderBy;
        try {
            listFields = await sp.web.lists.getByTitle(listTitle).fields
                .select(selectFields)
                .expand(expandVal)
                .filter(filterCatVal)
                .orderBy(orderByVal, orderByDir)
                .get();
        } catch (err) {
            console.log(err);
            return null;
        }
        return listFields;
    }

    private _getFieldsByListId = async (listId: string, selectFields: string, filter: string, expand: string, orderBy: string, orderByDir: boolean): Promise<any> => {
        let listFields: any = null;
        let filterCatVal: string = filter.length == 0 ? "Hidden eq false" : filter;
        let expandVal: string = expand.length == 0 ? "" : expand;
        let orderByVal: string = orderBy.length == 0 ? "Title" : orderBy;
        try {
            listFields = await sp.web.lists.getById(listId).fields
                .select(selectFields)
                .expand(expandVal)
                .filter(filterCatVal)
                .orderBy(orderByVal, orderByDir)
                .get();
        } catch (err) {
            console.log(err);
            return null;
        }
        return listFields;
    }


    private _getAllLists = async (selectFields: string, filter: string, expand: string, orderBy: string, orderByDir: boolean): Promise<any> => {
        let allLists: any = null;
        let filterCatVal: string = filter.length == 0 ? "Hidden eq false" : filter;
        let expandVal: string = expand.length == 0 ? "" : expand;
        let orderByVal: string = orderBy.length == 0 ? "Created" : orderBy;
        try {
            allLists = await sp.web.lists
                .select(selectFields)
                .expand(expandVal)
                .filter(filterCatVal)
                .orderBy(orderByVal, orderByDir)
                .get();
        }
        catch (err) {
            console.log(err);
            return null;
        }
        return allLists;
    }

    private _getLatestItemId = async (listTitle: string): Promise<any> => {
        let latestItemId: any = null;
        try {
            latestItemId = await sp.web.lists.getByTitle(listTitle).items.orderBy('Id', false).top(1).select('Id').get();
        }
        catch (err) {
            console.log(err);
            return null;
        }
        return latestItemId;
    }

    private _createListItem = async (listName: string, items: any): Promise<any> => {
        let listItemsRes: any = null;
        try {
            listItemsRes = await sp.web.lists.getById(listName).items.add(
                items
            );
        }
        catch (err) {
            console.log(err);
            return null;
        }
        return listItemsRes;
    }

    private _addListAttachments = async (listId: string, itemId: number, attachmentItems: any): Promise<any> => {
        let listItemsRes: any = null;
        try {
            const list: IList = sp.web.lists.getById(listId);

            listItemsRes = await list.items.getById(itemId).attachmentFiles.addMultiple(attachmentItems);
        }
        catch (err) {
            console.log(err);
            return null;
        }
        return listItemsRes;
    }

    private _deleteListAttachments = async (listId: string, itemId: number, fileList: any): Promise<any> => {
        let listItemsRes: any = null;
        try {

            const list: IList = sp.web.lists.getById(listId);

            if (fileList.length > 0) {

                let attachItems = fileList.map(a => a);
                listItemsRes = await list.items.getById(itemId).attachmentFiles.deleteMultiple(...attachItems);
            }
        }
        catch (err) {

            console.log(err);
            return null;
        }
        return listItemsRes;
    }

    public _updateListItem = async (listName: string, items: any, listId: number): Promise<any> => {
        let listItemsRes: any = null;
        try {
            listItemsRes = await sp.web.lists.getById(listName).items.getById(listId).update(
                items
            );
        }
        catch (err) {
            console.log(err);
            return null;
        }
        return listItemsRes;
    }

    public _updateListItemByTitle = async (listName: string, items: any, listId: number): Promise<any> => {
        let listItemsRes: any = null;
        try {
            listItemsRes = await sp.web.lists.getByTitle(listName).items.getById(listId).update(
                items
            );
        }
        catch (err) {
            console.log(err);
            return null;
        }
        return listItemsRes;
    }


    public _addMultipleTransactionItems = async (listId: string, items: any, jvRequestId: string, ID: number, Year: string): Promise<any> => {
        let listItemsRes: any = null;
        try {
            let list = sp.web.lists.getById(listId);
            const entityTypeFullName = await list.getListItemEntityTypeFullName();
            let batch = sp.web.createBatch();
            if (items.length > 0) {
                items.forEach((element, i) => {
                    let transacationUniqueId = "";
                    i++;
                    if (i.toString().length == 1) {
                        transacationUniqueId = jvRequestId + "-0" + i.toString();
                    }
                    else {
                        transacationUniqueId = jvRequestId + "-" + i.toString();
                    }
                    let updateTransactionDetails = {};
                    updateTransactionDetails["JV_Req_IDId"] = ID;
                    updateTransactionDetails["Transaction_Name"] = element.TransactionName;
                    updateTransactionDetails["Posting_Type"] = element.PostingType;
                    updateTransactionDetails["IsExternalTransaction"] = element.IsExternalTransaction;
                    updateTransactionDetails["Debit_Account_Number"] = element.DebitAccNumber;
                    updateTransactionDetails["Credit_Account_Number"] = element.CreditAccNumber;
                    updateTransactionDetails["Amount"] = element.Amount.replace(/,/g, '');
                    updateTransactionDetails["Transaction_Code"] = transacationUniqueId;
                    if (element.TransactionDescription) {
                        updateTransactionDetails["Transaction_Description"] = element.TransactionDescription;
                    }
                    updateTransactionDetails["Year"] = Year;
                    list.items.inBatch(batch).add(updateTransactionDetails, entityTypeFullName).then(b => { console.log(b); });

                });
                listItemsRes = await batch.execute();
            }
        }
        catch (err) {
            console.log(err);
            return null;
        }
        return listItemsRes;
    }

    public _addMultipleTransactionItemsForEdit = async (listId: string, items: any): Promise<any> => {
        let listItemsRes: any = null;
        try {
            let list = sp.web.lists.getById(listId);
            const entityTypeFullName = await list.getListItemEntityTypeFullName();
            let batch = sp.web.createBatch();
            if (items.length > 0) {
                items.forEach((element) => {

                    list.items.inBatch(batch).add(element, entityTypeFullName).then(b => { console.log(b); });

                });
                listItemsRes = await batch.execute();
            }
        }
        catch (err) {
            console.log(err);
            return null;
        }
        return listItemsRes;
    }

    public _updateMultipleTransactionItems = async (listId: string, items: any, IsDelete: boolean): Promise<any> => {
        let listItemsRes: any = null;
        try {
            let list = sp.web.lists.getById(listId);
            const entityTypeFullName = await list.getListItemEntityTypeFullName();
            let batch = sp.web.createBatch();
            if (items.length > 0) {
                items.forEach((element, i) => {

                    let updateTransactionDetails = {};
                    if (!IsDelete) {

                        updateTransactionDetails["Transaction_Name"] = element.TransactionName;
                        updateTransactionDetails["Posting_Type"] = element.PostingType;
                        updateTransactionDetails["IsExternalTransaction"] = element.IsExternalTransaction;
                        updateTransactionDetails["Debit_Account_Number"] = element.DebitAccNumber;
                        updateTransactionDetails["Credit_Account_Number"] = element.CreditAccNumber;
                        updateTransactionDetails["Amount"] = element.Amount;
                        if (element.TransactionDescription) {
                            updateTransactionDetails["Transaction_Description"] = element.TransactionDescription;
                        }
                    }
                    else {
                        updateTransactionDetails["IsActive"] = false;
                    }
                    list.items.getById(element.Id).inBatch(batch).update(updateTransactionDetails, "*", entityTypeFullName).then(b => { console.log(b); });

                });
                listItemsRes = await batch.execute();
            }
        }
        catch (err) {
            console.log(err);
            return null;
        }
        return listItemsRes;
    }

    private _getListItemByItemId = async (listTitle: string, itemId: number, selectFields: string, expand: string): Promise<any> => {
        let expandVal: string = expand.length == 0 ? "" : expand;
        let Items: any = null;
        try {
            Items = await sp.web.lists.getByTitle(listTitle).items
                .getById(itemId)
                .select(selectFields)
                .expand(expandVal)
                .get();
        }
        catch (err) {
            console.log(err);
            return null;
        }
        return Items;
    }

    public sortByOrder(x: any, y: any) {
        if (x.EWB_DisplayOrder && y.EWB_DisplayOrder) {
            var xVal = parseInt(x.EWB_DisplayOrder) || 0;
            var yVal = parseInt(y.EWB_DisplayOrder) || 0;
            return xVal - yVal;
        }
        else if (x.EWB_DisplayOrder) {
            return -x.EWB_DisplayOrder;
        }
        else if (y.EWB_DisplayOrder) {
            return y.EWB_DisplayOrder;
        }
        return 0;
    }
    public recursiveSort(arr: any[]) {
        arr.sort(this.sortByOrder);
        arr.forEach((item) => {
            if (item.children && Array.isArray(item.children)) {
                this.recursiveSort(item.children);
            }
        });
        return arr;
    }

    public _getUserId = async (email: string): Promise<any> => {
        let Items: any = null;
        try {
            Items = await sp.web.siteUsers.getByEmail(email).get();
        }
        catch (err) {
            console.log(err);
            return null;
        }
        return Items;
    }


    public _addMultipleItems = async (listId: string, items: any): Promise<any> => {
        let listItemsRes: any = [];
        try {
            let list = sp.web.lists.getById(listId);
            const entityTypeFullName = await list.getListItemEntityTypeFullName();
            let batch = sp.web.createBatch();
            if (items.length > 0) {
                items.forEach((element, i) => {
                    list.items.inBatch(batch).add(element, entityTypeFullName).then(b => {
                        console.log(b);
                        listItemsRes.push(b.data.ID);
                    });
                });
                await batch.execute();
            }
        }
        catch (err) {
            console.log(err);
            return null;
        }
        return listItemsRes;
    }

    public _fileUploda = (serverRelativeURL: string, file: any): Promise<any> => {
        let item: any = null;
        if (file.size <= 10485760) {
            sp.web.getFolderByServerRelativeUrl(serverRelativeURL).files.add(file.name, file, true).then(f => {
                console.log("File Uploaded");
                f.file.getItem().then(item => {
                    item.update({
                        Title: "Metadata Updated"
                    }).then((myupdate) => {
                        console.log(myupdate);
                        console.log("Metadata Updated");
                    });
                });
            });
        }
        else {
            sp.web.getFolderByServerRelativeUrl(serverRelativeURL)
                .files.addChunked(file.name, file)
                .then(({ file }) => file.getItem()).then((item: any) => {
                    console.log("File Uploaded");
                    return item.update({
                        Title: 'Metadata Updated'
                    }).then((myupdate) => {
                        console.log(myupdate);
                        console.log("Metadata Updated");
                    });
                }).catch(console.log);
        }

        return item;
    }

    public _getListDetailsbyID = (listID: string, select: string, expand: string): Promise<any> => {
        let item: any = null;
        item = sp.web.lists.getById(listID).select(select).expand(expand).get();
        return item;
    }

    public _getFileBuffer = async (fileURL: string): Promise<any> => {
        let filebuffer = null;
        try {
            filebuffer = await sp.web.getFileByServerRelativeUrl(fileURL).getBuffer().then((buffer: ArrayBuffer) => {
                return btoa(String.fromCharCode.apply(null, new Uint8Array(buffer)));
            });

        }
        catch (error) {
            console.log("Error in file Reading" + error)
        }

        return filebuffer;
    }



    public _createFolder = async (libName: string, folderName: string): Promise<any> => {
        let filePath = await sp.web.folders.getByName(libName).folders.add(folderName)
        return filePath.data.ServerRelativeUrl;
    }

    public _uploadFiletoLibrary = async (serverrelativeURL: string, fileData: any): Promise<any> => {
        let metaData = fileData.metaData;
        sp.web.getFolderByServerRelativeUrl(serverrelativeURL).files.add(fileData.file.name, fileData.file.content, false).then(f => {
            console.log("File Uploaded");
            f.file.getItem().then(item => {
                item.update(
                    metaData
                ).then((myupdate) => {
                    console.log(myupdate);
                    console.log("Metadata Updated");
                });
            });
        });

    }

    public _copyFiletoLibrary = async (srcPath: string, destPath: string, name: string): Promise<any> => {
        //let metaData = fileData.metaData;

        await sp.web.getFileByServerRelativePath(srcPath).copyByPath(`${destPath}/${name}`, false, false);
        // .then(f => {
        //     console.log("File Uploaded");
        //     f.file.getItem().then(item => {
        //         item.update(
        //             metaData
        //         ).then((myupdate) => {
        //             console.log(myupdate);
        //             console.log("Metadata Updated");
        //         });
        //     });
        // });

    }



    public _deleteFile = async (serverrelavtiveURL: string): Promise<any> => {
        await sp.web.getFileByServerRelativeUrl(serverrelavtiveURL).delete();
    }

    public _getFilesfromFolder = async (serverrelavtiveURL: string): Promise<any> => {
        let allFiles: any = [];
        let files: any = null;
        try {
            files = await sp.web.getFolderByServerRelativePath(serverrelavtiveURL).files();

            // if (files.length > 0) {
            //     files.map(async file => {
            //         let fileContent = await this._getFileMetadata(file.ServerRelativeUrl);
            //         allFiles.push(fileContent);
            //     });

            //     return allFiles;
            // }
        } catch (err) {
            console.log(err);
            return null;
        }

        return files;

    }

    public _getFileMetadata = async (Name: string, serverrelavtiveURL: string): Promise<any> => {
        let fileData: any = [];
        let file = null;
        fileData = await sp.web.getFolderByServerRelativePath(serverrelavtiveURL).listItemAllFields();
        if (fileData) {
            let file = {
                ServerRelativeUrl: serverrelavtiveURL,
                Title: fileData.Title,
                Company: fileData.Company,
                Index: fileData.Index,
                Created: fileData.Created,
                FileType: fileData.FileType,
                UserTitle: fileData.UserTitle,
                PaymentOption: fileData.PaymentOption,
                Name: Name

            }
            return file;
        }


    }



    private _getOtherSiteItemsByListId = async (siteURL: string, listId: string, selectFields: string, categoryFilter: string, expand: string, orderBy: string, orderByDir: boolean, top?: number): Promise<any> => {
        let _listItems: any = null;

        let newWeb = Web(siteURL);

        let filterCatVal: string = categoryFilter.length == 0 ? "" : categoryFilter;
        let expandVal: string = expand.length == 0 ? "" : expand;
        let orderByVal: string = orderBy.length == 0 ? "Created" : orderBy;
        let selectFieldsVal: string = selectFields.length == 0 ? "" : selectFields;
        let topRecords = top ? top : 5000;
        try {
            _listItems = await newWeb.lists.getById(listId).items
                .select(selectFieldsVal)
                .expand(expandVal)
                .filter(filterCatVal)
                .top(topRecords)
                .orderBy(orderByVal, orderByDir)
                .get();
        }
        catch (err) {
            console.log(err);
            return null;
        }
        return _listItems;
    }

    public _getfileFromLibrary = async (folderPath: string, selectFields: string, filter: string): Promise<any> => {
        let _files: any = null;
        let filterCatVal: string = filter.length == 0 ? "" : filter;
        let selectFieldsVal: string = selectFields.length == 0 ? "" : selectFields;
        _files = await sp.web.getFolderByServerRelativeUrl(folderPath).files
        .select(selectFields)
        .expand("ListItemAllFields")
        .filter(filterCatVal)
        .get();

        return _files;

    }



}