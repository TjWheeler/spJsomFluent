import common from "../common"

export default class ListHelper {
    constructor(context: SP.ClientContext) {
        this.context = context;
    }
    private context: SP.ClientContext;
    public createList(web: SP.Web, listName: string, template: string): JQueryPromise<SP.List> {
        var deferred = $.Deferred();
        var list = null as SP.List;
        //TODO:
        common.reject(deferred, "Not implemented");
        deferred.resolve(list);
        return deferred.promise() as JQueryPromise<SP.List>;
    }
    public setListItemProperties(listItem: SP.ListItem, properties: any): JQueryPromise<SP.ListItem> {
        var deferred = $.Deferred();
        for (var propertyName in properties) {
            if (typeof (properties[propertyName]) !== typeof (undefined)) {
                listItem.set_item(propertyName, properties[propertyName]);
            }
        }
        listItem.update();
        this.context.load(listItem);
        common.executeQuery(this.context)
            .fail((sender, args) => { deferred.reject(sender, args); })
            .done(() => {
                deferred.resolve(listItem);
            });
        return deferred.promise() as JQueryPromise<SP.ListItem>;
    }
    public createListItemWithContentTypeName(web: SP.Web, listName: string, contentTypeName: string, properties: any): JQueryPromise<SP.ListItem> {
        var deferred = $.Deferred();
        var clientContext = this.context;
        var list = web.get_lists().getByTitle(listName) as SP.List;
        var itemCreateInfo = new SP.ListItemCreationInformation();

        var listContentTypes = list.get_contentTypes();
        clientContext.load(listContentTypes);
        common.executeQuery(clientContext)
            .fail((sender, args) => { deferred.reject(sender, args); })
            .done(() => {
                var contentTypeId = null;
                for (let i = 0; i < listContentTypes.get_count(); i++) {
                    let contentType = listContentTypes.getItemAtIndex(i) as SP.ContentType;
                    if (contentType.get_name() === contentTypeName) {
                        contentTypeId = contentType.get_id().get_stringValue();
                        break;
                    }
                }
                var listItem = list.addItem(itemCreateInfo);
                if (contentTypeId) {
                    listItem.set_item('ContentTypeId', contentTypeId);
                }
                for (var propertyName in properties) {
                    if (typeof (properties[propertyName]) !== typeof (undefined)) {
                        listItem.set_item(propertyName, properties[propertyName]);
                    }
                }
                listItem.update();
                clientContext.load(listItem);
                common.executeQuery(clientContext)
                    .fail((sender, args) => { deferred.reject(sender, args); })
                    .done(() => {
                        deferred.resolve(listItem);
                    });
            });
        return deferred.promise() as JQueryPromise<SP.ListItem>;

    }
    public createListItem(web: SP.Web, listName: string, properties: any): JQueryPromise<SP.ListItem> {
        var deferred = $.Deferred();
        var clientContext = this.context;
        var list = web.get_lists().getByTitle(listName) as SP.List;
        var itemCreateInfo = new SP.ListItemCreationInformation();
        var contentTypeId = null;
        var listItem = list.addItem(itemCreateInfo);
        //TODO: validate this works with People, taxonomy and lookup fields.
        for (var propertyName in properties) {
            if (typeof (properties[propertyName]) !== typeof (undefined)) {
                listItem.set_item(propertyName, properties[propertyName]);
            }
        }
        listItem.update();
        clientContext.load(listItem);
        common.executeQuery(clientContext)
            .fail((sender, args) => { deferred.reject(sender, args); })
            .done(() => {
                deferred.resolve(listItem);
            });
        return deferred.promise() as JQueryPromise<SP.ListItem>;
    }
    public getFileListItem(serverRelativeUrl: string): JQueryPromise<SP.ListItem> {
        var deferred = $.Deferred();
        var file = this.context.get_site().get_rootWeb().getFileByServerRelativeUrl(serverRelativeUrl);
        var listItem = file.get_listItemAllFields();
        this.context.load(listItem);
        common.executeQuery(this.context)
            .fail((sender, args) => { deferred.reject(sender, args); })
            .done(() => {
                deferred.resolve(listItem);
            });
        return deferred.promise() as JQueryPromise<SP.ListItem>;
    }
    public getFile(serverRelativeUrl: string): JQueryPromise<SP.File> {
        var deferred = $.Deferred();
        var file = this.context.get_site().get_rootWeb().getFileByServerRelativeUrl(serverRelativeUrl);
        this.context.load(file);
        common.executeQuery(this.context)
            .fail((sender, args) => { deferred.reject(sender, args); })
            .done(() => {
                deferred.resolve(file);
            });
        return deferred.promise() as JQueryPromise<SP.File>;
    }
    public checkInFile(web: SP.Web, serverRelativeUrl: string, comment: string, checkInType: SP.CheckinType): JQueryPromise<SP.File> {
        var deferred = $.Deferred();
        var file = web.getFileByServerRelativeUrl(serverRelativeUrl);
        this.context.load(file);
        file.checkIn(comment, checkInType);
        common.executeQuery(this.context)
            .fail((sender, args) => { deferred.reject(sender, args); })
            .done(() => {
                deferred.resolve(file);
            });
        return deferred.promise() as JQueryPromise<SP.File>;
    }
    public getList(web: SP.Web, listName: string): JQueryPromise<SP.List> {
        //TODO: test
        var deferred = $.Deferred();
        var list = web.get_lists().getByTitle(listName);
        this.context.load(list);
        common.executeQuery(this.context)
            .fail((sender, args) => { deferred.reject(sender, args); })
            .done(() => {
                deferred.resolve(list);
            });
        return deferred.promise() as JQueryPromise<SP.List>;
    }
    public getListItemById(web: SP.Web, listName: string, id: number): JQueryPromise<SP.ListItem> {
        //TODO: test
        var deferred = $.Deferred();
        var listItem = web.get_lists().getByTitle(listName).getItemById(id);
        this.context.load(listItem);
        common.executeQuery(this.context)
            .fail((sender, args) => { deferred.reject(sender, args); })
            .done(() => {
                deferred.resolve(listItem);
            });
        return deferred.promise() as JQueryPromise<SP.ListItem>;
    }

    public deleteListItemById(web: SP.Web, listName: string, id: number): JQueryPromise<any> {
        //TODO: test
        var deferred = $.Deferred();
        var listItem = web.get_lists().getByTitle(listName).getItemById(id);
        this.context.load(listItem);
        listItem.deleteObject(); //TODO: verify this works before executeQuery
        common.executeQuery(this.context)
            .fail((sender, args) => { deferred.reject(sender, args); })
            .done(() => {
                deferred.resolve();
            });
        return deferred.promise();
    }
    public getListItems(web: SP.Web, listName: string, viewXml: string): JQueryPromise<SP.ListItemCollection> {
        //TODO: test
        var deferred = $.Deferred();
        var list = web.get_lists().getByTitle(listName);
        var query = new SP.CamlQuery();
        query.set_viewXml(viewXml);
        var listItems = list.getItems(query);
        this.context.load(listItems);
        common.executeQuery(this.context)
            .fail((sender, args) => { deferred.reject(sender, args); })
            .done(() => {
                deferred.resolve(listItems);
            });
        return deferred.promise() as JQueryPromise<SP.ListItemCollection>;
    }
    public getListItemCount(web: SP.Web, listName: string, viewXml: string): JQueryPromise<any> {
        var deferred = $.Deferred();
        //TODO:
        common.reject(deferred, "Not implemented");
        return deferred.promise();
    }
    public addContentTypeListAssociation(web: SP.Web, listName: string, contentTypeName: string): JQueryPromise<SP.ContentType> {
        var deferred = $.Deferred();
        var list = web.get_lists().getByTitle(listName);
        
        var listCts = list.get_contentTypes();
        this.context.load(list);
        this.context.load(listCts);
        var contentTypes = this.context.get_site().get_rootWeb().get_contentTypes();
        var findContentType = (collection: SP.ContentTypeCollection, name: string) => {
            for (let i = 0; i < collection.get_count(); i++) {
                if (collection.getItemAtIndex(i).get_name().toLowerCase() === name.toLowerCase()) {
                    return collection.getItemAtIndex(i);
                }
            }
            return null;
        };
        this.context.load(contentTypes);
        common.executeQuery(this.context)
            .fail((sender, args) => { deferred.reject(sender, args); })
            .done(() => {
                let contentType = findContentType(contentTypes, contentTypeName) as SP.ContentType;
                if (!contentType) {
                    common.reject(deferred, `Content Type ${contentTypeName} not found`);
                    return;
                }
                //check if the CT is already associated
                let listCt = findContentType(listCts, contentTypeName) as SP.ContentType;
                if (listCt) {
                    //already associated
                    deferred.resolve(listCt);
                }
                else {
                    if (!list.get_contentTypesEnabled()) {
                        list.set_contentTypesEnabled(true);//enable custom cts on the list.
                    }
                    var associatedCt = list.get_contentTypes().addExistingContentType(contentType);
                    list.update();
                    common.executeQuery(this.context)
                        .fail((sender, args) => { deferred.reject(sender, args); })
                        .done(() => {
                            deferred.resolve(associatedCt);
                        });
                }
            });
        return deferred.promise() as JQueryPromise<SP.ContentType>;
    }
    public removeContentTypeListAssociation(web: SP.Web, listName: string, contentTypeName: string): JQueryPromise<any> {
        var deferred = $.Deferred();
        var list = web.get_lists().getByTitle(listName);
        var listCts = list.get_contentTypes();
        this.context.load(list);
        this.context.load(listCts);
        var findContentType = (collection: SP.ContentTypeCollection, name: string) => {
            for (let i = 0; i < collection.get_count(); i++) {
                if (collection.getItemAtIndex(i).get_name().toLowerCase() === name.toLowerCase()) {
                    return collection.getItemAtIndex(i);
                }
            }
            return null;
        };
        common.executeQuery(this.context)
            .fail((sender, args) => { deferred.reject(sender, args); })
            .done(() => {
                let listCt = findContentType(listCts, contentTypeName) as SP.ContentType;
                if (listCt) {
                    listCt.deleteObject();
                    list.update();
                    common.executeQuery(this.context)
                        .fail((sender, args) => { deferred.reject(sender, args); })
                        .done(() => {
                            deferred.resolve();
                        });
                }
                else {
                    deferred.resolve(); //not found, nothing to do.
                }
            });
        return deferred.promise();
    }
    public setDefaultValueOnList(web: SP.Web, listName: string, fieldInternalName: string, defaultValue: any): JQueryPromise<any> {
        var deferred = $.Deferred();
        var list = web.get_lists().getByTitle(listName);
        var fields = list.get_fields();
        this.context.load(list);
        this.context.load(fields);
        var findField = (collection: SP.FieldCollection, name: string) => {
            for (let i = 0; i < collection.get_count(); i++) {
                if (collection.getItemAtIndex(i).get_internalName().toLowerCase() === name.toLowerCase()) {
                    return collection.getItemAtIndex(i);
                }
            }
            return null;
        };
        common.executeQuery(this.context)
            .fail((sender, args) => { deferred.reject(sender, args); })
            .done(() => {
                let field = findField(fields, fieldInternalName) as SP.Field;
                if (field) {
                    field.set_defaultValue(defaultValue);
                    field.update()
                    common.executeQuery(this.context)
                        .fail((sender, args) => { deferred.reject(sender, args); })
                        .done(() => {
                            deferred.resolve();
                        });
                }
                else {
                    common.reject(deferred, `Field ${fieldInternalName} not found`);
                }
            });
        return deferred.promise();
    }
}