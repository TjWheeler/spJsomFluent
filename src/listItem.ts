import ListHelper from "./helper/listHelper";
import { Fluent } from "./fluent"
export default class ListItem {
    constructor(fluent: Fluent) {
        this.fluent = fluent;
        this.listHelper = new ListHelper(fluent.context);
    }
    private fluent = null as Fluent;
    private listHelper = null as ListHelper;
    private readonly _helperName: string = "listItem";
    public update(listItem: SP.ListItem, properties: any): Fluent {
        this.fluent.chainAction(`${this._helperName}.update`, () => {
            return this.listHelper.setListItemProperties(listItem, properties);
        });
        return this.fluent;
    }
    public createWithContentType(web: SP.Web, listName: string, contentTypeName: string, properties: any): Fluent {
        this.fluent.chainAction(`${this._helperName}.createWithContentType`, () => {
            return this.listHelper.createListItemWithContentTypeName(web, listName, contentTypeName, properties);
        });
        return this.fluent;
    }
    public create(web: SP.Web, listName: string, properties: any): Fluent {
        this.fluent.chainAction(`${this._helperName}.create`, () => {
            return this.listHelper.createListItem(web, listName, properties);
        });
        return this.fluent;
    }
    
    public getListItemById(web: SP.Web, listName: string, id: number): Fluent {
        //TODO: test
        this.fluent.chainAction(`${this._helperName}.getListItemById`, () => {
            return this.listHelper.getListItemById(web, listName, id);
        });
        return this.fluent;
    }
    public getFileListItem(serverRelativeUrl: string): Fluent {
        this.fluent.chainAction(`${this._helperName}.getFileListItem`, () => {
            return this.listHelper.getFileListItem(serverRelativeUrl);
        });
        return this.fluent;
    }
    public deleteById(web: SP.Web, listName: string, id: number): Fluent {
        //TODO: test
        this.fluent.chainAction(`${this._helperName}.deleteById`, () => {
            return this.listHelper.deleteListItemById(web, listName, id);
        });
        return this.fluent;
    }
    public query(web: SP.Web, listName: string, viewXml: string): Fluent {
        //TODO: test
        this.fluent.chainAction(`${this._helperName}.query`, () => {
            return this.listHelper.getListItems(web, listName, viewXml);
        });
        return this.fluent;
    }
    public count(web: SP.Web, listName: string, viewXml: string): Fluent {
        this.fluent.chainAction(`${this._helperName}.count`, () => {
            return this.listHelper.getListItemCount(web, listName, viewXml);
        });
        return this.fluent;
    }
}