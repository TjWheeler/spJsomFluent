import common from "./common"
import ListHelper from "./helper/listHelper";
import { Fluent } from "./fluent"
export default class List {
    constructor(fluent: Fluent) {
        this.fluent = fluent;
        this.listHelper = new ListHelper(fluent.context);
    }
    private fluent = null as Fluent;
    private listHelper = null as ListHelper;
    private readonly _helperName: string = "list";
    public create(web: SP.Web, listName: string, template: string): Fluent {
        this.fluent.chainAction(`${this._helperName}.create`, () => {
            return this.listHelper.createList(web, listName, template);
        });
        return this.fluent;
    }
    public delete(web: SP.Web, listName: string, template: string): Fluent {
        this.fluent.chainAction(`${this._helperName}.delete`, () => {
            return common.notImplementedPromise();
        });
        return this.fluent;
    }
    
    public get(web: SP.Web, listName: string): Fluent {
        //TODO: test
        this.fluent.chainAction(`${this._helperName}.get`, () => {
            return this.listHelper.getList(web, listName);
        });
        return this.fluent;
    }
    
    public addContentTypeListAssociation(web: SP.Web, listName: string, contentTypeName: string): Fluent {
        this.fluent.chainAction(`${this._helperName}.addContentTypeListAssociation`, () => {
            return this.listHelper.addContentTypeListAssociation(web, listName, contentTypeName);
        });
        return this.fluent;
    }
    public removeContentTypeListAssociation(web: SP.Web, listName: string, contentTypeName: string): Fluent {
        this.fluent.chainAction(`${this._helperName}.removeContentTypeListAssociation`, () => {
            return this.listHelper.removeContentTypeListAssociation(web, listName, contentTypeName);
        });
        return this.fluent;
    }
    public setDefaultValueOnList(web: SP.Web, listName: string, fieldInternalName: string, defaultValue: any): Fluent {
        this.fluent.chainAction(`${this._helperName}.setDefaultValueOnList`, () => {
            return this.listHelper.setDefaultValueOnList(web, listName, fieldInternalName, defaultValue);
        });
        return this.fluent;
    }
}