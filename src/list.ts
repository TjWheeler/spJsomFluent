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
    /**
    * Create new list 
    * Example: create(context.get_web(), "My Task List", 107)
    * templateId - See: https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-csom/ee541191(v%3Doffice.15)#remarks
    */
    public create(web: SP.Web, listName: string, templateId: number): Fluent {
        this.fluent.chainAction(`${this._helperName}.create`, () => {
            return this.listHelper.createList(web, listName, templateId);
        });
        return this.fluent;
    }
    /**
    * 
    * Result: 
    * Example: 
    */
    public exists(web: SP.Web, listName: string): Fluent {
        this.fluent.chainAction(`${this._helperName}.exists`, () => {
            return this.listHelper.exists(web, listName);
        });
        return this.fluent;
    }
    /**
    * 
    * Result: 
    * Example: 
    */
    public delete(web: SP.Web, listName: string, template: string): Fluent {
        this.fluent.chainAction(`${this._helperName}.delete`, () => {
            return common.notImplementedPromise();
        });
        return this.fluent;
    }
    /**
    * 
    * Result: 
    * Example: 
    */
    public get(web: SP.Web, listName: string): Fluent {
        //TODO: test
        this.fluent.chainAction(`${this._helperName}.get`, () => {
            return this.listHelper.getList(web, listName);
        });
        return this.fluent;
    }
    /**
    * 
    * Result: 
    * Example: 
    */
    public addContentTypeListAssociation(web: SP.Web, listName: string, contentTypeName: string): Fluent {
        this.fluent.chainAction(`${this._helperName}.addContentTypeListAssociation`, () => {
            return this.listHelper.addContentTypeListAssociation(web, listName, contentTypeName);
        });
        return this.fluent;
    }
    /**
    * 
    * Result: 
    * Example: 
    */
    public removeContentTypeListAssociation(web: SP.Web, listName: string, contentTypeName: string): Fluent {
        this.fluent.chainAction(`${this._helperName}.removeContentTypeListAssociation`, () => {
            return this.listHelper.removeContentTypeListAssociation(web, listName, contentTypeName);
        });
        return this.fluent;
    }
    /**
    * 
    * Result: 
    * Example: 
    */
    public setDefaultValueOnList(web: SP.Web, listName: string, fieldInternalName: string, defaultValue: any): Fluent {
        this.fluent.chainAction(`${this._helperName}.setDefaultValueOnList`, () => {
            return this.listHelper.setDefaultValueOnList(web, listName, fieldInternalName, defaultValue);
        });
        return this.fluent;
    }
     /**
    * Enable email alerts on a list 
    * Example: setAlerts(context.get_web(), "My Task List", true)
    * Note: will not work for 2010/2013
    */
    public setAlerts(web: SP.Web, listName: string, enabled: boolean): Fluent {
        this.fluent.chainAction(`${this._helperName}.setAlerts`, () => {
            return this.listHelper.setAlerts(web, listName, enabled);
        });
        return this.fluent;
    }
}