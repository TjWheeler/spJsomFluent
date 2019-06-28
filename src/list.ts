import { ListHelper } from "./helper/listHelper";
import { Fluent } from "./fluent"
export class List {
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
        return this.fluent.chainAction(`${this._helperName}.create`, () => {
            return this.listHelper.createList(web, listName, templateId);
        });
        
    }
    /**
    * Check if the list exists
    * Result: boolean
    * Example: exists(context.get_web(), "My List")
    */
    public exists(web: SP.Web, listName: string): Fluent {
        return this.fluent.chainAction(`${this._helperName}.exists`, () => {
            return this.listHelper.exists(web, listName);
        });
        
    }
    /**
    * Delete a list
    * Result: void
    * Example: delete(context.get_web(), "My List")
    */
    public delete(web: SP.Web, listName: string): Fluent {
        return this.fluent.chainAction(`${this._helperName}.delete`, () => {
            return this.listHelper.deleteList(web, listName);
        });
        
    }
    /**
    * Get a list
    * Result: SP.List
    * Example: get(context.get_web(), "My List")
    */
    public get(web: SP.Web, listName: string): Fluent {
        //TODO: test
        return this.fluent.chainAction(`${this._helperName}.get`, () => {
            return this.listHelper.getList(web, listName);
        });
        
    }
    /**
    * Adds a Content Type to a List.  Resolves the new associated list content type.
    * Result: SP.ContentType
    * Example: addContentTypeListAssociation(context.get_web(), "My List", "My Content Type Name")
    */
    public addContentTypeListAssociation(web: SP.Web, listName: string, contentTypeName: string): Fluent {
        return this.fluent.chainAction(`${this._helperName}.addContentTypeListAssociation`, () => {
            return this.listHelper.addContentTypeListAssociation(web, listName, contentTypeName);
        });
        
    }
    /**
    * Removes a Content Type associated from a list.
    * Result: void
    * Example: removeContentTypeListAssociation(context.get_web(), "My List", "My Content Type Name")
    */
    public removeContentTypeListAssociation(web: SP.Web, listName: string, contentTypeName: string): Fluent {
        return this.fluent.chainAction(`${this._helperName}.removeContentTypeListAssociation`, () => {
            return this.listHelper.removeContentTypeListAssociation(web, listName, contentTypeName);
        });
        
    }
    /**
    * Sets a default value on a field in a list
    * Result: void
    * Example: setDefaultValueOnList(context.get_web(), "My Task List", "ClientId", 123)
    */
    public setDefaultValueOnList(web: SP.Web, listName: string, fieldInternalName: string, defaultValue: any): Fluent {
        return this.fluent.chainAction(`${this._helperName}.setDefaultValueOnList`, () => {
            return this.listHelper.setDefaultValueOnList(web, listName, fieldInternalName, defaultValue);
        });
        
    }
     /**
    * Enable email alerts on a list 
    * Result: void
    * Example: setAlerts(context.get_web(), "My Task List", true)
    * Note: will not work for 2010/2013
    */
    public setAlerts(web: SP.Web, listName: string, enabled: boolean): Fluent {
        return this.fluent.chainAction(`${this._helperName}.setAlerts`, () => {
            return this.listHelper.setAlerts(web, listName, enabled);
        });
        
    }
}