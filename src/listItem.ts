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
    /**
    * Update a list item with properties
    * Result: SP.ListItem
    * Example: update(listItem, { 'Title': 'Title value here', ClientId: 123 })
    */
    public update(listItem: SP.ListItem, properties: any): Fluent {
        return this.fluent.chainAction(`${this._helperName}.update`, () => {
            return this.listHelper.setListItemProperties(listItem, properties);
        });
        
    }
    /**
    * Create a list item, specifying the Content Type
    * Result: SP.ListItem
    * Example: createWithContentType(context.get_web(), "My list", "My Content Type Name", { 'Title': 'Title value here', ClientId: 123 })
    */
    public createWithContentType(web: SP.Web, listName: string, contentTypeName: string, properties: any): Fluent {
        return this.fluent.chainAction(`${this._helperName}.createWithContentType`, () => {
            return this.listHelper.createListItemWithContentTypeName(web, listName, contentTypeName, properties);
        });
        
    }
    /**
    * Create new list item with optional property values
    * Example: create(context.get_web(), "MyList", properties)
    * Note: Properties are an object with Property/Value, where property
    *       is the internal field name.
    *       Eg; var properties = {
                Title: "My title",
                PersonOrGroupField: personValue,
                MultiChoiceField: ["Second", "Third"],
                ChoiceField: "Second",
                NumberField: 1234
            };
            For personValue you can pass through the user Id or a SP.UserFieldValue such as:
            var personValue = new SP.FieldUserValue();
            personValue.set_lookupId(_spPageContextInfo.userId);
    */
    public create(web: SP.Web, listName: string, properties: any): Fluent {
        return this.fluent.chainAction(`${this._helperName}.create`, () => {
            return this.listHelper.createListItem(web, listName, properties);
        });
        
    }
    
    /**
    * Get the listitem using the ID with optional view fields
    * Result: SP.ListItem
    * Example: get(context.get_web(), "MyList", 1, ["ID", "Title"])
    */
    public get(web: SP.Web, listName: string, id: number, viewFields: Array<string> = null): Fluent {
        return this.fluent.chainAction(`${this._helperName}.get`, () => {
            return this.listHelper.getListItemById(web, listName, id, viewFields);
        });
        
    }
    /**
    * Get the listitem for a File using the ID with optional view fields
    * Result: SP.ListItem
    * Example: getFileListItem("/sites/site/documents/mydoc.docx", ["ID", "Title", "FileLeafRef"])
    */
    public getFileListItem(serverRelativeUrl: string, viewFields: Array<string> = null): Fluent {
        return this.fluent.chainAction(`${this._helperName}.getFileListItem`, () => {
            return this.listHelper.getFileListItem(serverRelativeUrl, viewFields);
        });
        
    }
    
    /**
    * Delete List Item
    * Result: void
    * Example: deleteById(context.get_web(), "MyList", 7)
    */
    public deleteById(web: SP.Web, listName: string, id: number): Fluent {
        return this.fluent.chainAction(`${this._helperName}.deleteById`, () => {
            return this.listHelper.deleteListItemById(web, listName, id);
        });
        
    }
    /**
    * Execute a query using supplied CAML.  
    * Returns: SP.ListItemCollection
    * Example: query(context.get_web(), "MyList", "<View><Query><Where><In><FieldRef Name='ID' /><Values><Value Type='Number'>1</Value><Value Type='Number'>2</Value></Values></In></View></Query></Where>")
    */
    public query(web: SP.Web, listName: string, viewXml: string): Fluent {
        return this.fluent.chainAction(`${this._helperName}.query`, () => {
            return this.listHelper.getListItems(web, listName, viewXml);
        });
    }
    
}