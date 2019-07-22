import { ListHelper } from "./helper/listHelper";
import { Fluent } from "./fluent"
export class File {
    constructor(fluent: Fluent) {
        this.fluent = fluent;
        this.listHelper = new ListHelper(fluent.context);
    }
    private fluent = null as Fluent;
    private listHelper = null as ListHelper;
    private readonly _helperName: string = "file";

    /**
    * Get a file from a document library
    * Result: SP.File
    * Example: get(_spPageContextInfo.siteServerRelativeUrl + "/documents/doc.docx")
    */
    public get(serverRelativeUrl: string): Fluent {
        return this.fluent.chainAction(`${this._helperName}.getListItem`, () => {
            return this.listHelper.getFile(serverRelativeUrl);
        });
        
    }
    /**
    * Get the list item for the file
    * Result: SP.ListItem
    * Example: getListItem(_spPageContextInfo.siteServerRelativeUrl + "/documents/doc.docx")
    */
    public getListItem(serverRelativeUrl: string): Fluent {
        return this.fluent.chainAction(`${this._helperName}.getListItem`, () => {
            return this.listHelper.getFileListItem(serverRelativeUrl);
        });
        
    }
    /**
    * Check in a file
    * Result: SP.File
    * Example: checkIn(SP.ClientContext.get_current().get_web(), _spPageContextInfo.webServerRelativeUrl + '/pages/mypage.aspx', "Checked in by spJsomFluent", SP.CheckinType.majorCheckIn)
    */
    public checkIn(web: SP.Web, serverRelativeUrl: string, comment: string, checkInType: SP.CheckinType): Fluent {
        return this.fluent.chainAction(`${this._helperName}.checkInFile`, () => {
            return this.listHelper.checkInFile(web, serverRelativeUrl, comment, checkInType);
        });
        
    }
    /**
    * Check out a file
    * Result: SP.File
    * Example: checkOut(SP.ClientContext.get_current().get_web(), _spPageContextInfo.webServerRelativeUrl + '/pages/mypage.aspx', "Checked in by spJsomFluent", SP.CheckinType.majorCheckIn)
    */
    public checkOut(web: SP.Web, serverRelativeUrl: string): Fluent {
        return this.fluent.chainAction(`${this._helperName}.checkOutFile`, () => {
            return this.listHelper.checkOutFile(web, serverRelativeUrl);
        });

    }
    /**
    * Upload a file to a library
    * Result: SP.File
    * Example: uploadFile(SP.ClientContext.get_current().get_web(), "Documents", file)
    */
    public uploadFile(web: SP.Web, listName: string, file: any, filename?:string): Fluent {
        return this.fluent.chainAction(`${this._helperName}.uploadFile`, () => {
            return this.listHelper.uploadFile(web, listName, file, filename);
        });

    }
}