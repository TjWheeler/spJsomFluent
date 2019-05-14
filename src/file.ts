import ListHelper from "./helper/listHelper";
import { Fluent } from "./fluent"
export default class File {
    constructor(fluent: Fluent) {
        this.fluent = fluent;
        this.listHelper = new ListHelper(fluent.context);
    }
    private fluent = null as Fluent;
    private listHelper = null as ListHelper;
    private readonly _helperName: string = "file";

    public getFile(serverRelativeUrl: string): Fluent {
        this.fluent.chainAction(`${this._helperName}.getFile`, () => {
            return this.listHelper.getFileListItem(serverRelativeUrl);
        });
        return this.fluent;
    }
    public checkInFile(web: SP.Web, serverRelativeUrl: string, comment: string, checkInType: SP.CheckinType): Fluent {
        this.fluent.chainAction(`${this._helperName}.checkInFile`, () => {
            return this.listHelper.checkInFile(web, serverRelativeUrl, comment, checkInType);
        });
        return this.fluent;
    }
}