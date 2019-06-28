import { Fluent } from "./fluent"
import { PermissionHelper } from "./helper/permissionHelper";

export class Permission {
    constructor(fluent: Fluent) {
        this._fluent = fluent;
        this.permissionHelper = new PermissionHelper(fluent.context);
    }
    private _fluent = null as Fluent;
    private readonly _helperName: string = "permission";
    private permissionHelper = null as PermissionHelper;
    /**
    * Check if the current user has specified Web permission
    * Result: boolean
    * Example: hasWebPermission(SP.PermissionKind.createSSCSite, context.get_web())
    */
    public hasWebPermission(permission: SP.PermissionKind, web: SP.Web): Fluent {
        return this._fluent.chainAction(`${this._helperName}.hasWebPermission`, () => {
            return this.permissionHelper.hasWebPermission(permission, web);
        });
        
    }
    /**
    * Check if the current user has specified List permission
    * Result: boolean
    * Example: hasListPermission(SP.PermissionKind.addListItems, context.get_web().get_lists().getByTitle("MyList"))
    */
    public hasListPermission(permission: SP.PermissionKind, list: SP.List): Fluent {
        return this._fluent.chainAction(`${this._helperName}.hasListPermission`, () => {
            return this.permissionHelper.hasListPermission(permission, list);
        });
        
    }
    /**
    * Check if the current user has specified ListItem permission
    * Result: boolean
    * Example: hasItemPermission(SP.PermissionKind.editListItems, context.get_web().get_lists().getByTitle("MyList").getItemById(0))
    */
    public hasItemPermission(permission: SP.PermissionKind, item: SP.ListItem): Fluent {
        return this._fluent.chainAction(`${this._helperName}.hasItemPermission`, () => {
            return this.permissionHelper.hasItemPermission(permission, item);
        });
        
    }
}