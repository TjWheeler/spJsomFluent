import { Fluent } from "./fluent"
import PermissionHelper from "./helper/permissionHelper";

export default class Permission {
    constructor(fluent: Fluent) {
        this._fluent = fluent;
        this.permissionHelper = new PermissionHelper(fluent.context);
    }
    private _fluent = null as Fluent;
    private readonly _helperName: string = "permission";
    private permissionHelper = null as PermissionHelper;
    public hasWebPermission(permission: SP.PermissionKind, web: SP.Web): Fluent {
        this._fluent.chainAction(`${this._helperName}.hasWebPermission`, () => {
            return this.permissionHelper.hasWebPermission(permission, web);
        });
        return this._fluent;
    }
    public hasListPermission(permission: SP.PermissionKind, list: SP.List): Fluent {
        this._fluent.chainAction(`${this._helperName}.hasSitePermission`, () => {
            return this.permissionHelper.hasListPermission(permission, list);
        });
        return this._fluent;
    }
    public hasItemPermission(permission: SP.PermissionKind, item: SP.ListItem): Fluent {
        this._fluent.chainAction(`${this._helperName}.hasItemPermission`, () => {
            return this.permissionHelper.hasItemPermission(permission, item);
        });
        return this._fluent;
    }
}