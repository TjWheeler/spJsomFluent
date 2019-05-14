import * as core from "./core"
import UserHelper from "./helper/userHelper";
import { Fluent } from "./fluent"

export class User {
    constructor(fluent: Fluent) {
        this._fluent = fluent;
        this.userHelper = new UserHelper(fluent.context);
    }
    private _fluent = null as Fluent;
    private readonly _helperName: string = "userProfile";
    private userHelper = null as UserHelper;
    public getCurrentUserProfileProperties(): Fluent {
        this._fluent.registerDependency(core.Dependency.UserProfile);
        this._fluent.chainAction(`${this._helperName}.getCurrentUserProfile`, () => {
            return this.userHelper.getCurrentUserProfileProperties();
        });
        return this._fluent;
    }
    public getCurrentUserManager(): Fluent {
        this._fluent.registerDependency(core.Dependency.UserProfile);
        this._fluent.chainAction(`${this._helperName}.getCurrentUserManager`, () => {
            return this.userHelper.getCurrentUserManager();
        });
        return this._fluent;
    }
}