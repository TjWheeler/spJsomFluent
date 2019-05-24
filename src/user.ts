import * as core from "./core"
import UserHelper from "./helper/userHelper";
import { Fluent, Dependency } from "./fluent"
import common from "./common";

export default class User {
    constructor(fluent: Fluent) {
        this._fluent = fluent;
        this.userHelper = new UserHelper(fluent.context);
    }
    private _fluent = null as Fluent;
    private readonly _helperName: string = "userProfile";
    private userHelper = null as UserHelper;
    /**
    * Get User by email or account name
    * Example: get("my@email.address")
    * Example: get("i:0#.f|membership|my@email.address")
    */
    public get(emailOrAccountName:string): Fluent {
        this._fluent.chainAction(`${this._helperName}.get`, () => {
            return this.userHelper.getUserByEmail(emailOrAccountName);
        });
        return this._fluent;
    }
    public getById(id: number): Fluent {
        this._fluent.chainAction(`${this._helperName}.getById`, () => {
            return this.userHelper.getUserById(id);
        });
        return this._fluent;
    }
    public getCurrentUser(): Fluent {
        this._fluent.chainAction(`${this._helperName}.getCurrentUser`, () => {
            return this.userHelper.getCurrentUser();
        });
        return this._fluent;
    }
    public getCurrentUserProfileProperties(): Fluent {
        this._fluent.registerDependency(Dependency.UserProfile);
        this._fluent.chainAction(`${this._helperName}.getCurrentUserProfileProperties`, () => {
            return this.userHelper.getCurrentUserProfileProperties();
        });
        return this._fluent;
    }
    public getCurrentUserManager(): Fluent {
        this._fluent.registerDependency(Dependency.UserProfile);
        this._fluent.chainAction(`${this._helperName}.getCurrentUserManager`, () => {
            return this.userHelper.getCurrentUserManager();
        });
        return this._fluent;
    }
}