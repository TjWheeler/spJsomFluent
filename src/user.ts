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
    * Result: SP.User
    * Example: get("my@email.address")
    * Example: get("i:0#.f|membership|my@email.address")
    */
    public get(emailOrAccountName:string): Fluent {
        return this._fluent.chainAction(`${this._helperName}.get`, () => {
            return this.userHelper.getUserByEmail(emailOrAccountName);
        });
        
    }
    /**
    * Get a user by their Id
    * Result: SP.User
    * Example: getById(15)
    */
    public getById(id: number): Fluent {
        return this._fluent.chainAction(`${this._helperName}.getById`, () => {
            return this.userHelper.getUserById(id);
        });
        
    }
    /**
    * Get the current user
    * Result: SP.User
    * Example: getCurrentUser()
    */
    public getCurrentUser(): Fluent {
        return this._fluent.chainAction(`${this._helperName}.getCurrentUser`, () => {
            return this.userHelper.getCurrentUser();
        });
        
    }
    /**
    * Get the profile properties for the current user
    * Result: SP.UserProfiles.PersonProperties
    * Example: getCurrentUserProfileProperties()
    */
    public getCurrentUserProfileProperties(): Fluent {
        this._fluent.registerDependency(Dependency.UserProfile);
        return this._fluent.chainAction(`${this._helperName}.getCurrentUserProfileProperties`, () => {
            return this.userHelper.getCurrentUserProfileProperties();
        });
        
    }
    /**
    * Get the manager for the current user
    * Result: SP.User
    * Example: getCurrentUserManager()
    */
    public getCurrentUserManager(): Fluent {
        this._fluent.registerDependency(Dependency.UserProfile);
        return this._fluent.chainAction(`${this._helperName}.getCurrentUserManager`, () => {
            return this.userHelper.getCurrentUserManager();
        });
        
    }
}