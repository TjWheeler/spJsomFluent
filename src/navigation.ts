import common from "./common"
import * as core from "./core"
import NavigationHelper from "./helper/navigationHelper";
import { Fluent, NavigationLocation, NavigationType, Dependency } from "./fluent"
export default class Navigation {
    constructor(fluent: Fluent) {
        this.fluent = fluent;
        this.navigationHelper = new NavigationHelper(fluent.context);
    }
    private fluent = null as Fluent;
    private navigationHelper = null as NavigationHelper;
    private readonly _helperName: string = "navigation";

    /**
    * Create new navigation node
    * Example: createNode(context.get_web(), NavigationLocation.Quicklaunch, "Test Node", "/sites/mysite/pages/default.aspx")
    */
    public createNode(web: SP.Web, location: NavigationLocation, title: string, url: string, asLastNode: boolean = true): Fluent {
        this.fluent.chainAction(`${this._helperName}.createNode`, () => {
            if (location == NavigationLocation.Quicklaunch) {
                return this.navigationHelper.createQuicklaunchNode(web, title, url, asLastNode);
            }
            else if (location == NavigationLocation.TopNavigation) {
                return this.navigationHelper.createTopNavigationNode(web, title, url, asLastNode);
            } else {
                throw "Unknown location " + location;
            }
        });
        return this.fluent;
    }
    /**
    * Delete all navigation nodes for the web
    * Example: deleteAllNodes(context.get_web(), NavigationLocation.Quicklaunch)
    */
    public deleteAllNodes(web: SP.Web, location: NavigationLocation): Fluent {
        this.fluent.chainAction(`${this._helperName}.deleteAllNodes`, () => {
            if (location == NavigationLocation.Quicklaunch) {
                return this.navigationHelper.deleteQuicklaunchNodes(web);
            }
            else if (location == NavigationLocation.TopNavigation) {
                return this.navigationHelper.deleteTopNavigationNodes(web);
            } else {
                throw "Unknown location " + location;
            }


        });
        return this.fluent;
    }
    /**
    * Delete all navigation nodes that match the supplied title for the web
    * Example: deleteNode(context.get_web(), NavigationLocation.Quicklaunch, "My link title")
    */
    public deleteNode(web: SP.Web, location: NavigationLocation, title:string ): Fluent {
        this.fluent.chainAction(`${this._helperName}.deleteNode`, () => {
            if (location == NavigationLocation.Quicklaunch) {
                return this.navigationHelper.deleteQuicklaunchNode(web, title);
            }
            else if (location == NavigationLocation.TopNavigation) {
                return this.navigationHelper.deleteTopQuicklaunchNode(web, title);
            } else {
                throw "Unknown location " + location;
            }
        });
        return this.fluent;
    }
    /**
    * Set navigation for the web
    * Example: setCurrentNavigation(context.get_web(), 3, true, true)
    * Note: showSubsites and showPages is only applicable for NavigationType.StructuralChildrenOnly (3)
    */
    public setCurrentNavigation(web: SP.Web, type: NavigationType, showSubsites: boolean = false, showPages:boolean = false): Fluent {
        this.fluent.registerDependency(Dependency.Publishing);
        this.fluent.chainAction(`${this._helperName}.setCurrentNavigation`, () => {
            return this.navigationHelper.setCurrentNavigation(web, type, showSubsites, showPages);
        });
        return this.fluent;
    }
}