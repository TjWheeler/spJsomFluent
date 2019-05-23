/// <reference path="../typings/@types/all.ts" />
import WebHelper from "./helper/webHelper";
import { Fluent } from "./fluent"

export default class Web {
    constructor(fluent: Fluent) {
        this._fluent = fluent;
        this.webHelper = new WebHelper(fluent.context);
    }
    private _fluent = null as Fluent;
    private readonly _helperName: string = "web";
    private webHelper = null as WebHelper;
    /**
      * Set the web welcome page.  The url should be relative to the root folder of the web. 
      * Example: setWelcomePage(context.get_web(), "pages/default.aspx")
      */
    public setWelcomePage(web: SP.Web, url: string): Fluent {
        this._fluent.chainAction(`${this._helperName}.setWelcomePage`, () => {
            return this.webHelper.setWelcomePage(web, url);
        });
        return this._fluent;
    }
    /**
      * Create a new sub web.
      * name: Forms part of the URL, don't include slash
      * template: template id and configuration value
      * Example: create("SubWebName", SP.ClientContext.get_current().get_rootWeb(), "My Web Name", "CMSPUBLISHING#0")
      * 
      * @remarks
      * For templates See: https://blogs.technet.microsoft.com/praveenh/2010/10/21/sharepoint-templates-and-their-ids/
      */
    public create(name: string, parentWeb: SP.Web, title: string, template: string, useSamePermissionsAsParent: boolean = true): Fluent {
        this._fluent.chainAction(`${this._helperName}.create`, () => {
            return this.webHelper.createWeb(name, parentWeb, title, template, useSamePermissionsAsParent);
        });
        return this._fluent;
    }
    /**
      * Check if a web exists
      * url: server relative url
      * Example: .exists("/sites/mysubweb")
      */
    public exists(url: string): Fluent {
        this._fluent.chainAction(`${this._helperName}.exists`, () => {
            return this.webHelper.doesWebExist(url);
        });
        return this._fluent;
    }
    /**
      * Get and load all webs starting from a web
      * Example: getWebs(SP.ClientContext.get_current().get_rootWeb()))
      */
    public getWebs(fromWeb: SP.Web): Fluent {
        this._fluent.chainAction(`${this._helperName}.getWebs`, () => {
            return this.webHelper.getWebs(fromWeb);
        });
        return this._fluent;
    }
    
}