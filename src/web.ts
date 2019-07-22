import { WebHelper } from "./helper/webHelper";
import { Fluent } from "./fluent"

export class Web {
    constructor(fluent: Fluent) {
        this._fluent = fluent;
        this.webHelper = new WebHelper(fluent.context);
    }
    private _fluent = null as Fluent;
    private readonly _helperName: string = "web";
    private webHelper = null as WebHelper;
    /**
      * Set the web welcome page.  The url should be relative to the root folder of the web. 
      * Result: SP.Folder
      * Example: setWelcomePage(context.get_web(), "pages/default.aspx")
      */
    public setWelcomePage(web: SP.Web, url: string): Fluent {
        return this._fluent.chainAction(`${this._helperName}.setWelcomePage`, () => {
            return this.webHelper.setWelcomePage(web, url);
        });
        
    }
    /**
      * Create a new sub web.
      * Result: SP.Web
      * name: Forms part of the URL, don't include slash
      * template: template id and configuration value
      * Example: create("SubWebName", SP.ClientContext.get_current().get_rootWeb(), "My Web Name", "CMSPUBLISHING#0")
      * 
      * @remarks
      * For templates See: https://blogs.technet.microsoft.com/praveenh/2010/10/21/sharepoint-templates-and-their-ids/
      */
    public create(name: string, parentWeb: SP.Web, title: string, template: string, useSamePermissionsAsParent: boolean = true): Fluent {
        return this._fluent.chainAction(`${this._helperName}.create`, () => {
            return this.webHelper.createWeb(name, parentWeb, title, template, useSamePermissionsAsParent);
        });
        
    }
    /**
      * Check if a web exists
      * Result: boolean
      * url: server relative url
      * Example: .exists("/sites/mysubweb")
      */
    public exists(url: string): Fluent {
        return this._fluent.chainAction(`${this._helperName}.exists`, () => {
            return this.webHelper.doesWebExist(url);
        });
        
    }
    /**
      * Get and load all webs starting from a web and its children
      * Result: Array<SP.Web>
      * Example: getWebs(SP.ClientContext.get_current().get_rootWeb())
      */
    public getWebs(fromWeb: SP.Web): Fluent {
        return this._fluent.chainAction(`${this._helperName}.getWebs`, () => {
            return this.webHelper.getWebs(fromWeb);
        });
        
    }

    /**
      * Set site title
      * Result: void
      * Example: setTitle(SP.ClientContext.get_current().get_web(), "My Title")
      */
    public setTitle(web: SP.Web, title:string): Fluent {
        return this._fluent.chainAction(`${this._helperName}.setTitle`, () => {
            return this.webHelper.setTitle(web, title);
        });

    }

    /**
      * Set site logo url
      * Result: void
      * Example: setLogoUrl(SP.ClientContext.get_current().get_web(), "/site/images/logo.jpg")
      */
    public setLogoUrl(web: SP.Web, url: string): Fluent {
        return this._fluent.chainAction(`${this._helperName}.setLogoUrl`, () => {
            return this.webHelper.setLogoUrl(web, url);
        });

    }

    /**
      * Activate feature
      * Result: void
      * Example: activateFeature(SP.ClientContext.get_current().get_web(), new SP.Guid('15a572c6-e545-4d32-897a-bab6f5846e18'))
      */
    public activateFeature(web: SP.Web, featureId: SP.Guid, force:boolean = true): Fluent {
        return this._fluent.chainAction(`${this._helperName}.activateFeature`, () => {
            return this.webHelper.activateFeature(web, featureId, force);
        });
    }
    
}