import { Fluent, Dependency } from "./fluent"
import { PageHelper } from "./helper/pageHelper"

export class PublishingPage {
    constructor(fluent: Fluent) {
        this.fluent = fluent;
        this.pageHelper = new PageHelper(fluent.context);
    }
    private fluent = null as Fluent;
    private readonly _helperName: string = "publishingPage";
    private pageHelper: PageHelper;
    /**
    * Creates a new publishing page.
    * Result: SP.Publishing.PublishingPage
    * Example: .publishingPage.create(SP.ClientContext.get_current().get_web(), "Home.aspx", _spPageContextInfo.siteServerRelativeUrl + "/_catalogs/masterpage/BlankWebPartPage.aspx")
    */
    public create(web: SP.Web, name: string, layoutUrl: string): Fluent {
        this.fluent.registerDependency(Dependency.Publishing);
        return this.fluent.chainAction(`${this._helperName}.create`, () => {
            return this.pageHelper.createPublishingPage(web, name, layoutUrl);
        });
        
    }
    /**
    * Gets a page layout.
    * Result: SP.ListItem
    * Example: .publishingPage.getLayout(_spPageContextInfo.siteServerRelativeUrl + "/_catalogs/masterpage/BlankWebPartPage.aspx")
    * 
    */
    public getLayout(serverRelativeUrl:string): Fluent {
        return this.fluent.chainAction(`${this._helperName}.getLayout`, () => {
            return this.pageHelper.getPublishingLayout(serverRelativeUrl);
        });
        
    }

}