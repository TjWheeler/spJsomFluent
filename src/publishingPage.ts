import * as core from "./core"
import { Fluent, Dependency } from "./fluent"
import PageHelper from "./helper/pageHelper"

export default class PublishingPage {
    constructor(fluent: Fluent) {
        this.fluent = fluent;
        this.pageHelper = new PageHelper(fluent.context);
    }
    private fluent = null as Fluent;
    private readonly _helperName: string = "publishingPage";
    private pageHelper: PageHelper;
    public create(web: SP.Web, name: string, layoutUrl: string): Fluent {
        this.fluent.registerDependency(Dependency.Publishing);
        this.fluent.chainAction(`${this._helperName}.create`, () => {
            return this.pageHelper.createPublishingPage(web, name, layoutUrl);
        });
        return this.fluent;
    }
    public getLayout(serverRelativeUrl:string): Fluent {
        this.fluent.chainAction(`${this._helperName}.getLayout`, () => {
            return this.pageHelper.getPublishingLayout(serverRelativeUrl);
        });
        return this.fluent;
    }

}