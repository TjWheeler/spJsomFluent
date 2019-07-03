import { common } from "../common"
import * as spData from "./Client.CamlBuilder"
export class PageHelper {
    constructor(context: SP.ClientContext) {
        this.context = context;
    }
    private context: SP.ClientContext;
    public createPublishingPage(web: SP.Web, name: string, layoutUrl: string): JQueryPromise<SP.Publishing.PublishingPage> {
        var deferred = $.Deferred();
        var file = this.context.get_site().get_rootWeb().getFileByServerRelativeUrl(layoutUrl);
        var listItem = file.get_listItemAllFields();
        this.context.load(listItem);
        var pageInfo = new SP.Publishing.PublishingPageInformation();
        pageInfo.set_name(name);
        pageInfo.set_pageLayoutListItem(listItem);
        var publishingWeb = SP.Publishing.PublishingWeb.getPublishingWeb(this.context, web);
        this.context.load(publishingWeb);
        var publishingPage = publishingWeb.addPublishingPage(pageInfo);
        this.context.load(publishingPage);
        common.executeQuery(this.context)
            .fail((sender, args) => { deferred.reject(sender, args); })
            .done(() => {
                deferred.resolve(publishingPage);
            });
        return deferred.promise() as JQueryPromise<SP.Publishing.PublishingPage>;
    }
    public getPublishingLayout(serverRelativeUrl: string): JQueryPromise<SP.ListItem> {
        var deferred = $.Deferred();
        var file = this.context.get_site().get_rootWeb().getFileByServerRelativeUrl(serverRelativeUrl);
        var listItem = file.get_listItemAllFields();
        this.context.load(listItem);
        common.executeQuery(this.context)
            .fail((sender, args) => { deferred.reject(sender, args); })
            .done(() => {
                deferred.resolve(listItem);
            });
        return deferred.promise() as JQueryPromise<SP.ListItem>;
    }
    public setLayout(web: SP.Web, name: string, layoutUrl: string): JQueryPromise<any> {
        var deferred = $.Deferred();
        var file = this.context.get_site().get_rootWeb().getFileByServerRelativeUrl(layoutUrl);
        var listItem = file.get_listItemAllFields();
        this.context.load(listItem);

        var camlBuilder = new spData.CamlBuilder();
        camlBuilder.addTextClause(spData.CamlOperator.Eq, "FileLeafRef", name);
        var list = web.get_lists().getByTitle("Pages");
        var query = new SP.CamlQuery();
        query.set_viewXml(camlBuilder.viewXml);
        var pages = list.getItems(query);
        this.context.load(pages);
        common.executeQuery(this.context)
            .fail((sender, args) => { deferred.reject(sender, args); })
            .done(() => {
                if (pages.get_count() === 0) {
                    deferred.reject(this, { get_message: () => { return "Page not found" } });
                }
                var page = pages.getItemAtIndex(0);
                page.set_item("PublishingPageLayout", layoutUrl);
                page.update();
                common.executeQuery(this.context)
                    .fail((sender, args) => { deferred.reject(sender, args); })
                    .done(() => {
                        deferred.resolve();
                    });
            });
        return deferred.promise() as JQueryPromise<any>;
    }

}