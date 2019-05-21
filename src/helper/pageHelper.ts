import common from "../common"

export default class PageHelper {
    constructor(context: SP.ClientContext) {
        this.context = context;
    }
    private context: SP.ClientContext;
    public createPublishingPage(web: SP.Web, name: string, layoutUrl: string): JQueryPromise<SP.ListItem> {
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
        return deferred.promise() as JQueryPromise<SP.ListItem>;
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
    

}