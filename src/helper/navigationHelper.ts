import common from "../common"

export default class NavigationHelper {
    constructor(context: SP.ClientContext) {
        this.context = context;
    }
    private context: SP.ClientContext;
    public deleteQuicklaunchNodes(web: SP.Web): JQueryPromise<any> {
        return this.deleteNodes(web.get_navigation().get_quickLaunch());
    }
    public deleteTopNavigationNodes(web: SP.Web): JQueryPromise<any> {
        return this.deleteNodes(web.get_navigation().get_topNavigationBar());
    }
    public deleteQuicklaunchNode(web: SP.Web, title: string): JQueryPromise<any> {
        return this.deleteNode(web.get_navigation().get_quickLaunch(), title);
    }
    public deleteTopQuicklaunchNode(web: SP.Web, title: string): JQueryPromise<any> {
        return this.deleteNode(web.get_navigation().get_topNavigationBar(), title);
    }
    private deleteNodes(nav: SP.NavigationNodeCollection): JQueryPromise<any> {
        var deferred = $.Deferred();
        this.context.load(nav);
        common.executeQuery(this.context)
            .fail((sender, args) => { deferred.reject(sender, args); })
            .done(() => {
                var enumerator = nav.getEnumerator();
                var itemsToDelete = [];
                while (enumerator.moveNext()) {
                    itemsToDelete.push(enumerator.get_current());
                }
                for (let i = 0; i < itemsToDelete.length; i++) {
                    itemsToDelete[i].deleteObject();
                }
                common.executeQuery(this.context)
                    .fail((sender, args) => { deferred.reject(sender, args); })
                    .done(() => {
                        deferred.resolve();
                    });
            });
        return deferred.promise() as JQueryPromise<any>;
    }
    private deleteNode(nav: SP.NavigationNodeCollection, title: string): JQueryPromise<any> {
        var deferred = $.Deferred();
        this.context.load(nav);
        common.executeQuery(this.context)
            .fail((sender, args) => { deferred.reject(sender, args); })
            .done(() => {
                var enumerator = nav.getEnumerator();
                var itemsToDelete = [];
                while (enumerator.moveNext()) {
                    var node = enumerator.get_current();
                    if (node.get_title() === title) {
                        itemsToDelete.push(node);
                    }
                }
                for (let i = 0; i < itemsToDelete.length; i++) {
                    itemsToDelete[i].deleteObject();
                }
                common.executeQuery(this.context)
                    .fail((sender, args) => { deferred.reject(sender, args); })
                    .done(() => {
                        deferred.resolve();
                    });
            });
        return deferred.promise() as JQueryPromise<any>;
    }
    public createQuicklaunchNode(web: SP.Web, title:string, url:string, asLastNode:boolean = true): JQueryPromise<SP.NavigationNode> {
        return this.createNode(web.get_navigation().get_quickLaunch(), title, url, asLastNode);
    }
    public createTopNavigationNode(web: SP.Web, title: string, url: string, asLastNode: boolean = true): JQueryPromise<SP.NavigationNode> {
        return this.createNode(web.get_navigation().get_topNavigationBar(), title, url, asLastNode);
    }
    private createNode(nav:SP.NavigationNodeCollection, title: string, url: string, asLastNode: boolean = true): JQueryPromise<SP.NavigationNode> {
        var deferred = $.Deferred();
        this.context.load(nav);
        var info = new SP.NavigationNodeCreationInformation();
        info.set_title(title);
        info.set_url(url);
        info.set_asLastNode(asLastNode);
        var node = nav.add(info);
        this.context.load(nav);
        common.executeQuery(this.context)
            .fail((sender, args) => { deferred.reject(sender, args); })
            .done(() => {
                deferred.resolve(node);
            });
        return deferred.promise() as JQueryPromise<SP.NavigationNode>;
    }
}