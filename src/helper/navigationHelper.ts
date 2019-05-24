import common from "../common"
import { NavigationType } from "../fluent";

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
    
    public setCurrentNavigation(web: SP.Web, navigationType: NavigationType, showSubsites: boolean = true, showPages:boolean = true): JQueryPromise<any> {
        var deferred = $.Deferred();
        const siblingsPropertyName = "__NavigationShowSiblings";
        const includeTypesPropertyName = "__CurrentNavigationIncludeTypes";
        var allProperties = web.get_allProperties();
        this.context.load(web);
        this.context.load(allProperties);
        var setOptions = () => {
            if (showPages && showSubsites) {
                allProperties.set_item(includeTypesPropertyName, "3");
            } else if (showPages && !showSubsites) {
                allProperties.set_item(includeTypesPropertyName, "2");
            }
            else if (!showPages && showSubsites) {
                allProperties.set_item(includeTypesPropertyName, "1");
            }
            else if (!showPages && !showSubsites) {
                allProperties.set_item(includeTypesPropertyName, "0");
            }
        };
        common.executeQuery(this.context)
            .fail((sender, args) => { deferred.reject(sender, args); })
            .done(() => {
                var nav = new SP.Publishing.Navigation.WebNavigationSettings(this.context, web);
                switch (navigationType) {
                    case NavigationType.Inherit:
                        nav.get_currentNavigation().set_source(SP.Publishing.Navigation.StandardNavigationSource.inheritFromParentWeb);
                        break;
                    case NavigationType.Managed:
                        common.reject(deferred, "Not implemented");
                        break;
                    case NavigationType.StructuralWithSiblings:
                        nav.get_currentNavigation().set_source(SP.Publishing.Navigation.StandardNavigationSource.portalProvider);
                        allProperties.set_item(siblingsPropertyName, "True");
                        setOptions();
                        web.update();
                        break;
                    case NavigationType.StructuralChildrenOnly:
                        nav.get_currentNavigation().set_source(SP.Publishing.Navigation.StandardNavigationSource.portalProvider);
                        allProperties.set_item(siblingsPropertyName, "False");
                        setOptions();
                        web.update();
                        break;
                    default:
                        common.reject(deferred, "Unknown Navigation Type");
                }   
                nav.update(null);
                common.executeQuery(this.context)
                    .fail((sender, args) => { deferred.reject(sender, args); })
                    .done(() => {
                        deferred.resolve();
                    });
            });
        return deferred.promise() as JQueryPromise<any>;
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