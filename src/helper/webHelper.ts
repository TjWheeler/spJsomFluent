import { common } from "../common"

export class WebHelper {
    constructor(context: SP.ClientContext) {
        this.context = context;
    }
    private context: SP.ClientContext;
    public setWelcomePage(web: SP.Web, url: string): JQueryPromise<SP.Folder> {
            var deferred = $.Deferred();
            var rootFolder = web.get_rootFolder();
            this.context.load(web);
            this.context.load(rootFolder);
            rootFolder.set_welcomePage(url);
            rootFolder.update();
            common.executeQuery(this.context)
                .fail((sender, args) => { deferred.reject(sender, args); })
                .done(() => {
                    deferred.resolve(rootFolder);
                });
        return deferred.promise() as JQueryPromise<SP.Folder>;
    }
    public setTitle(web: SP.Web, title: string): JQueryPromise<any> {
        var deferred = $.Deferred();
        this.context.load(web);
        web.set_title(title);
        web.update();
        common.executeQuery(this.context)
            .fail((sender, args) => { deferred.reject(sender, args); })
            .done(() => {
                deferred.resolve();
            });
        return deferred.promise() as JQueryPromise<any>;
    }
    public setLogoUrl(web: SP.Web, url: string): JQueryPromise<any> {
        var deferred = $.Deferred();
        this.context.load(web);
        (<any>web).set_siteLogoUrl(url); //TS definition missing here.
        web.update();
        common.executeQuery(this.context)
            .fail((sender, args) => { deferred.reject(sender, args); })
            .done(() => {
                deferred.resolve();
            });
        return deferred.promise() as JQueryPromise<any>;
    }
    public activateFeature(web: SP.Web, featureId: SP.Guid, force:boolean): JQueryPromise<any> {
        var deferred = $.Deferred();
        this.context.load(web);
        web.get_features().add(featureId, force, SP.FeatureDefinitionScope.web);
        web.update();
        common.executeQuery(this.context)
            .fail((sender, args) => { deferred.reject(sender, args); })
            .done(() => {
                deferred.resolve();
            });
        return deferred.promise() as JQueryPromise<any>;
    }
    public createWeb(name: string, parentWeb: SP.Web, title: string, template: string, useSamePermissionsAsParent:boolean = true): JQueryPromise<SP.Web> {
            var deferred = $.Deferred();
            var info = new SP.WebCreationInformation();
            info.set_url(name);
            info.set_title(title);
            info.set_webTemplate(template);
            info.set_useSamePermissionsAsParentSite(useSamePermissionsAsParent);
            var newWeb = parentWeb.get_webs().add(info);
            this.context.load(newWeb);
            common.executeQuery(this.context)
                .fail((sender, args) => { deferred.reject(sender, args); })
                .done(() => {
                    deferred.resolve(newWeb);
                });
        return deferred.promise() as JQueryPromise<SP.Web>;
    }
    public doesWebExist(url: string): JQueryPromise<any> {
            var deferred = $.Deferred();
            var output = [] as Array<SP.Web>;
            this.getAllWebs(this.context, this.context.get_site().get_rootWeb(), output)
                .fail((sender, args) => { deferred.reject(sender, args); })
                .done(() => {
                    for (var i = 0; i < output.length; i++) {
                        if (url.toLowerCase() === output[i].get_serverRelativeUrl() || url.toLowerCase() === output[i].get_url()) {
                            deferred.resolve(true);
                            return;
                        }
                    }
                    deferred.resolve(false);
                });
            return deferred.promise() as JQueryPromise<any>;
    }
    public getWebs(fromWeb: SP.Web): JQueryPromise<Array<SP.Web>> {
            var deferred = $.Deferred();
            var output = [] as Array<SP.Web>;
            this.getAllWebs(this.context, fromWeb, output)
                .fail((sender, args) => { deferred.reject(sender, args); })
                .done(() => {
                    deferred.resolve(output);
                });
        return deferred.promise() as JQueryPromise<Array<SP.Web>>;
    }
    private getAllWebs(context: SP.ClientContext, web: SP.Web, output: Array<SP.Web>): JQueryPromise<Array<SP.Web>> {
        var deferred = $.Deferred();
        var webs = web.get_webs();
        context.load(webs);
        common.executeQuery(context)
            .fail((sender, args) => {
                deferred.reject(sender, args);
            })
            .done(() => {
                var promises = [];
                for (var i = 0; i < webs.get_count(); i++) {
                    var subWeb = webs.getItemAtIndex(i);
                    output.push(subWeb);
                    promises.push(this.getAllWebs(context, subWeb, output));
                }
                if (promises.length) {
                    $.when.apply($, promises)
                        .fail((sender, args) => {
                            deferred.reject(sender, args);
                        })
                        .done(() => {
                            deferred.resolve();
                        });
                } else {
                    deferred.resolve();
                }
            });
        return deferred.promise() as JQueryPromise<Array<SP.Web>>;
    }

}