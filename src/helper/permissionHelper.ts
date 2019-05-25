import common from "../common"

export default class PermissionHelper {
    constructor(context: SP.ClientContext) {
        this.context = context;
    }
    private context: SP.ClientContext;
    public hasWebPermission(permission: SP.PermissionKind, web: SP.Web): JQueryPromise<boolean> {
        var deferred = $.Deferred();
        this.context.load(web, 'EffectiveBasePermissions');
        common.executeQuery(this.context)
            .fail((sender, args) => { deferred.reject(sender, args); })
            .done(() => {
                deferred.resolve(web.get_effectiveBasePermissions().has(permission));
            });
        return deferred.promise() as JQueryPromise<boolean>;
    }
    public hasListPermission(permission: SP.PermissionKind, list: SP.List): JQueryPromise<boolean> {
        var deferred = $.Deferred();
        this.context.load(list, 'EffectiveBasePermissions');
        common.executeQuery(this.context)
            .fail((sender, args) => { deferred.reject(sender, args); })
            .done(() => {
                deferred.resolve(list.get_effectiveBasePermissions().has(permission));
            });
        return deferred.promise() as JQueryPromise<boolean>;
    }
    public hasItemPermission(permission: SP.PermissionKind, item: SP.ListItem): JQueryPromise<boolean> {
            var deferred = $.Deferred();
            this.context.load(item, 'EffectiveBasePermissions');
            common.executeQuery(this.context)
                .fail((sender, args) => { deferred.reject(sender, args); })
                .done(() => {
                    deferred.resolve(item.get_effectiveBasePermissions().has(permission));
                });
        return deferred.promise() as JQueryPromise<boolean>;
    }
}