
export default class common {
    public static reject(deferred: JQueryDeferred<any>, reason:string) {
        deferred.reject(this, { get_message: function () { return reason; } } );
    }
    public static notImplementedPromise() : JQueryPromise<any> {
        var deferred = $.Deferred();
        common.reject(deferred, "Not Implemented");
        return deferred.promise();
    }
    public static executeQuery = (clientContext: SP.ClientContext) => {
        var deferred = $.Deferred();
        clientContext.executeQueryAsync(
            (sender, args) => {
                deferred.resolve(sender, args);
            },
            (sender, args) => {
                deferred.reject(sender, args);
            }
        );
        return deferred.promise();
    };
    public static FilterArray<T>(items: Array<T>, predicate: any): Array<T> {
        let output = [];
        for (let i = 0; i < items.length; i++) {
            if (predicate(items[i], i)) {
                output.push(items[i]);
            }
        }
        return output;
    }
}