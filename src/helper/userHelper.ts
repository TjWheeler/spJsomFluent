import { common } from "../common"

export class UserHelper {
    constructor(context: SP.ClientContext) {
        this.context = context;
    }
    private context: SP.ClientContext;
    
    public getUserByEmail(email: string): JQueryPromise<SP.User> {
        return this.loadUser(this.context.get_web().ensureUser(email));
    }
    public getUserById(id: number): JQueryPromise<SP.User> {
        return this.loadUser(this.context.get_web().get_siteUsers().getById(id));
    }
    private loadUser(user: SP.User): JQueryPromise<SP.User> {
        var deferred = $.Deferred();
        this.context.load(user);
        this.context.executeQueryAsync(
            (sender, args) => {
                deferred.resolve(user);
            },
            (sender, args) => {
                console.error(args.get_message());
                deferred.reject(sender, args);
            }
        );
        return deferred.promise() as JQueryPromise<SP.User>;
    }
    public getCurrentUser(): JQueryPromise<SP.UserProfiles.PersonProperties> {
        var deferred = $.Deferred();
        var user = this.context.get_web().get_currentUser();
        this.context.load(user);
        this.context.executeQueryAsync(
            (sender, args) => {
                deferred.resolve(user);
            },
            (sender, args) => {
                console.error(args.get_message());
                deferred.reject(sender, args);
            }
        );
        return deferred.promise() as JQueryPromise<SP.UserProfiles.PersonProperties>;
    }
    public getCurrentUserProfileProperties(): JQueryPromise<SP.UserProfiles.PersonProperties> {
            var deferred = $.Deferred();
            SP.SOD.executeFunc('userprofile', 'SP.UserProfiles.PeopleManager', () => {
                var clientContext = this.context;
                var currentUser = clientContext.get_web().get_currentUser();
                var peopleManager = new SP.UserProfiles.PeopleManager(clientContext);
                var profile = peopleManager.getMyProperties();
                clientContext.load(currentUser);
                clientContext.load(profile);
                clientContext.executeQueryAsync(
                    (sender, args) => {
                        deferred.resolve(profile.get_userProfileProperties());
                    },
                    (sender, args) => {
                        console.error(args.get_message());
                        deferred.reject(sender, args);
                    }
                );
            });
        return deferred.promise() as JQueryPromise<SP.UserProfiles.PersonProperties>;
    }
    public getCurrentUserManager(): JQueryPromise<SP.User> {
            var deferred = $.Deferred();
            var peopleManager = new SP.UserProfiles.PeopleManager(this.context);
            var profilePropertyNames = ["Manager"];
            var user_email = this.context.get_web().get_currentUser().get_email();
            var userProfilePropertiesForUser = new SP.UserProfiles.UserProfilePropertiesForUser(this.context,
                "i:0#.f|membership|" + user_email,
                profilePropertyNames);  //TODO: check for better mechanism to constructure login
            var userProfileProps = peopleManager.getUserProfilePropertiesFor(userProfilePropertiesForUser);

            this.context.load(userProfilePropertiesForUser);
            common.executeQuery(this.context)
                .fail((sender, args) => { deferred.reject(sender, args); })
                .done(() => {
                    if (userProfileProps[0]) {
                        var user = this.context.get_web().ensureUser(userProfileProps[0]);
                        this.context.load(user);
                        common.executeQuery(this.context)
                            .fail((sender, args) => { deferred.reject(sender, args); })
                            .done(() => {
                                deferred.resolve(user);
                            });
                    } else {
                        deferred.resolve(null);
                    }
            })
        return deferred.promise() as JQueryPromise<SP.User>;
    }
}