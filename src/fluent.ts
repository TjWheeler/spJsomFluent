// spJsom - Fluent
import Web from "./web"
import Permission from "./permission"
import List from "./list"
import ListItem from "./listitem"
import File from "./file"
import User from "./user"
import Navigation from "./navigation"
import PublishingPage from "./publishingpage"
import common from "./common"

export class Fluent  {
    public context: SP.ClientContext;
    private commands: Array<FluentCommand> = [];
    private results: Array<ActionResult> = [];
    private resultPromise: JQueryDeferred<any>;
    private onExecuting: any;
    private onExecuted: any;
    private dependencies: Array<Dependency> = [];
    public settings: ISettings = { timeoutMilliseconds: 5000, enableDependencyTimeout: true };
    private totalCommands = 0;
    withContext(context: SP.ClientContext): Fluent {
        this.context = context;
        return this;
    }
    withSettings(settings: ISettings): Fluent {
        for (var setting in settings) {
            if (typeof (this.settings[setting]) !== typeof (undefined)) {
                this.settings[setting] = settings[setting];
            }
        }
        return this;
    }
    get promise(): JQueryPromise<any> {
        return this.resultPromise.promise();
    }
    get permission(): Permission {
        return new Permission(this);
    }
    get list(): List {
        return new List(this);
    }
    get listItem(): ListItem {
        return new ListItem(this);
    }
    get file(): File {
        return new File(this);
    }
    get publishingPage(): PublishingPage {
        return new PublishingPage(this);
    }
    get web(): Web {
        return new Web(this);
    }
    get user(): User{
        return new User(this);
    }
    get navigation(): Navigation {
        return new Navigation(this);
    }

    public execute(): JQueryPromise<Array<ActionResult>> {
        this.resultPromise = $.Deferred();
        //The dependency timeout is fired if it takes too long to load to avoid promise never completing.
        if (this.settings.enableDependencyTimeout) {
            var expiry = setTimeout(() => {
                common.reject(this.resultPromise, "Timeout waiting for dependencies to load");
            }, this.settings.timeoutMilliseconds);
        }
        //To track progress
        this.totalCommands = this.getAvailableActionCommandCount();
        this.loadDependencies()
            .done(() => {
                this.continue();
            })
            .always(() => {
                if (this.settings.enableDependencyTimeout) {
                    clearTimeout(expiry);
                }
            });
        return this.resultPromise.promise() as JQueryPromise<Array<ActionResult>>;
    }
    public onActionExecuted(onExecuted: any): Fluent {
        this.onExecuted = onExecuted;
        return this;
    }
    public onActionExecuting(onExecuting: any): Fluent {
        this.onExecuting = onExecuting;
        return this;
    }
    public when(predicate: any) {
        if (this.peekLastCommand() && this.peekLastCommand().constructor !== ActionCommand) {
            throw "Illegal operation: The last command must be an ActionCommand";
        }
        var command = new WhenCommand();
        command.action = () => {
            var deferred = $.Deferred();
            var command = this.peekLastResult();
            if (command) {
                if (!command.success) {
                    throw "Illegal operation: The last command should have succeeded for this call to have been made.  This is an issue in the fluent api";
                }
                if (predicate(...command.result)) {
                    deferred.resolve();
                }
                else {
                    this.reject(deferred, this, "Predicate returned false");
                }
            } else {
                this.failChain(this, "No action to process for when");
            }
            return deferred.promise();
        };
        this.commands.push(command);
        return this;
    }
    //Similar to when, but all previous results are passed into the predicate
    public whenAll(predicate: any) : Fluent {
        if (this.peekLastCommand() && this.peekLastCommand().constructor !== ActionCommand) {
            throw "Illegal operation: The last command must be an ActionCommand";
        }
        var command = new WhenCommand();
        command.action = () => {
            var deferred = $.Deferred();
            var command = this.peekLastResult();
            if (command) {
                if (!command.success) {
                    throw "Illegal operation: The last command should have succeeded for this call to have been made.  This is an issue in the fluent api";
                }
                if (predicate(this.results)) {
                    deferred.resolve();
                }
                else {
                    this.reject(deferred, this, "Predicate returned false");
                }
            } else {
                this.failChain(this, "No action to process for whenAll");
            }
            return deferred.promise();
        };
        this.commands.push(command);
        return this;
    }
    public whenTrue(): Fluent {
        if (this.peekLastCommand() && this.peekLastCommand().constructor !== ActionCommand) {
            throw "Illegal operation: The last command must be an ActionCommand";
        }
        var command = new WhenCommand();
        command.action = () => {
            var deferred = $.Deferred();
            var command = this.peekLastResult();
            if (command) {
                if (!command.success) {
                    throw "Illegal operation: The last command should have succeeded for this call to have been made.  This is an issue in the fluent api";
                }
                if (command.result && command.result.length && command.result[0]) {
                    deferred.resolve();
                }
                else {
                    this.reject(deferred, this, "Result is not true");
                }
            } else {
                this.failChain(this, "No action to process for whenTrue");
            }
            return deferred.promise();
        };
        this.commands.push(command);
        return this;
    }
    public whenFalse(): Fluent {
        if (this.peekLastCommand() && this.peekLastCommand().constructor !== ActionCommand) {
            throw "Illegal operation: The last command must be an ActionCommand";
        }
        var command = new WhenCommand();
        command.action = () => {
            var deferred = $.Deferred();
            var command = this.peekLastResult();
            if (command) {
                if (!command.success) {
                    throw "Illegal operation: The last command should have succeeded for this call to have been made.  This is an issue in the fluent api";
                }
                if (command.result && command.result.length && !command.result[0]) {
                    deferred.resolve();
                }
                else {
                    this.reject(deferred, this, "Result is not false");
                }
            } else {
                this.failChain(this, "No action to process for whenFalse");
            }
            return deferred.promise();
        };
        this.commands.push(command);
        return this;
    }
    public chainAction(name: string, action: any) : Fluent {
        var command = new ActionCommand();
        command.name = name;
        command.action = action;
        this.commands.push(command);
        return this;
    }
    public chain(command: ActionCommand): Fluent {
        this.commands.push(command);
        return this;
    }
    public static executeQuery(context: SP.ClientContext) {
        return common.executeQuery(context);
    }
    public registerDependency(dependency: Dependency) {
        if (this.dependencies.indexOf(dependency) === -1) {
            this.dependencies.push(dependency);
        }
    }
    private continue() {
        if (!this.context) {
            throw "context not set, you must call withContext";
        }
        var command = this.commands.shift();
        if (command && command.action) {
            if (command.constructor === ActionCommand) {
                if (this.onExecuting) {
                    let step = this.totalCommands - (this.getAvailableActionCommandCount());
                    this.onExecuting((command as ActionCommand).name, step, this.totalCommands);
                }
            }
            var promise = command.action() as JQueryPromise<any>;
            promise.done((arg1, arg2, arg3, arg4, arg5, arg6, arg7) => {
                if (command.constructor === ActionCommand) {
                    let results = [];
                    this.storeResult(arg1, results);
                    this.storeResult(arg2, results);
                    this.storeResult(arg3, results);
                    this.storeResult(arg4, results);
                    this.storeResult(arg5, results);
                    this.storeResult(arg6, results);
                    this.storeResult(arg7, results);
                    this.addResult(command as ActionCommand, true, results);
                    if (this.onExecuted) {
                        this.onExecuted((command as ActionCommand).name, true, results);
                    }
                }
                if (this.commands.length) {
                    this.continue();
                }
                else {
                    this.resolveChain();
                }
            })
                .fail((sender, args) => {
                    if (command.constructor === ActionCommand) {
                        let results = [];
                        this.storeResult(sender, results);
                        this.storeResult(args, results);
                        this.addResult(command as ActionCommand, false, results);
                        if (this.onExecuted) {
                            this.onExecuted((command as ActionCommand).name, false, {sender:sender, args:args});
                        }
                        this.failChain(sender, args);
                        return;
                    }
                    this.resolveChain();
                });
        } else {
            this.resolveChain();
        }
    }
    private storeResult(arg: any, results: Array<any>) {
        if (typeof (arg) !== typeof (undefined)) {
            results.push(arg);
        }
    }
    private resolveChain() {
        this.resultPromise.resolve(this.results);
    }
    private failChain(sender, args) {
        if (typeof (args) == "string") {
            args = { get_message: () => { return args } };
        }
        this.resultPromise.reject(sender, args);
    }
    private reject(deferred: JQueryDeferred<any>, sender, args) {
        if (typeof (args) == "string") {
            args = { get_message: () => { return args } };
        }
        deferred.reject(sender, args);
    }
    private addResult(command: ActionCommand, success: boolean, results: Array<any>) {
        var result = new ActionResult();
        result.name = command.name;
        result.success = success;
        result.result = results;
        this.results.push(result);
    }
    private loadDependencies(): JQueryPromise<any> {
        var deferred = $.Deferred();
        let spDependencies = ["SP.js", "SP.Runtime.js"];
        for (let i = 0; i < this.dependencies.length; i++) {
            switch (this.dependencies[i]) {
                case Dependency.UserProfile:
                    (<any>window).LoadSodByKey("userprofile");
                    spDependencies.push("userprofile");
                    break;
                case Dependency.Publishing:
                    SP.SOD.registerSod('SP.Publishing.js', SP.Utilities.Utility.getLayoutsPageUrl('sp.publishing.js'));
                    spDependencies.push("SP.Publishing.js");
                    break;
                case Dependency.Taxonomy:
                    SP.SOD.registerSod('sp.taxonomy.js', SP.Utilities.Utility.getLayoutsPageUrl('sp.taxonomy.js'));
                    spDependencies.push("sp.taxonomy.js");
                    break;
                default:
                    throw "Depenency not implemented";
            }
        }
        SP.SOD.loadMultiple(spDependencies, () => {
            deferred.resolve();
        });
        return deferred.promise();
    }
    private peekLastCommand() {
        if (this.commands.length) {
            return this.commands[this.commands.length - 1];
        }
        return null;
    }
    private peekLastResult() {
        if (this.results.length) {
            return this.results[this.results.length - 1];
        }
        return null;
    }
    private getAvailableActionCommandCount(): number {
        var count = 0;
        for (var i = 0; i < this.commands.length; i++) {
            if (this.commands[i].constructor === ActionCommand) {
                count++;
            }
        }
        return count;
    }
}
export enum NavigationLocation {
    TopNavigation,
    Quicklaunch
}
export enum NavigationType {
    Inherit,
    Managed,
    StructuralWithSiblings,
    StructuralChildrenOnly
}
export enum Dependency {
    Publishing,
    UserProfile,
    Taxonomy
}
export interface ISettings {
    timeoutMilliseconds?: number,
    enableDependencyTimeout?: boolean
}
export class FluentCommand {
    action: any;
}
export class ActionCommand extends FluentCommand {
    name: string;
}
export class ActionResult {
    name: string;
    success: boolean;
    result: Array<any>;
}
export interface KeyValuePair {
    key: string;
    value: any;
}
export class WhenCommand extends FluentCommand { }

