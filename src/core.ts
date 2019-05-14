/// <reference path="../typings/@types/all.ts" />

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
export enum Dependency {
    Publishing,
    UserProfile,
    Taxonomy
}
//export interface IFluentInternal extends IFluent {
//    chainAction(name: string, action: any);
//    registerDependency(dependency: Dependency);
//}
//export interface IFluent {
//    context: SP.ClientContext;
//    withContext(context: SP.ClientContext): IFluent;
//    withSettings(settings: ISettings): IFluent;
//    settings: ISettings;
//    promise: JQueryPromise<any>;
//    execute(): JQueryPromise<Array<ActionResult>>;
//    onActionExecuted(onExecuted: any): IFluent;
//    onActionExecuting(onExecuting: any): IFluent;
//    when(predicate: any);
//    whenAll(predicate: any);
//    whenTrue(): IFluent;
//    whenFalse(): IFluent;
//}
export interface ISettings {
    timeoutMilliseconds?: number,
    enableDependencyTimeout?: boolean
}