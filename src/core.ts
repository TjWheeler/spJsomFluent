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

