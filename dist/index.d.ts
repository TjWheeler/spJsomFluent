// Generated by dts-bundle v0.7.3

export class common {
    static reject(deferred: JQueryDeferred<any>, reason: string): void;
    static notImplementedPromise(): JQueryPromise<any>;
    static executeQuery: (clientContext: SP.ClientContext) => JQuery.Promise<any, any, any>;
    static FilterArray<T>(items: Array<T>, predicate: any): Array<T>;
}

export class File {
    constructor(fluent: Fluent);
    /**
     * Get a file from a document library
     * Result: SP.File
     * Example: get(_spPageContextInfo.siteServerRelativeUrl + "/documents/doc.docx")
     */
    get(serverRelativeUrl: string): Fluent;
    /**
     * Get the list item for the file
     * Result: SP.ListItem
     * Example: getListItem(_spPageContextInfo.siteServerRelativeUrl + "/documents/doc.docx")
     */
    getListItem(web: SP.Web, serverRelativeUrl: string): Fluent;
    /**
     * Check in a file
     * Result: SP.File
     * Example: checkIn(SP.ClientContext.get_current().get_web(), _spPageContextInfo.webServerRelativeUrl + '/pages/mypage.aspx', "Checked in by spJsomFluent", SP.CheckinType.majorCheckIn)
     */
    checkIn(web: SP.Web, serverRelativeUrl: string, comment: string, checkInType: SP.CheckinType): Fluent;
    /**
     * Check out a file
     * Result: SP.File
     * Example: checkOut(SP.ClientContext.get_current().get_web(), _spPageContextInfo.webServerRelativeUrl + '/pages/mypage.aspx', "Checked in by spJsomFluent", SP.CheckinType.majorCheckIn)
     */
    checkOut(web: SP.Web, serverRelativeUrl: string): Fluent;
    /**
     * Upload a file to a library
     * Result: SP.File
     * Example: uploadFile(SP.ClientContext.get_current().get_web(), "Documents", file)
     */
    uploadFile(web: SP.Web, listName: string, file: any, filename?: string): Fluent;
}

export class Fluent {
    context: SP.ClientContext;
    settings: ISettings;
    withContext(context: SP.ClientContext): Fluent;
    withSettings(settings: ISettings): Fluent;
    withDependency(dependency: Dependency): Fluent;
    readonly promise: JQueryPromise<any>;
    readonly permission: Permission;
    readonly list: List;
    readonly listItem: ListItem;
    readonly file: File;
    readonly publishingPage: PublishingPage;
    readonly web: Web;
    readonly user: User;
    readonly navigation: Navigation;
    execute(): JQueryPromise<Array<ActionResult>>;
    onActionExecuted(onExecuted: OnExecuted): Fluent;
    onActionExecuting(onExecuting: OnExecuting): Fluent;
    update(predicate: any): this;
    when(predicate: any): this;
    whenAll(predicate: any): Fluent;
    whenTrue(): Fluent;
    whenFalse(): Fluent;
    chainAction(name: string, action: any): Fluent;
    chain(command: ActionCommand): Fluent;
    static executeQuery(context: SP.ClientContext): JQuery.Promise<any, any, any>;
    registerDependency(dependency: Dependency): void;
}
export enum NavigationLocation {
    TopNavigation = 0,
    Quicklaunch = 1
}
export enum NavigationType {
    Inherit = 0,
    Managed = 1,
    StructuralWithSiblings = 2,
    StructuralChildrenOnly = 3
}
export enum Dependency {
    Publishing = 0,
    UserProfile = 1,
    Taxonomy = 2
}
export interface ISettings {
    timeoutMilliseconds?: number;
    enableDependencyTimeout?: boolean;
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
export class WhenCommand extends FluentCommand {
}
export type OnExecuting = (commandName: string, step: number, total: number) => any;
export type OnExecuted = (commandName: string, success: boolean, result: Array<any>) => any;

export interface ICamlBuilder {
    query: string;
    viewXml: string;
    viewScope: string;
    viewFields: string;
    totalClauses: number;
    rowLimit: number;
    begin(requireAll: boolean): any;
    addNullClause(fieldRef: string): any;
    addTextClause(operation: CamlOperator, fieldRef: string, value: string): any;
    addNumberClause(operation: CamlOperator, fieldRef: string, value: number): any;
    addDateTimeClause(operation: CamlOperator, fieldRef: string, value: string, includeTime: boolean): any;
    addLookupClause(operation: CamlOperator, fieldRef: string, value: string, valueType: string): any;
    recurrenceQuery(overlapType: EventRecurranceOverlap): any;
    addViewField(fieldRef: string): any;
    addViewFields(fieldRefs: Array<string>): any;
    addGroupByField(fieldRef: string): any;
    addGroupByFields(fieldRefs: Array<string>): any;
    addOrderBy(fieldRef: string, ascending: boolean): any;
    addDaysToDate(date: any, days: any): any;
}
export enum CamlOperator {
    Contains = 0,
    BeginsWith = 1,
    Eq = 2,
    Geq = 3,
    Leq = 4,
    Lt = 5,
    Gt = 6,
    Neq = 7,
    DateRangesOverlap = 8,
    IsNotNull = 9,
    IsNull = 10
}
export enum EventRecurranceOverlap {
    Now = 0,
    Today = 1,
    Week = 2,
    Month = 3,
    Year = 4
}
export enum AggregationType {
    SUM = 0,
    COUNT = 1,
    AVG = 2,
    MAX = 3,
    MIN = 4,
    STDEV = 5,
    VAR = 6
}
export class StringBuilder {
    value: string;
    add(value: string): void;
    toString(): string;
}
export interface OdataItem {
    property: string;
    name: string;
    value: any;
}
export interface OdataPair {
    left: OdataPair;
    right: OdataPair;
    type: string;
    func: string;
    property: string;
    name: string;
    value: any;
    args: Array<OdataPair>;
}
export class CamlBuilder implements ICamlBuilder {
    constructor();
    rowLimit: number;
    viewScope: string;
    readonly query: string;
    readonly viewFields: string;
    readonly totalClauses: number;
    setFilter(caml: string): void;
    readonly viewXml: string;
    begin(requireAll: boolean): void;
    getNullClause(fieldRef: string): string;
    getClause(operation: CamlOperator, fieldRef: string, value: string, valueType: string): string;
    getDateTimeClause(operation: CamlOperator, fieldRef: string, value: string, includeTime: boolean): string;
    addNullClause(fieldRef: string): void;
    addTextClause(operation: CamlOperator, fieldRef: string, value: string): void;
    addBooleanClause(operation: CamlOperator, fieldRef: string, value: boolean): void;
    addNumberClause(operation: CamlOperator, fieldRef: string, value: number): void;
    addDateTimeClause(operation: CamlOperator, fieldRef: string, value: string, includeTime: boolean): void;
    addLookupClause(operation: CamlOperator, fieldRef: string, value: string, valueType: string): void;
    addAggregation(fieldRef: string, aggregationType: AggregationType): void;
    recurrenceQuery(overlapType: EventRecurranceOverlap): void;
    addViewField(fieldRef: string): void;
    addViewFields(fieldRefs: Array<string>): void;
    addGroupByField(fieldRef: string): void;
    addGroupByFields(fieldRefs: Array<string>): void;
    addOrderBy(fieldRef: string, ascending: boolean): void;
    addDaysToDate(date: any, days: any): any;
}
export class OrderBy {
    fieldRef: string;
    ascending: boolean;
}

export class ListHelper {
    constructor(context: SP.ClientContext);
    createList(web: SP.Web, listName: string, templateId: number): JQueryPromise<SP.List>;
    setAlerts(web: SP.Web, listName: string, enabled: boolean): JQueryPromise<any>;
    setListItemProperties(listItem: SP.ListItem, properties: any): JQueryPromise<SP.ListItem>;
    exists(web: SP.Web, listName: string): JQueryPromise<SP.ListItem>;
    createListItemWithContentTypeName(web: SP.Web, listName: string, contentTypeName: string, properties: any): JQueryPromise<SP.ListItem>;
    createListItem(web: SP.Web, listName: string, properties: any): JQueryPromise<SP.ListItem>;
    getFile(serverRelativeUrl: string): JQueryPromise<SP.File>;
    checkInFile(web: SP.Web, serverRelativeUrl: string, comment: string, checkInType: SP.CheckinType): JQueryPromise<SP.File>;
    uploadFile(web: SP.Web, listname: string, file: File, filename: string): JQueryPromise<SP.File>;
    checkOutFile(web: SP.Web, serverRelativeUrl: string): JQueryPromise<SP.File>;
    getList(web: SP.Web, listName: string): JQueryPromise<SP.List>;
    deleteList(web: SP.Web, listName: string): JQueryPromise<SP.List>;
    getFileListItem(web: SP.Web, serverRelativeUrl: string, viewFields?: Array<string>): JQueryPromise<SP.ListItem>;
    getListItemById(web: SP.Web, listName: string, id: number, viewFields?: Array<string>): JQueryPromise<SP.ListItem>;
    deleteListItemById(web: SP.Web, listName: string, id: number): JQueryPromise<any>;
    getListItems(web: SP.Web, listName: string, viewXml: string): JQueryPromise<SP.ListItemCollection>;
    addContentTypeListAssociation(web: SP.Web, listName: string, contentTypeName: string): JQueryPromise<SP.ContentType>;
    removeContentTypeListAssociation(web: SP.Web, listName: string, contentTypeName: string): JQueryPromise<any>;
    setDefaultValueOnList(web: SP.Web, listName: string, fieldInternalName: string, defaultValue: any): JQueryPromise<any>;
}

export class NavigationHelper {
    constructor(context: SP.ClientContext);
    deleteQuicklaunchNodes(web: SP.Web): JQueryPromise<any>;
    deleteTopNavigationNodes(web: SP.Web): JQueryPromise<any>;
    deleteQuicklaunchNode(web: SP.Web, title: string): JQueryPromise<any>;
    deleteTopQuicklaunchNode(web: SP.Web, title: string): JQueryPromise<any>;
    setCurrentNavigation(web: SP.Web, navigationType: NavigationType, showSubsites?: boolean, showPages?: boolean): JQueryPromise<any>;
    createQuicklaunchNode(web: SP.Web, title: string, url: string, asLastNode?: boolean): JQueryPromise<SP.NavigationNode>;
    createTopNavigationNode(web: SP.Web, title: string, url: string, asLastNode?: boolean): JQueryPromise<SP.NavigationNode>;
}

export class PageHelper {
    constructor(context: SP.ClientContext);
    createPublishingPage(web: SP.Web, name: string, layoutUrl: string, title: string): JQueryPromise<SP.Publishing.PublishingPage>;
    getPublishingLayout(serverRelativeUrl: string): JQueryPromise<SP.ListItem>;
    setLayout(web: SP.Web, name: string, layoutUrl: string): JQueryPromise<any>;
}

export class PermissionHelper {
    constructor(context: SP.ClientContext);
    hasWebPermission(permission: SP.PermissionKind, web: SP.Web): JQueryPromise<boolean>;
    hasListPermission(permission: SP.PermissionKind, list: SP.List): JQueryPromise<boolean>;
    hasItemPermission(permission: SP.PermissionKind, item: SP.ListItem): JQueryPromise<boolean>;
}

interface IProgressCallback {
    (percentComplete: number): void;
}
export class SPRestFileUploader {
    uploadFileAsArrayBuffer(file: File, webUrl: string, listname: string, filename: string, progressCallback?: IProgressCallback): JQueryPromise<any>;
    updateListItem(webUrl: string, serverRelativeUrl: string, properties: any, etag?: string): JQuery.jqXHR<any>;
}
export {};

export class UserHelper {
    constructor(context: SP.ClientContext);
    getUserByEmail(email: string): JQueryPromise<SP.User>;
    getUserById(id: number): JQueryPromise<SP.User>;
    getCurrentUser(): JQueryPromise<SP.UserProfiles.PersonProperties>;
    getCurrentUserProfileProperties(): JQueryPromise<SP.UserProfiles.PersonProperties>;
    getCurrentUserManager(): JQueryPromise<SP.User>;
}

export class WebHelper {
    constructor(context: SP.ClientContext);
    setWelcomePage(web: SP.Web, url: string): JQueryPromise<SP.Folder>;
    setTitle(web: SP.Web, title: string): JQueryPromise<any>;
    setLogoUrl(web: SP.Web, url: string): JQueryPromise<any>;
    activateFeature(web: SP.Web, featureId: SP.Guid, force: boolean): JQueryPromise<any>;
    createWeb(name: string, parentWeb: SP.Web, title: string, template: string, useSamePermissionsAsParent?: boolean): JQueryPromise<SP.Web>;
    doesWebExist(url: string): JQueryPromise<any>;
    getWebs(fromWeb: SP.Web): JQueryPromise<Array<SP.Web>>;
}

export class List {
    constructor(fluent: Fluent);
    /**
     * Create new list
     * Example: create(context.get_web(), "My Task List", 107)
     * templateId - See: https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-csom/ee541191(v%3Doffice.15)#remarks
     */
    create(web: SP.Web, listName: string, templateId: number): Fluent;
    /**
     * Check if the list exists
     * Result: boolean
     * Example: exists(context.get_web(), "My List")
     */
    exists(web: SP.Web, listName: string): Fluent;
    /**
     * Delete a list
     * Result: void
     * Example: delete(context.get_web(), "My List")
     */
    delete(web: SP.Web, listName: string): Fluent;
    /**
     * Get a list
     * Result: SP.List
     * Example: get(context.get_web(), "My List")
     */
    get(web: SP.Web, listName: string): Fluent;
    /**
     * Adds a Content Type to a List.  Resolves the new associated list content type.
     * Result: SP.ContentType
     * Example: addContentTypeListAssociation(context.get_web(), "My List", "My Content Type Name")
     */
    addContentTypeListAssociation(web: SP.Web, listName: string, contentTypeName: string): Fluent;
    /**
     * Removes a Content Type associated from a list.
     * Result: void
     * Example: removeContentTypeListAssociation(context.get_web(), "My List", "My Content Type Name")
     */
    removeContentTypeListAssociation(web: SP.Web, listName: string, contentTypeName: string): Fluent;
    /**
     * Sets a default value on a field in a list
     * Result: void
     * Example: setDefaultValueOnList(context.get_web(), "My Task List", "ClientId", 123)
     */
    setDefaultValueOnList(web: SP.Web, listName: string, fieldInternalName: string, defaultValue: any): Fluent;
    /**
    * Enable email alerts on a list
    * Result: void
    * Example: setAlerts(context.get_web(), "My Task List", true)
    * Note: will not work for 2010/2013
    */
    setAlerts(web: SP.Web, listName: string, enabled: boolean): Fluent;
}

export class ListItem {
    constructor(fluent: Fluent);
    /**
     * Update a list item with properties
     * Result: SP.ListItem
     * Example: update(listItem, { 'Title': 'Title value here', ClientId: 123 })
     */
    update(listItem: SP.ListItem, properties: any): Fluent;
    /**
     * Create a list item, specifying the Content Type
     * Result: SP.ListItem
     * Example: createWithContentType(context.get_web(), "My list", "My Content Type Name", { 'Title': 'Title value here', ClientId: 123 })
     */
    createWithContentType(web: SP.Web, listName: string, contentTypeName: string, properties: any): Fluent;
    /**
     * Create new list item with optional property values
     * Example: create(context.get_web(), "MyList", properties)
     * Note: Properties are an object with Property/Value, where property
     *       is the internal field name.
     *       Eg; var properties = {
                Title: "My title",
                PersonOrGroupField: personValue,
                MultiChoiceField: ["Second", "Third"],
                ChoiceField: "Second",
                NumberField: 1234
            };
            For personValue you can pass through the user Id or a SP.UserFieldValue such as:
            var personValue = new SP.FieldUserValue();
            personValue.set_lookupId(_spPageContextInfo.userId);
     */
    create(web: SP.Web, listName: string, properties: any): Fluent;
    /**
     * Get the listitem using the ID with optional view fields
     * Result: SP.ListItem
     * Example: get(context.get_web(), "MyList", 1, ["ID", "Title"])
     */
    get(web: SP.Web, listName: string, id: number, viewFields?: Array<string>): Fluent;
    /**
     * Get the listitem for a File using the ID with optional view fields
     * Result: SP.ListItem
     * Example: getFileListItem(context.get_web(), "/sites/site/documents/mydoc.docx", ["ID", "Title", "FileLeafRef"])
     */
    getFileListItem(web: SP.Web, serverRelativeUrl: string, viewFields?: Array<string>): Fluent;
    /**
     * Delete List Item
     * Result: void
     * Example: deleteById(context.get_web(), "MyList", 7)
     */
    deleteById(web: SP.Web, listName: string, id: number): Fluent;
    /**
     * Execute a query using supplied CAML.
     * Returns: SP.ListItemCollection
     * Example: query(context.get_web(), "MyList", "<View><Query><Where><In><FieldRef Name='ID' /><Values><Value Type='Number'>1</Value><Value Type='Number'>2</Value></Values></In></View></Query></Where>")
     */
    query(web: SP.Web, listName: string, viewXml: string): Fluent;
    /**
     * Get all list items in a list
     * Returns: SP.ListItemCollection
     * Example: getAll(context.get_web(), "MyList")
     */
    getAll(web: SP.Web, listName: string): Fluent;
}

export class Navigation {
    constructor(fluent: Fluent);
    /**
     * Create new navigation node
     * Result: SP.NavigationNode
     * Example: createNode(context.get_web(), NavigationLocation.Quicklaunch, "Test Node", "/sites/mysite/pages/default.aspx")
     */
    createNode(web: SP.Web, location: NavigationLocation, title: string, url: string, asLastNode?: boolean): Fluent;
    /**
     * Delete all navigation nodes for the web
     * Result: void
     * Example: deleteAllNodes(context.get_web(), NavigationLocation.Quicklaunch)
     */
    deleteAllNodes(web: SP.Web, location: NavigationLocation): Fluent;
    /**
     * Delete all navigation nodes that match the supplied title for the web
     * Result: void
     * Example: deleteNode(context.get_web(), NavigationLocation.Quicklaunch, "My link title")
     */
    deleteNode(web: SP.Web, location: NavigationLocation, title: string): Fluent;
    /**
     * Set navigation for the web
     * Result: void
     * Example: setCurrentNavigation(context.get_web(), 3, true, true)
     * Note: showSubsites and showPages is only applicable for NavigationType.StructuralChildrenOnly (3)
     */
    setCurrentNavigation(web: SP.Web, type: NavigationType, showSubsites?: boolean, showPages?: boolean): Fluent;
}

export class Permission {
    constructor(fluent: Fluent);
    /**
     * Check if the current user has specified Web permission
     * Result: boolean
     * Example: hasWebPermission(SP.PermissionKind.createSSCSite, context.get_web())
     */
    hasWebPermission(permission: SP.PermissionKind, web: SP.Web): Fluent;
    /**
     * Check if the current user has specified List permission
     * Result: boolean
     * Example: hasListPermission(SP.PermissionKind.addListItems, context.get_web().get_lists().getByTitle("MyList"))
     */
    hasListPermission(permission: SP.PermissionKind, list: SP.List): Fluent;
    /**
     * Check if the current user has specified ListItem permission
     * Result: boolean
     * Example: hasItemPermission(SP.PermissionKind.editListItems, context.get_web().get_lists().getByTitle("MyList").getItemById(0))
     */
    hasItemPermission(permission: SP.PermissionKind, item: SP.ListItem): Fluent;
}

export class PublishingPage {
    constructor(fluent: Fluent);
    /**
     * Creates a new publishing page.
     * Result: SP.Publishing.PublishingPage
     * Example: .publishingPage.create(SP.ClientContext.get_current().get_web(), "Home.aspx", _spPageContextInfo.siteServerRelativeUrl + "/_catalogs/masterpage/BlankWebPartPage.aspx")
     */
    create(web: SP.Web, name: string, layoutUrl: string, title: string): Fluent;
    /**
     * Gets a page layout.
     * Result: SP.ListItem
     * Example: .publishingPage.getLayout(_spPageContextInfo.siteServerRelativeUrl + "/_catalogs/masterpage/BlankWebPartPage.aspx")
     *
     */
    getLayout(serverRelativeUrl: string): Fluent;
    /**
     * Changes layout for a publishing page.
     * Result: void
     * Example: .publishingPage.setLayout(SP.ClientContext.get_current().get_web(), "Home.aspx", _spPageContextInfo.siteServerRelativeUrl + "/_catalogs/masterpage/BlankWebPartPage.aspx")
     */
    setLayout(web: SP.Web, name: string, layoutUrl: string): Fluent;
}

export class User {
    constructor(fluent: Fluent);
    /**
     * Get User by email or account name
     * Result: SP.User
     * Example: get("my@email.address")
     * Example: get("i:0#.f|membership|my@email.address")
     */
    get(emailOrAccountName: string): Fluent;
    /**
     * Get a user by their Id
     * Result: SP.User
     * Example: getById(15)
     */
    getById(id: number): Fluent;
    /**
     * Get the current user
     * Result: SP.User
     * Example: getCurrentUser()
     */
    getCurrentUser(): Fluent;
    /**
     * Get the profile properties for the current user
     * Result: SP.UserProfiles.PersonProperties
     * Example: getCurrentUserProfileProperties()
     */
    getCurrentUserProfileProperties(): Fluent;
    /**
     * Get the manager for the current user
     * Result: SP.User
     * Example: getCurrentUserManager()
     */
    getCurrentUserManager(): Fluent;
}

export class Web {
        constructor(fluent: Fluent);
        /**
             * Set the web welcome page.  The url should be relative to the root folder of the web.
             * Result: SP.Folder
             * Example: setWelcomePage(context.get_web(), "pages/default.aspx")
             */
        setWelcomePage(web: SP.Web, url: string): Fluent;
        /**
             * Create a new sub web.
             * Result: SP.Web
             * name: Forms part of the URL, don't include slash
             * template: template id and configuration value
             * Example: create("SubWebName", SP.ClientContext.get_current().get_rootWeb(), "My Web Name", "CMSPUBLISHING#0")
             *
             * @remarks
             * For templates See: https://blogs.technet.microsoft.com/praveenh/2010/10/21/sharepoint-templates-and-their-ids/
             */
        create(name: string, parentWeb: SP.Web, title: string, template: string, useSamePermissionsAsParent?: boolean): Fluent;
        /**
             * Check if a web exists
             * Result: boolean
             * url: server relative url
             * Example: .exists("/sites/mysubweb")
             */
        exists(url: string): Fluent;
        /**
             * Get and load all webs starting from a web and its children
             * Result: Array<SP.Web>
             * Example: getWebs(SP.ClientContext.get_current().get_rootWeb())
             */
        getWebs(fromWeb: SP.Web): Fluent;
        /**
             * Set site title
             * Result: void
             * Example: setTitle(SP.ClientContext.get_current().get_web(), "My Title")
             */
        setTitle(web: SP.Web, title: string): Fluent;
        /**
             * Set site logo url
             * Result: void
             * Example: setLogoUrl(SP.ClientContext.get_current().get_web(), "/site/images/logo.jpg")
             */
        setLogoUrl(web: SP.Web, url: string): Fluent;
        /**
             * Activate feature
             * Result: void
             * Example: activateFeature(SP.ClientContext.get_current().get_web(), new SP.Guid('15a572c6-e545-4d32-897a-bab6f5846e18'))
             */
        activateFeature(web: SP.Web, featureId: SP.Guid, force?: boolean): Fluent;
}

