# spJsomFluent

spJsom is a library that simplies usage of the SharePoint Client Side Object Model (JavaScript/JSOM).
Using a fluent API, the library allows chaining multiple commands together with conditions and progress monitoring. 
Goals of this library are:
- Simply using the CSOM
- Avoid nesting promises
- Reduce boiler plate code
- Provide an easy to read syntax
- Speed up development

This solution can be used in TypeScript or plain JavaScript projects.

Repository: https://github.com/TjWheeler/spJsomFluent

# Examples
## A quick example

In this example we are get getting the current user profile properties.  All the difficulty behind loading the correct scripts and calling the PeopleManager are handled for you.

```javascript
var context = SP.ClientContext.get_current();
new spJsom.Fluent()
    .withContext(context)
    .user.getCurrentUserProfileProperties()
    .execute()
    .fail(function(sender, args) {
        console.error(args.get_message());
    })
    .done(function(results) {
        console.log(results);
        console.log("Your profile properties", results[0].result[0]);
    });
```

## TypeScript Examples
For more examples, including more complex scenarios with multiple commands, 
see [TypeScript examples](https://github.com/TjWheeler/spJsomFluent/blob/master/examples/spJsomExamples-typescript.ts)

## JavaScript Examples

Coming soon!

# Dependencies

jQuery must be loaded as the library uses jQuery Promises.

# Commands

## whenTrue & whenFalse

These commands .whenTrue() and .whenFalse() are a conditional check.  They take the value from the previous step then compare it with True/False to determine if execution of the chain should continue.
For example:
```
.permission.hasWebPermission(SP.PermissionKind.createSSCSite, context.get_web())
.whenTrue() //stops executing the chain if the user doesn't have permission
```
Note that the when commands do not reject a promise even when their condition is not met.  The chain of execution simply stops at that point.

## when

You can supply your own predicate to determine if execution of the chain should continue.  
For example:
```javascript
.listItem.get(context.get_web(), "MyList", 1)
.when(function(listItem) {
	if(listItem.get_item("Title") === "My Title") {
		return true; //the chain execution will continue
	} 
	return false; //the chain execution will stop
})
```
The result of the previous step is passed into the *when* predicate allowing you the opportunity to evaluate the result and continue or stop.

## whenAll

Similar to *when*, but instead of the previous steps result being passed in, an array with all previous steps are passed to the predicate.
```javascript
...
.listItem.get(context.get_web(), "MyList", 1)
.whenAll(function(results) {
	var listItem = results[0];
	if(listItem.get_item("Title") === "My Title") {
		return true; //the chain execution will continue
	} 
	return false; //the chain execution will stop
})
```
## File Commands

Fluent API `spJsom.file`

| Command        | Result        | Description | Example |
| ------------- | ------------- | ------------- | ------------- |
| get(serverRelativeUrl: string)	| SP.File	| Get a file from a document library	| get(_spPageContextInfo.siteServerRelativeUrl + "/documents/doc.docx")	|
| getListItem(serverRelativeUrl: string)	| SP.ListItem | Get the list item for the file | getListItem(_spPageContextInfo.siteServerRelativeUrl + "/documents/doc.docx") |
| checkIn(web: SP.Web, serverRelativeUrl: string, comment: string, checkInType: SP.CheckinType) | SP.File | Check in a file | checkIn(SP.ClientContext.get_current().get_web(), _spPageContextInfo.webServerRelativeUrl + '/pages/mypage.aspx', "Checked in by spJsomFluent", SP.CheckinType.majorCheckIn)	|

## List Commands

Fluent API `spJsom.list`

| Command        | Result        | Description | Example |
| ------------- | ------------- | ------------- | ------------- |
| create(web: SP.Web, listName: string, templateId: number)	| SP.List | Create new list | create(context.get_web(), "My Task List", 107) |
| exists(web: SP.Web, listName: string) | boolean | Check if the list exists | exists(context.get_web(), "My List") |
| delete(web: SP.Web, listName: string)	| void | Delete a list | delete(context.get_web(), "My List") |
| get(web: SP.Web, listName: string) | SP.List | Get a list | get(context.get_web(), "My List") |
| addContentTypeListAssociation(web: SP.Web, listName: string, contentTypeName: string)| SP.ContentType | Adds a Content Type to a List.  Resolves the new associated list content type. | addContentTypeListAssociation(context.get_web(), "My List", "My Content Type Name") |
| removeContentTypeListAssociation(web: SP.Web, listName: string, contentTypeName: string)	| void | Removes a Content Type associated from a list. | removeContentTypeListAssociation(context.get_web(), "My List", "My Content Type Name") |
| setDefaultValueOnList(web: SP.Web, listName: string, fieldInternalName: string, defaultValue: any) | void | Sets a default value on a field in a list | setDefaultValueOnList(context.get_web(), "My Task List", "ClientId", 123)|
| setAlerts(web: SP.Web, listName: string, enabled: boolean)| void | Enable email alerts on a list.  Note: will not work for 2010/2013.  | setAlerts(context.get_web(), "My Task List", true) |

## ListItem Commands

Fluent API `spJsom.listItem`

| Command        | Result        | Description | Example |
| ------------- | ------------- | ------------- | ------------- |
| update(listItem: SP.ListItem, properties: any) | SP.ListItem | Update a list item with properties| update(listItem, { 'Title': 'Title value here', ClientId: 123 }) |
| createWithContentType(web: SP.Web, listName: string, contentTypeName: string, properties: any) | SP.ListItem | Create a list item, specifying the Content Type | createWithContentType(context.get_web(), "My list", "My Content Type Name", { 'Title': 'Title value here', ClientId: 123 })|
| create(web:SP.Web, listName: string, properties: any) | SP.ListItem | Create a list item with properties| update(listItem, { 'Title': 'Title value here', ClientId: 123 }) |
| get(web: SP.Web, listName: string, id: number, viewFields: Array<string> = null) | SP.ListItem | Get the listitem using the ID with optional view fields | get(context.get_web(), "MyList", 1, ["ID", "Title"]) |
| getFileListItem(serverRelativeUrl: string, viewFields: Array<string> = null) | SP.ListItem | Get the listitem for a File using the ID with optional view fields | getFileListItem("/sites/site/documents/mydoc.docx", ["ID", "Title", "FileLeafRef"]) |
| deleteById(web: SP.Web, listName: string, id: number) | void | Delete List Item | deleteById(context.get_web(), "MyList", 7) |
| query(web: SP.Web, listName: string, viewXml: string) | SP.ListItemCollection | Execute a query using supplied CAML.| query(context.get_web(), "MyList", "<View><Query><Where><In><FieldRef Name='ID' /><Values><Value Type='Number'>1</Value><Value Type='Number'>2</Value></Values></In></View></Query></Where>") |

## Permission Commands

Fluent API `spJsom.permission`

| Command        | Result        | Description | Example |
| ------------- | ------------- | ------------- | ------------- |
|||||
|||||

Fluent API `spJsom.web`

## Web Commands

| Command        | Result        | Description | Example |
| ------------- | ------------- | ------------- | ------------- |
|||||
|||||



## Navigation Commands

Fluent API `spJsom.navigation`

| Command        | Result        | Description | Example |
| ------------- | ------------- | ------------- | ------------- |
|||||
|||||

## Publishing Page Commands

Fluent API `spJsom.publishingPage`

| Command        | Result        | Description | Example |
| ------------- | ------------- | ------------- | ------------- |
| create(web: SP.Web, name: string, layoutUrl: string) | SP.Publishing.PublishingPage |Creates a new publishing page.   |  .publishingPage.create(SP.ClientContext.GetCurrent().get_web(), "Home.aspx", _spPageContextInfo.siteServerRelativeUrl + "/_catalogs/masterpage/BlankWebPartPage.aspx") |
| getLayout(serverRelativeUrl:string)		| SP.ListItem | Gets a page layout | |

# Events

To monitor activity there are 2 events you can wire up to.  These are:
1. onExecuting(commandName, step, totalSteps)
2. onExecuted(commandName, success, results)

## onExecuting
To show progress, this method will give you the details you need.
```typescript
.onActionExecuting((actionName:string, current:number, total:number) => {
    let progress = Math.ceil(100 * current / total);
    console.log(`Executing: ${actionName}, we are ${progress}% through.`);
})
            
```

## onExecuted
```typescript
.onActionExecuted((actionName:string, success:boolean, results: Array<any>) => {
    console.log(`Executed: ${actionName}:${success}`);
    if (!success) {
        console.error(actionName, results);
    }
})
```



# Usage 

## Install - npm
`npm i spjsomfluent --save`

## Importing spJsom
```
import * as spJsom from "./node_modules/spjsomfluent/src/fluent"
```

## Development commands - NodeJs

### Install modules
`npm install`

### Run the dev build
`npm run debug`

### Run the production build
`npm run release`

### Watch the dev configuration, and recompile as files change
`npm run dev`


