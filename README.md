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

## List Commands

Coming soon

## Permission Commands

Coming soon

## Web Commands

Coming soon

## ListItem Commands

Coming soon

## Navigation Commands

Coming soon

## Publishing Page Commands

| Fluent API | .publishingpages |

| Command        | Example  |
| ------------- |:-------------:|
| col 3 is      | right-aligned |


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


