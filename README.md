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

In this example, we want to create a list, then add a list item.  We also want to check to make sure the list doesn't already exist and the user has permission.
Each command is one *link* in the chain.  The *when* commands will stop execution of the chain if they fail to meet their criteria.
```javascript
	var context = SP.ClientContext.get_current();
    var listName = "MyList1";
    new spJsom.Fluent()
        .withContext(context)
        .list.exists(context.get_web(), listName)
        .whenFalse()  //stops here if the list exists
        .permission.hasListPermission(SP.PermissionKind.addListItems, context.get_web().get_lists().getByTitle("MyList"))
        .whenTrue() //stops here if the user doesn't have permission
        .list.create(context.get_web(), listName, 100)
        .listItem.create(context.get_web(), listName, { Title: "Created by spJsom" })
        .execute()
			.fail((sender, args) => {
				console.error(args.get_message());
			})
			.done((results) => {
				console.log(results);
			});
```

## TypeScript Examples

See [https://github.com/TjWheeler/spJsomFluent/examples/spJsomExamples-typescript.ts](TypeScript examples)

## JavaScript Examples

Coming soon!

#Usage 

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


