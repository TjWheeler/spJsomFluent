# spJsomFluent

Note: This library is not yet ready for use and is still in initial development.

Repository: https://github.com/TjWheeler/spJsomFluent

## Install - npm
`npm i spjsomfluent --save`

## Usage
```
import * as spJsom from "./node_modules/spjsomfluent/fluent"
```

# TypeScript Examples

## Set Welcome Page
```
	new spJsom.Fluent()
		.withContext(SP.ClientContext.get_current())
		.web.setWelcomePage(SP.ClientContext.get_current().get_web(), `pages/MyPage.aspx`)
		.execute()
		.fail((sender, args) => {
			console.error(args.get_message());
		})
		.done((results) => {
			console.log(results);
		});
```

## Create Publishing Page & set the web welcome page
```
	var context = SP.ClientContext.get_current();
    var webUrl = _spPageContextInfo.webServerRelativeUrl;
    var filename = "MyPage.aspx";
    new spJsom.Fluent()
        .withContext(context)
        .publishingPage.create(context.get_web(), filename, `${webUrl}/_catalogs/masterpage/blankwebpartpage.aspx`)
        .file.checkIn(context.get_web(), `${webUrl}/pages/${filename}`, "Checked in by spJsomFluent", SP.CheckinType.majorCheckIn)
        .web.setWelcomePage(context.get_web(), `pages/${filename}`)
        .execute()
        .fail((sender, args) => {
            console.error(args.get_message());
        })
        .done((results) => {
            console.log(results);
        });
```


## Development commands

### Install modules
`npm install`

### Run the dev build
`npm run debug`

### Run the production build
`npm run release`

### Watch the dev configuration, and redeploy as files change
`npm run dev`


