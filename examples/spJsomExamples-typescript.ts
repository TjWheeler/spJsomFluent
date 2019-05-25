import * as spJsom from "../src/fluent"
var examples = {
    run: () => {
        examples.getUser();
        //examples.createListAndItem();
        //examples.provisionWebWithProgress();
        //examples.query();
        //examples.provisionWebWithProgress();
        //examples.getListItem();
        //examples.createListAndItem();
        //examples.createListItem();
        //examples.setWelcomePage();
        //examples.createPublishingPage();
    },
    deleteListItem: () => {
        var context = SP.ClientContext.get_current();
        new spJsom.Fluent()
            .withContext(context)
            .listItem.deleteById(context.get_web(), "MyList", 7)
            .execute()
            .fail((sender, args) => {
                console.error(args.get_message());
            })
            .done((results) => {
                console.log(results);
            });
    },
    query: () => {
        var context = SP.ClientContext.get_current();
        new spJsom.Fluent()
            .withContext(context)
            .listItem.query(context.get_web(), "MyList", "<View><Query><Where><In><FieldRef Name='ID' /><Values><Value Type='Number'>1</Value><Value Type='Number'>2</Value></Values></In></View></Query></Where>")
            .execute()
            .fail((sender, args) => {
                console.error(args.get_message());
            })
            .done((results) => {
                console.log(results);
            });
    },
    provisionWebWithProgress: () => {
        var subSitePath = "testsite1";
        var url = (`${_spPageContextInfo.webServerRelativeUrl}/${subSitePath}`).replace("//", "/");
        var context = SP.ClientContext.get_current();
        var newSiteContext = new SP.ClientContext(url);  //Web doesn't exist yet, but we can create a context for it

        new spJsom.Fluent()
            .withContext(context)
            .onActionExecuting((actionName, current, total) => {
                let progress = Math.ceil(100 * current / total);
                console.log(`Executing: ${actionName}, we are ${progress}% through.`);
            })
            .onActionExecuted((actionName, success, results: Array<any>) => {
                console.log(`Executed: ${actionName}:${success}`);
                if (!success) {
                    console.error(actionName);
                    console.error(results);
                }
            })
            .permission.hasWebPermission(SP.PermissionKind.createSSCSite, context.get_web())
            .whenTrue() //cstops executing the chain if the user doesn't have permission
            .web.exists(url)
            .whenFalse() //stops executing the chain if the web already exists
            .web.create(subSitePath, context.get_web(), "Test Site", "CMSPUBLISHING#0") //Create a sub site
            .withContext(newSiteContext) //Switch to the new site context
            .publishingPage.create(newSiteContext.get_web(), "Home.aspx", _spPageContextInfo.siteServerRelativeUrl + "/_catalogs/masterpage/BlankWebPartPage.aspx")
            .web.setWelcomePage(newSiteContext.get_web(), "pages/home.aspx")
            .execute().done((results: Array<spJsom.ActionResult>) => {
                var lastAction = results[results.length - 1];
                console.log("The last action was " + lastAction.name + " and success was: " + lastAction.success, lastAction.result);
                if (lastAction.name === "web.exists" && lastAction.result[0] === true) {
                    console.warn("The web already exists.");
                }
            })
            .fail((sender, args) => {
                console.error(args.get_message());
            });
    },
    getListItem: () => {
        var context = SP.ClientContext.get_current();
        new spJsom.Fluent()
            .withContext(context)
            .listItem.get(context.get_web(), "MyList", 7, ["ID", "Title"])
            .execute()
            .fail((sender, args) => {
                console.error(args.get_message());
            })
            .done((results) => {
                console.log(results);
            });
    },
    createListAndItem: () => {
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
    },
    createListItem: () => {
        var context = SP.ClientContext.get_current();
        var personValue = new SP.FieldUserValue();
        personValue.set_lookupId(_spPageContextInfo.userId);
        var properties = {
            Title: "My title",
            PersonOrGroupField: personValue,
            MultiChoiceField: ["Second", "Third"],
            ChoiceField: "Second",
            NumberField: 12.45
        };
        new spJsom.Fluent()
            .withContext(context)
            .listItem.create(context.get_web(), "SomeList", properties)
            .execute()
            .fail((sender, args) => {
                console.error(args.get_message());
            })
            .done((results) => {
                console.log(results);
            });
    },
    getUser: () => {
        var context = SP.ClientContext.get_current();
        new spJsom.Fluent()
            .withContext(context)
            .user.getById(_spPageContextInfo.userId)
            .user.get(_spPageContextInfo.userEmail)
            .user.getCurrentUser()
            .user.getCurrentUserProfileProperties()
            .execute()
            .fail((sender, args) => {
                console.error(args.get_message());
            })
            .done((results) => {
                console.log(results);
                console.log(`Your name is: ${results[0].result[0].get_title()}`);
                console.log("Your profile properties", results[3].result[0]);
            });
    },
    setWelcomePage: () => {
        new spJsom.Fluent()
            .withContext(SP.ClientContext.get_current())
            .web.setWelcomePage(SP.ClientContext.get_current().get_web(), 'pages/MyPage.aspx')
            .execute()
            .fail((sender, args) => {
                console.error(args.get_message());
            })
            .done((results) => {
                console.log(results);
            });
    },
    createPublishingPage: () => {
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
    }
};
SP.SOD.loadMultiple(["SP.js"], () => {
    console.log("Executing examples");
    examples.run();
});
