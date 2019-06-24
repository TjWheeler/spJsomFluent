import * as spJsom from "spjsomfluent"
export class myApp {
    public run() {
        this.createListAndItem();
    }
    private createListAndItem() {
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
    }
}

SP.SOD.loadMultiple(["SP.js"], () => {
    console.log("Executing sample app");
    new myApp().run();
});
