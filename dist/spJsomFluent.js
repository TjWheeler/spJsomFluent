/*! spJsomFluent v0.1.0 - https://github.com/TjWheeler/spJsomFluent */
var spJsom=function(e){var t={};function n(i){if(t[i])return t[i].exports;var r=t[i]={i:i,l:!1,exports:{}};return e[i].call(r.exports,r,r.exports,n),r.l=!0,r.exports}return n.m=e,n.c=t,n.d=function(e,t,i){n.o(e,t)||Object.defineProperty(e,t,{enumerable:!0,get:i})},n.r=function(e){"undefined"!=typeof Symbol&&Symbol.toStringTag&&Object.defineProperty(e,Symbol.toStringTag,{value:"Module"}),Object.defineProperty(e,"__esModule",{value:!0})},n.t=function(e,t){if(1&t&&(e=n(e)),8&t)return e;if(4&t&&"object"==typeof e&&e&&e.__esModule)return e;var i=Object.create(null);if(n.r(i),Object.defineProperty(i,"default",{enumerable:!0,value:e}),2&t&&"string"!=typeof e)for(var r in e)n.d(i,r,function(t){return e[t]}.bind(null,r));return i},n.n=function(e){var t=e&&e.__esModule?function(){return e.default}:function(){return e};return n.d(t,"a",t),t},n.o=function(e,t){return Object.prototype.hasOwnProperty.call(e,t)},n.p="",n(n.s=0)}([function(e,t,n){"use strict";n.r(t);var i,r=function(){function e(){}return e.reject=function(e,t){e.reject(this,{get_message:function(){return t}})},e.notImplementedPromise=function(){var t=$.Deferred();return e.reject(t,"Not Implemented"),t.promise()},e.FilterArray=function(e,t){for(var n=[],i=0;i<e.length;i++)t(e[i],i)&&n.push(e[i]);return n},e.executeQuery=function(e){var t=$.Deferred();return e.executeQueryAsync(function(e,n){t.resolve(e,n)},function(e,n){t.reject(e,n)}),t.promise()},e}(),o=function(){function e(e){this.context=e}return e.prototype.setWelcomePage=function(e,t){var n=$.Deferred(),i=e.get_rootFolder();return this.context.load(e),this.context.load(i),i.set_welcomePage(t),i.update(),r.executeQuery(this.context).fail(function(e,t){n.reject(e,t)}).done(function(){n.resolve(i)}),n.promise()},e.prototype.createWeb=function(e,t,n,i,o){void 0===o&&(o=!0);var s=$.Deferred(),u=new SP.WebCreationInformation;u.set_url(e),u.set_title(n),u.set_webTemplate(i),u.set_useSamePermissionsAsParentSite(o);var c=t.get_webs().add(u);return this.context.load(c),r.executeQuery(this.context).fail(function(e,t){s.reject(e,t)}).done(function(){s.resolve(c)}),s.promise()},e.prototype.doesWebExist=function(e){var t=$.Deferred(),n=[];return this.getAllWebs(this.context,this.context.get_site().get_rootWeb(),n).fail(function(e,n){t.reject(e,n)}).done(function(){for(var i=0;i<n.length;i++)if(e.toLowerCase()===n[i].get_serverRelativeUrl()||e.toLowerCase()===n[i].get_url())return void t.resolve(!0);t.resolve(!1)}),t.promise()},e.prototype.getWebs=function(e){var t=$.Deferred(),n=[];return this.getAllWebs(this.context,e,n).fail(function(e,n){t.reject(e,n)}).done(function(){t.resolve(n)}),t.promise()},e.prototype.getAllWebs=function(e,t,n){var i=this,o=$.Deferred(),s=t.get_webs();return e.load(s),r.executeQuery(e).fail(function(e,t){o.reject(e,t)}).done(function(){for(var t=[],r=0;r<s.get_count();r++){var u=s.getItemAtIndex(r);n.push(u),t.push(i.getAllWebs(e,u,n))}t.length?$.when.apply($,t).fail(function(e,t){o.reject(e,t)}).done(function(){o.resolve()}):o.resolve()}),o.promise()},e}(),s=function(){function e(e){this._fluent=null,this._helperName="web",this.webHelper=null,this._fluent=e,this.webHelper=new o(e.context)}return e.prototype.setWelcomePage=function(e,t){var n=this;return this._fluent.chainAction(this._helperName+".setWelcomePage",function(){return n.webHelper.setWelcomePage(e,t)}),this._fluent},e.prototype.create=function(e,t,n,i,r){var o=this;return void 0===r&&(r=!0),this._fluent.chainAction(this._helperName+".create",function(){return o.webHelper.createWeb(e,t,n,i,r)}),this._fluent},e.prototype.exists=function(e){var t=this;return this._fluent.chainAction(this._helperName+".exists",function(){return t.webHelper.doesWebExist(e)}),this._fluent},e.prototype.getWebs=function(e){var t=this;return this._fluent.chainAction(this._helperName+".getWebs",function(){return t.webHelper.getWebs(e)}),this._fluent},e}(),u=function(){function e(e){this.context=e}return e.prototype.hasWebPermission=function(e,t){var n=$.Deferred();return this.context.load(t,"EffectiveBasePermissions"),r.executeQuery(this.context).fail(function(e,t){n.reject(e,t)}).done(function(){n.resolve(t.get_effectiveBasePermissions().has(e))}),n.promise()},e.prototype.hasListPermission=function(e,t){var n=$.Deferred();return this.context.load(t,"EffectiveBasePermissions"),r.executeQuery(this.context).fail(function(e,t){n.reject(e,t)}).done(function(){n.resolve(t.get_effectiveBasePermissions().has(e))}),n.promise()},e.prototype.hasItemPermission=function(e,t){var n=$.Deferred();return this.context.load(t,"EffectiveBasePermissions"),r.executeQuery(this.context).fail(function(e,t){n.reject(e,t)}).done(function(){n.resolve(t.get_effectiveBasePermissions().has(e))}),n.promise()},e}(),c=function(){function e(e){this._fluent=null,this._helperName="permission",this.permissionHelper=null,this._fluent=e,this.permissionHelper=new u(e.context)}return e.prototype.hasWebPermission=function(e,t){var n=this;return this._fluent.chainAction(this._helperName+".hasWebPermission",function(){return n.permissionHelper.hasWebPermission(e,t)}),this._fluent},e.prototype.hasListPermission=function(e,t){var n=this;return this._fluent.chainAction(this._helperName+".hasSitePermission",function(){return n.permissionHelper.hasListPermission(e,t)}),this._fluent},e.prototype.hasItemPermission=function(e,t){var n=this;return this._fluent.chainAction(this._helperName+".hasItemPermission",function(){return n.permissionHelper.hasItemPermission(e,t)}),this._fluent},e}(),a=function(){function e(e){this.context=e}return e.prototype.createList=function(e,t,n){var i=$.Deferred(),o=new SP.ListCreationInformation;o.set_title(t),o.set_templateType(n);var s=e.get_lists().add(o);return this.context.load(s),r.executeQuery(this.context).fail(function(e,t){i.reject(e,t)}).done(function(){i.resolve(s)}),i.promise()},e.prototype.setAlerts=function(e,t,n){var i=$.Deferred(),o=e.get_lists().getByTitle(t);return this.context.load(o),o.set_enableAssignToEmail(n),o.update(),r.executeQuery(this.context).fail(function(e,t){i.reject(e,t)}).done(function(){i.resolve()}),i.promise()},e.prototype.setListItemProperties=function(e,t){var n=$.Deferred();for(var i in t)void 0!==t[i]&&e.set_item(i,t[i]);return e.update(),this.context.load(e),r.executeQuery(this.context).fail(function(e,t){n.reject(e,t)}).done(function(){n.resolve(e)}),n.promise()},e.prototype.createListItemWithContentTypeName=function(e,t,n,i){var o=$.Deferred(),s=this.context,u=e.get_lists().getByTitle(t),c=new SP.ListItemCreationInformation,a=u.get_contentTypes();return s.load(a),r.executeQuery(s).fail(function(e,t){o.reject(e,t)}).done(function(){for(var e=null,t=0;t<a.get_count();t++){var l=a.getItemAtIndex(t);if(l.get_name()===n){e=l.get_id().get_stringValue();break}}var f=u.addItem(c);for(var h in e&&f.set_item("ContentTypeId",e),i)void 0!==i[h]&&f.set_item(h,i[h]);f.update(),s.load(f),r.executeQuery(s).fail(function(e,t){o.reject(e,t)}).done(function(){o.resolve(f)})}),o.promise()},e.prototype.createListItem=function(e,t,n){var i=$.Deferred(),o=this.context,s=e.get_lists().getByTitle(t),u=new SP.ListItemCreationInformation,c=s.addItem(u);for(var a in n)void 0!==n[a]&&c.set_item(a,n[a]);return c.update(),o.load(c),r.executeQuery(o).fail(function(e,t){i.reject(e,t)}).done(function(){i.resolve(c)}),i.promise()},e.prototype.loadListItem=function(e,t){void 0===t&&(t=null);var n=$.Deferred();return t&&t.length?this.context.load(e,t):this.context.load(e),r.executeQuery(this.context).fail(function(e,t){n.reject(e,t)}).done(function(){n.resolve(e)}),n.promise()},e.prototype.getFile=function(e){var t=$.Deferred(),n=this.context.get_site().get_rootWeb().getFileByServerRelativeUrl(e);return this.context.load(n),r.executeQuery(this.context).fail(function(e,n){t.reject(e,n)}).done(function(){t.resolve(n)}),t.promise()},e.prototype.checkInFile=function(e,t,n,i){var o=$.Deferred(),s=e.getFileByServerRelativeUrl(t);return this.context.load(s),s.checkIn(n,i),r.executeQuery(this.context).fail(function(e,t){o.reject(e,t)}).done(function(){o.resolve(s)}),o.promise()},e.prototype.getList=function(e,t){var n=$.Deferred(),i=e.get_lists().getByTitle(t);return this.context.load(i),r.executeQuery(this.context).fail(function(e,t){n.reject(e,t)}).done(function(){n.resolve(i)}),n.promise()},e.prototype.getFileListItem=function(e,t){void 0===t&&(t=null);var n=this.context.get_site().get_rootWeb().getFileByServerRelativeUrl(e).get_listItemAllFields();return this.loadListItem(n)},e.prototype.getListItemById=function(e,t,n,i){void 0===i&&(i=null);var r=e.get_lists().getByTitle(t).getItemById(n);return this.loadListItem(r,i)},e.prototype.deleteListItemById=function(e,t,n){var i=$.Deferred(),o=e.get_lists().getByTitle(t).getItemById(n);return this.context.load(o),o.deleteObject(),r.executeQuery(this.context).fail(function(e,t){i.reject(e,t)}).done(function(){i.resolve()}),i.promise()},e.prototype.getListItems=function(e,t,n){var i=$.Deferred(),o=e.get_lists().getByTitle(t),s=new SP.CamlQuery;n||(n="<View><Query></Query></Where>"),s.set_viewXml(n);var u=o.getItems(s);return this.context.load(u),r.executeQuery(this.context).fail(function(e,t){i.reject(e,t)}).done(function(){i.resolve(u)}),i.promise()},e.prototype.addContentTypeListAssociation=function(e,t,n){var i=this,o=$.Deferred(),s=e.get_lists().getByTitle(t),u=s.get_contentTypes();this.context.load(s),this.context.load(u);var c=this.context.get_site().get_rootWeb().get_contentTypes(),a=function(e,t){for(var n=0;n<e.get_count();n++)if(e.getItemAtIndex(n).get_name().toLowerCase()===t.toLowerCase())return e.getItemAtIndex(n);return null};return this.context.load(c),r.executeQuery(this.context).fail(function(e,t){o.reject(e,t)}).done(function(){var e=a(c,n);if(e){var t=a(u,n);if(t)o.resolve(t);else{s.get_contentTypesEnabled()||s.set_contentTypesEnabled(!0);var l=s.get_contentTypes().addExistingContentType(e);s.update(),r.executeQuery(i.context).fail(function(e,t){o.reject(e,t)}).done(function(){o.resolve(l)})}}else r.reject(o,"Content Type "+n+" not found")}),o.promise()},e.prototype.removeContentTypeListAssociation=function(e,t,n){var i=this,o=$.Deferred(),s=e.get_lists().getByTitle(t),u=s.get_contentTypes();this.context.load(s),this.context.load(u);return r.executeQuery(this.context).fail(function(e,t){o.reject(e,t)}).done(function(){var e=function(e,t){for(var n=0;n<e.get_count();n++)if(e.getItemAtIndex(n).get_name().toLowerCase()===t.toLowerCase())return e.getItemAtIndex(n);return null}(u,n);e?(e.deleteObject(),s.update(),r.executeQuery(i.context).fail(function(e,t){o.reject(e,t)}).done(function(){o.resolve()})):o.resolve()}),o.promise()},e.prototype.setDefaultValueOnList=function(e,t,n,i){var o=this,s=$.Deferred(),u=e.get_lists().getByTitle(t),c=u.get_fields();this.context.load(u),this.context.load(c);return r.executeQuery(this.context).fail(function(e,t){s.reject(e,t)}).done(function(){var e=function(e,t){for(var n=0;n<e.get_count();n++)if(e.getItemAtIndex(n).get_internalName().toLowerCase()===t.toLowerCase())return e.getItemAtIndex(n);return null}(c,n);e?(e.set_defaultValue(i),e.update(),r.executeQuery(o.context).fail(function(e,t){s.reject(e,t)}).done(function(){s.resolve()})):r.reject(s,"Field "+n+" not found")}),s.promise()},e}(),l=function(){function e(e){this.fluent=null,this.listHelper=null,this._helperName="list",this.fluent=e,this.listHelper=new a(e.context)}return e.prototype.create=function(e,t,n){var i=this;return this.fluent.chainAction(this._helperName+".create",function(){return i.listHelper.createList(e,t,n)}),this.fluent},e.prototype.delete=function(e,t,n){return this.fluent.chainAction(this._helperName+".delete",function(){return r.notImplementedPromise()}),this.fluent},e.prototype.get=function(e,t){var n=this;return this.fluent.chainAction(this._helperName+".get",function(){return n.listHelper.getList(e,t)}),this.fluent},e.prototype.addContentTypeListAssociation=function(e,t,n){var i=this;return this.fluent.chainAction(this._helperName+".addContentTypeListAssociation",function(){return i.listHelper.addContentTypeListAssociation(e,t,n)}),this.fluent},e.prototype.removeContentTypeListAssociation=function(e,t,n){var i=this;return this.fluent.chainAction(this._helperName+".removeContentTypeListAssociation",function(){return i.listHelper.removeContentTypeListAssociation(e,t,n)}),this.fluent},e.prototype.setDefaultValueOnList=function(e,t,n,i){var r=this;return this.fluent.chainAction(this._helperName+".setDefaultValueOnList",function(){return r.listHelper.setDefaultValueOnList(e,t,n,i)}),this.fluent},e.prototype.setAlerts=function(e,t,n){var i=this;return this.fluent.chainAction(this._helperName+".setAlerts",function(){return i.listHelper.setAlerts(e,t,n)}),this.fluent},e}(),f=function(){function e(e){this.fluent=null,this.listHelper=null,this._helperName="listItem",this.fluent=e,this.listHelper=new a(e.context)}return e.prototype.update=function(e,t){var n=this;return this.fluent.chainAction(this._helperName+".update",function(){return n.listHelper.setListItemProperties(e,t)}),this.fluent},e.prototype.createWithContentType=function(e,t,n,i){var r=this;return this.fluent.chainAction(this._helperName+".createWithContentType",function(){return r.listHelper.createListItemWithContentTypeName(e,t,n,i)}),this.fluent},e.prototype.create=function(e,t,n){var i=this;return this.fluent.chainAction(this._helperName+".create",function(){return i.listHelper.createListItem(e,t,n)}),this.fluent},e.prototype.get=function(e,t,n,i){var r=this;return void 0===i&&(i=null),this.fluent.chainAction(this._helperName+".get",function(){return r.listHelper.getListItemById(e,t,n,i)}),this.fluent},e.prototype.getFileListItem=function(e,t){var n=this;return void 0===t&&(t=null),this.fluent.chainAction(this._helperName+".getFileListItem",function(){return n.listHelper.getFileListItem(e,t)}),this.fluent},e.prototype.deleteById=function(e,t,n){var i=this;return this.fluent.chainAction(this._helperName+".deleteById",function(){return i.listHelper.deleteListItemById(e,t,n)}),this.fluent},e.prototype.query=function(e,t,n){var i=this;return this.fluent.chainAction(this._helperName+".query",function(){return i.listHelper.getListItems(e,t,n)}),this.fluent},e}(),h=function(){function e(e){this.fluent=null,this.listHelper=null,this._helperName="file",this.fluent=e,this.listHelper=new a(e.context)}return e.prototype.getFile=function(e){var t=this;return this.fluent.chainAction(this._helperName+".getFile",function(){return t.listHelper.getFileListItem(e)}),this.fluent},e.prototype.checkIn=function(e,t,n,i){var r=this;return this.fluent.chainAction(this._helperName+".checkInFile",function(){return r.listHelper.checkInFile(e,t,n,i)}),this.fluent},e}(),p=function(){function e(e){this.context=e}return e.prototype.getUserByEmail=function(e){return this.loadUser(this.context.get_web().ensureUser(e))},e.prototype.getUserById=function(e){return this.loadUser(this.context.get_web().get_siteUsers().getById(e))},e.prototype.loadUser=function(e){var t=$.Deferred();return this.context.load(e),this.context.executeQueryAsync(function(n,i){t.resolve(e)},function(e,n){console.error(n.get_message()),t.reject(e,n)}),t.promise()},e.prototype.getCurrentUser=function(){var e=$.Deferred(),t=this.context.get_web().get_currentUser();return this.context.load(t),this.context.executeQueryAsync(function(n,i){e.resolve(t)},function(t,n){console.error(n.get_message()),e.reject(t,n)}),e.promise()},e.prototype.getCurrentUserProfileProperties=function(){var e=this,t=$.Deferred();return SP.SOD.executeFunc("userprofile","SP.UserProfiles.PeopleManager",function(){var n=e.context,i=n.get_web().get_currentUser(),r=new SP.UserProfiles.PeopleManager(n).getMyProperties();n.load(i),n.load(r),n.executeQueryAsync(function(e,n){t.resolve(r.get_userProfileProperties())},function(e,n){console.error(n.get_message()),t.reject(e,n)})}),t.promise()},e.prototype.getCurrentUserManager=function(){var e=this,t=$.Deferred(),n=new SP.UserProfiles.PeopleManager(this.context),i=this.context.get_web().get_currentUser().get_email(),o=new SP.UserProfiles.UserProfilePropertiesForUser(this.context,"i:0#.f|membership|"+i,["Manager"]),s=n.getUserProfilePropertiesFor(o);return this.context.load(o),r.executeQuery(this.context).fail(function(e,n){t.reject(e,n)}).done(function(){if(s[0]){var n=e.context.get_web().ensureUser(s[0]);e.context.load(n),r.executeQuery(e.context).fail(function(e,n){t.reject(e,n)}).done(function(){t.resolve(n)})}else t.resolve(null)}),t.promise()},e}(),d=function(){function e(e){this._fluent=null,this._helperName="userProfile",this.userHelper=null,this._fluent=e,this.userHelper=new p(e.context)}return e.prototype.get=function(e){var t=this;return this._fluent.chainAction(this._helperName+".get",function(){return t.userHelper.getUserByEmail(e)}),this._fluent},e.prototype.getById=function(e){var t=this;return this._fluent.chainAction(this._helperName+".getById",function(){return t.userHelper.getUserById(e)}),this._fluent},e.prototype.getCurrentUser=function(){var e=this;return this._fluent.chainAction(this._helperName+".getCurrentUser",function(){return e.userHelper.getCurrentUser()}),this._fluent},e.prototype.getCurrentUserProfileProperties=function(){var e=this;return this._fluent.registerDependency(j.UserProfile),this._fluent.chainAction(this._helperName+".getCurrentUserProfileProperties",function(){return e.userHelper.getCurrentUserProfileProperties()}),this._fluent},e.prototype.getCurrentUserManager=function(){var e=this;return this._fluent.registerDependency(j.UserProfile),this._fluent.chainAction(this._helperName+".getCurrentUserManager",function(){return e.userHelper.getCurrentUserManager()}),this._fluent},e}(),g=function(){function e(e){this.context=e}return e.prototype.deleteQuicklaunchNodes=function(e){return this.deleteNodes(e.get_navigation().get_quickLaunch())},e.prototype.deleteTopNavigationNodes=function(e){return this.deleteNodes(e.get_navigation().get_topNavigationBar())},e.prototype.deleteQuicklaunchNode=function(e,t){return this.deleteNode(e.get_navigation().get_quickLaunch(),t)},e.prototype.deleteTopQuicklaunchNode=function(e,t){return this.deleteNode(e.get_navigation().get_topNavigationBar(),t)},e.prototype.setCurrentNavigation=function(e,t,n,i){var o=this;void 0===n&&(n=!0),void 0===i&&(i=!0);var s=$.Deferred(),u=e.get_allProperties();this.context.load(e),this.context.load(u);var c=function(){i&&n?u.set_item("__CurrentNavigationIncludeTypes","3"):i&&!n?u.set_item("__CurrentNavigationIncludeTypes","2"):!i&&n?u.set_item("__CurrentNavigationIncludeTypes","1"):i||n||u.set_item("__CurrentNavigationIncludeTypes","0")};return r.executeQuery(this.context).fail(function(e,t){s.reject(e,t)}).done(function(){var n=new SP.Publishing.Navigation.WebNavigationSettings(o.context,e);switch(t){case I.Inherit:n.get_currentNavigation().set_source(SP.Publishing.Navigation.StandardNavigationSource.inheritFromParentWeb);break;case I.Managed:r.reject(s,"Not implemented");break;case I.StructuralWithSiblings:n.get_currentNavigation().set_source(SP.Publishing.Navigation.StandardNavigationSource.portalProvider),u.set_item("__NavigationShowSiblings","True"),c(),e.update();break;case I.StructuralChildrenOnly:n.get_currentNavigation().set_source(SP.Publishing.Navigation.StandardNavigationSource.portalProvider),u.set_item("__NavigationShowSiblings","False"),c(),e.update();break;default:r.reject(s,"Unknown Navigation Type")}n.update(null),r.executeQuery(o.context).fail(function(e,t){s.reject(e,t)}).done(function(){s.resolve()})}),s.promise()},e.prototype.deleteNodes=function(e){var t=this,n=$.Deferred();return this.context.load(e),r.executeQuery(this.context).fail(function(e,t){n.reject(e,t)}).done(function(){for(var i=e.getEnumerator(),o=[];i.moveNext();)o.push(i.get_current());for(var s=0;s<o.length;s++)o[s].deleteObject();r.executeQuery(t.context).fail(function(e,t){n.reject(e,t)}).done(function(){n.resolve()})}),n.promise()},e.prototype.deleteNode=function(e,t){var n=this,i=$.Deferred();return this.context.load(e),r.executeQuery(this.context).fail(function(e,t){i.reject(e,t)}).done(function(){for(var o=e.getEnumerator(),s=[];o.moveNext();){var u=o.get_current();u.get_title()===t&&s.push(u)}for(var c=0;c<s.length;c++)s[c].deleteObject();r.executeQuery(n.context).fail(function(e,t){i.reject(e,t)}).done(function(){i.resolve()})}),i.promise()},e.prototype.createQuicklaunchNode=function(e,t,n,i){return void 0===i&&(i=!0),this.createNode(e.get_navigation().get_quickLaunch(),t,n,i)},e.prototype.createTopNavigationNode=function(e,t,n,i){return void 0===i&&(i=!0),this.createNode(e.get_navigation().get_topNavigationBar(),t,n,i)},e.prototype.createNode=function(e,t,n,i){void 0===i&&(i=!0);var o=$.Deferred();this.context.load(e);var s=new SP.NavigationNodeCreationInformation;s.set_title(t),s.set_url(n),s.set_asLastNode(i);var u=e.add(s);return this.context.load(e),r.executeQuery(this.context).fail(function(e,t){o.reject(e,t)}).done(function(){o.resolve(u)}),o.promise()},e}(),m=function(){function e(e){this.fluent=null,this.navigationHelper=null,this._helperName="navigation",this.fluent=e,this.navigationHelper=new g(e.context)}return e.prototype.createNode=function(e,t,n,i,r){var o=this;return void 0===r&&(r=!0),this.fluent.chainAction(this._helperName+".createNode",function(){if(t==w.Quicklaunch)return o.navigationHelper.createQuicklaunchNode(e,n,i,r);if(t==w.TopNavigation)return o.navigationHelper.createTopNavigationNode(e,n,i,r);throw"Unknown location "+t}),this.fluent},e.prototype.deleteAllNodes=function(e,t){var n=this;return this.fluent.chainAction(this._helperName+".deleteAllNodes",function(){if(t==w.Quicklaunch)return n.navigationHelper.deleteQuicklaunchNodes(e);if(t==w.TopNavigation)return n.navigationHelper.deleteTopNavigationNodes(e);throw"Unknown location "+t}),this.fluent},e.prototype.deleteNode=function(e,t,n){var i=this;return this.fluent.chainAction(this._helperName+".deleteNode",function(){if(t==w.Quicklaunch)return i.navigationHelper.deleteQuicklaunchNode(e,n);if(t==w.TopNavigation)return i.navigationHelper.deleteTopQuicklaunchNode(e,n);throw"Unknown location "+t}),this.fluent},e.prototype.setCurrentNavigation=function(e,t,n,i){var r=this;return void 0===n&&(n=!1),void 0===i&&(i=!1),this.fluent.registerDependency(j.Publishing),this.fluent.chainAction(this._helperName+".setCurrentNavigation",function(){return r.navigationHelper.setCurrentNavigation(e,t,n,i)}),this.fluent},e}(),v=function(){function e(e){this.context=e}return e.prototype.createPublishingPage=function(e,t,n){var i=$.Deferred(),o=this.context.get_site().get_rootWeb().getFileByServerRelativeUrl(n).get_listItemAllFields();this.context.load(o);var s=new SP.Publishing.PublishingPageInformation;s.set_name(t),s.set_pageLayoutListItem(o);var u=SP.Publishing.PublishingWeb.getPublishingWeb(this.context,e);this.context.load(u);var c=u.addPublishingPage(s);return this.context.load(c),r.executeQuery(this.context).fail(function(e,t){i.reject(e,t)}).done(function(){i.resolve(c)}),i.promise()},e.prototype.getPublishingLayout=function(e){var t=$.Deferred(),n=this.context.get_site().get_rootWeb().getFileByServerRelativeUrl(e).get_listItemAllFields();return this.context.load(n),r.executeQuery(this.context).fail(function(e,n){t.reject(e,n)}).done(function(){t.resolve(n)}),t.promise()},e}(),y=function(){function e(e){this.fluent=null,this._helperName="publishingPage",this.fluent=e,this.pageHelper=new v(e.context)}return e.prototype.create=function(e,t,n){var i=this;return this.fluent.registerDependency(j.Publishing),this.fluent.chainAction(this._helperName+".create",function(){return i.pageHelper.createPublishingPage(e,t,n)}),this.fluent},e.prototype.getLayout=function(e){var t=this;return this.fluent.chainAction(this._helperName+".getLayout",function(){return t.pageHelper.getPublishingLayout(e)}),this.fluent},e}(),_=(i=function(e,t){return(i=Object.setPrototypeOf||{__proto__:[]}instanceof Array&&function(e,t){e.__proto__=t}||function(e,t){for(var n in t)t.hasOwnProperty(n)&&(e[n]=t[n])})(e,t)},function(e,t){function n(){this.constructor=e}i(e,t),e.prototype=null===t?Object.create(t):(n.prototype=t.prototype,new n)}),x=function(){return function(){}}(),b=function(e){function t(){return null!==e&&e.apply(this,arguments)||this}return _(t,e),t}(x),P=function(){return function(){}}(),N=function(e){function t(){return null!==e&&e.apply(this,arguments)||this}return _(t,e),t}(x);n.d(t,"Fluent",function(){return A}),n.d(t,"NavigationLocation",function(){return w}),n.d(t,"NavigationType",function(){return I}),n.d(t,"Dependency",function(){return j});var w,I,j,A=function(){function e(){this.commands=[],this.results=[],this.dependencies=[],this.settings={timeoutMilliseconds:5e3,enableDependencyTimeout:!0}}return e.prototype.withContext=function(e){return this.context=e,this},e.prototype.withSettings=function(e){for(var t in e)void 0!==this.settings[t]&&(this.settings[t]=e[t]);return this},Object.defineProperty(e.prototype,"promise",{get:function(){return this.resultPromise.promise()},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"permission",{get:function(){return new c(this)},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"list",{get:function(){return new l(this)},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"listItem",{get:function(){return new f(this)},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"file",{get:function(){return new h(this)},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"publishingPage",{get:function(){return new y(this)},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"web",{get:function(){return new s(this)},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"user",{get:function(){return new d(this)},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"navigation",{get:function(){return new m(this)},enumerable:!0,configurable:!0}),e.prototype.execute=function(){var e=this;if(this.resultPromise=$.Deferred(),this.settings.enableDependencyTimeout)var t=setTimeout(function(){r.reject(e.resultPromise,"Timeout waiting for dependencies to load")},this.settings.timeoutMilliseconds);return this.loadDependencies().done(function(){e.continue()}).always(function(){e.settings.enableDependencyTimeout&&clearTimeout(t)}),this.resultPromise.promise()},e.prototype.onActionExecuted=function(e){return this.onExecuted=e,this},e.prototype.onActionExecuting=function(e){return this.onExecuting=e,this},e.prototype.when=function(e){var t=this;if(this.peekLastCommand()&&this.peekLastCommand().constructor!==b)throw"Illegal operation: The last command must be an ActionCommand";var n=new N;return n.action=function(){var n=$.Deferred(),i=t.peekLastResult();if(i){if(!i.success)throw"Illegal operation: The last command should have succeeded for this call to have been made.  This is an issue in the fluent api";e.apply(void 0,i.result)?n.resolve():t.reject(n,t,"Predicate returned false")}else t.failChain(t,"No action to process for when");return n.promise()},this.commands.push(n),this},e.prototype.whenAll=function(e){var t=this;if(this.peekLastCommand()&&this.peekLastCommand().constructor!==b)throw"Illegal operation: The last command must be an ActionCommand";var n=new N;return n.action=function(){var n=$.Deferred(),i=t.peekLastResult();if(i){if(!i.success)throw"Illegal operation: The last command should have succeeded for this call to have been made.  This is an issue in the fluent api";e(t.results)?n.resolve():t.reject(n,t,"Predicate returned false")}else t.failChain(t,"No action to process for whenAll");return n.promise()},this.commands.push(n),this},e.prototype.whenTrue=function(){var e=this;if(this.peekLastCommand()&&this.peekLastCommand().constructor!==b)throw"Illegal operation: The last command must be an ActionCommand";var t=new N;return t.action=function(){var t=$.Deferred(),n=e.peekLastResult();if(n){if(!n.success)throw"Illegal operation: The last command should have succeeded for this call to have been made.  This is an issue in the fluent api";n.result&&n.result.length&&n.result[0]?t.resolve():e.reject(t,e,"Result is not true")}else e.failChain(e,"No action to process for whenTrue");return t.promise()},this.commands.push(t),this},e.prototype.whenFalse=function(){var e=this;if(this.peekLastCommand()&&this.peekLastCommand().constructor!==b)throw"Illegal operation: The last command must be an ActionCommand";var t=new N;return t.action=function(){var t=$.Deferred(),n=e.peekLastResult();if(n){if(!n.success)throw"Illegal operation: The last command should have succeeded for this call to have been made.  This is an issue in the fluent api";n.result&&n.result.length&&!n.result[0]?t.resolve():e.reject(t,e,"Result is not false")}else e.failChain(e,"No action to process for whenFalse");return t.promise()},this.commands.push(t),this},e.prototype.chainAction=function(e,t){var n=new b;n.name=e,n.action=t,this.commands.push(n)},e.prototype.registerDependency=function(e){-1===this.dependencies.indexOf(e)&&this.dependencies.push(e)},e.prototype.continue=function(){var e=this;if(!this.context)throw"context not set, you must call withContext";var t=this.commands.shift();t&&t.action?(t.constructor===b&&this.onExecuting&&this.onExecuting(t.name),t.action().done(function(n,i,r,o,s,u,c){if(t.constructor===b){var a=[];e.storeResult(n,a),e.storeResult(i,a),e.storeResult(r,a),e.storeResult(o,a),e.storeResult(s,a),e.storeResult(u,a),e.storeResult(c,a),e.addResult(t,!0,a),e.onExecuted&&e.onExecuted(t.name,!0,a)}e.commands.length?e.continue():e.resolveChain()}).fail(function(n,i){if(t.constructor===b){var r=[];return e.storeResult(n,r),e.storeResult(i,r),e.addResult(t,!1,r),void e.failChain(n,i)}e.resolveChain()})):this.resolveChain()},e.prototype.storeResult=function(e,t){void 0!==e&&t.push(e)},e.prototype.resolveChain=function(){this.resultPromise.resolve(this.results)},e.prototype.failChain=function(e,t){"string"==typeof t&&(t={get_message:function(){return t}}),this.resultPromise.reject(e,t)},e.prototype.reject=function(e,t,n){"string"==typeof n&&(n={get_message:function(){return n}}),e.reject(t,n)},e.prototype.addResult=function(e,t,n){var i=new P;i.name=e.name,i.success=t,i.result=n,this.results.push(i)},e.prototype.loadDependencies=function(){for(var e=$.Deferred(),t=["SP.js","SP.Runtime.js"],n=0;n<this.dependencies.length;n++)switch(this.dependencies[n]){case j.UserProfile:window.LoadSodByKey("userprofile"),t.push("userprofile");break;case j.Publishing:SP.SOD.registerSod("SP.Publishing.js",SP.Utilities.Utility.getLayoutsPageUrl("sp.publishing.js")),t.push("SP.Publishing.js");break;case j.Taxonomy:SP.SOD.registerSod("sp.taxonomy.js",SP.Utilities.Utility.getLayoutsPageUrl("sp.taxonomy.js")),t.push("sp.taxonomy.js");break;default:throw"Depenency not implemented"}return SP.SOD.loadMultiple(t,function(){e.resolve()}),e.promise()},e.prototype.peekLastCommand=function(){return this.commands.length?this.commands[this.commands.length-1]:null},e.prototype.peekLastResult=function(){return this.results.length?this.results[this.results.length-1]:null},e}();!function(e){e[e.TopNavigation=0]="TopNavigation",e[e.Quicklaunch=1]="Quicklaunch"}(w||(w={})),function(e){e[e.Inherit=0]="Inherit",e[e.Managed=1]="Managed",e[e.StructuralWithSiblings=2]="StructuralWithSiblings",e[e.StructuralChildrenOnly=3]="StructuralChildrenOnly"}(I||(I={})),function(e){e[e.Publishing=0]="Publishing",e[e.UserProfile=1]="UserProfile",e[e.Taxonomy=2]="Taxonomy"}(j||(j={}))}]);
//# sourceMappingURL=spJsomFluent.js.map