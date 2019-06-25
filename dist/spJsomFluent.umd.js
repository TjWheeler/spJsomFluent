/*! spJsomFluent "0.1.13" - https://github.com/TjWheeler/spJsomFluent */
!function(e,t){"object"==typeof exports&&"object"==typeof module?module.exports=t():"function"==typeof define&&define.amd?define([],t):"object"==typeof exports?exports.spJsom=t():e.spJsom=t()}(window,function(){return function(e){var t={};function n(r){if(t[r])return t[r].exports;var i=t[r]={i:r,l:!1,exports:{}};return e[r].call(i.exports,i,i.exports,n),i.l=!0,i.exports}return n.m=e,n.c=t,n.d=function(e,t,r){n.o(e,t)||Object.defineProperty(e,t,{enumerable:!0,get:r})},n.r=function(e){"undefined"!=typeof Symbol&&Symbol.toStringTag&&Object.defineProperty(e,Symbol.toStringTag,{value:"Module"}),Object.defineProperty(e,"__esModule",{value:!0})},n.t=function(e,t){if(1&t&&(e=n(e)),8&t)return e;if(4&t&&"object"==typeof e&&e&&e.__esModule)return e;var r=Object.create(null);if(n.r(r),Object.defineProperty(r,"default",{enumerable:!0,value:e}),2&t&&"string"!=typeof e)for(var i in e)n.d(r,i,function(t){return e[t]}.bind(null,i));return r},n.n=function(e){var t=e&&e.__esModule?function(){return e.default}:function(){return e};return n.d(t,"a",t),t},n.o=function(e,t){return Object.prototype.hasOwnProperty.call(e,t)},n.p="",n(n.s=0)}([function(e,t,n){"use strict";n.r(t);var r=function(){function e(){}return e.reject=function(e,t){e.reject(this,{get_message:function(){return t}})},e.notImplementedPromise=function(){var t=$.Deferred();return e.reject(t,"Not Implemented"),t.promise()},e.FilterArray=function(e,t){for(var n=[],r=0;r<e.length;r++)t(e[r],r)&&n.push(e[r]);return n},e.executeQuery=function(e){var t=$.Deferred();return e.executeQueryAsync(function(e,n){t.resolve(e,n)},function(e,n){t.reject(e,n)}),t.promise()},e}(),i=function(){function e(e){this.context=e}return e.prototype.setWelcomePage=function(e,t){var n=$.Deferred(),i=e.get_rootFolder();return this.context.load(e),this.context.load(i),i.set_welcomePage(t),i.update(),r.executeQuery(this.context).fail(function(e,t){n.reject(e,t)}).done(function(){n.resolve(i)}),n.promise()},e.prototype.createWeb=function(e,t,n,i,o){void 0===o&&(o=!0);var s=$.Deferred(),u=new SP.WebCreationInformation;u.set_url(e),u.set_title(n),u.set_webTemplate(i),u.set_useSamePermissionsAsParentSite(o);var c=t.get_webs().add(u);return this.context.load(c),r.executeQuery(this.context).fail(function(e,t){s.reject(e,t)}).done(function(){s.resolve(c)}),s.promise()},e.prototype.doesWebExist=function(e){var t=$.Deferred(),n=[];return this.getAllWebs(this.context,this.context.get_site().get_rootWeb(),n).fail(function(e,n){t.reject(e,n)}).done(function(){for(var r=0;r<n.length;r++)if(e.toLowerCase()===n[r].get_serverRelativeUrl()||e.toLowerCase()===n[r].get_url())return void t.resolve(!0);t.resolve(!1)}),t.promise()},e.prototype.getWebs=function(e){var t=$.Deferred(),n=[];return this.getAllWebs(this.context,e,n).fail(function(e,n){t.reject(e,n)}).done(function(){t.resolve(n)}),t.promise()},e.prototype.getAllWebs=function(e,t,n){var i=this,o=$.Deferred(),s=t.get_webs();return e.load(s),r.executeQuery(e).fail(function(e,t){o.reject(e,t)}).done(function(){for(var t=[],r=0;r<s.get_count();r++){var u=s.getItemAtIndex(r);n.push(u),t.push(i.getAllWebs(e,u,n))}t.length?$.when.apply($,t).fail(function(e,t){o.reject(e,t)}).done(function(){o.resolve()}):o.resolve()}),o.promise()},e}(),o=function(){function e(e){this._fluent=null,this._helperName="web",this.webHelper=null,this._fluent=e,this.webHelper=new i(e.context)}return e.prototype.setWelcomePage=function(e,t){var n=this;return this._fluent.chainAction(this._helperName+".setWelcomePage",function(){return n.webHelper.setWelcomePage(e,t)})},e.prototype.create=function(e,t,n,r,i){var o=this;return void 0===i&&(i=!0),this._fluent.chainAction(this._helperName+".create",function(){return o.webHelper.createWeb(e,t,n,r,i)})},e.prototype.exists=function(e){var t=this;return this._fluent.chainAction(this._helperName+".exists",function(){return t.webHelper.doesWebExist(e)})},e.prototype.getWebs=function(e){var t=this;return this._fluent.chainAction(this._helperName+".getWebs",function(){return t.webHelper.getWebs(e)})},e}(),s=function(){function e(e){this.context=e}return e.prototype.hasWebPermission=function(e,t){var n=$.Deferred();return this.context.load(t,"EffectiveBasePermissions"),r.executeQuery(this.context).fail(function(e,t){n.reject(e,t)}).done(function(){n.resolve(t.get_effectiveBasePermissions().has(e))}),n.promise()},e.prototype.hasListPermission=function(e,t){var n=$.Deferred();return this.context.load(t,"EffectiveBasePermissions"),r.executeQuery(this.context).fail(function(e,t){n.reject(e,t)}).done(function(){n.resolve(t.get_effectiveBasePermissions().has(e))}),n.promise()},e.prototype.hasItemPermission=function(e,t){var n=$.Deferred();return this.context.load(t,"EffectiveBasePermissions"),r.executeQuery(this.context).fail(function(e,t){n.reject(e,t)}).done(function(){n.resolve(t.get_effectiveBasePermissions().has(e))}),n.promise()},e}(),u=function(){function e(e){this._fluent=null,this._helperName="permission",this.permissionHelper=null,this._fluent=e,this.permissionHelper=new s(e.context)}return e.prototype.hasWebPermission=function(e,t){var n=this;return this._fluent.chainAction(this._helperName+".hasWebPermission",function(){return n.permissionHelper.hasWebPermission(e,t)})},e.prototype.hasListPermission=function(e,t){var n=this;return this._fluent.chainAction(this._helperName+".hasListPermission",function(){return n.permissionHelper.hasListPermission(e,t)})},e.prototype.hasItemPermission=function(e,t){var n=this;return this._fluent.chainAction(this._helperName+".hasItemPermission",function(){return n.permissionHelper.hasItemPermission(e,t)})},e}(),c=function(){function e(e){this.context=e}return e.prototype.createList=function(e,t,n){var i=$.Deferred(),o=new SP.ListCreationInformation;o.set_title(t),o.set_templateType(n);var s=e.get_lists().add(o);return this.context.load(s),r.executeQuery(this.context).fail(function(e,t){i.reject(e,t)}).done(function(){i.resolve(s)}),i.promise()},e.prototype.setAlerts=function(e,t,n){var i=$.Deferred(),o=e.get_lists().getByTitle(t);return this.context.load(o),o.set_enableAssignToEmail(n),o.update(),r.executeQuery(this.context).fail(function(e,t){i.reject(e,t)}).done(function(){i.resolve()}),i.promise()},e.prototype.setListItemProperties=function(e,t){var n=$.Deferred();for(var i in t)void 0!==t[i]&&e.set_item(i,t[i]);return e.update(),this.context.load(e),r.executeQuery(this.context).fail(function(e,t){n.reject(e,t)}).done(function(){n.resolve(e)}),n.promise()},e.prototype.exists=function(e,t){var n=$.Deferred(),i=e.get_lists();return this.context.load(i),r.executeQuery(this.context).fail(function(e,t){n.reject(e,t)}).done(function(){for(var e=0;e<i.get_count();e++)if(i.getItemAtIndex(e).get_title().toLowerCase()===t.toLowerCase())return void n.resolve(!0);n.resolve(!1)}),n.promise()},e.prototype.createListItemWithContentTypeName=function(e,t,n,i){var o=$.Deferred(),s=this.context,u=e.get_lists().getByTitle(t),c=new SP.ListItemCreationInformation,a=u.get_contentTypes();return s.load(a),r.executeQuery(s).fail(function(e,t){o.reject(e,t)}).done(function(){for(var e=null,t=0;t<a.get_count();t++){var l=a.getItemAtIndex(t);if(l.get_name()===n){e=l.get_id().get_stringValue();break}}var f=u.addItem(c);for(var h in e&&f.set_item("ContentTypeId",e),i)void 0!==i[h]&&f.set_item(h,i[h]);f.update(),s.load(f),r.executeQuery(s).fail(function(e,t){o.reject(e,t)}).done(function(){o.resolve(f)})}),o.promise()},e.prototype.createListItem=function(e,t,n){var i=$.Deferred(),o=this.context,s=e.get_lists().getByTitle(t),u=new SP.ListItemCreationInformation,c=s.addItem(u);for(var a in n)void 0!==n[a]&&c.set_item(a,n[a]);return c.update(),o.load(c),r.executeQuery(o).fail(function(e,t){i.reject(e,t)}).done(function(){i.resolve(c)}),i.promise()},e.prototype.loadListItem=function(e,t){void 0===t&&(t=null);var n=$.Deferred();return t&&t.length?this.context.load(e,t):this.context.load(e),r.executeQuery(this.context).fail(function(e,t){n.reject(e,t)}).done(function(){n.resolve(e)}),n.promise()},e.prototype.getFile=function(e){var t=$.Deferred(),n=this.context.get_site().get_rootWeb().getFileByServerRelativeUrl(e);return this.context.load(n),r.executeQuery(this.context).fail(function(e,n){t.reject(e,n)}).done(function(){t.resolve(n)}),t.promise()},e.prototype.checkInFile=function(e,t,n,i){var o=$.Deferred(),s=e.getFileByServerRelativeUrl(t);return this.context.load(s),s.checkIn(n,i),r.executeQuery(this.context).fail(function(e,t){o.reject(e,t)}).done(function(){o.resolve(s)}),o.promise()},e.prototype.getList=function(e,t){var n=$.Deferred(),i=e.get_lists().getByTitle(t);return this.context.load(i),r.executeQuery(this.context).fail(function(e,t){n.reject(e,t)}).done(function(){n.resolve(i)}),n.promise()},e.prototype.deleteList=function(e,t){var n=$.Deferred(),i=e.get_lists().getByTitle(t);return this.context.load(i),i.deleteObject(),r.executeQuery(this.context).fail(function(e,t){n.reject(e,t)}).done(function(){n.resolve()}),n.promise()},e.prototype.getFileListItem=function(e,t){void 0===t&&(t=null);var n=this.context.get_site().get_rootWeb().getFileByServerRelativeUrl(e).get_listItemAllFields();return this.loadListItem(n)},e.prototype.getListItemById=function(e,t,n,r){void 0===r&&(r=null);var i=e.get_lists().getByTitle(t).getItemById(n);return this.loadListItem(i,r)},e.prototype.deleteListItemById=function(e,t,n){var i=$.Deferred(),o=e.get_lists().getByTitle(t).getItemById(n);return this.context.load(o),o.deleteObject(),r.executeQuery(this.context).fail(function(e,t){i.reject(e,t)}).done(function(){i.resolve()}),i.promise()},e.prototype.getListItems=function(e,t,n){var i=$.Deferred(),o=e.get_lists().getByTitle(t),s=new SP.CamlQuery;n||(n="<View><Query></Query></Where>"),s.set_viewXml(n);var u=o.getItems(s);return this.context.load(u),r.executeQuery(this.context).fail(function(e,t){i.reject(e,t)}).done(function(){i.resolve(u)}),i.promise()},e.prototype.addContentTypeListAssociation=function(e,t,n){var i=this,o=$.Deferred(),s=e.get_lists().getByTitle(t),u=s.get_contentTypes();this.context.load(s),this.context.load(u);var c=this.context.get_site().get_rootWeb().get_contentTypes(),a=function(e,t){for(var n=0;n<e.get_count();n++)if(e.getItemAtIndex(n).get_name().toLowerCase()===t.toLowerCase())return e.getItemAtIndex(n);return null};return this.context.load(c),r.executeQuery(this.context).fail(function(e,t){o.reject(e,t)}).done(function(){var e=a(c,n);if(e){var t=a(u,n);if(t)o.resolve(t);else{s.get_contentTypesEnabled()||s.set_contentTypesEnabled(!0);var l=s.get_contentTypes().addExistingContentType(e);s.update(),r.executeQuery(i.context).fail(function(e,t){o.reject(e,t)}).done(function(){o.resolve(l)})}}else r.reject(o,"Content Type "+n+" not found")}),o.promise()},e.prototype.removeContentTypeListAssociation=function(e,t,n){var i=this,o=$.Deferred(),s=e.get_lists().getByTitle(t),u=s.get_contentTypes();this.context.load(s),this.context.load(u);return r.executeQuery(this.context).fail(function(e,t){o.reject(e,t)}).done(function(){var e=function(e,t){for(var n=0;n<e.get_count();n++)if(e.getItemAtIndex(n).get_name().toLowerCase()===t.toLowerCase())return e.getItemAtIndex(n);return null}(u,n);e?(e.deleteObject(),s.update(),r.executeQuery(i.context).fail(function(e,t){o.reject(e,t)}).done(function(){o.resolve()})):o.resolve()}),o.promise()},e.prototype.setDefaultValueOnList=function(e,t,n,i){var o=this,s=$.Deferred(),u=e.get_lists().getByTitle(t),c=u.get_fields();this.context.load(u),this.context.load(c);return r.executeQuery(this.context).fail(function(e,t){s.reject(e,t)}).done(function(){var e=function(e,t){for(var n=0;n<e.get_count();n++)if(e.getItemAtIndex(n).get_internalName().toLowerCase()===t.toLowerCase())return e.getItemAtIndex(n);return null}(c,n);e?(e.set_defaultValue(i),e.update(),r.executeQuery(o.context).fail(function(e,t){s.reject(e,t)}).done(function(){s.resolve()})):r.reject(s,"Field "+n+" not found")}),s.promise()},e}(),a=function(){function e(e){this.fluent=null,this.listHelper=null,this._helperName="list",this.fluent=e,this.listHelper=new c(e.context)}return e.prototype.create=function(e,t,n){var r=this;return this.fluent.chainAction(this._helperName+".create",function(){return r.listHelper.createList(e,t,n)})},e.prototype.exists=function(e,t){var n=this;return this.fluent.chainAction(this._helperName+".exists",function(){return n.listHelper.exists(e,t)})},e.prototype.delete=function(e,t){var n=this;return this.fluent.chainAction(this._helperName+".delete",function(){return n.listHelper.deleteList(e,t)})},e.prototype.get=function(e,t){var n=this;return this.fluent.chainAction(this._helperName+".get",function(){return n.listHelper.getList(e,t)})},e.prototype.addContentTypeListAssociation=function(e,t,n){var r=this;return this.fluent.chainAction(this._helperName+".addContentTypeListAssociation",function(){return r.listHelper.addContentTypeListAssociation(e,t,n)})},e.prototype.removeContentTypeListAssociation=function(e,t,n){var r=this;return this.fluent.chainAction(this._helperName+".removeContentTypeListAssociation",function(){return r.listHelper.removeContentTypeListAssociation(e,t,n)})},e.prototype.setDefaultValueOnList=function(e,t,n,r){var i=this;return this.fluent.chainAction(this._helperName+".setDefaultValueOnList",function(){return i.listHelper.setDefaultValueOnList(e,t,n,r)})},e.prototype.setAlerts=function(e,t,n){var r=this;return this.fluent.chainAction(this._helperName+".setAlerts",function(){return r.listHelper.setAlerts(e,t,n)})},e}(),l=function(){function e(e){this.fluent=null,this.listHelper=null,this._helperName="listItem",this.fluent=e,this.listHelper=new c(e.context)}return e.prototype.update=function(e,t){var n=this;return this.fluent.chainAction(this._helperName+".update",function(){return n.listHelper.setListItemProperties(e,t)})},e.prototype.createWithContentType=function(e,t,n,r){var i=this;return this.fluent.chainAction(this._helperName+".createWithContentType",function(){return i.listHelper.createListItemWithContentTypeName(e,t,n,r)})},e.prototype.create=function(e,t,n){var r=this;return this.fluent.chainAction(this._helperName+".create",function(){return r.listHelper.createListItem(e,t,n)})},e.prototype.get=function(e,t,n,r){var i=this;return void 0===r&&(r=null),this.fluent.chainAction(this._helperName+".get",function(){return i.listHelper.getListItemById(e,t,n,r)})},e.prototype.getFileListItem=function(e,t){var n=this;return void 0===t&&(t=null),this.fluent.chainAction(this._helperName+".getFileListItem",function(){return n.listHelper.getFileListItem(e,t)})},e.prototype.deleteById=function(e,t,n){var r=this;return this.fluent.chainAction(this._helperName+".deleteById",function(){return r.listHelper.deleteListItemById(e,t,n)})},e.prototype.query=function(e,t,n){var r=this;return this.fluent.chainAction(this._helperName+".query",function(){return r.listHelper.getListItems(e,t,n)})},e}(),f=function(){function e(e){this.fluent=null,this.listHelper=null,this._helperName="file",this.fluent=e,this.listHelper=new c(e.context)}return e.prototype.get=function(e){var t=this;return this.fluent.chainAction(this._helperName+".getListItem",function(){return t.listHelper.getFile(e)})},e.prototype.getListItem=function(e){var t=this;return this.fluent.chainAction(this._helperName+".getListItem",function(){return t.listHelper.getFileListItem(e)})},e.prototype.checkIn=function(e,t,n,r){var i=this;return this.fluent.chainAction(this._helperName+".checkInFile",function(){return i.listHelper.checkInFile(e,t,n,r)})},e}(),h=function(){function e(e){this.context=e}return e.prototype.getUserByEmail=function(e){return this.loadUser(this.context.get_web().ensureUser(e))},e.prototype.getUserById=function(e){return this.loadUser(this.context.get_web().get_siteUsers().getById(e))},e.prototype.loadUser=function(e){var t=$.Deferred();return this.context.load(e),this.context.executeQueryAsync(function(n,r){t.resolve(e)},function(e,n){console.error(n.get_message()),t.reject(e,n)}),t.promise()},e.prototype.getCurrentUser=function(){var e=$.Deferred(),t=this.context.get_web().get_currentUser();return this.context.load(t),this.context.executeQueryAsync(function(n,r){e.resolve(t)},function(t,n){console.error(n.get_message()),e.reject(t,n)}),e.promise()},e.prototype.getCurrentUserProfileProperties=function(){var e=this,t=$.Deferred();return SP.SOD.executeFunc("userprofile","SP.UserProfiles.PeopleManager",function(){var n=e.context,r=n.get_web().get_currentUser(),i=new SP.UserProfiles.PeopleManager(n).getMyProperties();n.load(r),n.load(i),n.executeQueryAsync(function(e,n){t.resolve(i.get_userProfileProperties())},function(e,n){console.error(n.get_message()),t.reject(e,n)})}),t.promise()},e.prototype.getCurrentUserManager=function(){var e=this,t=$.Deferred(),n=new SP.UserProfiles.PeopleManager(this.context),i=this.context.get_web().get_currentUser().get_email(),o=new SP.UserProfiles.UserProfilePropertiesForUser(this.context,"i:0#.f|membership|"+i,["Manager"]),s=n.getUserProfilePropertiesFor(o);return this.context.load(o),r.executeQuery(this.context).fail(function(e,n){t.reject(e,n)}).done(function(){if(s[0]){var n=e.context.get_web().ensureUser(s[0]);e.context.load(n),r.executeQuery(e.context).fail(function(e,n){t.reject(e,n)}).done(function(){t.resolve(n)})}else t.resolve(null)}),t.promise()},e}(),p=function(){function e(e){this._fluent=null,this._helperName="userProfile",this.userHelper=null,this._fluent=e,this.userHelper=new h(e.context)}return e.prototype.get=function(e){var t=this;return this._fluent.chainAction(this._helperName+".get",function(){return t.userHelper.getUserByEmail(e)})},e.prototype.getById=function(e){var t=this;return this._fluent.chainAction(this._helperName+".getById",function(){return t.userHelper.getUserById(e)})},e.prototype.getCurrentUser=function(){var e=this;return this._fluent.chainAction(this._helperName+".getCurrentUser",function(){return e.userHelper.getCurrentUser()})},e.prototype.getCurrentUserProfileProperties=function(){var e=this;return this._fluent.registerDependency(b.UserProfile),this._fluent.chainAction(this._helperName+".getCurrentUserProfileProperties",function(){return e.userHelper.getCurrentUserProfileProperties()})},e.prototype.getCurrentUserManager=function(){var e=this;return this._fluent.registerDependency(b.UserProfile),this._fluent.chainAction(this._helperName+".getCurrentUserManager",function(){return e.userHelper.getCurrentUserManager()})},e}(),d=function(){function e(e){this.context=e}return e.prototype.deleteQuicklaunchNodes=function(e){return this.deleteNodes(e.get_navigation().get_quickLaunch())},e.prototype.deleteTopNavigationNodes=function(e){return this.deleteNodes(e.get_navigation().get_topNavigationBar())},e.prototype.deleteQuicklaunchNode=function(e,t){return this.deleteNode(e.get_navigation().get_quickLaunch(),t)},e.prototype.deleteTopQuicklaunchNode=function(e,t){return this.deleteNode(e.get_navigation().get_topNavigationBar(),t)},e.prototype.setCurrentNavigation=function(e,t,n,i){var o=this;void 0===n&&(n=!0),void 0===i&&(i=!0);var s=$.Deferred(),u=e.get_allProperties();this.context.load(e),this.context.load(u);var c=function(){i&&n?u.set_item("__CurrentNavigationIncludeTypes","3"):i&&!n?u.set_item("__CurrentNavigationIncludeTypes","2"):!i&&n?u.set_item("__CurrentNavigationIncludeTypes","1"):i||n||u.set_item("__CurrentNavigationIncludeTypes","0")};return r.executeQuery(this.context).fail(function(e,t){s.reject(e,t)}).done(function(){var n=new SP.Publishing.Navigation.WebNavigationSettings(o.context,e);switch(t){case x.Inherit:n.get_currentNavigation().set_source(SP.Publishing.Navigation.StandardNavigationSource.inheritFromParentWeb);break;case x.Managed:r.reject(s,"Not implemented");break;case x.StructuralWithSiblings:n.get_currentNavigation().set_source(SP.Publishing.Navigation.StandardNavigationSource.portalProvider),u.set_item("__NavigationShowSiblings","True"),c(),e.update();break;case x.StructuralChildrenOnly:n.get_currentNavigation().set_source(SP.Publishing.Navigation.StandardNavigationSource.portalProvider),u.set_item("__NavigationShowSiblings","False"),c(),e.update();break;default:r.reject(s,"Unknown Navigation Type")}n.update(null),r.executeQuery(o.context).fail(function(e,t){s.reject(e,t)}).done(function(){s.resolve()})}),s.promise()},e.prototype.deleteNodes=function(e){var t=this,n=$.Deferred();return this.context.load(e),r.executeQuery(this.context).fail(function(e,t){n.reject(e,t)}).done(function(){for(var i=e.getEnumerator(),o=[];i.moveNext();)o.push(i.get_current());for(var s=0;s<o.length;s++)o[s].deleteObject();r.executeQuery(t.context).fail(function(e,t){n.reject(e,t)}).done(function(){n.resolve()})}),n.promise()},e.prototype.deleteNode=function(e,t){var n=this,i=$.Deferred();return this.context.load(e),r.executeQuery(this.context).fail(function(e,t){i.reject(e,t)}).done(function(){for(var o=e.getEnumerator(),s=[];o.moveNext();){var u=o.get_current();u.get_title()===t&&s.push(u)}for(var c=0;c<s.length;c++)s[c].deleteObject();r.executeQuery(n.context).fail(function(e,t){i.reject(e,t)}).done(function(){i.resolve()})}),i.promise()},e.prototype.createQuicklaunchNode=function(e,t,n,r){return void 0===r&&(r=!0),this.createNode(e.get_navigation().get_quickLaunch(),t,n,r)},e.prototype.createTopNavigationNode=function(e,t,n,r){return void 0===r&&(r=!0),this.createNode(e.get_navigation().get_topNavigationBar(),t,n,r)},e.prototype.createNode=function(e,t,n,i){void 0===i&&(i=!0);var o=$.Deferred();this.context.load(e);var s=new SP.NavigationNodeCreationInformation;s.set_title(t),s.set_url(n),s.set_asLastNode(i);var u=e.add(s);return this.context.load(e),r.executeQuery(this.context).fail(function(e,t){o.reject(e,t)}).done(function(){o.resolve(u)}),o.promise()},e}(),g=function(){function e(e){this.fluent=null,this.navigationHelper=null,this._helperName="navigation",this.fluent=e,this.navigationHelper=new d(e.context)}return e.prototype.createNode=function(e,t,n,r,i){var o=this;return void 0===i&&(i=!0),this.fluent.chainAction(this._helperName+".createNode",function(){if(t==_.Quicklaunch)return o.navigationHelper.createQuicklaunchNode(e,n,r,i);if(t==_.TopNavigation)return o.navigationHelper.createTopNavigationNode(e,n,r,i);throw"Unknown location "+t})},e.prototype.deleteAllNodes=function(e,t){var n=this;return this.fluent.chainAction(this._helperName+".deleteAllNodes",function(){if(t==_.Quicklaunch)return n.navigationHelper.deleteQuicklaunchNodes(e);if(t==_.TopNavigation)return n.navigationHelper.deleteTopNavigationNodes(e);throw"Unknown location "+t})},e.prototype.deleteNode=function(e,t,n){var r=this;return this.fluent.chainAction(this._helperName+".deleteNode",function(){if(t==_.Quicklaunch)return r.navigationHelper.deleteQuicklaunchNode(e,n);if(t==_.TopNavigation)return r.navigationHelper.deleteTopQuicklaunchNode(e,n);throw"Unknown location "+t})},e.prototype.setCurrentNavigation=function(e,t,n,r){var i=this;return void 0===n&&(n=!1),void 0===r&&(r=!1),this.fluent.registerDependency(b.Publishing),this.fluent.chainAction(this._helperName+".setCurrentNavigation",function(){return i.navigationHelper.setCurrentNavigation(e,t,n,r)})},e}(),m=function(){function e(e){this.context=e}return e.prototype.createPublishingPage=function(e,t,n){var i=$.Deferred(),o=this.context.get_site().get_rootWeb().getFileByServerRelativeUrl(n).get_listItemAllFields();this.context.load(o);var s=new SP.Publishing.PublishingPageInformation;s.set_name(t),s.set_pageLayoutListItem(o);var u=SP.Publishing.PublishingWeb.getPublishingWeb(this.context,e);this.context.load(u);var c=u.addPublishingPage(s);return this.context.load(c),r.executeQuery(this.context).fail(function(e,t){i.reject(e,t)}).done(function(){i.resolve(c)}),i.promise()},e.prototype.getPublishingLayout=function(e){var t=$.Deferred(),n=this.context.get_site().get_rootWeb().getFileByServerRelativeUrl(e).get_listItemAllFields();return this.context.load(n),r.executeQuery(this.context).fail(function(e,n){t.reject(e,n)}).done(function(){t.resolve(n)}),t.promise()},e}(),v=function(){function e(e){this.fluent=null,this._helperName="publishingPage",this.fluent=e,this.pageHelper=new m(e.context)}return e.prototype.create=function(e,t,n){var r=this;return this.fluent.registerDependency(b.Publishing),this.fluent.chainAction(this._helperName+".create",function(){return r.pageHelper.createPublishingPage(e,t,n)})},e.prototype.getLayout=function(e){var t=this;return this.fluent.chainAction(this._helperName+".getLayout",function(){return t.pageHelper.getPublishingLayout(e)})},e}();n.d(t,"Fluent",function(){return N}),n.d(t,"NavigationLocation",function(){return _}),n.d(t,"NavigationType",function(){return x}),n.d(t,"Dependency",function(){return b}),n.d(t,"FluentCommand",function(){return w}),n.d(t,"ActionCommand",function(){return I}),n.d(t,"ActionResult",function(){return j}),n.d(t,"WhenCommand",function(){return C});var y,_,x,b,P=(y=function(e,t){return(y=Object.setPrototypeOf||{__proto__:[]}instanceof Array&&function(e,t){e.__proto__=t}||function(e,t){for(var n in t)t.hasOwnProperty(n)&&(e[n]=t[n])})(e,t)},function(e,t){function n(){this.constructor=e}y(e,t),e.prototype=null===t?Object.create(t):(n.prototype=t.prototype,new n)}),N=function(){function e(){this.commands=[],this.results=[],this.dependencies=[],this.settings={timeoutMilliseconds:5e3,enableDependencyTimeout:!0},this.totalCommands=0}return e.prototype.withContext=function(e){return this.context=e,this},e.prototype.withSettings=function(e){for(var t in e)void 0!==this.settings[t]&&(this.settings[t]=e[t]);return this},e.prototype.withDependency=function(e){return this.registerDependency(e),this},Object.defineProperty(e.prototype,"promise",{get:function(){return this.resultPromise.promise()},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"permission",{get:function(){return new u(this)},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"list",{get:function(){return new a(this)},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"listItem",{get:function(){return new l(this)},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"file",{get:function(){return new f(this)},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"publishingPage",{get:function(){return new v(this)},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"web",{get:function(){return new o(this)},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"user",{get:function(){return new p(this)},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"navigation",{get:function(){return new g(this)},enumerable:!0,configurable:!0}),e.prototype.execute=function(){var e=this;if(this.resultPromise=$.Deferred(),this.settings.enableDependencyTimeout)var t=setTimeout(function(){r.reject(e.resultPromise,"Timeout waiting for dependencies to load")},this.settings.timeoutMilliseconds);return this.totalCommands=this.getAvailableActionCommandCount(),this.loadDependencies().done(function(){e.continue()}).always(function(){e.settings.enableDependencyTimeout&&clearTimeout(t)}),this.resultPromise.promise()},e.prototype.onActionExecuted=function(e){return this.onExecuted=e,this},e.prototype.onActionExecuting=function(e){return this.onExecuting=e,this},e.prototype.when=function(e){var t=this;if(this.peekLastCommand()&&this.peekLastCommand().constructor!==I)throw"Illegal operation: The last command must be an ActionCommand";var n=new C;return n.action=function(){var n=$.Deferred(),r=t.peekLastResult();if(r){if(!r.success)throw"Illegal operation: The last command should have succeeded for this call to have been made.  This is an issue in the fluent api";e.apply(void 0,r.result)?n.resolve():t.reject(n,t,"Predicate returned false")}else t.failChain(t,"No action to process for when");return n.promise()},this.commands.push(n),this},e.prototype.whenAll=function(e){var t=this;if(this.peekLastCommand()&&this.peekLastCommand().constructor!==I)throw"Illegal operation: The last command must be an ActionCommand";var n=new C;return n.action=function(){var n=$.Deferred(),r=t.peekLastResult();if(r){if(!r.success)throw"Illegal operation: The last command should have succeeded for this call to have been made.  This is an issue in the fluent api";e(t.results)?n.resolve():t.reject(n,t,"Predicate returned false")}else t.failChain(t,"No action to process for whenAll");return n.promise()},this.commands.push(n),this},e.prototype.whenTrue=function(){var e=this;if(this.peekLastCommand()&&this.peekLastCommand().constructor!==I)throw"Illegal operation: The last command must be an ActionCommand";var t=new C;return t.action=function(){var t=$.Deferred(),n=e.peekLastResult();if(n){if(!n.success)throw"Illegal operation: The last command should have succeeded for this call to have been made.  This is an issue in the fluent api";n.result&&n.result.length&&n.result[0]?t.resolve():e.reject(t,e,"Result is not true")}else e.failChain(e,"No action to process for whenTrue");return t.promise()},this.commands.push(t),this},e.prototype.whenFalse=function(){var e=this;if(this.peekLastCommand()&&this.peekLastCommand().constructor!==I)throw"Illegal operation: The last command must be an ActionCommand";var t=new C;return t.action=function(){var t=$.Deferred(),n=e.peekLastResult();if(n){if(!n.success)throw"Illegal operation: The last command should have succeeded for this call to have been made.  This is an issue in the fluent api";n.result&&n.result.length&&!n.result[0]?t.resolve():e.reject(t,e,"Result is not false")}else e.failChain(e,"No action to process for whenFalse");return t.promise()},this.commands.push(t),this},e.prototype.chainAction=function(e,t){var n=new I;return n.name=e,n.action=t,this.commands.push(n),this},e.prototype.chain=function(e){return this.commands.push(e),this},e.executeQuery=function(e){return r.executeQuery(e)},e.prototype.registerDependency=function(e){-1===this.dependencies.indexOf(e)&&this.dependencies.push(e)},e.prototype.continue=function(){var e=this;if(!this.context)throw"context not set, you must call withContext";var t=this.commands.shift();if(t&&t.action){if(t.constructor===I&&this.onExecuting){var n=this.totalCommands-this.getAvailableActionCommandCount();this.onExecuting(t.name,n,this.totalCommands)}t.action().done(function(n,r,i,o,s,u,c){if(t.constructor===I){var a=[];e.storeResult(n,a),e.storeResult(r,a),e.storeResult(i,a),e.storeResult(o,a),e.storeResult(s,a),e.storeResult(u,a),e.storeResult(c,a),e.addResult(t,!0,a),e.onExecuted&&e.onExecuted(t.name,!0,a)}e.commands.length?e.continue():e.resolveChain()}).fail(function(n,r){if(t.constructor===I){var i=[];return e.storeResult(n,i),e.storeResult(r,i),e.addResult(t,!1,i),e.onExecuted&&e.onExecuted(t.name,!1,{sender:n,args:r}),void e.failChain(n,r)}e.resolveChain()})}else this.resolveChain()},e.prototype.storeResult=function(e,t){void 0!==e&&t.push(e)},e.prototype.resolveChain=function(){this.resultPromise.resolve(this.results)},e.prototype.failChain=function(e,t){"string"==typeof t&&(t={get_message:function(){return t}}),this.resultPromise.reject(e,t)},e.prototype.reject=function(e,t,n){"string"==typeof n&&(n={get_message:function(){return n}}),e.reject(t,n)},e.prototype.addResult=function(e,t,n){var r=new j;r.name=e.name,r.success=t,r.result=n,this.results.push(r)},e.prototype.loadDependencies=function(){for(var e=$.Deferred(),t=["SP.js","SP.Runtime.js"],n=0;n<this.dependencies.length;n++)switch(this.dependencies[n]){case b.UserProfile:window.LoadSodByKey("userprofile"),t.push("userprofile");break;case b.Publishing:SP.SOD.registerSod("SP.Publishing.js",SP.Utilities.Utility.getLayoutsPageUrl("sp.publishing.js")),t.push("SP.Publishing.js");break;case b.Taxonomy:SP.SOD.registerSod("sp.taxonomy.js",SP.Utilities.Utility.getLayoutsPageUrl("sp.taxonomy.js")),t.push("sp.taxonomy.js");break;default:throw"Depenency not implemented"}return SP.SOD.loadMultiple(t,function(){e.resolve()}),e.promise()},e.prototype.peekLastCommand=function(){return this.commands.length?this.commands[this.commands.length-1]:null},e.prototype.peekLastResult=function(){return this.results.length?this.results[this.results.length-1]:null},e.prototype.getAvailableActionCommandCount=function(){for(var e=0,t=0;t<this.commands.length;t++)this.commands[t].constructor===I&&e++;return e},e}();!function(e){e[e.TopNavigation=0]="TopNavigation",e[e.Quicklaunch=1]="Quicklaunch"}(_||(_={})),function(e){e[e.Inherit=0]="Inherit",e[e.Managed=1]="Managed",e[e.StructuralWithSiblings=2]="StructuralWithSiblings",e[e.StructuralChildrenOnly=3]="StructuralChildrenOnly"}(x||(x={})),function(e){e[e.Publishing=0]="Publishing",e[e.UserProfile=1]="UserProfile",e[e.Taxonomy=2]="Taxonomy"}(b||(b={}));var w=function(){return function(){}}(),I=function(e){function t(){return null!==e&&e.apply(this,arguments)||this}return P(t,e),t}(w),j=function(){return function(){}}(),C=function(e){function t(){return null!==e&&e.apply(this,arguments)||this}return P(t,e),t}(w)}])});
//# sourceMappingURL=spJsomFluent.umd.js.map