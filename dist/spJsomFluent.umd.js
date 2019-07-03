/*! spJsomFluent "0.1.18" - https://github.com/TjWheeler/spJsomFluent */
(function webpackUniversalModuleDefinition(root, factory) {
	if(typeof exports === 'object' && typeof module === 'object')
		module.exports = factory();
	else if(typeof define === 'function' && define.amd)
		define([], factory);
	else if(typeof exports === 'object')
		exports["spJsom"] = factory();
	else
		root["spJsom"] = factory();
})(window, function() {
return /******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId]) {
/******/ 			return installedModules[moduleId].exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			i: moduleId,
/******/ 			l: false,
/******/ 			exports: {}
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.l = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// define getter function for harmony exports
/******/ 	__webpack_require__.d = function(exports, name, getter) {
/******/ 		if(!__webpack_require__.o(exports, name)) {
/******/ 			Object.defineProperty(exports, name, { enumerable: true, get: getter });
/******/ 		}
/******/ 	};
/******/
/******/ 	// define __esModule on exports
/******/ 	__webpack_require__.r = function(exports) {
/******/ 		if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 			Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 		}
/******/ 		Object.defineProperty(exports, '__esModule', { value: true });
/******/ 	};
/******/
/******/ 	// create a fake namespace object
/******/ 	// mode & 1: value is a module id, require it
/******/ 	// mode & 2: merge all properties of value into the ns
/******/ 	// mode & 4: return value when already ns object
/******/ 	// mode & 8|1: behave like require
/******/ 	__webpack_require__.t = function(value, mode) {
/******/ 		if(mode & 1) value = __webpack_require__(value);
/******/ 		if(mode & 8) return value;
/******/ 		if((mode & 4) && typeof value === 'object' && value && value.__esModule) return value;
/******/ 		var ns = Object.create(null);
/******/ 		__webpack_require__.r(ns);
/******/ 		Object.defineProperty(ns, 'default', { enumerable: true, value: value });
/******/ 		if(mode & 2 && typeof value != 'string') for(var key in value) __webpack_require__.d(ns, key, function(key) { return value[key]; }.bind(null, key));
/******/ 		return ns;
/******/ 	};
/******/
/******/ 	// getDefaultExport function for compatibility with non-harmony modules
/******/ 	__webpack_require__.n = function(module) {
/******/ 		var getter = module && module.__esModule ?
/******/ 			function getDefault() { return module['default']; } :
/******/ 			function getModuleExports() { return module; };
/******/ 		__webpack_require__.d(getter, 'a', getter);
/******/ 		return getter;
/******/ 	};
/******/
/******/ 	// Object.prototype.hasOwnProperty.call
/******/ 	__webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";
/******/
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = "./src/fluent.ts");
/******/ })
/************************************************************************/
/******/ ({

/***/ "./src/common.ts":
/*!***********************!*\
  !*** ./src/common.ts ***!
  \***********************/
/*! exports provided: common */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "common", function() { return common; });
var common = /** @class */ (function () {
    function common() {
    }
    common.reject = function (deferred, reason) {
        deferred.reject(this, { get_message: function () { return reason; } });
    };
    common.notImplementedPromise = function () {
        var deferred = $.Deferred();
        common.reject(deferred, "Not Implemented");
        return deferred.promise();
    };
    common.FilterArray = function (items, predicate) {
        var output = [];
        for (var i = 0; i < items.length; i++) {
            if (predicate(items[i], i)) {
                output.push(items[i]);
            }
        }
        return output;
    };
    common.executeQuery = function (clientContext) {
        var deferred = $.Deferred();
        clientContext.executeQueryAsync(function (sender, args) {
            deferred.resolve(sender, args);
        }, function (sender, args) {
            deferred.reject(sender, args);
        });
        return deferred.promise();
    };
    return common;
}());



/***/ }),

/***/ "./src/file.ts":
/*!*********************!*\
  !*** ./src/file.ts ***!
  \*********************/
/*! exports provided: File */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "File", function() { return File; });
/* harmony import */ var _helper_listHelper__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./helper/listHelper */ "./src/helper/listHelper.ts");

var File = /** @class */ (function () {
    function File(fluent) {
        this.fluent = null;
        this.listHelper = null;
        this._helperName = "file";
        this.fluent = fluent;
        this.listHelper = new _helper_listHelper__WEBPACK_IMPORTED_MODULE_0__["ListHelper"](fluent.context);
    }
    /**
    * Get a file from a document library
    * Result: SP.File
    * Example: get(_spPageContextInfo.siteServerRelativeUrl + "/documents/doc.docx")
    */
    File.prototype.get = function (serverRelativeUrl) {
        var _this = this;
        return this.fluent.chainAction(this._helperName + ".getListItem", function () {
            return _this.listHelper.getFile(serverRelativeUrl);
        });
    };
    /**
    * Get the list item for the file
    * Result: SP.ListItem
    * Example: getListItem(_spPageContextInfo.siteServerRelativeUrl + "/documents/doc.docx")
    */
    File.prototype.getListItem = function (serverRelativeUrl) {
        var _this = this;
        return this.fluent.chainAction(this._helperName + ".getListItem", function () {
            return _this.listHelper.getFileListItem(serverRelativeUrl);
        });
    };
    /**
    * Check in a file
    * Result: SP.File
    * Example: checkIn(SP.ClientContext.get_current().get_web(), _spPageContextInfo.webServerRelativeUrl + '/pages/mypage.aspx', "Checked in by spJsomFluent", SP.CheckinType.majorCheckIn)
    */
    File.prototype.checkIn = function (web, serverRelativeUrl, comment, checkInType) {
        var _this = this;
        return this.fluent.chainAction(this._helperName + ".checkInFile", function () {
            return _this.listHelper.checkInFile(web, serverRelativeUrl, comment, checkInType);
        });
    };
    return File;
}());



/***/ }),

/***/ "./src/fluent.ts":
/*!***********************!*\
  !*** ./src/fluent.ts ***!
  \***********************/
/*! exports provided: Fluent, NavigationLocation, NavigationType, Dependency, FluentCommand, ActionCommand, ActionResult, WhenCommand */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Fluent", function() { return Fluent; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "NavigationLocation", function() { return NavigationLocation; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "NavigationType", function() { return NavigationType; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Dependency", function() { return Dependency; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "FluentCommand", function() { return FluentCommand; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "ActionCommand", function() { return ActionCommand; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "ActionResult", function() { return ActionResult; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "WhenCommand", function() { return WhenCommand; });
/* harmony import */ var _web__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./web */ "./src/web.ts");
/* harmony import */ var _permission__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./permission */ "./src/permission.ts");
/* harmony import */ var _list__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./list */ "./src/list.ts");
/* harmony import */ var _listitem__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./listitem */ "./src/listitem.ts");
/* harmony import */ var _file__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ./file */ "./src/file.ts");
/* harmony import */ var _user__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ./user */ "./src/user.ts");
/* harmony import */ var _navigation__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ./navigation */ "./src/navigation.ts");
/* harmony import */ var _publishingpage__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ./publishingpage */ "./src/publishingpage.ts");
/* harmony import */ var _common__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ./common */ "./src/common.ts");
var __extends = (undefined && undefined.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
// spJsom - Fluent









var Fluent = /** @class */ (function () {
    function Fluent() {
        this.commands = [];
        this.results = [];
        this.dependencies = [];
        this.settings = { timeoutMilliseconds: 5000, enableDependencyTimeout: true };
        this.totalCommands = 0;
    }
    Fluent.prototype.withContext = function (context) {
        this.context = context;
        return this;
    };
    Fluent.prototype.withSettings = function (settings) {
        for (var setting in settings) {
            if (typeof (this.settings[setting]) !== typeof (undefined)) {
                this.settings[setting] = settings[setting];
            }
        }
        return this;
    };
    Fluent.prototype.withDependency = function (dependency) {
        this.registerDependency(dependency);
        return this;
    };
    Object.defineProperty(Fluent.prototype, "promise", {
        get: function () {
            return this.resultPromise.promise();
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Fluent.prototype, "permission", {
        get: function () {
            return new _permission__WEBPACK_IMPORTED_MODULE_1__["Permission"](this);
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Fluent.prototype, "list", {
        get: function () {
            return new _list__WEBPACK_IMPORTED_MODULE_2__["List"](this);
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Fluent.prototype, "listItem", {
        get: function () {
            return new _listitem__WEBPACK_IMPORTED_MODULE_3__["ListItem"](this);
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Fluent.prototype, "file", {
        get: function () {
            return new _file__WEBPACK_IMPORTED_MODULE_4__["File"](this);
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Fluent.prototype, "publishingPage", {
        get: function () {
            return new _publishingpage__WEBPACK_IMPORTED_MODULE_7__["PublishingPage"](this);
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Fluent.prototype, "web", {
        get: function () {
            return new _web__WEBPACK_IMPORTED_MODULE_0__["Web"](this);
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Fluent.prototype, "user", {
        get: function () {
            return new _user__WEBPACK_IMPORTED_MODULE_5__["User"](this);
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Fluent.prototype, "navigation", {
        get: function () {
            return new _navigation__WEBPACK_IMPORTED_MODULE_6__["Navigation"](this);
        },
        enumerable: true,
        configurable: true
    });
    Fluent.prototype.execute = function () {
        var _this = this;
        this.resultPromise = $.Deferred();
        //The dependency timeout is fired if it takes too long to load to avoid promise never completing.
        if (this.settings.enableDependencyTimeout) {
            var expiry = setTimeout(function () {
                _common__WEBPACK_IMPORTED_MODULE_8__["common"].reject(_this.resultPromise, "Timeout waiting for dependencies to load");
            }, this.settings.timeoutMilliseconds);
        }
        //To track progress
        this.totalCommands = this.getAvailableActionCommandCount();
        this.loadDependencies()
            .done(function () {
            _this.continue();
        })
            .always(function () {
            if (_this.settings.enableDependencyTimeout) {
                clearTimeout(expiry);
            }
        });
        return this.resultPromise.promise();
    };
    Fluent.prototype.onActionExecuted = function (onExecuted) {
        this.onExecuted = onExecuted;
        return this;
    };
    Fluent.prototype.onActionExecuting = function (onExecuting) {
        this.onExecuting = onExecuting;
        return this;
    };
    Fluent.prototype.when = function (predicate) {
        var _this = this;
        if (this.peekLastCommand() && this.peekLastCommand().constructor !== ActionCommand) {
            throw "Illegal operation: The last command must be an ActionCommand";
        }
        var command = new WhenCommand();
        command.action = function () {
            var deferred = $.Deferred();
            var command = _this.peekLastResult();
            if (command) {
                if (!command.success) {
                    throw "Illegal operation: The last command should have succeeded for this call to have been made.  This is an issue in the fluent api";
                }
                if (predicate.apply(void 0, command.result)) {
                    deferred.resolve();
                }
                else {
                    _this.reject(deferred, _this, "Predicate returned false");
                }
            }
            else {
                _this.failChain(_this, "No action to process for when");
            }
            return deferred.promise();
        };
        this.commands.push(command);
        return this;
    };
    //Similar to when, but all previous results are passed into the predicate
    Fluent.prototype.whenAll = function (predicate) {
        var _this = this;
        if (this.peekLastCommand() && this.peekLastCommand().constructor !== ActionCommand) {
            throw "Illegal operation: The last command must be an ActionCommand";
        }
        var command = new WhenCommand();
        command.action = function () {
            var deferred = $.Deferred();
            var command = _this.peekLastResult();
            if (command) {
                if (!command.success) {
                    throw "Illegal operation: The last command should have succeeded for this call to have been made.  This is an issue in the fluent api";
                }
                if (predicate(_this.results)) {
                    deferred.resolve();
                }
                else {
                    _this.reject(deferred, _this, "Predicate returned false");
                }
            }
            else {
                _this.failChain(_this, "No action to process for whenAll");
            }
            return deferred.promise();
        };
        this.commands.push(command);
        return this;
    };
    Fluent.prototype.whenTrue = function () {
        var _this = this;
        if (this.peekLastCommand() && this.peekLastCommand().constructor !== ActionCommand) {
            throw "Illegal operation: The last command must be an ActionCommand";
        }
        var command = new WhenCommand();
        command.action = function () {
            var deferred = $.Deferred();
            var command = _this.peekLastResult();
            if (command) {
                if (!command.success) {
                    throw "Illegal operation: The last command should have succeeded for this call to have been made.  This is an issue in the fluent api";
                }
                if (command.result && command.result.length && command.result[0]) {
                    deferred.resolve();
                }
                else {
                    _this.reject(deferred, _this, "Result is not true");
                }
            }
            else {
                _this.failChain(_this, "No action to process for whenTrue");
            }
            return deferred.promise();
        };
        this.commands.push(command);
        return this;
    };
    Fluent.prototype.whenFalse = function () {
        var _this = this;
        if (this.peekLastCommand() && this.peekLastCommand().constructor !== ActionCommand) {
            throw "Illegal operation: The last command must be an ActionCommand";
        }
        var command = new WhenCommand();
        command.action = function () {
            var deferred = $.Deferred();
            var command = _this.peekLastResult();
            if (command) {
                if (!command.success) {
                    throw "Illegal operation: The last command should have succeeded for this call to have been made.  This is an issue in the fluent api";
                }
                if (command.result && command.result.length && !command.result[0]) {
                    deferred.resolve();
                }
                else {
                    _this.reject(deferred, _this, "Result is not false");
                }
            }
            else {
                _this.failChain(_this, "No action to process for whenFalse");
            }
            return deferred.promise();
        };
        this.commands.push(command);
        return this;
    };
    Fluent.prototype.chainAction = function (name, action) {
        var command = new ActionCommand();
        command.name = name;
        command.action = action;
        this.commands.push(command);
        return this;
    };
    Fluent.prototype.chain = function (command) {
        this.commands.push(command);
        return this;
    };
    Fluent.executeQuery = function (context) {
        return _common__WEBPACK_IMPORTED_MODULE_8__["common"].executeQuery(context);
    };
    Fluent.prototype.registerDependency = function (dependency) {
        if (this.dependencies.indexOf(dependency) === -1) {
            this.dependencies.push(dependency);
        }
    };
    Fluent.prototype.continue = function () {
        var _this = this;
        if (!this.context) {
            throw "context not set, you must call withContext";
        }
        var command = this.commands.shift();
        if (command && command.action) {
            if (command.constructor === ActionCommand) {
                if (this.onExecuting) {
                    var step = this.totalCommands - (this.getAvailableActionCommandCount());
                    this.onExecuting(command.name, step, this.totalCommands);
                }
            }
            var promise = command.action();
            promise.done(function (arg1, arg2, arg3, arg4, arg5, arg6, arg7) {
                if (command.constructor === ActionCommand) {
                    var results = [];
                    _this.storeResult(arg1, results);
                    _this.storeResult(arg2, results);
                    _this.storeResult(arg3, results);
                    _this.storeResult(arg4, results);
                    _this.storeResult(arg5, results);
                    _this.storeResult(arg6, results);
                    _this.storeResult(arg7, results);
                    _this.addResult(command, true, results);
                    if (_this.onExecuted) {
                        _this.onExecuted(command.name, true, results);
                    }
                }
                if (_this.commands.length) {
                    _this.continue();
                }
                else {
                    _this.resolveChain();
                }
            })
                .fail(function (sender, args) {
                if (command.constructor === ActionCommand) {
                    var results = [];
                    _this.storeResult(sender, results);
                    _this.storeResult(args, results);
                    _this.addResult(command, false, results);
                    if (_this.onExecuted) {
                        _this.onExecuted(command.name, false, { sender: sender, args: args });
                    }
                    _this.failChain(sender, args);
                    return;
                }
                _this.resolveChain();
            });
        }
        else {
            this.resolveChain();
        }
    };
    Fluent.prototype.storeResult = function (arg, results) {
        if (typeof (arg) !== typeof (undefined)) {
            results.push(arg);
        }
    };
    Fluent.prototype.resolveChain = function () {
        this.resultPromise.resolve(this.results);
    };
    Fluent.prototype.failChain = function (sender, args) {
        if (typeof (args) == "string") {
            args = { get_message: function () { return args; } };
        }
        this.resultPromise.reject(sender, args);
    };
    Fluent.prototype.reject = function (deferred, sender, args) {
        if (typeof (args) == "string") {
            args = { get_message: function () { return args; } };
        }
        deferred.reject(sender, args);
    };
    Fluent.prototype.addResult = function (command, success, results) {
        var result = new ActionResult();
        result.name = command.name;
        result.success = success;
        result.result = results;
        this.results.push(result);
    };
    Fluent.prototype.loadDependencies = function () {
        var deferred = $.Deferred();
        var spDependencies = ["SP.js", "SP.Runtime.js"];
        for (var i = 0; i < this.dependencies.length; i++) {
            switch (this.dependencies[i]) {
                case Dependency.UserProfile:
                    window.LoadSodByKey("userprofile");
                    spDependencies.push("userprofile");
                    break;
                case Dependency.Publishing:
                    SP.SOD.registerSod('SP.Publishing.js', SP.Utilities.Utility.getLayoutsPageUrl('sp.publishing.js'));
                    spDependencies.push("SP.Publishing.js");
                    break;
                case Dependency.Taxonomy:
                    SP.SOD.registerSod('sp.taxonomy.js', SP.Utilities.Utility.getLayoutsPageUrl('sp.taxonomy.js'));
                    spDependencies.push("sp.taxonomy.js");
                    break;
                default:
                    throw "Depenency not implemented";
            }
        }
        SP.SOD.loadMultiple(spDependencies, function () {
            deferred.resolve();
        });
        return deferred.promise();
    };
    Fluent.prototype.peekLastCommand = function () {
        if (this.commands.length) {
            return this.commands[this.commands.length - 1];
        }
        return null;
    };
    Fluent.prototype.peekLastResult = function () {
        if (this.results.length) {
            return this.results[this.results.length - 1];
        }
        return null;
    };
    Fluent.prototype.getAvailableActionCommandCount = function () {
        var count = 0;
        for (var i = 0; i < this.commands.length; i++) {
            if (this.commands[i].constructor === ActionCommand) {
                count++;
            }
        }
        return count;
    };
    return Fluent;
}());

var NavigationLocation;
(function (NavigationLocation) {
    NavigationLocation[NavigationLocation["TopNavigation"] = 0] = "TopNavigation";
    NavigationLocation[NavigationLocation["Quicklaunch"] = 1] = "Quicklaunch";
})(NavigationLocation || (NavigationLocation = {}));
var NavigationType;
(function (NavigationType) {
    NavigationType[NavigationType["Inherit"] = 0] = "Inherit";
    NavigationType[NavigationType["Managed"] = 1] = "Managed";
    NavigationType[NavigationType["StructuralWithSiblings"] = 2] = "StructuralWithSiblings";
    NavigationType[NavigationType["StructuralChildrenOnly"] = 3] = "StructuralChildrenOnly";
})(NavigationType || (NavigationType = {}));
var Dependency;
(function (Dependency) {
    Dependency[Dependency["Publishing"] = 0] = "Publishing";
    Dependency[Dependency["UserProfile"] = 1] = "UserProfile";
    Dependency[Dependency["Taxonomy"] = 2] = "Taxonomy";
})(Dependency || (Dependency = {}));
var FluentCommand = /** @class */ (function () {
    function FluentCommand() {
    }
    return FluentCommand;
}());

var ActionCommand = /** @class */ (function (_super) {
    __extends(ActionCommand, _super);
    function ActionCommand() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    return ActionCommand;
}(FluentCommand));

var ActionResult = /** @class */ (function () {
    function ActionResult() {
    }
    return ActionResult;
}());

var WhenCommand = /** @class */ (function (_super) {
    __extends(WhenCommand, _super);
    function WhenCommand() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    return WhenCommand;
}(FluentCommand));



/***/ }),

/***/ "./src/helper/Client.CamlBuilder.ts":
/*!******************************************!*\
  !*** ./src/helper/Client.CamlBuilder.ts ***!
  \******************************************/
/*! exports provided: CamlOperator, EventRecurranceOverlap, AggregationType, StringBuilder, CamlBuilder, OrderBy */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "CamlOperator", function() { return CamlOperator; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "EventRecurranceOverlap", function() { return EventRecurranceOverlap; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "AggregationType", function() { return AggregationType; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "StringBuilder", function() { return StringBuilder; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "CamlBuilder", function() { return CamlBuilder; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "OrderBy", function() { return OrderBy; });

/// <summary>
///     CAML Operations
/// </summary>
var CamlOperator;
(function (CamlOperator) {
    CamlOperator[CamlOperator["Contains"] = 0] = "Contains";
    CamlOperator[CamlOperator["BeginsWith"] = 1] = "BeginsWith";
    CamlOperator[CamlOperator["Eq"] = 2] = "Eq";
    CamlOperator[CamlOperator["Geq"] = 3] = "Geq";
    CamlOperator[CamlOperator["Leq"] = 4] = "Leq";
    CamlOperator[CamlOperator["Lt"] = 5] = "Lt";
    CamlOperator[CamlOperator["Gt"] = 6] = "Gt";
    CamlOperator[CamlOperator["Neq"] = 7] = "Neq";
    CamlOperator[CamlOperator["DateRangesOverlap"] = 8] = "DateRangesOverlap";
    CamlOperator[CamlOperator["IsNotNull"] = 9] = "IsNotNull";
    CamlOperator[CamlOperator["IsNull"] = 10] = "IsNull";
})(CamlOperator || (CamlOperator = {}));
var EventRecurranceOverlap;
(function (EventRecurranceOverlap) {
    EventRecurranceOverlap[EventRecurranceOverlap["Now"] = 0] = "Now";
    EventRecurranceOverlap[EventRecurranceOverlap["Today"] = 1] = "Today";
    EventRecurranceOverlap[EventRecurranceOverlap["Week"] = 2] = "Week";
    EventRecurranceOverlap[EventRecurranceOverlap["Month"] = 3] = "Month";
    EventRecurranceOverlap[EventRecurranceOverlap["Year"] = 4] = "Year";
})(EventRecurranceOverlap || (EventRecurranceOverlap = {}));
var AggregationType;
(function (AggregationType) {
    AggregationType[AggregationType["SUM"] = 0] = "SUM";
    AggregationType[AggregationType["COUNT"] = 1] = "COUNT";
    AggregationType[AggregationType["AVG"] = 2] = "AVG";
    AggregationType[AggregationType["MAX"] = 3] = "MAX";
    AggregationType[AggregationType["MIN"] = 4] = "MIN";
    AggregationType[AggregationType["STDEV"] = 5] = "STDEV";
    AggregationType[AggregationType["VAR"] = 6] = "VAR";
})(AggregationType || (AggregationType = {}));
var StringBuilder = /** @class */ (function () {
    function StringBuilder() {
        this.value = "";
    }
    StringBuilder.prototype.add = function (value) {
        this.value += value;
    };
    StringBuilder.prototype.toString = function () {
        return this.value;
    };
    return StringBuilder;
}());

var CamlBuilder = /** @class */ (function () {
    function CamlBuilder() {
        this.camlClauses = [];
        this.orderByFields = [];
        this.requireAll = true;
        this.viewFieldsString = "";
        this.groupByFieldsString = "";
        this.aggregationFieldsString = "";
        this.begin(true);
    }
    Object.defineProperty(CamlBuilder.prototype, "query", {
        get: function () {
            return this.buildQuery();
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(CamlBuilder.prototype, "viewFields", {
        get: function () {
            return this.viewFieldsString;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(CamlBuilder.prototype, "totalClauses", {
        get: function () {
            return this.camlClauses.length;
        },
        enumerable: true,
        configurable: true
    });
    CamlBuilder.prototype.setFilter = function (caml) {
        var hasWhereClause = caml.indexOf("<Where>") === 0;
        var sb = new StringBuilder();
        if (!hasWhereClause) {
            sb.add("<Where>");
        }
        sb.add(caml);
        if (!hasWhereClause) {
            sb.add("</Where>");
        }
        this.directCaml = sb.toString();
    };
    CamlBuilder.prototype.buildQuery = function () {
        if (this.recurrence) {
            return "<Where>" + this.recurrence + "</Where>";
        }
        var query = "";
        var openOperators = 0;
        if (this.directCaml) {
            query += this.directCaml;
        }
        else if (!this.directCaml && this.camlClauses.length > 0) {
            query += "<Where>";
            //When we have just one clause we can't use AND or OR.
            if (this.camlClauses.length > 1) {
                var totalCamlPairs = this.camlClauses.length - 1;
                while (totalCamlPairs > 0) {
                    query += (this.requireAll ? "<And>" : "<Or>");
                    totalCamlPairs--;
                    openOperators++;
                }
            }
            var clausesAdded = 0;
            for (var i = 0; i < this.camlClauses.length; i++) {
                var clause = this.camlClauses[i];
                query += clause;
                clausesAdded++;
                if (clausesAdded > 1) {
                    query += this.requireAll ? "</And>" : "</Or>";
                    openOperators--;
                }
            }
            query += "</Where>";
        }
        if (this.orderByFields.length > 0) {
            query += "<OrderBy>";
            for (var i = 0; i < this.orderByFields.length; i++) {
                var item = this.orderByFields[i];
                query += "<FieldRef Name=\"" + item.fieldRef + "\" Ascending=\"" + (item.ascending ? "TRUE" : "FALSE") + "\" />";
            }
            query += "</OrderBy>";
        }
        return query;
    };
    Object.defineProperty(CamlBuilder.prototype, "viewXml", {
        get: function () {
            var viewFields = "";
            if (this.viewFields) {
                viewFields = "<ViewFields>" + this.viewFields + "</ViewFields>";
            }
            var groupBy = "";
            if (this.groupByFieldsString) {
                groupBy = "<GroupBy Collapse=\"TRUE\" GroupLimit=\"1999\">" + this.groupByFieldsString + "</GroupBy>";
            }
            var orderBy = "";
            if (this.orderByFields && this.orderByFields.length) {
                orderBy = "<OrderBy>";
                for (var i = 0; i < this.orderByFields.length; i++) {
                    orderBy += "<FieldRef Name=\"" + this.orderByFields[i].fieldRef + "\" Ascending=\"" + (this.orderByFields[i].ascending ? 'TRUE' : 'FALSE') + "\"/>";
                }
                orderBy += "</OrderBy>";
            }
            var aggregations = "";
            if (this.aggregationFieldsString) {
                aggregations = "<Aggregations Value=\"On\">" + this.aggregationFieldsString + "</Aggregations>";
            }
            var queryOptionText = "";
            if (this.queryOptions) {
                queryOptionText = "<QueryOptions>\n                " + this.queryOptions + "\n            </QueryOptions>";
            }
            return "<View" + (this.viewScope ? " " + this.viewScope : "") + ">" + viewFields + "<Query>" + groupBy + ((this.orderByFields && this.orderByFields.length > 0) || this.totalClauses > 0 || this.directCaml ? this.query : "") + orderBy + "</Query>" + queryOptionText + aggregations + (this.rowLimit ? "<RowLimit>" + this.rowLimit + "</RowLimit>" : "") + "</View>";
        },
        enumerable: true,
        configurable: true
    });
    CamlBuilder.prototype.begin = function (requireAll) {
        this.camlClauses = [];
        this.requireAll = requireAll;
        this.recurrence = null;
        this.viewFieldsString = "";
        this.orderByFields = [];
        this.directCaml = null;
    };
    CamlBuilder.prototype.getNullClause = function (fieldRef) {
        var retVal = "";
        if (fieldRef) {
            retVal = "<IsNull><FieldRef Name=\"" + fieldRef + "\" /></IsNull>";
        }
        return retVal;
    };
    CamlBuilder.prototype.getClause = function (operation, fieldRef, value, valueType) {
        var retVal = "<" + CamlOperator[operation] + "><FieldRef Name=\"" + fieldRef + "\" /><Value Type=\"" + valueType + "\">" + value + "</Value></" + CamlOperator[operation] + ">";
        return retVal;
    };
    CamlBuilder.prototype.getDateTimeClause = function (operation, fieldRef, value, includeTime) {
        var retVal = "<" + CamlOperator[operation] + "><FieldRef Name=\"" + fieldRef + "\" /><Value Type=\"DateTime\" " + (includeTime ? "IncludeTimeValue='TRUE'" : "") + ">" + value + "</Value></" + CamlOperator[operation] + ">";
        return retVal;
    };
    CamlBuilder.prototype.addNullClause = function (fieldRef) {
        this.camlClauses.push(this.getNullClause(fieldRef));
    };
    CamlBuilder.prototype.addTextClause = function (operation, fieldRef, value) {
        this.camlClauses.push(this.getClause(operation, fieldRef, value, "Text"));
    };
    CamlBuilder.prototype.addBooleanClause = function (operation, fieldRef, value) {
        this.camlClauses.push(this.getClause(operation, fieldRef, value ? "1" : "0", "Integer"));
    };
    CamlBuilder.prototype.addNumberClause = function (operation, fieldRef, value) {
        this.camlClauses.push(this.getClause(operation, fieldRef, value.toString(), "Number")); //TODO: verify type
    };
    CamlBuilder.prototype.addDateTimeClause = function (operation, fieldRef, value, includeTime) {
        this.camlClauses.push(this.getDateTimeClause(operation, fieldRef, value, includeTime));
    };
    CamlBuilder.prototype.addLookupClause = function (operation, fieldRef, value, valueType) {
        var clause = "<" + CamlOperator[operation] + "><FieldRef Name=\"" + fieldRef + "\" LookupId=\"True\" /><Value Type=\"" + valueType + "\">" + value + "</Value></" + CamlOperator[operation] + ">";
        this.camlClauses.push(clause);
    };
    CamlBuilder.prototype.addAggregation = function (fieldRef, aggregationType) {
        this.aggregationFieldsString = this.aggregationFieldsString + ("<FieldRef Name=\"" + fieldRef.replace(" ", "_x0020_") + "\" Type=\"" + AggregationType[aggregationType] + "\" />");
    };
    CamlBuilder.prototype.recurrenceQuery = function (overlapType) {
        this.recurrence = "<DateRangesOverlap><FieldRef Name=\"EventDate\"/><FieldRef Name=\"EndDate\"/><FieldRef Name=\"RecurrenceID\"/><Value><" + EventRecurranceOverlap[overlapType] + "/></Value></DateRangesOverlap>";
        this.queryOptions +=
            "<CalendarDate>2018-01-01T12:00:00Z</CalendarDate>\n            <RecurrencePatternXMLVersion>v3</RecurrencePatternXMLVersion>\n            <ExpandRecurrence>TRUE</ExpandRecurrence>";
    };
    CamlBuilder.prototype.addViewField = function (fieldRef) {
        this.viewFieldsString = this.viewFieldsString + ("<FieldRef Name=\"" + fieldRef.replace(" ", "_x0020_") + "\"/>");
    };
    CamlBuilder.prototype.addViewFields = function (fieldRefs) {
        for (var i = 0; i < fieldRefs.length; i++) {
            this.addViewField(fieldRefs[i]);
        }
    };
    CamlBuilder.prototype.addGroupByField = function (fieldRef) {
        this.groupByFieldsString = this.groupByFieldsString + ("<FieldRef Name=\"" + fieldRef.replace(" ", "_x0020_") + "\"/>");
    };
    CamlBuilder.prototype.addGroupByFields = function (fieldRefs) {
        for (var i = 0; i < fieldRefs.length; i++) {
            this.addGroupByField(fieldRefs[i]);
        }
    };
    CamlBuilder.prototype.addOrderBy = function (fieldRef, ascending) {
        var orderBy = new OrderBy();
        orderBy.fieldRef = fieldRef;
        orderBy.ascending = ascending;
        this.orderByFields.push(orderBy);
    };
    CamlBuilder.prototype.addDaysToDate = function (date, days) {
        date.setDate(date.getDate() + days);
        return date;
    };
    return CamlBuilder;
}());

var OrderBy = /** @class */ (function () {
    function OrderBy() {
    }
    return OrderBy;
}());



/***/ }),

/***/ "./src/helper/listHelper.ts":
/*!**********************************!*\
  !*** ./src/helper/listHelper.ts ***!
  \**********************************/
/*! exports provided: ListHelper */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "ListHelper", function() { return ListHelper; });
/* harmony import */ var _common__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../common */ "./src/common.ts");

var ListHelper = /** @class */ (function () {
    function ListHelper(context) {
        this.context = context;
    }
    ListHelper.prototype.createList = function (web, listName, templateId) {
        var deferred = $.Deferred();
        var list = null;
        var info = new SP.ListCreationInformation();
        info.set_title(listName);
        info.set_templateType(templateId);
        var list = web.get_lists().add(info);
        this.context.load(list);
        _common__WEBPACK_IMPORTED_MODULE_0__["common"].executeQuery(this.context)
            .fail(function (sender, args) { deferred.reject(sender, args); })
            .done(function () {
            deferred.resolve(list);
        });
        return deferred.promise();
    };
    ListHelper.prototype.setAlerts = function (web, listName, enabled) {
        var deferred = $.Deferred();
        var list = web.get_lists().getByTitle(listName);
        this.context.load(list);
        list.set_enableAssignToEmail(enabled); //Not in typescript definitions currently.  //TODO: test 
        list.update();
        _common__WEBPACK_IMPORTED_MODULE_0__["common"].executeQuery(this.context)
            .fail(function (sender, args) { deferred.reject(sender, args); })
            .done(function () {
            deferred.resolve();
        });
        return deferred.promise();
    };
    ListHelper.prototype.setListItemProperties = function (listItem, properties) {
        var deferred = $.Deferred();
        for (var propertyName in properties) {
            if (typeof (properties[propertyName]) !== typeof (undefined)) {
                listItem.set_item(propertyName, properties[propertyName]);
            }
        }
        listItem.update();
        this.context.load(listItem);
        _common__WEBPACK_IMPORTED_MODULE_0__["common"].executeQuery(this.context)
            .fail(function (sender, args) { deferred.reject(sender, args); })
            .done(function () {
            deferred.resolve(listItem);
        });
        return deferred.promise();
    };
    ListHelper.prototype.exists = function (web, listName) {
        var deferred = $.Deferred();
        var lists = web.get_lists();
        this.context.load(lists);
        _common__WEBPACK_IMPORTED_MODULE_0__["common"].executeQuery(this.context)
            .fail(function (sender, args) { deferred.reject(sender, args); })
            .done(function () {
            for (var i = 0; i < lists.get_count(); i++) {
                if (lists.getItemAtIndex(i).get_title().toLowerCase() === listName.toLowerCase()) {
                    deferred.resolve(true);
                    return;
                }
            }
            deferred.resolve(false);
        });
        return deferred.promise();
    };
    ListHelper.prototype.createListItemWithContentTypeName = function (web, listName, contentTypeName, properties) {
        var deferred = $.Deferred();
        var clientContext = this.context;
        var list = web.get_lists().getByTitle(listName);
        var itemCreateInfo = new SP.ListItemCreationInformation();
        var listContentTypes = list.get_contentTypes();
        clientContext.load(listContentTypes);
        _common__WEBPACK_IMPORTED_MODULE_0__["common"].executeQuery(clientContext)
            .fail(function (sender, args) { deferred.reject(sender, args); })
            .done(function () {
            var contentTypeId = null;
            for (var i = 0; i < listContentTypes.get_count(); i++) {
                var contentType = listContentTypes.getItemAtIndex(i);
                if (contentType.get_name() === contentTypeName) {
                    contentTypeId = contentType.get_id().get_stringValue();
                    break;
                }
            }
            var listItem = list.addItem(itemCreateInfo);
            if (contentTypeId) {
                listItem.set_item('ContentTypeId', contentTypeId);
            }
            for (var propertyName in properties) {
                if (typeof (properties[propertyName]) !== typeof (undefined)) {
                    listItem.set_item(propertyName, properties[propertyName]);
                }
            }
            listItem.update();
            clientContext.load(listItem);
            _common__WEBPACK_IMPORTED_MODULE_0__["common"].executeQuery(clientContext)
                .fail(function (sender, args) { deferred.reject(sender, args); })
                .done(function () {
                deferred.resolve(listItem);
            });
        });
        return deferred.promise();
    };
    ListHelper.prototype.createListItem = function (web, listName, properties) {
        var deferred = $.Deferred();
        var clientContext = this.context;
        var list = web.get_lists().getByTitle(listName);
        var itemCreateInfo = new SP.ListItemCreationInformation();
        var contentTypeId = null;
        var listItem = list.addItem(itemCreateInfo);
        //TODO: validate this works with People, taxonomy and lookup fields.
        for (var propertyName in properties) {
            if (typeof (properties[propertyName]) !== typeof (undefined)) {
                listItem.set_item(propertyName, properties[propertyName]);
            }
        }
        listItem.update();
        clientContext.load(listItem);
        _common__WEBPACK_IMPORTED_MODULE_0__["common"].executeQuery(clientContext)
            .fail(function (sender, args) { deferred.reject(sender, args); })
            .done(function () {
            deferred.resolve(listItem);
        });
        return deferred.promise();
    };
    ListHelper.prototype.loadListItem = function (listItem, viewFields) {
        if (viewFields === void 0) { viewFields = null; }
        var deferred = $.Deferred();
        if (viewFields && viewFields.length) {
            this.context.load(listItem, viewFields);
        }
        else {
            this.context.load(listItem);
        }
        _common__WEBPACK_IMPORTED_MODULE_0__["common"].executeQuery(this.context)
            .fail(function (sender, args) { deferred.reject(sender, args); })
            .done(function () {
            deferred.resolve(listItem);
        });
        return deferred.promise();
    };
    ListHelper.prototype.getFile = function (serverRelativeUrl) {
        var deferred = $.Deferred();
        var file = this.context.get_site().get_rootWeb().getFileByServerRelativeUrl(serverRelativeUrl);
        this.context.load(file);
        _common__WEBPACK_IMPORTED_MODULE_0__["common"].executeQuery(this.context)
            .fail(function (sender, args) { deferred.reject(sender, args); })
            .done(function () {
            deferred.resolve(file);
        });
        return deferred.promise();
    };
    ListHelper.prototype.checkInFile = function (web, serverRelativeUrl, comment, checkInType) {
        var deferred = $.Deferred();
        var file = web.getFileByServerRelativeUrl(serverRelativeUrl);
        this.context.load(file);
        file.checkIn(comment, checkInType);
        _common__WEBPACK_IMPORTED_MODULE_0__["common"].executeQuery(this.context)
            .fail(function (sender, args) { deferred.reject(sender, args); })
            .done(function () {
            deferred.resolve(file);
        });
        return deferred.promise();
    };
    ListHelper.prototype.getList = function (web, listName) {
        var deferred = $.Deferred();
        var list = web.get_lists().getByTitle(listName);
        this.context.load(list);
        _common__WEBPACK_IMPORTED_MODULE_0__["common"].executeQuery(this.context)
            .fail(function (sender, args) { deferred.reject(sender, args); })
            .done(function () {
            deferred.resolve(list);
        });
        return deferred.promise();
    };
    ListHelper.prototype.deleteList = function (web, listName) {
        var deferred = $.Deferred();
        var list = web.get_lists().getByTitle(listName);
        this.context.load(list);
        list.deleteObject();
        _common__WEBPACK_IMPORTED_MODULE_0__["common"].executeQuery(this.context)
            .fail(function (sender, args) { deferred.reject(sender, args); })
            .done(function () {
            deferred.resolve();
        });
        return deferred.promise();
    };
    ListHelper.prototype.getFileListItem = function (serverRelativeUrl, viewFields) {
        if (viewFields === void 0) { viewFields = null; }
        var file = this.context.get_site().get_rootWeb().getFileByServerRelativeUrl(serverRelativeUrl);
        var listItem = file.get_listItemAllFields();
        return this.loadListItem(listItem);
    };
    ListHelper.prototype.getListItemById = function (web, listName, id, viewFields) {
        if (viewFields === void 0) { viewFields = null; }
        var listItem = web.get_lists().getByTitle(listName).getItemById(id);
        return this.loadListItem(listItem, viewFields);
    };
    ListHelper.prototype.deleteListItemById = function (web, listName, id) {
        var deferred = $.Deferred();
        var listItem = web.get_lists().getByTitle(listName).getItemById(id);
        this.context.load(listItem);
        listItem.deleteObject();
        _common__WEBPACK_IMPORTED_MODULE_0__["common"].executeQuery(this.context)
            .fail(function (sender, args) { deferred.reject(sender, args); })
            .done(function () {
            deferred.resolve();
        });
        return deferred.promise();
    };
    ListHelper.prototype.getListItems = function (web, listName, viewXml) {
        var deferred = $.Deferred();
        var list = web.get_lists().getByTitle(listName);
        var query = new SP.CamlQuery();
        if (!viewXml) {
            viewXml = "<View><Query></Query></Where>";
        }
        query.set_viewXml(viewXml);
        var listItems = list.getItems(query);
        this.context.load(listItems);
        _common__WEBPACK_IMPORTED_MODULE_0__["common"].executeQuery(this.context)
            .fail(function (sender, args) { deferred.reject(sender, args); })
            .done(function () {
            deferred.resolve(listItems);
        });
        return deferred.promise();
    };
    ListHelper.prototype.addContentTypeListAssociation = function (web, listName, contentTypeName) {
        var _this = this;
        var deferred = $.Deferred();
        var list = web.get_lists().getByTitle(listName);
        var listCts = list.get_contentTypes();
        this.context.load(list);
        this.context.load(listCts);
        var contentTypes = this.context.get_site().get_rootWeb().get_contentTypes();
        var findContentType = function (collection, name) {
            for (var i = 0; i < collection.get_count(); i++) {
                if (collection.getItemAtIndex(i).get_name().toLowerCase() === name.toLowerCase()) {
                    return collection.getItemAtIndex(i);
                }
            }
            return null;
        };
        this.context.load(contentTypes);
        _common__WEBPACK_IMPORTED_MODULE_0__["common"].executeQuery(this.context)
            .fail(function (sender, args) { deferred.reject(sender, args); })
            .done(function () {
            var contentType = findContentType(contentTypes, contentTypeName);
            if (!contentType) {
                _common__WEBPACK_IMPORTED_MODULE_0__["common"].reject(deferred, "Content Type " + contentTypeName + " not found");
                return;
            }
            //check if the CT is already associated
            var listCt = findContentType(listCts, contentTypeName);
            if (listCt) {
                //already associated
                deferred.resolve(listCt);
            }
            else {
                if (!list.get_contentTypesEnabled()) {
                    list.set_contentTypesEnabled(true); //enable custom cts on the list.
                }
                var associatedCt = list.get_contentTypes().addExistingContentType(contentType);
                list.update();
                _common__WEBPACK_IMPORTED_MODULE_0__["common"].executeQuery(_this.context)
                    .fail(function (sender, args) { deferred.reject(sender, args); })
                    .done(function () {
                    deferred.resolve(associatedCt);
                });
            }
        });
        return deferred.promise();
    };
    ListHelper.prototype.removeContentTypeListAssociation = function (web, listName, contentTypeName) {
        var _this = this;
        var deferred = $.Deferred();
        var list = web.get_lists().getByTitle(listName);
        var listCts = list.get_contentTypes();
        this.context.load(list);
        this.context.load(listCts);
        var findContentType = function (collection, name) {
            for (var i = 0; i < collection.get_count(); i++) {
                if (collection.getItemAtIndex(i).get_name().toLowerCase() === name.toLowerCase()) {
                    return collection.getItemAtIndex(i);
                }
            }
            return null;
        };
        _common__WEBPACK_IMPORTED_MODULE_0__["common"].executeQuery(this.context)
            .fail(function (sender, args) { deferred.reject(sender, args); })
            .done(function () {
            var listCt = findContentType(listCts, contentTypeName);
            if (listCt) {
                listCt.deleteObject();
                list.update();
                _common__WEBPACK_IMPORTED_MODULE_0__["common"].executeQuery(_this.context)
                    .fail(function (sender, args) { deferred.reject(sender, args); })
                    .done(function () {
                    deferred.resolve();
                });
            }
            else {
                deferred.resolve(); //not found, nothing to do.
            }
        });
        return deferred.promise();
    };
    ListHelper.prototype.setDefaultValueOnList = function (web, listName, fieldInternalName, defaultValue) {
        var _this = this;
        var deferred = $.Deferred();
        var list = web.get_lists().getByTitle(listName);
        var fields = list.get_fields();
        this.context.load(list);
        this.context.load(fields);
        var findField = function (collection, name) {
            for (var i = 0; i < collection.get_count(); i++) {
                if (collection.getItemAtIndex(i).get_internalName().toLowerCase() === name.toLowerCase()) {
                    return collection.getItemAtIndex(i);
                }
            }
            return null;
        };
        _common__WEBPACK_IMPORTED_MODULE_0__["common"].executeQuery(this.context)
            .fail(function (sender, args) { deferred.reject(sender, args); })
            .done(function () {
            var field = findField(fields, fieldInternalName);
            if (field) {
                field.set_defaultValue(defaultValue);
                field.update();
                _common__WEBPACK_IMPORTED_MODULE_0__["common"].executeQuery(_this.context)
                    .fail(function (sender, args) { deferred.reject(sender, args); })
                    .done(function () {
                    deferred.resolve();
                });
            }
            else {
                _common__WEBPACK_IMPORTED_MODULE_0__["common"].reject(deferred, "Field " + fieldInternalName + " not found");
            }
        });
        return deferred.promise();
    };
    return ListHelper;
}());



/***/ }),

/***/ "./src/helper/navigationHelper.ts":
/*!****************************************!*\
  !*** ./src/helper/navigationHelper.ts ***!
  \****************************************/
/*! exports provided: NavigationHelper */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "NavigationHelper", function() { return NavigationHelper; });
/* harmony import */ var _common__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../common */ "./src/common.ts");
/* harmony import */ var _fluent__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../fluent */ "./src/fluent.ts");


var NavigationHelper = /** @class */ (function () {
    function NavigationHelper(context) {
        this.context = context;
    }
    NavigationHelper.prototype.deleteQuicklaunchNodes = function (web) {
        return this.deleteNodes(web.get_navigation().get_quickLaunch());
    };
    NavigationHelper.prototype.deleteTopNavigationNodes = function (web) {
        return this.deleteNodes(web.get_navigation().get_topNavigationBar());
    };
    NavigationHelper.prototype.deleteQuicklaunchNode = function (web, title) {
        return this.deleteNode(web.get_navigation().get_quickLaunch(), title);
    };
    NavigationHelper.prototype.deleteTopQuicklaunchNode = function (web, title) {
        return this.deleteNode(web.get_navigation().get_topNavigationBar(), title);
    };
    NavigationHelper.prototype.setCurrentNavigation = function (web, navigationType, showSubsites, showPages) {
        var _this = this;
        if (showSubsites === void 0) { showSubsites = true; }
        if (showPages === void 0) { showPages = true; }
        var deferred = $.Deferred();
        var siblingsPropertyName = "__NavigationShowSiblings";
        var includeTypesPropertyName = "__CurrentNavigationIncludeTypes";
        var allProperties = web.get_allProperties();
        this.context.load(web);
        this.context.load(allProperties);
        var setOptions = function () {
            if (showPages && showSubsites) {
                allProperties.set_item(includeTypesPropertyName, "3");
            }
            else if (showPages && !showSubsites) {
                allProperties.set_item(includeTypesPropertyName, "2");
            }
            else if (!showPages && showSubsites) {
                allProperties.set_item(includeTypesPropertyName, "1");
            }
            else if (!showPages && !showSubsites) {
                allProperties.set_item(includeTypesPropertyName, "0");
            }
        };
        _common__WEBPACK_IMPORTED_MODULE_0__["common"].executeQuery(this.context)
            .fail(function (sender, args) { deferred.reject(sender, args); })
            .done(function () {
            var nav = new SP.Publishing.Navigation.WebNavigationSettings(_this.context, web);
            switch (navigationType) {
                case _fluent__WEBPACK_IMPORTED_MODULE_1__["NavigationType"].Inherit:
                    nav.get_currentNavigation().set_source(SP.Publishing.Navigation.StandardNavigationSource.inheritFromParentWeb);
                    break;
                case _fluent__WEBPACK_IMPORTED_MODULE_1__["NavigationType"].Managed:
                    _common__WEBPACK_IMPORTED_MODULE_0__["common"].reject(deferred, "Not implemented");
                    break;
                case _fluent__WEBPACK_IMPORTED_MODULE_1__["NavigationType"].StructuralWithSiblings:
                    nav.get_currentNavigation().set_source(SP.Publishing.Navigation.StandardNavigationSource.portalProvider);
                    allProperties.set_item(siblingsPropertyName, "True");
                    setOptions();
                    web.update();
                    break;
                case _fluent__WEBPACK_IMPORTED_MODULE_1__["NavigationType"].StructuralChildrenOnly:
                    nav.get_currentNavigation().set_source(SP.Publishing.Navigation.StandardNavigationSource.portalProvider);
                    allProperties.set_item(siblingsPropertyName, "False");
                    setOptions();
                    web.update();
                    break;
                default:
                    _common__WEBPACK_IMPORTED_MODULE_0__["common"].reject(deferred, "Unknown Navigation Type");
            }
            nav.update(null);
            _common__WEBPACK_IMPORTED_MODULE_0__["common"].executeQuery(_this.context)
                .fail(function (sender, args) { deferred.reject(sender, args); })
                .done(function () {
                deferred.resolve();
            });
        });
        return deferred.promise();
    };
    NavigationHelper.prototype.deleteNodes = function (nav) {
        var _this = this;
        var deferred = $.Deferred();
        this.context.load(nav);
        _common__WEBPACK_IMPORTED_MODULE_0__["common"].executeQuery(this.context)
            .fail(function (sender, args) { deferred.reject(sender, args); })
            .done(function () {
            var enumerator = nav.getEnumerator();
            var itemsToDelete = [];
            while (enumerator.moveNext()) {
                itemsToDelete.push(enumerator.get_current());
            }
            for (var i = 0; i < itemsToDelete.length; i++) {
                itemsToDelete[i].deleteObject();
            }
            _common__WEBPACK_IMPORTED_MODULE_0__["common"].executeQuery(_this.context)
                .fail(function (sender, args) { deferred.reject(sender, args); })
                .done(function () {
                deferred.resolve();
            });
        });
        return deferred.promise();
    };
    NavigationHelper.prototype.deleteNode = function (nav, title) {
        var _this = this;
        var deferred = $.Deferred();
        this.context.load(nav);
        _common__WEBPACK_IMPORTED_MODULE_0__["common"].executeQuery(this.context)
            .fail(function (sender, args) { deferred.reject(sender, args); })
            .done(function () {
            var enumerator = nav.getEnumerator();
            var itemsToDelete = [];
            while (enumerator.moveNext()) {
                var node = enumerator.get_current();
                if (node.get_title() === title) {
                    itemsToDelete.push(node);
                }
            }
            for (var i = 0; i < itemsToDelete.length; i++) {
                itemsToDelete[i].deleteObject();
            }
            _common__WEBPACK_IMPORTED_MODULE_0__["common"].executeQuery(_this.context)
                .fail(function (sender, args) { deferred.reject(sender, args); })
                .done(function () {
                deferred.resolve();
            });
        });
        return deferred.promise();
    };
    NavigationHelper.prototype.createQuicklaunchNode = function (web, title, url, asLastNode) {
        if (asLastNode === void 0) { asLastNode = true; }
        return this.createNode(web.get_navigation().get_quickLaunch(), title, url, asLastNode);
    };
    NavigationHelper.prototype.createTopNavigationNode = function (web, title, url, asLastNode) {
        if (asLastNode === void 0) { asLastNode = true; }
        return this.createNode(web.get_navigation().get_topNavigationBar(), title, url, asLastNode);
    };
    NavigationHelper.prototype.createNode = function (nav, title, url, asLastNode) {
        if (asLastNode === void 0) { asLastNode = true; }
        var deferred = $.Deferred();
        this.context.load(nav);
        var info = new SP.NavigationNodeCreationInformation();
        info.set_title(title);
        info.set_url(url);
        info.set_asLastNode(asLastNode);
        var node = nav.add(info);
        this.context.load(nav);
        _common__WEBPACK_IMPORTED_MODULE_0__["common"].executeQuery(this.context)
            .fail(function (sender, args) { deferred.reject(sender, args); })
            .done(function () {
            deferred.resolve(node);
        });
        return deferred.promise();
    };
    return NavigationHelper;
}());



/***/ }),

/***/ "./src/helper/pageHelper.ts":
/*!**********************************!*\
  !*** ./src/helper/pageHelper.ts ***!
  \**********************************/
/*! exports provided: PageHelper */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "PageHelper", function() { return PageHelper; });
/* harmony import */ var _common__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../common */ "./src/common.ts");
/* harmony import */ var _Client_CamlBuilder__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./Client.CamlBuilder */ "./src/helper/Client.CamlBuilder.ts");


var PageHelper = /** @class */ (function () {
    function PageHelper(context) {
        this.context = context;
    }
    PageHelper.prototype.createPublishingPage = function (web, name, layoutUrl) {
        var deferred = $.Deferred();
        var file = this.context.get_site().get_rootWeb().getFileByServerRelativeUrl(layoutUrl);
        var listItem = file.get_listItemAllFields();
        this.context.load(listItem);
        var pageInfo = new SP.Publishing.PublishingPageInformation();
        pageInfo.set_name(name);
        pageInfo.set_pageLayoutListItem(listItem);
        var publishingWeb = SP.Publishing.PublishingWeb.getPublishingWeb(this.context, web);
        this.context.load(publishingWeb);
        var publishingPage = publishingWeb.addPublishingPage(pageInfo);
        this.context.load(publishingPage);
        _common__WEBPACK_IMPORTED_MODULE_0__["common"].executeQuery(this.context)
            .fail(function (sender, args) { deferred.reject(sender, args); })
            .done(function () {
            deferred.resolve(publishingPage);
        });
        return deferred.promise();
    };
    PageHelper.prototype.getPublishingLayout = function (serverRelativeUrl) {
        var deferred = $.Deferred();
        var file = this.context.get_site().get_rootWeb().getFileByServerRelativeUrl(serverRelativeUrl);
        var listItem = file.get_listItemAllFields();
        this.context.load(listItem);
        _common__WEBPACK_IMPORTED_MODULE_0__["common"].executeQuery(this.context)
            .fail(function (sender, args) { deferred.reject(sender, args); })
            .done(function () {
            deferred.resolve(listItem);
        });
        return deferred.promise();
    };
    PageHelper.prototype.setLayout = function (web, name, layoutUrl) {
        var _this = this;
        var deferred = $.Deferred();
        var file = this.context.get_site().get_rootWeb().getFileByServerRelativeUrl(layoutUrl);
        var listItem = file.get_listItemAllFields();
        this.context.load(listItem);
        var camlBuilder = new _Client_CamlBuilder__WEBPACK_IMPORTED_MODULE_1__["CamlBuilder"]();
        camlBuilder.addTextClause(_Client_CamlBuilder__WEBPACK_IMPORTED_MODULE_1__["CamlOperator"].Eq, "FileLeafRef", name);
        var list = web.get_lists().getByTitle("Pages");
        var query = new SP.CamlQuery();
        query.set_viewXml(camlBuilder.viewXml);
        var pages = list.getItems(query);
        this.context.load(pages);
        _common__WEBPACK_IMPORTED_MODULE_0__["common"].executeQuery(this.context)
            .fail(function (sender, args) { deferred.reject(sender, args); })
            .done(function () {
            if (pages.get_count() === 0) {
                deferred.reject(_this, { get_message: function () { return "Page not found"; } });
            }
            var page = pages.getItemAtIndex(0);
            page.set_item("PublishingPageLayout", layoutUrl);
            page.update();
            _common__WEBPACK_IMPORTED_MODULE_0__["common"].executeQuery(_this.context)
                .fail(function (sender, args) { deferred.reject(sender, args); })
                .done(function () {
                deferred.resolve();
            });
        });
        return deferred.promise();
    };
    return PageHelper;
}());



/***/ }),

/***/ "./src/helper/permissionHelper.ts":
/*!****************************************!*\
  !*** ./src/helper/permissionHelper.ts ***!
  \****************************************/
/*! exports provided: PermissionHelper */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "PermissionHelper", function() { return PermissionHelper; });
/* harmony import */ var _common__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../common */ "./src/common.ts");

var PermissionHelper = /** @class */ (function () {
    function PermissionHelper(context) {
        this.context = context;
    }
    PermissionHelper.prototype.hasWebPermission = function (permission, web) {
        var deferred = $.Deferred();
        this.context.load(web, 'EffectiveBasePermissions');
        _common__WEBPACK_IMPORTED_MODULE_0__["common"].executeQuery(this.context)
            .fail(function (sender, args) { deferred.reject(sender, args); })
            .done(function () {
            deferred.resolve(web.get_effectiveBasePermissions().has(permission));
        });
        return deferred.promise();
    };
    PermissionHelper.prototype.hasListPermission = function (permission, list) {
        var deferred = $.Deferred();
        this.context.load(list, 'EffectiveBasePermissions');
        _common__WEBPACK_IMPORTED_MODULE_0__["common"].executeQuery(this.context)
            .fail(function (sender, args) { deferred.reject(sender, args); })
            .done(function () {
            deferred.resolve(list.get_effectiveBasePermissions().has(permission));
        });
        return deferred.promise();
    };
    PermissionHelper.prototype.hasItemPermission = function (permission, item) {
        var deferred = $.Deferred();
        this.context.load(item, 'EffectiveBasePermissions');
        _common__WEBPACK_IMPORTED_MODULE_0__["common"].executeQuery(this.context)
            .fail(function (sender, args) { deferred.reject(sender, args); })
            .done(function () {
            deferred.resolve(item.get_effectiveBasePermissions().has(permission));
        });
        return deferred.promise();
    };
    return PermissionHelper;
}());



/***/ }),

/***/ "./src/helper/userHelper.ts":
/*!**********************************!*\
  !*** ./src/helper/userHelper.ts ***!
  \**********************************/
/*! exports provided: UserHelper */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "UserHelper", function() { return UserHelper; });
/* harmony import */ var _common__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../common */ "./src/common.ts");

var UserHelper = /** @class */ (function () {
    function UserHelper(context) {
        this.context = context;
    }
    UserHelper.prototype.getUserByEmail = function (email) {
        return this.loadUser(this.context.get_web().ensureUser(email));
    };
    UserHelper.prototype.getUserById = function (id) {
        return this.loadUser(this.context.get_web().get_siteUsers().getById(id));
    };
    UserHelper.prototype.loadUser = function (user) {
        var deferred = $.Deferred();
        this.context.load(user);
        this.context.executeQueryAsync(function (sender, args) {
            deferred.resolve(user);
        }, function (sender, args) {
            console.error(args.get_message());
            deferred.reject(sender, args);
        });
        return deferred.promise();
    };
    UserHelper.prototype.getCurrentUser = function () {
        var deferred = $.Deferred();
        var user = this.context.get_web().get_currentUser();
        this.context.load(user);
        this.context.executeQueryAsync(function (sender, args) {
            deferred.resolve(user);
        }, function (sender, args) {
            console.error(args.get_message());
            deferred.reject(sender, args);
        });
        return deferred.promise();
    };
    UserHelper.prototype.getCurrentUserProfileProperties = function () {
        var _this = this;
        var deferred = $.Deferred();
        SP.SOD.executeFunc('userprofile', 'SP.UserProfiles.PeopleManager', function () {
            var clientContext = _this.context;
            var currentUser = clientContext.get_web().get_currentUser();
            var peopleManager = new SP.UserProfiles.PeopleManager(clientContext);
            var profile = peopleManager.getMyProperties();
            clientContext.load(currentUser);
            clientContext.load(profile);
            clientContext.executeQueryAsync(function (sender, args) {
                deferred.resolve(profile.get_userProfileProperties());
            }, function (sender, args) {
                console.error(args.get_message());
                deferred.reject(sender, args);
            });
        });
        return deferred.promise();
    };
    UserHelper.prototype.getCurrentUserManager = function () {
        var _this = this;
        var deferred = $.Deferred();
        var peopleManager = new SP.UserProfiles.PeopleManager(this.context);
        var profilePropertyNames = ["Manager"];
        var user_email = this.context.get_web().get_currentUser().get_email();
        var userProfilePropertiesForUser = new SP.UserProfiles.UserProfilePropertiesForUser(this.context, "i:0#.f|membership|" + user_email, profilePropertyNames); //TODO: check for better mechanism to constructure login
        var userProfileProps = peopleManager.getUserProfilePropertiesFor(userProfilePropertiesForUser);
        this.context.load(userProfilePropertiesForUser);
        _common__WEBPACK_IMPORTED_MODULE_0__["common"].executeQuery(this.context)
            .fail(function (sender, args) { deferred.reject(sender, args); })
            .done(function () {
            if (userProfileProps[0]) {
                var user = _this.context.get_web().ensureUser(userProfileProps[0]);
                _this.context.load(user);
                _common__WEBPACK_IMPORTED_MODULE_0__["common"].executeQuery(_this.context)
                    .fail(function (sender, args) { deferred.reject(sender, args); })
                    .done(function () {
                    deferred.resolve(user);
                });
            }
            else {
                deferred.resolve(null);
            }
        });
        return deferred.promise();
    };
    return UserHelper;
}());



/***/ }),

/***/ "./src/helper/webHelper.ts":
/*!*********************************!*\
  !*** ./src/helper/webHelper.ts ***!
  \*********************************/
/*! exports provided: WebHelper */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "WebHelper", function() { return WebHelper; });
/* harmony import */ var _common__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../common */ "./src/common.ts");

var WebHelper = /** @class */ (function () {
    function WebHelper(context) {
        this.context = context;
    }
    WebHelper.prototype.setWelcomePage = function (web, url) {
        var deferred = $.Deferred();
        var rootFolder = web.get_rootFolder();
        this.context.load(web);
        this.context.load(rootFolder);
        rootFolder.set_welcomePage(url);
        rootFolder.update();
        _common__WEBPACK_IMPORTED_MODULE_0__["common"].executeQuery(this.context)
            .fail(function (sender, args) { deferred.reject(sender, args); })
            .done(function () {
            deferred.resolve(rootFolder);
        });
        return deferred.promise();
    };
    WebHelper.prototype.createWeb = function (name, parentWeb, title, template, useSamePermissionsAsParent) {
        if (useSamePermissionsAsParent === void 0) { useSamePermissionsAsParent = true; }
        var deferred = $.Deferred();
        var info = new SP.WebCreationInformation();
        info.set_url(name);
        info.set_title(title);
        info.set_webTemplate(template);
        info.set_useSamePermissionsAsParentSite(useSamePermissionsAsParent);
        var newWeb = parentWeb.get_webs().add(info);
        this.context.load(newWeb);
        _common__WEBPACK_IMPORTED_MODULE_0__["common"].executeQuery(this.context)
            .fail(function (sender, args) { deferred.reject(sender, args); })
            .done(function () {
            deferred.resolve(newWeb);
        });
        return deferred.promise();
    };
    WebHelper.prototype.doesWebExist = function (url) {
        var deferred = $.Deferred();
        var output = [];
        this.getAllWebs(this.context, this.context.get_site().get_rootWeb(), output)
            .fail(function (sender, args) { deferred.reject(sender, args); })
            .done(function () {
            for (var i = 0; i < output.length; i++) {
                if (url.toLowerCase() === output[i].get_serverRelativeUrl() || url.toLowerCase() === output[i].get_url()) {
                    deferred.resolve(true);
                    return;
                }
            }
            deferred.resolve(false);
        });
        return deferred.promise();
    };
    WebHelper.prototype.getWebs = function (fromWeb) {
        var deferred = $.Deferred();
        var output = [];
        this.getAllWebs(this.context, fromWeb, output)
            .fail(function (sender, args) { deferred.reject(sender, args); })
            .done(function () {
            deferred.resolve(output);
        });
        return deferred.promise();
    };
    WebHelper.prototype.getAllWebs = function (context, web, output) {
        var _this = this;
        var deferred = $.Deferred();
        var webs = web.get_webs();
        context.load(webs);
        _common__WEBPACK_IMPORTED_MODULE_0__["common"].executeQuery(context)
            .fail(function (sender, args) {
            deferred.reject(sender, args);
        })
            .done(function () {
            var promises = [];
            for (var i = 0; i < webs.get_count(); i++) {
                var subWeb = webs.getItemAtIndex(i);
                output.push(subWeb);
                promises.push(_this.getAllWebs(context, subWeb, output));
            }
            if (promises.length) {
                $.when.apply($, promises)
                    .fail(function (sender, args) {
                    deferred.reject(sender, args);
                })
                    .done(function () {
                    deferred.resolve();
                });
            }
            else {
                deferred.resolve();
            }
        });
        return deferred.promise();
    };
    return WebHelper;
}());



/***/ }),

/***/ "./src/list.ts":
/*!*********************!*\
  !*** ./src/list.ts ***!
  \*********************/
/*! exports provided: List */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "List", function() { return List; });
/* harmony import */ var _helper_listHelper__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./helper/listHelper */ "./src/helper/listHelper.ts");

var List = /** @class */ (function () {
    function List(fluent) {
        this.fluent = null;
        this.listHelper = null;
        this._helperName = "list";
        this.fluent = fluent;
        this.listHelper = new _helper_listHelper__WEBPACK_IMPORTED_MODULE_0__["ListHelper"](fluent.context);
    }
    /**
    * Create new list
    * Example: create(context.get_web(), "My Task List", 107)
    * templateId - See: https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-csom/ee541191(v%3Doffice.15)#remarks
    */
    List.prototype.create = function (web, listName, templateId) {
        var _this = this;
        return this.fluent.chainAction(this._helperName + ".create", function () {
            return _this.listHelper.createList(web, listName, templateId);
        });
    };
    /**
    * Check if the list exists
    * Result: boolean
    * Example: exists(context.get_web(), "My List")
    */
    List.prototype.exists = function (web, listName) {
        var _this = this;
        return this.fluent.chainAction(this._helperName + ".exists", function () {
            return _this.listHelper.exists(web, listName);
        });
    };
    /**
    * Delete a list
    * Result: void
    * Example: delete(context.get_web(), "My List")
    */
    List.prototype.delete = function (web, listName) {
        var _this = this;
        return this.fluent.chainAction(this._helperName + ".delete", function () {
            return _this.listHelper.deleteList(web, listName);
        });
    };
    /**
    * Get a list
    * Result: SP.List
    * Example: get(context.get_web(), "My List")
    */
    List.prototype.get = function (web, listName) {
        var _this = this;
        //TODO: test
        return this.fluent.chainAction(this._helperName + ".get", function () {
            return _this.listHelper.getList(web, listName);
        });
    };
    /**
    * Adds a Content Type to a List.  Resolves the new associated list content type.
    * Result: SP.ContentType
    * Example: addContentTypeListAssociation(context.get_web(), "My List", "My Content Type Name")
    */
    List.prototype.addContentTypeListAssociation = function (web, listName, contentTypeName) {
        var _this = this;
        return this.fluent.chainAction(this._helperName + ".addContentTypeListAssociation", function () {
            return _this.listHelper.addContentTypeListAssociation(web, listName, contentTypeName);
        });
    };
    /**
    * Removes a Content Type associated from a list.
    * Result: void
    * Example: removeContentTypeListAssociation(context.get_web(), "My List", "My Content Type Name")
    */
    List.prototype.removeContentTypeListAssociation = function (web, listName, contentTypeName) {
        var _this = this;
        return this.fluent.chainAction(this._helperName + ".removeContentTypeListAssociation", function () {
            return _this.listHelper.removeContentTypeListAssociation(web, listName, contentTypeName);
        });
    };
    /**
    * Sets a default value on a field in a list
    * Result: void
    * Example: setDefaultValueOnList(context.get_web(), "My Task List", "ClientId", 123)
    */
    List.prototype.setDefaultValueOnList = function (web, listName, fieldInternalName, defaultValue) {
        var _this = this;
        return this.fluent.chainAction(this._helperName + ".setDefaultValueOnList", function () {
            return _this.listHelper.setDefaultValueOnList(web, listName, fieldInternalName, defaultValue);
        });
    };
    /**
   * Enable email alerts on a list
   * Result: void
   * Example: setAlerts(context.get_web(), "My Task List", true)
   * Note: will not work for 2010/2013
   */
    List.prototype.setAlerts = function (web, listName, enabled) {
        var _this = this;
        return this.fluent.chainAction(this._helperName + ".setAlerts", function () {
            return _this.listHelper.setAlerts(web, listName, enabled);
        });
    };
    return List;
}());



/***/ }),

/***/ "./src/listitem.ts":
/*!*************************!*\
  !*** ./src/listitem.ts ***!
  \*************************/
/*! exports provided: ListItem */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "ListItem", function() { return ListItem; });
/* harmony import */ var _helper_listHelper__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./helper/listHelper */ "./src/helper/listHelper.ts");

var ListItem = /** @class */ (function () {
    function ListItem(fluent) {
        this.fluent = null;
        this.listHelper = null;
        this._helperName = "listItem";
        this.fluent = fluent;
        this.listHelper = new _helper_listHelper__WEBPACK_IMPORTED_MODULE_0__["ListHelper"](fluent.context);
    }
    /**
    * Update a list item with properties
    * Result: SP.ListItem
    * Example: update(listItem, { 'Title': 'Title value here', ClientId: 123 })
    */
    ListItem.prototype.update = function (listItem, properties) {
        var _this = this;
        return this.fluent.chainAction(this._helperName + ".update", function () {
            return _this.listHelper.setListItemProperties(listItem, properties);
        });
    };
    /**
    * Create a list item, specifying the Content Type
    * Result: SP.ListItem
    * Example: createWithContentType(context.get_web(), "My list", "My Content Type Name", { 'Title': 'Title value here', ClientId: 123 })
    */
    ListItem.prototype.createWithContentType = function (web, listName, contentTypeName, properties) {
        var _this = this;
        return this.fluent.chainAction(this._helperName + ".createWithContentType", function () {
            return _this.listHelper.createListItemWithContentTypeName(web, listName, contentTypeName, properties);
        });
    };
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
    ListItem.prototype.create = function (web, listName, properties) {
        var _this = this;
        return this.fluent.chainAction(this._helperName + ".create", function () {
            return _this.listHelper.createListItem(web, listName, properties);
        });
    };
    /**
    * Get the listitem using the ID with optional view fields
    * Result: SP.ListItem
    * Example: get(context.get_web(), "MyList", 1, ["ID", "Title"])
    */
    ListItem.prototype.get = function (web, listName, id, viewFields) {
        var _this = this;
        if (viewFields === void 0) { viewFields = null; }
        return this.fluent.chainAction(this._helperName + ".get", function () {
            return _this.listHelper.getListItemById(web, listName, id, viewFields);
        });
    };
    /**
    * Get the listitem for a File using the ID with optional view fields
    * Result: SP.ListItem
    * Example: getFileListItem("/sites/site/documents/mydoc.docx", ["ID", "Title", "FileLeafRef"])
    */
    ListItem.prototype.getFileListItem = function (serverRelativeUrl, viewFields) {
        var _this = this;
        if (viewFields === void 0) { viewFields = null; }
        return this.fluent.chainAction(this._helperName + ".getFileListItem", function () {
            return _this.listHelper.getFileListItem(serverRelativeUrl, viewFields);
        });
    };
    /**
    * Delete List Item
    * Result: void
    * Example: deleteById(context.get_web(), "MyList", 7)
    */
    ListItem.prototype.deleteById = function (web, listName, id) {
        var _this = this;
        return this.fluent.chainAction(this._helperName + ".deleteById", function () {
            return _this.listHelper.deleteListItemById(web, listName, id);
        });
    };
    /**
    * Execute a query using supplied CAML.
    * Returns: SP.ListItemCollection
    * Example: query(context.get_web(), "MyList", "<View><Query><Where><In><FieldRef Name='ID' /><Values><Value Type='Number'>1</Value><Value Type='Number'>2</Value></Values></In></View></Query></Where>")
    */
    ListItem.prototype.query = function (web, listName, viewXml) {
        var _this = this;
        return this.fluent.chainAction(this._helperName + ".query", function () {
            return _this.listHelper.getListItems(web, listName, viewXml);
        });
    };
    /**
    * Get all list items in a list
    * Returns: SP.ListItemCollection
    * Example: getAll(context.get_web(), "MyList")
    */
    ListItem.prototype.getAll = function (web, listName) {
        var _this = this;
        return this.fluent.chainAction(this._helperName + ".query", function () {
            return _this.listHelper.getListItems(web, listName, "<View></View>");
        });
    };
    return ListItem;
}());



/***/ }),

/***/ "./src/navigation.ts":
/*!***************************!*\
  !*** ./src/navigation.ts ***!
  \***************************/
/*! exports provided: Navigation */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Navigation", function() { return Navigation; });
/* harmony import */ var _helper_navigationHelper__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./helper/navigationHelper */ "./src/helper/navigationHelper.ts");
/* harmony import */ var _fluent__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./fluent */ "./src/fluent.ts");


var Navigation = /** @class */ (function () {
    function Navigation(fluent) {
        this.fluent = null;
        this.navigationHelper = null;
        this._helperName = "navigation";
        this.fluent = fluent;
        this.navigationHelper = new _helper_navigationHelper__WEBPACK_IMPORTED_MODULE_0__["NavigationHelper"](fluent.context);
    }
    /**
    * Create new navigation node
    * Result: SP.NavigationNode
    * Example: createNode(context.get_web(), NavigationLocation.Quicklaunch, "Test Node", "/sites/mysite/pages/default.aspx")
    */
    Navigation.prototype.createNode = function (web, location, title, url, asLastNode) {
        var _this = this;
        if (asLastNode === void 0) { asLastNode = true; }
        return this.fluent.chainAction(this._helperName + ".createNode", function () {
            if (location == _fluent__WEBPACK_IMPORTED_MODULE_1__["NavigationLocation"].Quicklaunch) {
                return _this.navigationHelper.createQuicklaunchNode(web, title, url, asLastNode);
            }
            else if (location == _fluent__WEBPACK_IMPORTED_MODULE_1__["NavigationLocation"].TopNavigation) {
                return _this.navigationHelper.createTopNavigationNode(web, title, url, asLastNode);
            }
            else {
                throw "Unknown location " + location;
            }
        });
    };
    /**
    * Delete all navigation nodes for the web
    * Result: void
    * Example: deleteAllNodes(context.get_web(), NavigationLocation.Quicklaunch)
    */
    Navigation.prototype.deleteAllNodes = function (web, location) {
        var _this = this;
        return this.fluent.chainAction(this._helperName + ".deleteAllNodes", function () {
            if (location == _fluent__WEBPACK_IMPORTED_MODULE_1__["NavigationLocation"].Quicklaunch) {
                return _this.navigationHelper.deleteQuicklaunchNodes(web);
            }
            else if (location == _fluent__WEBPACK_IMPORTED_MODULE_1__["NavigationLocation"].TopNavigation) {
                return _this.navigationHelper.deleteTopNavigationNodes(web);
            }
            else {
                throw "Unknown location " + location;
            }
        });
    };
    /**
    * Delete all navigation nodes that match the supplied title for the web
    * Result: void
    * Example: deleteNode(context.get_web(), NavigationLocation.Quicklaunch, "My link title")
    */
    Navigation.prototype.deleteNode = function (web, location, title) {
        var _this = this;
        return this.fluent.chainAction(this._helperName + ".deleteNode", function () {
            if (location == _fluent__WEBPACK_IMPORTED_MODULE_1__["NavigationLocation"].Quicklaunch) {
                return _this.navigationHelper.deleteQuicklaunchNode(web, title);
            }
            else if (location == _fluent__WEBPACK_IMPORTED_MODULE_1__["NavigationLocation"].TopNavigation) {
                return _this.navigationHelper.deleteTopQuicklaunchNode(web, title);
            }
            else {
                throw "Unknown location " + location;
            }
        });
    };
    /**
    * Set navigation for the web
    * Result: void
    * Example: setCurrentNavigation(context.get_web(), 3, true, true)
    * Note: showSubsites and showPages is only applicable for NavigationType.StructuralChildrenOnly (3)
    */
    Navigation.prototype.setCurrentNavigation = function (web, type, showSubsites, showPages) {
        var _this = this;
        if (showSubsites === void 0) { showSubsites = false; }
        if (showPages === void 0) { showPages = false; }
        this.fluent.registerDependency(_fluent__WEBPACK_IMPORTED_MODULE_1__["Dependency"].Publishing);
        return this.fluent.chainAction(this._helperName + ".setCurrentNavigation", function () {
            return _this.navigationHelper.setCurrentNavigation(web, type, showSubsites, showPages);
        });
    };
    return Navigation;
}());



/***/ }),

/***/ "./src/permission.ts":
/*!***************************!*\
  !*** ./src/permission.ts ***!
  \***************************/
/*! exports provided: Permission */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Permission", function() { return Permission; });
/* harmony import */ var _helper_permissionHelper__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./helper/permissionHelper */ "./src/helper/permissionHelper.ts");

var Permission = /** @class */ (function () {
    function Permission(fluent) {
        this._fluent = null;
        this._helperName = "permission";
        this.permissionHelper = null;
        this._fluent = fluent;
        this.permissionHelper = new _helper_permissionHelper__WEBPACK_IMPORTED_MODULE_0__["PermissionHelper"](fluent.context);
    }
    /**
    * Check if the current user has specified Web permission
    * Result: boolean
    * Example: hasWebPermission(SP.PermissionKind.createSSCSite, context.get_web())
    */
    Permission.prototype.hasWebPermission = function (permission, web) {
        var _this = this;
        return this._fluent.chainAction(this._helperName + ".hasWebPermission", function () {
            return _this.permissionHelper.hasWebPermission(permission, web);
        });
    };
    /**
    * Check if the current user has specified List permission
    * Result: boolean
    * Example: hasListPermission(SP.PermissionKind.addListItems, context.get_web().get_lists().getByTitle("MyList"))
    */
    Permission.prototype.hasListPermission = function (permission, list) {
        var _this = this;
        return this._fluent.chainAction(this._helperName + ".hasListPermission", function () {
            return _this.permissionHelper.hasListPermission(permission, list);
        });
    };
    /**
    * Check if the current user has specified ListItem permission
    * Result: boolean
    * Example: hasItemPermission(SP.PermissionKind.editListItems, context.get_web().get_lists().getByTitle("MyList").getItemById(0))
    */
    Permission.prototype.hasItemPermission = function (permission, item) {
        var _this = this;
        return this._fluent.chainAction(this._helperName + ".hasItemPermission", function () {
            return _this.permissionHelper.hasItemPermission(permission, item);
        });
    };
    return Permission;
}());



/***/ }),

/***/ "./src/publishingpage.ts":
/*!*******************************!*\
  !*** ./src/publishingpage.ts ***!
  \*******************************/
/*! exports provided: PublishingPage */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "PublishingPage", function() { return PublishingPage; });
/* harmony import */ var _fluent__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./fluent */ "./src/fluent.ts");
/* harmony import */ var _helper_pageHelper__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./helper/pageHelper */ "./src/helper/pageHelper.ts");


var PublishingPage = /** @class */ (function () {
    function PublishingPage(fluent) {
        this.fluent = null;
        this._helperName = "publishingPage";
        this.fluent = fluent;
        this.pageHelper = new _helper_pageHelper__WEBPACK_IMPORTED_MODULE_1__["PageHelper"](fluent.context);
    }
    /**
    * Creates a new publishing page.
    * Result: SP.Publishing.PublishingPage
    * Example: .publishingPage.create(SP.ClientContext.get_current().get_web(), "Home.aspx", _spPageContextInfo.siteServerRelativeUrl + "/_catalogs/masterpage/BlankWebPartPage.aspx")
    */
    PublishingPage.prototype.create = function (web, name, layoutUrl) {
        var _this = this;
        this.fluent.registerDependency(_fluent__WEBPACK_IMPORTED_MODULE_0__["Dependency"].Publishing);
        return this.fluent.chainAction(this._helperName + ".create", function () {
            return _this.pageHelper.createPublishingPage(web, name, layoutUrl);
        });
    };
    /**
    * Gets a page layout.
    * Result: SP.ListItem
    * Example: .publishingPage.getLayout(_spPageContextInfo.siteServerRelativeUrl + "/_catalogs/masterpage/BlankWebPartPage.aspx")
    *
    */
    PublishingPage.prototype.getLayout = function (serverRelativeUrl) {
        var _this = this;
        return this.fluent.chainAction(this._helperName + ".getLayout", function () {
            return _this.pageHelper.getPublishingLayout(serverRelativeUrl);
        });
    };
    /**
    * Changes layout for a publishing page.
    * Result: void
    * Example: .publishingPage.setLayout(SP.ClientContext.get_current().get_web(), "Home.aspx", _spPageContextInfo.siteServerRelativeUrl + "/_catalogs/masterpage/BlankWebPartPage.aspx")
    */
    PublishingPage.prototype.setLayout = function (web, name, layoutUrl) {
        var _this = this;
        this.fluent.registerDependency(_fluent__WEBPACK_IMPORTED_MODULE_0__["Dependency"].Publishing);
        return this.fluent.chainAction(this._helperName + ".setLayout", function () {
            return _this.pageHelper.setLayout(web, name, layoutUrl);
        });
    };
    return PublishingPage;
}());



/***/ }),

/***/ "./src/user.ts":
/*!*********************!*\
  !*** ./src/user.ts ***!
  \*********************/
/*! exports provided: User */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "User", function() { return User; });
/* harmony import */ var _helper_userHelper__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./helper/userHelper */ "./src/helper/userHelper.ts");
/* harmony import */ var _fluent__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./fluent */ "./src/fluent.ts");


var User = /** @class */ (function () {
    function User(fluent) {
        this._fluent = null;
        this._helperName = "userProfile";
        this.userHelper = null;
        this._fluent = fluent;
        this.userHelper = new _helper_userHelper__WEBPACK_IMPORTED_MODULE_0__["UserHelper"](fluent.context);
    }
    /**
    * Get User by email or account name
    * Result: SP.User
    * Example: get("my@email.address")
    * Example: get("i:0#.f|membership|my@email.address")
    */
    User.prototype.get = function (emailOrAccountName) {
        var _this = this;
        return this._fluent.chainAction(this._helperName + ".get", function () {
            return _this.userHelper.getUserByEmail(emailOrAccountName);
        });
    };
    /**
    * Get a user by their Id
    * Result: SP.User
    * Example: getById(15)
    */
    User.prototype.getById = function (id) {
        var _this = this;
        return this._fluent.chainAction(this._helperName + ".getById", function () {
            return _this.userHelper.getUserById(id);
        });
    };
    /**
    * Get the current user
    * Result: SP.User
    * Example: getCurrentUser()
    */
    User.prototype.getCurrentUser = function () {
        var _this = this;
        return this._fluent.chainAction(this._helperName + ".getCurrentUser", function () {
            return _this.userHelper.getCurrentUser();
        });
    };
    /**
    * Get the profile properties for the current user
    * Result: SP.UserProfiles.PersonProperties
    * Example: getCurrentUserProfileProperties()
    */
    User.prototype.getCurrentUserProfileProperties = function () {
        var _this = this;
        this._fluent.registerDependency(_fluent__WEBPACK_IMPORTED_MODULE_1__["Dependency"].UserProfile);
        return this._fluent.chainAction(this._helperName + ".getCurrentUserProfileProperties", function () {
            return _this.userHelper.getCurrentUserProfileProperties();
        });
    };
    /**
    * Get the manager for the current user
    * Result: SP.User
    * Example: getCurrentUserManager()
    */
    User.prototype.getCurrentUserManager = function () {
        var _this = this;
        this._fluent.registerDependency(_fluent__WEBPACK_IMPORTED_MODULE_1__["Dependency"].UserProfile);
        return this._fluent.chainAction(this._helperName + ".getCurrentUserManager", function () {
            return _this.userHelper.getCurrentUserManager();
        });
    };
    return User;
}());



/***/ }),

/***/ "./src/web.ts":
/*!********************!*\
  !*** ./src/web.ts ***!
  \********************/
/*! exports provided: Web */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Web", function() { return Web; });
/* harmony import */ var _helper_webHelper__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./helper/webHelper */ "./src/helper/webHelper.ts");

var Web = /** @class */ (function () {
    function Web(fluent) {
        this._fluent = null;
        this._helperName = "web";
        this.webHelper = null;
        this._fluent = fluent;
        this.webHelper = new _helper_webHelper__WEBPACK_IMPORTED_MODULE_0__["WebHelper"](fluent.context);
    }
    /**
      * Set the web welcome page.  The url should be relative to the root folder of the web.
      * Result: SP.Folder
      * Example: setWelcomePage(context.get_web(), "pages/default.aspx")
      */
    Web.prototype.setWelcomePage = function (web, url) {
        var _this = this;
        return this._fluent.chainAction(this._helperName + ".setWelcomePage", function () {
            return _this.webHelper.setWelcomePage(web, url);
        });
    };
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
    Web.prototype.create = function (name, parentWeb, title, template, useSamePermissionsAsParent) {
        var _this = this;
        if (useSamePermissionsAsParent === void 0) { useSamePermissionsAsParent = true; }
        return this._fluent.chainAction(this._helperName + ".create", function () {
            return _this.webHelper.createWeb(name, parentWeb, title, template, useSamePermissionsAsParent);
        });
    };
    /**
      * Check if a web exists
      * Result: boolean
      * url: server relative url
      * Example: .exists("/sites/mysubweb")
      */
    Web.prototype.exists = function (url) {
        var _this = this;
        return this._fluent.chainAction(this._helperName + ".exists", function () {
            return _this.webHelper.doesWebExist(url);
        });
    };
    /**
      * Get and load all webs starting from a web and its children
      * Result: Array<SP.Web>
      * Example: getWebs(SP.ClientContext.get_current().get_rootWeb()))
      */
    Web.prototype.getWebs = function (fromWeb) {
        var _this = this;
        return this._fluent.chainAction(this._helperName + ".getWebs", function () {
            return _this.webHelper.getWebs(fromWeb);
        });
    };
    return Web;
}());



/***/ })

/******/ });
});
//# sourceMappingURL=spJsomFluent.umd.js.map