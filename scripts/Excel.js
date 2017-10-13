var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var Excel;
(function (Excel) {
    function lowerCaseFirst(str) {
        return str[0].toLowerCase() + str.slice(1);
    }
    var iconSets = ["ThreeArrows",
        "ThreeArrowsGray",
        "ThreeFlags",
        "ThreeTrafficLights1",
        "ThreeTrafficLights2",
        "ThreeSigns",
        "ThreeSymbols",
        "ThreeSymbols2",
        "FourArrows",
        "FourArrowsGray",
        "FourRedToBlack",
        "FourRating",
        "FourTrafficLights",
        "FiveArrows",
        "FiveArrowsGray",
        "FiveRating",
        "FiveQuarters",
        "ThreeStars",
        "ThreeTriangles",
        "FiveBoxes"];
    var iconNames = [["RedDownArrow", "YellowSideArrow", "GreenUpArrow"],
        ["GrayDownArrow", "GraySideArrow", "GrayUpArrow"],
        ["RedFlag", "YellowFlag", "GreenFlag"],
        ["RedCircleWithBorder", "YellowCircle", "GreenCircle"],
        ["RedTrafficLight", "YellowTrafficLight", "GreenTrafficLight"],
        ["RedDiamond", "YellowTriangle", "GreenCircle"],
        ["RedCrossSymbol", "YellowExclamationSymbol", "GreenCheckSymbol"],
        ["RedCross", "YellowExclamation", "GreenCheck"],
        ["RedDownArrow", "YellowDownInclineArrow", "YellowUpInclineArrow", "GreenUpArrow"],
        ["GrayDownArrow", "GrayDownInclineArrow", "GrayUpInclineArrow", "GrayUpArrow"],
        ["BlackCircle", "GrayCircle", "PinkCircle", "RedCircle"],
        ["OneBar", "TwoBars", "ThreeBars", "FourBars"],
        ["BlackCircleWithBorder", "RedCircleWithBorder", "YellowCircle", "GreenCircle"],
        ["RedDownArrow", "YellowDownInclineArrow", "YellowSideArrow", "YellowUpInclineArrow", "GreenUpArrow"],
        ["GrayDownArrow", "GrayDownInclineArrow", "GraySideArrow", "GrayUpInclineArrow", "GrayUpArrow"],
        ["NoBars", "OneBar", "TwoBars", "ThreeBars", "FourBars"],
        ["WhiteCircleAllWhiteQuarters", "CircleWithThreeWhiteQuarters", "CircleWithTwoWhiteQuarters", "CircleWithOneWhiteQuarter", "BlackCircle"],
        ["SilverStar", "HalfGoldStar", "GoldStar"],
        ["RedDownTriangle", "YellowDash", "GreenUpTriangle"],
        ["NoFilledBoxes", "OneFilledBox", "TwoFilledBoxes", "ThreeFilledBoxes", "FourFilledBoxes"],];
    Excel.icons = {};
    iconSets.map(function (title, i) {
        var camelTitle = lowerCaseFirst(title);
        Excel.icons[camelTitle] = [];
        iconNames[i].map(function (iconName, j) {
            iconName = lowerCaseFirst(iconName);
            var obj = { set: title, index: j };
            Excel.icons[camelTitle].push(obj);
            Excel.icons[camelTitle][iconName] = obj;
        });
    });
    function setRangePropertiesInBulk(range, propertyName, values) {
        var maxCellCount = 1500;
        if (Array.isArray(values) && values.length > 0 && Array.isArray(values[0]) && (values.length * values[0].length > maxCellCount) && isExcel1_3OrAbove()) {
            var maxRowCount = Math.max(1, Math.round(maxCellCount / values[0].length));
            range._ValidateArraySize(values.length, values[0].length);
            for (var startRowIndex = 0; startRowIndex < values.length; startRowIndex += maxRowCount) {
                var rowCount = maxRowCount;
                if (startRowIndex + rowCount > values.length) {
                    rowCount = values.length - startRowIndex;
                }
                var chunk = range.getRow(startRowIndex).getBoundingRect(range.getRow(startRowIndex + rowCount - 1));
                var valueSlice = values.slice(startRowIndex, startRowIndex + rowCount);
                _createSetPropertyAction(chunk.context, chunk, propertyName, valueSlice);
            }
            return true;
        }
        return false;
    }
    function isExcel1_3OrAbove() {
        if (typeof (window) !== "undefined" && window.Office && window.Office.context && window.Office.context.requirements) {
            return window.Office.context.requirements.isSetSupported("ExcelApi", 1.3);
        }
        else {
            return true;
        }
    }
    var Session = (function () {
        function Session(workbookUrl, requestHeaders, persisted) {
            this.m_workbookUrl = workbookUrl;
            this.m_requestHeaders = requestHeaders;
            if (!this.m_requestHeaders) {
                this.m_requestHeaders = {};
            }
            if (OfficeExtension.Utility.isNullOrUndefined(persisted)) {
                persisted = true;
            }
            this.m_persisted = persisted;
        }
        Session.prototype.close = function () {
            var _this = this;
            if (this.m_requestUrlAndHeaderInfo &&
                !OfficeExtension.Utility._isLocalDocumentUrl(this.m_requestUrlAndHeaderInfo.url)) {
                var url = this.m_requestUrlAndHeaderInfo.url;
                if (url.charAt(url.length - 1) != "/") {
                    url = url + "/";
                }
                url = url + "closeSession";
                var headers = this.m_requestUrlAndHeaderInfo;
                var req = { method: "POST", url: url, headers: this.m_requestUrlAndHeaderInfo.headers, body: "" };
                this.m_requestUrlAndHeaderInfo = null;
                return OfficeExtension.HttpUtility.sendRequest(req)
                    .then(function (resp) {
                    if (resp.statusCode != 204) {
                        var err = OfficeExtension.Utility._parseErrorResponse(resp);
                        throw OfficeExtension.Utility.createRuntimeError(err.errorCode, err.errorMessage, "Session.close");
                    }
                    _this.m_requestUrlAndHeaderInfo = null;
                    var foundSessionKey = null;
                    for (var key in _this.m_requestHeaders) {
                        if (key.toLowerCase() == Session.WorkbookSessionIdHeaderNameLower) {
                            foundSessionKey = key;
                            break;
                        }
                    }
                    if (foundSessionKey) {
                        delete _this.m_requestHeaders[foundSessionKey];
                    }
                });
            }
            else {
                return OfficeExtension.Utility._createPromiseFromResult(null);
            }
        };
        Session.prototype._resolveRequestUrlAndHeaderInfo = function () {
            var _this = this;
            if (this.m_requestUrlAndHeaderInfo) {
                return OfficeExtension.Utility._createPromiseFromResult(this.m_requestUrlAndHeaderInfo);
            }
            if (OfficeExtension.Utility.isNullOrEmptyString(this.m_workbookUrl) ||
                OfficeExtension.Utility._isLocalDocumentUrl(this.m_workbookUrl)) {
                this.m_requestUrlAndHeaderInfo = { url: this.m_workbookUrl, headers: this.m_requestHeaders };
                return OfficeExtension.Utility._createPromiseFromResult(this.m_requestUrlAndHeaderInfo);
            }
            var foundSessionId = false;
            for (var key in this.m_requestHeaders) {
                if (key.toLowerCase() == Session.WorkbookSessionIdHeaderNameLower) {
                    foundSessionId = true;
                    break;
                }
            }
            if (foundSessionId) {
                this.m_requestUrlAndHeaderInfo = { url: this.m_workbookUrl, headers: this.m_requestHeaders };
                return OfficeExtension.Utility._createPromiseFromResult(this.m_requestUrlAndHeaderInfo);
            }
            var url = this.m_workbookUrl;
            if (url.charAt(url.length - 1) != "/") {
                url = url + "/";
            }
            url = url + "createSession";
            var headers = {};
            OfficeExtension.Utility._copyHeaders(this.m_requestHeaders, headers);
            headers["Content-Type"] = "application/json";
            var body = {};
            body.persistChanges = this.m_persisted;
            var req = { method: "POST", url: url, headers: headers, body: JSON.stringify(body) };
            return OfficeExtension.HttpUtility.sendRequest(req)
                .then(function (resp) {
                if (resp.statusCode !== 201) {
                    var err = OfficeExtension.Utility._parseErrorResponse(resp);
                    throw OfficeExtension.Utility.createRuntimeError(err.errorCode, err.errorMessage, "Session.resolveRequestUrlAndHeaderInfo");
                }
                var session = JSON.parse(resp.body);
                var sessionId = session.id;
                headers = {};
                OfficeExtension.Utility._copyHeaders(_this.m_requestHeaders, headers);
                headers[Session.WorkbookSessionIdHeaderName] = sessionId;
                _this.m_requestUrlAndHeaderInfo = { url: _this.m_workbookUrl, headers: headers };
                return _this.m_requestUrlAndHeaderInfo;
            });
        };
        return Session;
    }());
    Session.WorkbookSessionIdHeaderName = "Workbook-Session-Id";
    Session.WorkbookSessionIdHeaderNameLower = "workbook-session-id";
    Excel.Session = Session;
    var RequestContext = (function (_super) {
        __extends(RequestContext, _super);
        function RequestContext(url) {
            var _this = _super.call(this, url) || this;
            _this.m_workbook = new Workbook(_this, OfficeExtension.ObjectPathFactory.createGlobalObjectObjectPath(_this));
            _this._rootObject = _this.m_workbook;
            return _this;
        }
        RequestContext.prototype._processOfficeJsErrorResponse = function (officeJsErrorCode, response) {
            var ooeInvalidApiCallInContext = 5004;
            if (officeJsErrorCode == ooeInvalidApiCallInContext) {
                response.ErrorCode = ErrorCodes.invalidOperationInCellEditMode;
                response.ErrorMessage = OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.invalidOperationInCellEditMode);
            }
        };
        Object.defineProperty(RequestContext.prototype, "workbook", {
            get: function () {
                return this.m_workbook;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RequestContext.prototype, "application", {
            get: function () {
                return this.workbook.application;
            },
            enumerable: true,
            configurable: true
        });
        return RequestContext;
    }(OfficeCore.RequestContext));
    Excel.RequestContext = RequestContext;
    function run(arg1, arg2, arg3) {
        return OfficeExtension.ClientRequestContext._runBatch("Excel.run", arguments, function (requestInfo) {
            var ret = new Excel.RequestContext(requestInfo);
            return ret;
        });
    }
    Excel.run = run;
    Excel._RedirectV1APIs = false;
    Excel._V1APIMap = {
        "GetDataAsync": {
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingGetData(callArgs); },
            postprocess: getDataCommonPostprocess
        },
        "GetSelectedDataAsync": {
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.getSelectedData(callArgs); },
            postprocess: getDataCommonPostprocess
        },
        "GoToByIdAsync": {
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.gotoById(callArgs); }
        },
        "AddColumnsAsync": {
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingAddColumns(callArgs); }
        },
        "AddFromSelectionAsync": {
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingAddFromSelection(callArgs); },
            postprocess: postprocessBindingDescriptor
        },
        "AddFromNamedItemAsync": {
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingAddFromNamedItem(callArgs); },
            postprocess: postprocessBindingDescriptor
        },
        "AddFromPromptAsync": {
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingAddFromPrompt(callArgs); },
            postprocess: postprocessBindingDescriptor
        },
        "AddRowsAsync": {
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingAddRows(callArgs); }
        },
        "GetByIdAsync": {
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingGetById(callArgs); },
            postprocess: postprocessBindingDescriptor
        },
        "ReleaseByIdAsync": {
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingReleaseById(callArgs); }
        },
        "GetAllAsync": {
            call: function (ctx) { return ctx.workbook._V1Api.bindingGetAll(); },
            postprocess: function (response) {
                return response.bindings.map(function (descriptor) { return postprocessBindingDescriptor(descriptor); });
            }
        },
        "DeleteAllDataValuesAsync": {
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingDeleteAllDataValues(callArgs); }
        },
        "SetSelectedDataAsync": {
            preprocess: function (callArgs) {
                var preimage = callArgs["cellFormat"];
                if (typeof (window) !== "undefined" && window.OSF.DDA.SafeArray) {
                    if (window.OSF.OUtil.listContainsKey(window.OSF.DDA.SafeArray.Delegate.ParameterMap.dynamicTypes, "cellFormat")) {
                        callArgs["cellFormat"] = window.OSF.DDA.SafeArray.Delegate.ParameterMap.dynamicTypes["cellFormat"]["toHost"](preimage);
                    }
                }
                else if (typeof (window) !== "undefined" && window.OSF.DDA.WAC) {
                    if (window.OSF.OUtil.listContainsKey(window.OSF.DDA.WAC.Delegate.ParameterMap.dynamicTypes, "cellFormat")) {
                        callArgs["cellFormat"] = window.OSF.DDA.WAC.Delegate.ParameterMap.dynamicTypes["cellFormat"]["toHost"](preimage);
                    }
                }
                return callArgs;
            },
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.setSelectedData(callArgs); }
        },
        "SetDataAsync": {
            preprocess: function (callArgs) {
                var preimage = callArgs["cellFormat"];
                if (typeof (window) !== "undefined" && window.OSF.DDA.SafeArray) {
                    if (window.OSF.OUtil.listContainsKey(window.OSF.DDA.SafeArray.Delegate.ParameterMap.dynamicTypes, "cellFormat")) {
                        callArgs["cellFormat"] = window.OSF.DDA.SafeArray.Delegate.ParameterMap.dynamicTypes["cellFormat"]["toHost"](preimage);
                    }
                }
                else if (typeof (window) !== "undefined" && window.OSF.DDA.WAC) {
                    if (window.OSF.OUtil.listContainsKey(window.OSF.DDA.WAC.Delegate.ParameterMap.dynamicTypes, "cellFormat")) {
                        callArgs["cellFormat"] = window.OSF.DDA.WAC.Delegate.ParameterMap.dynamicTypes["cellFormat"]["toHost"](preimage);
                    }
                }
                return callArgs;
            },
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingSetData(callArgs); }
        },
        "SetFormatsAsync": {
            preprocess: function (callArgs) {
                var preimage = callArgs["cellFormat"];
                if (typeof (window) !== "undefined" && window.OSF.DDA.SafeArray) {
                    if (window.OSF.OUtil.listContainsKey(window.OSF.DDA.SafeArray.Delegate.ParameterMap.dynamicTypes, "cellFormat")) {
                        callArgs["cellFormat"] = window.OSF.DDA.SafeArray.Delegate.ParameterMap.dynamicTypes["cellFormat"]["toHost"](preimage);
                    }
                }
                else if (typeof (window) !== "undefined" && window.OSF.DDA.WAC) {
                    if (window.OSF.OUtil.listContainsKey(window.OSF.DDA.WAC.Delegate.ParameterMap.dynamicTypes, "cellFormat")) {
                        callArgs["cellFormat"] = window.OSF.DDA.WAC.Delegate.ParameterMap.dynamicTypes["cellFormat"]["toHost"](preimage);
                    }
                }
                return callArgs;
            },
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingSetFormats(callArgs); }
        },
        "SetTableOptionsAsync": {
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingSetTableOptions(callArgs); }
        },
        "ClearFormatsAsync": {
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingClearFormats(callArgs); }
        },
        "GetFilePropertiesAsync": {
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.getFilePropertiesAsync(callArgs); }
        },
    };
    function postprocessBindingDescriptor(response) {
        var bindingDescriptor = {
            BindingColumnCount: response.bindingColumnCount,
            BindingId: response.bindingId,
            BindingRowCount: response.bindingRowCount,
            bindingType: response.bindingType,
            HasHeaders: response.hasHeaders
        };
        return window.OSF.DDA.OMFactory.manufactureBinding(bindingDescriptor, window.Microsoft.Office.WebExtension.context.document);
    }
    function getDataCommonPostprocess(response, callArgs) {
        var isPlainData = response.headers == null;
        var data;
        if (isPlainData) {
            data = response.rows;
        }
        else {
            data = response;
        }
        data = window.OSF.DDA.DataCoercion.coerceData(data, callArgs[window.Microsoft.Office.WebExtension.Parameters.CoercionType]);
        return data == undefined ? null : data;
    }
    Excel.Script = {
        CustomFunctions: {}
    };
    var _hostName = "Excel";
    var _defaultApiSetName = "ExcelApi";
    var _createPropertyObjectPath = OfficeExtension.ObjectPathFactory.createPropertyObjectPath;
    var _createMethodObjectPath = OfficeExtension.ObjectPathFactory.createMethodObjectPath;
    var _createIndexerObjectPath = OfficeExtension.ObjectPathFactory.createIndexerObjectPath;
    var _createNewObjectObjectPath = OfficeExtension.ObjectPathFactory.createNewObjectObjectPath;
    var _createChildItemObjectPathUsingIndexer = OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexer;
    var _createChildItemObjectPathUsingGetItemAt = OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingGetItemAt;
    var _createChildItemObjectPathUsingIndexerOrGetItemAt = OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexerOrGetItemAt;
    var _createMethodAction = OfficeExtension.ActionFactory.createMethodAction;
    var _createSetPropertyAction = OfficeExtension.ActionFactory.createSetPropertyAction;
    var _isNullOrUndefined = OfficeExtension.Utility.isNullOrUndefined;
    var _isUndefined = OfficeExtension.Utility.isUndefined;
    var _throwIfNotLoaded = OfficeExtension.Utility.throwIfNotLoaded;
    var _throwIfApiNotSupported = OfficeExtension.Utility.throwIfApiNotSupported;
    var _load = OfficeExtension.Utility.load;
    var _fixObjectPathIfNecessary = OfficeExtension.Utility.fixObjectPathIfNecessary;
    var _addActionResultHandler = OfficeExtension.Utility._addActionResultHandler;
    var _handleNavigationPropertyResults = OfficeExtension.Utility._handleNavigationPropertyResults;
    var _adjustToDateTime = OfficeExtension.Utility.adjustToDateTime;
    var Application = (function (_super) {
        __extends(Application, _super);
        function Application() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(Application.prototype, "_className", {
            get: function () {
                return "Application";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Application.prototype, "calculationMode", {
            get: function () {
                _throwIfNotLoaded("calculationMode", this.m_calculationMode, "Application", this._isNull);
                return this.m_calculationMode;
            },
            enumerable: true,
            configurable: true
        });
        Application.prototype.calculate = function (calculationType) {
            _createMethodAction(this.context, this, "Calculate", 0, [calculationType]);
        };
        Application.prototype.suspendApiCalculationUntilNextSync = function () {
            _throwIfApiNotSupported("Application.suspendApiCalculationUntilNextSync", _defaultApiSetName, "1.6", _hostName);
            _createMethodAction(this.context, this, "SuspendApiCalculationUntilNextSync", 0, []);
        };
        Application.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["CalculationMode"])) {
                this.m_calculationMode = obj["CalculationMode"];
            }
        };
        Application.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        Application.prototype.toJSON = function () {
            return {
                "calculationMode": this.m_calculationMode
            };
        };
        return Application;
    }(OfficeExtension.ClientObject));
    Excel.Application = Application;
    var Workbook = (function (_super) {
        __extends(Workbook, _super);
        function Workbook() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(Workbook.prototype, "_className", {
            get: function () {
                return "Workbook";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Workbook.prototype, "application", {
            get: function () {
                if (!this.m_application) {
                    this.m_application = new Excel.Application(this.context, _createPropertyObjectPath(this.context, this, "Application", false, false));
                }
                return this.m_application;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Workbook.prototype, "bindings", {
            get: function () {
                if (!this.m_bindings) {
                    this.m_bindings = new Excel.BindingCollection(this.context, _createPropertyObjectPath(this.context, this, "Bindings", true, false));
                }
                return this.m_bindings;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Workbook.prototype, "customFunctions", {
            get: function () {
                _throwIfApiNotSupported("Workbook.customFunctions", _defaultApiSetName, "1.7", _hostName);
                if (!this.m_customFunctions) {
                    this.m_customFunctions = new Excel.CustomFunctionCollection(this.context, _createPropertyObjectPath(this.context, this, "CustomFunctions", true, false));
                }
                return this.m_customFunctions;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Workbook.prototype, "customXmlParts", {
            get: function () {
                _throwIfApiNotSupported("Workbook.customXmlParts", _defaultApiSetName, "1.5", _hostName);
                if (!this.m_customXmlParts) {
                    this.m_customXmlParts = new Excel.CustomXmlPartCollection(this.context, _createPropertyObjectPath(this.context, this, "CustomXmlParts", true, false));
                }
                return this.m_customXmlParts;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Workbook.prototype, "functions", {
            get: function () {
                _throwIfApiNotSupported("Workbook.functions", _defaultApiSetName, "1.2", _hostName);
                if (!this.m_functions) {
                    this.m_functions = new Excel.Functions(this.context, _createPropertyObjectPath(this.context, this, "Functions", false, false));
                }
                return this.m_functions;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Workbook.prototype, "internalTest", {
            get: function () {
                _throwIfApiNotSupported("Workbook.internalTest", _defaultApiSetName, "1.6", _hostName);
                if (!this.m_internalTest) {
                    this.m_internalTest = new Excel.InternalTest(this.context, _createPropertyObjectPath(this.context, this, "InternalTest", false, false));
                }
                return this.m_internalTest;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Workbook.prototype, "names", {
            get: function () {
                if (!this.m_names) {
                    this.m_names = new Excel.NamedItemCollection(this.context, _createPropertyObjectPath(this.context, this, "Names", true, false));
                }
                return this.m_names;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Workbook.prototype, "pivotTables", {
            get: function () {
                _throwIfApiNotSupported("Workbook.pivotTables", _defaultApiSetName, "1.3", _hostName);
                if (!this.m_pivotTables) {
                    this.m_pivotTables = new Excel.PivotTableCollection(this.context, _createPropertyObjectPath(this.context, this, "PivotTables", true, false));
                }
                return this.m_pivotTables;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Workbook.prototype, "properties", {
            get: function () {
                _throwIfApiNotSupported("Workbook.properties", _defaultApiSetName, "1.7", _hostName);
                if (!this.m_properties) {
                    this.m_properties = new Excel.DocumentProperties(this.context, _createPropertyObjectPath(this.context, this, "Properties", false, false));
                }
                return this.m_properties;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Workbook.prototype, "settings", {
            get: function () {
                _throwIfApiNotSupported("Workbook.settings", _defaultApiSetName, "1.4", _hostName);
                if (!this.m_settings) {
                    this.m_settings = new Excel.SettingCollection(this.context, _createPropertyObjectPath(this.context, this, "Settings", true, false));
                }
                return this.m_settings;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Workbook.prototype, "tables", {
            get: function () {
                if (!this.m_tables) {
                    this.m_tables = new Excel.TableCollection(this.context, _createPropertyObjectPath(this.context, this, "Tables", true, false));
                }
                return this.m_tables;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Workbook.prototype, "worksheets", {
            get: function () {
                if (!this.m_worksheets) {
                    this.m_worksheets = new Excel.WorksheetCollection(this.context, _createPropertyObjectPath(this.context, this, "Worksheets", true, false));
                }
                return this.m_worksheets;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Workbook.prototype, "_V1Api", {
            get: function () {
                _throwIfApiNotSupported("Workbook._V1Api", _defaultApiSetName, "1.3", _hostName);
                if (!this.m__V1Api) {
                    this.m__V1Api = new Excel._V1Api(this.context, _createPropertyObjectPath(this.context, this, "_V1Api", false, false));
                }
                return this.m__V1Api;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Workbook.prototype, "pivotCaches", {
            get: function () {
                _throwIfApiNotSupported("Workbook.pivotCaches", "Pivot", "1.1", _hostName);
                if (!this.m_pivotCaches) {
                    this.m_pivotCaches = new Excel.PivotCacheCollection(this.context, _createPropertyObjectPath(this.context, this, "PivotCaches", true, false));
                }
                return this.m_pivotCaches;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Workbook.prototype, "name", {
            get: function () {
                _throwIfNotLoaded("name", this.m_name, "Workbook", this._isNull);
                _throwIfApiNotSupported("Workbook.name", _defaultApiSetName, "1.7", _hostName);
                return this.m_name;
            },
            enumerable: true,
            configurable: true
        });
        Workbook.prototype.getActiveCell = function () {
            _throwIfApiNotSupported("Workbook.getActiveCell", _defaultApiSetName, "1.7", _hostName);
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetActiveCell", 1, [], false, true, null));
        };
        Workbook.prototype.getSelectedRange = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetSelectedRange", 1, [], false, true, null));
        };
        Workbook.prototype._GetObjectByReferenceId = function (bstrReferenceId) {
            var action = _createMethodAction(this.context, this, "_GetObjectByReferenceId", 1, [bstrReferenceId]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        Workbook.prototype._GetObjectTypeNameByReferenceId = function (bstrReferenceId) {
            var action = _createMethodAction(this.context, this, "_GetObjectTypeNameByReferenceId", 1, [bstrReferenceId]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        Workbook.prototype._GetReferenceCount = function () {
            var action = _createMethodAction(this.context, this, "_GetReferenceCount", 1, []);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        Workbook.prototype._RemoveAllReferences = function () {
            _createMethodAction(this.context, this, "_RemoveAllReferences", 1, []);
        };
        Workbook.prototype._RemoveReference = function (bstrReferenceId) {
            _createMethodAction(this.context, this, "_RemoveReference", 1, [bstrReferenceId]);
        };
        Workbook.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Name"])) {
                this.m_name = obj["Name"];
            }
            _handleNavigationPropertyResults(this, obj, ["application", "Application", "bindings", "Bindings", "customFunctions", "CustomFunctions", "customXmlParts", "CustomXmlParts", "functions", "Functions", "internalTest", "InternalTest", "names", "Names", "pivotTables", "PivotTables", "properties", "Properties", "settings", "Settings", "tables", "Tables", "worksheets", "Worksheets", "_V1Api", "_V1Api", "pivotCaches", "PivotCaches"]);
        };
        Workbook.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        Object.defineProperty(Workbook.prototype, "onSelectionChanged", {
            get: function () {
                var _this = this;
                _throwIfApiNotSupported("Workbook.onSelectionChanged", _defaultApiSetName, "1.3", _hostName);
                if (!this.m_selectionChanged) {
                    this.m_selectionChanged = new OfficeExtension.EventHandlers(this.context, this, "SelectionChanged", {
                        registerFunc: function (handlerCallback) {
                            return _this.context.eventRegistration.register(2, "", handlerCallback);
                        },
                        unregisterFunc: function (handlerCallback) {
                            return _this.context.eventRegistration.unregister(2, "", handlerCallback);
                        },
                        eventArgsTransformFunc: function (args) {
                            return OfficeExtension.Utility._createPromiseFromResult({ workbook: _this });
                        }
                    });
                }
                return this.m_selectionChanged;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Workbook.prototype, "_onMessage", {
            get: function () {
                var _this = this;
                _throwIfApiNotSupported("Workbook._onMessage", _defaultApiSetName, "1.8", _hostName);
                if (!this.m__Message) {
                    this.m__Message = new OfficeExtension.EventHandlers(this.context, this, "_Message", {
                        registerFunc: function (handlerCallback) {
                            return _this.context.eventRegistration.register(5, "", handlerCallback);
                        },
                        unregisterFunc: function (handlerCallback) {
                            return _this.context.eventRegistration.unregister(5, "", handlerCallback);
                        },
                        eventArgsTransformFunc: function (args) {
                            return OfficeExtension.Utility._createPromiseFromResult({
                                entries: args.entries,
                                workbook: _this
                            });
                        }
                    });
                }
                return this.m__Message;
            },
            enumerable: true,
            configurable: true
        });
        Workbook.prototype.toJSON = function () {
            return {
                "name": this.m_name
            };
        };
        return Workbook;
    }(OfficeExtension.ClientObject));
    Excel.Workbook = Workbook;
    var Worksheet = (function (_super) {
        __extends(Worksheet, _super);
        function Worksheet() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(Worksheet.prototype, "_className", {
            get: function () {
                return "Worksheet";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Worksheet.prototype, "charts", {
            get: function () {
                if (!this.m_charts) {
                    this.m_charts = new Excel.ChartCollection(this.context, _createPropertyObjectPath(this.context, this, "Charts", true, false));
                }
                return this.m_charts;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Worksheet.prototype, "freezePanes", {
            get: function () {
                _throwIfApiNotSupported("Worksheet.freezePanes", _defaultApiSetName, "1.7", _hostName);
                if (!this.m_freezePanes) {
                    this.m_freezePanes = new Excel.WorksheetFreezePanes(this.context, _createPropertyObjectPath(this.context, this, "FreezePanes", false, false));
                }
                return this.m_freezePanes;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Worksheet.prototype, "names", {
            get: function () {
                _throwIfApiNotSupported("Worksheet.names", _defaultApiSetName, "1.4", _hostName);
                if (!this.m_names) {
                    this.m_names = new Excel.NamedItemCollection(this.context, _createPropertyObjectPath(this.context, this, "Names", true, false));
                }
                return this.m_names;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Worksheet.prototype, "pivotTables", {
            get: function () {
                _throwIfApiNotSupported("Worksheet.pivotTables", _defaultApiSetName, "1.3", _hostName);
                if (!this.m_pivotTables) {
                    this.m_pivotTables = new Excel.PivotTableCollection(this.context, _createPropertyObjectPath(this.context, this, "PivotTables", true, false));
                }
                return this.m_pivotTables;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Worksheet.prototype, "protection", {
            get: function () {
                _throwIfApiNotSupported("Worksheet.protection", _defaultApiSetName, "1.2", _hostName);
                if (!this.m_protection) {
                    this.m_protection = new Excel.WorksheetProtection(this.context, _createPropertyObjectPath(this.context, this, "Protection", false, false));
                }
                return this.m_protection;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Worksheet.prototype, "tables", {
            get: function () {
                if (!this.m_tables) {
                    this.m_tables = new Excel.TableCollection(this.context, _createPropertyObjectPath(this.context, this, "Tables", true, false));
                }
                return this.m_tables;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Worksheet.prototype, "gridlines", {
            get: function () {
                _throwIfNotLoaded("gridlines", this.m_gridlines, "Worksheet", this._isNull);
                _throwIfApiNotSupported("Worksheet.gridlines", _defaultApiSetName, "1.7", _hostName);
                return this.m_gridlines;
            },
            set: function (value) {
                this.m_gridlines = value;
                _createSetPropertyAction(this.context, this, "Gridlines", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Worksheet.prototype, "headings", {
            get: function () {
                _throwIfNotLoaded("headings", this.m_headings, "Worksheet", this._isNull);
                _throwIfApiNotSupported("Worksheet.headings", _defaultApiSetName, "1.7", _hostName);
                return this.m_headings;
            },
            set: function (value) {
                this.m_headings = value;
                _createSetPropertyAction(this.context, this, "Headings", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Worksheet.prototype, "id", {
            get: function () {
                _throwIfNotLoaded("id", this.m_id, "Worksheet", this._isNull);
                return this.m_id;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Worksheet.prototype, "name", {
            get: function () {
                _throwIfNotLoaded("name", this.m_name, "Worksheet", this._isNull);
                return this.m_name;
            },
            set: function (value) {
                this.m_name = value;
                _createSetPropertyAction(this.context, this, "Name", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Worksheet.prototype, "position", {
            get: function () {
                _throwIfNotLoaded("position", this.m_position, "Worksheet", this._isNull);
                return this.m_position;
            },
            set: function (value) {
                this.m_position = value;
                _createSetPropertyAction(this.context, this, "Position", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Worksheet.prototype, "tabColor", {
            get: function () {
                _throwIfNotLoaded("tabColor", this.m_tabColor, "Worksheet", this._isNull);
                _throwIfApiNotSupported("Worksheet.tabColor", _defaultApiSetName, "1.7", _hostName);
                return this.m_tabColor;
            },
            set: function (value) {
                this.m_tabColor = value;
                _createSetPropertyAction(this.context, this, "TabColor", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Worksheet.prototype, "visibility", {
            get: function () {
                _throwIfNotLoaded("visibility", this.m_visibility, "Worksheet", this._isNull);
                return this.m_visibility;
            },
            set: function (value) {
                this.m_visibility = value;
                _createSetPropertyAction(this.context, this, "Visibility", value);
            },
            enumerable: true,
            configurable: true
        });
        Worksheet.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["name", "position", "visibility", "tabColor", "gridlines", "headings"], [], [
                "charts",
                "freezePanes",
                "names",
                "pivotTables",
                "tables",
                "charts",
                "freezePanes",
                "names",
                "pivotTables",
                "protection",
                "tables"
            ]);
        };
        Worksheet.prototype.activate = function () {
            _createMethodAction(this.context, this, "Activate", 1, []);
        };
        Worksheet.prototype.calculate = function (markAllDirty) {
            _throwIfApiNotSupported("Worksheet.calculate", _defaultApiSetName, "1.6", _hostName);
            _createMethodAction(this.context, this, "Calculate", 0, [markAllDirty]);
        };
        Worksheet.prototype.delete = function () {
            _createMethodAction(this.context, this, "Delete", 0, []);
        };
        Worksheet.prototype.getCell = function (row, column) {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetCell", 1, [row, column], false, true, null));
        };
        Worksheet.prototype.getNext = function (visibleOnly) {
            _throwIfApiNotSupported("Worksheet.getNext", _defaultApiSetName, "1.5", _hostName);
            return new Excel.Worksheet(this.context, _createMethodObjectPath(this.context, this, "GetNext", 1, [visibleOnly], false, true, null));
        };
        Worksheet.prototype.getNextOrNullObject = function (visibleOnly) {
            _throwIfApiNotSupported("Worksheet.getNextOrNullObject", _defaultApiSetName, "1.5", _hostName);
            return new Excel.Worksheet(this.context, _createMethodObjectPath(this.context, this, "GetNextOrNullObject", 1, [visibleOnly], false, true, null));
        };
        Worksheet.prototype.getPrevious = function (visibleOnly) {
            _throwIfApiNotSupported("Worksheet.getPrevious", _defaultApiSetName, "1.5", _hostName);
            return new Excel.Worksheet(this.context, _createMethodObjectPath(this.context, this, "GetPrevious", 1, [visibleOnly], false, true, null));
        };
        Worksheet.prototype.getPreviousOrNullObject = function (visibleOnly) {
            _throwIfApiNotSupported("Worksheet.getPreviousOrNullObject", _defaultApiSetName, "1.5", _hostName);
            return new Excel.Worksheet(this.context, _createMethodObjectPath(this.context, this, "GetPreviousOrNullObject", 1, [visibleOnly], false, true, null));
        };
        Worksheet.prototype.getRange = function (address) {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetRange", 1, [address], false, true, null));
        };
        Worksheet.prototype.getRangeByIndexes = function (startRow, startColumn, rowCount, columnCount) {
            _throwIfApiNotSupported("Worksheet.getRangeByIndexes", _defaultApiSetName, "1.7", _hostName);
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetRangeByIndexes", 1, [startRow, startColumn, rowCount, columnCount], false, true, null));
        };
        Worksheet.prototype.getUsedRange = function (valuesOnly) {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetUsedRange", 1, [valuesOnly], false, true, null));
        };
        Worksheet.prototype.getUsedRangeOrNullObject = function (valuesOnly) {
            _throwIfApiNotSupported("Worksheet.getUsedRangeOrNullObject", _defaultApiSetName, "1.4", _hostName);
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetUsedRangeOrNullObject", 1, [valuesOnly], false, true, null));
        };
        Worksheet.prototype._RegisterDataChangedEvent = function () {
            _throwIfApiNotSupported("Worksheet._RegisterDataChangedEvent", _defaultApiSetName, "1.8", _hostName);
            _createMethodAction(this.context, this, "_RegisterDataChangedEvent", 0, []);
        };
        Worksheet.prototype._UnregisterDataChangedEvent = function () {
            _throwIfApiNotSupported("Worksheet._UnregisterDataChangedEvent", _defaultApiSetName, "1.8", _hostName);
            _createMethodAction(this.context, this, "_UnregisterDataChangedEvent", 0, []);
        };
        Worksheet.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Gridlines"])) {
                this.m_gridlines = obj["Gridlines"];
            }
            if (!_isUndefined(obj["Headings"])) {
                this.m_headings = obj["Headings"];
            }
            if (!_isUndefined(obj["Id"])) {
                this.m_id = obj["Id"];
            }
            if (!_isUndefined(obj["Name"])) {
                this.m_name = obj["Name"];
            }
            if (!_isUndefined(obj["Position"])) {
                this.m_position = obj["Position"];
            }
            if (!_isUndefined(obj["TabColor"])) {
                this.m_tabColor = obj["TabColor"];
            }
            if (!_isUndefined(obj["Visibility"])) {
                this.m_visibility = obj["Visibility"];
            }
            _handleNavigationPropertyResults(this, obj, ["charts", "Charts", "freezePanes", "FreezePanes", "names", "Names", "pivotTables", "PivotTables", "protection", "Protection", "tables", "Tables"]);
        };
        Worksheet.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        Worksheet.prototype._handleIdResult = function (value) {
            _super.prototype._handleIdResult.call(this, value);
            if (_isNullOrUndefined(value)) {
                return;
            }
            if (!_isUndefined(value["Id"])) {
                this.m_id = value["Id"];
            }
        };
        Object.defineProperty(Worksheet.prototype, "onActivated", {
            get: function () {
                _throwIfApiNotSupported("Worksheet.onActivated", _defaultApiSetName, "1.8", _hostName);
                if (!this.m_activated) {
                    this.m_activated = new OfficeExtension.GenericEventHandlers(this.context, this, "Activated", {
                        eventType: 0,
                        registerFunc: function () { return null; },
                        unregisterFunc: function () { return null; },
                        getTargetIdFunc: null,
                        eventArgsTransformFunc: function (value) {
                            return null;
                        }
                    });
                }
                return this.m_activated;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Worksheet.prototype, "onDataChanged", {
            get: function () {
                _throwIfApiNotSupported("Worksheet.onDataChanged", _defaultApiSetName, "1.8", _hostName);
                if (!this.m_dataChanged) {
                    this.m_dataChanged = new OfficeExtension.GenericEventHandlers(this.context, this, "DataChanged", {
                        eventType: 0,
                        registerFunc: function () { return null; },
                        unregisterFunc: function () { return null; },
                        getTargetIdFunc: null,
                        eventArgsTransformFunc: function (value) {
                            return null;
                        }
                    });
                }
                return this.m_dataChanged;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Worksheet.prototype, "onDeactivated", {
            get: function () {
                _throwIfApiNotSupported("Worksheet.onDeactivated", _defaultApiSetName, "1.8", _hostName);
                if (!this.m_deactivated) {
                    this.m_deactivated = new OfficeExtension.GenericEventHandlers(this.context, this, "Deactivated", {
                        eventType: 0,
                        registerFunc: function () { return null; },
                        unregisterFunc: function () { return null; },
                        getTargetIdFunc: null,
                        eventArgsTransformFunc: function (value) {
                            return null;
                        }
                    });
                }
                return this.m_deactivated;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Worksheet.prototype, "onSelectionChanged", {
            get: function () {
                _throwIfApiNotSupported("Worksheet.onSelectionChanged", _defaultApiSetName, "1.8", _hostName);
                if (!this.m_selectionChanged) {
                    this.m_selectionChanged = new OfficeExtension.GenericEventHandlers(this.context, this, "SelectionChanged", {
                        eventType: 0,
                        registerFunc: function () { return null; },
                        unregisterFunc: function () { return null; },
                        getTargetIdFunc: null,
                        eventArgsTransformFunc: function (value) {
                            return null;
                        }
                    });
                }
                return this.m_selectionChanged;
            },
            enumerable: true,
            configurable: true
        });
        Worksheet.prototype.toJSON = function () {
            return {
                "gridlines": this.m_gridlines,
                "headings": this.m_headings,
                "id": this.m_id,
                "name": this.m_name,
                "position": this.m_position,
                "protection": this.m_protection,
                "tabColor": this.m_tabColor,
                "visibility": this.m_visibility
            };
        };
        return Worksheet;
    }(OfficeExtension.ClientObject));
    Excel.Worksheet = Worksheet;
    var WorksheetCollection = (function (_super) {
        __extends(WorksheetCollection, _super);
        function WorksheetCollection() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(WorksheetCollection.prototype, "_className", {
            get: function () {
                return "WorksheetCollection";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(WorksheetCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, "WorksheetCollection", this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        WorksheetCollection.prototype.add = function (name) {
            return new Excel.Worksheet(this.context, _createMethodObjectPath(this.context, this, "Add", 0, [name], false, true, null));
        };
        WorksheetCollection.prototype.getActiveWorksheet = function () {
            return new Excel.Worksheet(this.context, _createMethodObjectPath(this.context, this, "GetActiveWorksheet", 1, [], false, false, null));
        };
        WorksheetCollection.prototype.getCount = function (visibleOnly) {
            _throwIfApiNotSupported("WorksheetCollection.getCount", _defaultApiSetName, "1.4", _hostName);
            var action = _createMethodAction(this.context, this, "GetCount", 1, [visibleOnly]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        WorksheetCollection.prototype.getFirst = function (visibleOnly) {
            _throwIfApiNotSupported("WorksheetCollection.getFirst", _defaultApiSetName, "1.5", _hostName);
            return new Excel.Worksheet(this.context, _createMethodObjectPath(this.context, this, "GetFirst", 1, [visibleOnly], false, true, null));
        };
        WorksheetCollection.prototype.getItem = function (key) {
            return new Excel.Worksheet(this.context, _createIndexerObjectPath(this.context, this, [key]));
        };
        WorksheetCollection.prototype.getItemOrNullObject = function (key) {
            _throwIfApiNotSupported("WorksheetCollection.getItemOrNullObject", _defaultApiSetName, "1.4", _hostName);
            return new Excel.Worksheet(this.context, _createMethodObjectPath(this.context, this, "GetItemOrNullObject", 1, [key], false, false, null));
        };
        WorksheetCollection.prototype.getLast = function (visibleOnly) {
            _throwIfApiNotSupported("WorksheetCollection.getLast", _defaultApiSetName, "1.5", _hostName);
            return new Excel.Worksheet(this.context, _createMethodObjectPath(this.context, this, "GetLast", 1, [visibleOnly], false, true, null));
        };
        WorksheetCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.Worksheet(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        WorksheetCollection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        Object.defineProperty(WorksheetCollection.prototype, "onActivated", {
            get: function () {
                _throwIfApiNotSupported("WorksheetCollection.onActivated", _defaultApiSetName, "1.8", _hostName);
                if (!this.m_activated) {
                    this.m_activated = new OfficeExtension.GenericEventHandlers(this.context, this, "Activated", {
                        eventType: 0,
                        registerFunc: function () { return null; },
                        unregisterFunc: function () { return null; },
                        getTargetIdFunc: null,
                        eventArgsTransformFunc: function (value) {
                            return null;
                        }
                    });
                }
                return this.m_activated;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(WorksheetCollection.prototype, "onAdded", {
            get: function () {
                _throwIfApiNotSupported("WorksheetCollection.onAdded", _defaultApiSetName, "1.8", _hostName);
                if (!this.m_added) {
                    this.m_added = new OfficeExtension.GenericEventHandlers(this.context, this, "Added", {
                        eventType: 0,
                        registerFunc: function () { return null; },
                        unregisterFunc: function () { return null; },
                        getTargetIdFunc: null,
                        eventArgsTransformFunc: function (value) {
                            return null;
                        }
                    });
                }
                return this.m_added;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(WorksheetCollection.prototype, "onDeactivated", {
            get: function () {
                _throwIfApiNotSupported("WorksheetCollection.onDeactivated", _defaultApiSetName, "1.8", _hostName);
                if (!this.m_deactivated) {
                    this.m_deactivated = new OfficeExtension.GenericEventHandlers(this.context, this, "Deactivated", {
                        eventType: 0,
                        registerFunc: function () { return null; },
                        unregisterFunc: function () { return null; },
                        getTargetIdFunc: null,
                        eventArgsTransformFunc: function (value) {
                            return null;
                        }
                    });
                }
                return this.m_deactivated;
            },
            enumerable: true,
            configurable: true
        });
        WorksheetCollection.prototype.toJSON = function () {
            return {};
        };
        return WorksheetCollection;
    }(OfficeExtension.ClientObject));
    Excel.WorksheetCollection = WorksheetCollection;
    var WorksheetProtection = (function (_super) {
        __extends(WorksheetProtection, _super);
        function WorksheetProtection() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(WorksheetProtection.prototype, "_className", {
            get: function () {
                return "WorksheetProtection";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(WorksheetProtection.prototype, "options", {
            get: function () {
                _throwIfNotLoaded("options", this.m_options, "WorksheetProtection", this._isNull);
                return this.m_options;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(WorksheetProtection.prototype, "protected", {
            get: function () {
                _throwIfNotLoaded("protected", this.m_protected, "WorksheetProtection", this._isNull);
                return this.m_protected;
            },
            enumerable: true,
            configurable: true
        });
        WorksheetProtection.prototype.protect = function (options) {
            _createMethodAction(this.context, this, "Protect", 0, [options]);
        };
        WorksheetProtection.prototype.unprotect = function () {
            _createMethodAction(this.context, this, "Unprotect", 0, []);
        };
        WorksheetProtection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Options"])) {
                this.m_options = obj["Options"];
            }
            if (!_isUndefined(obj["Protected"])) {
                this.m_protected = obj["Protected"];
            }
        };
        WorksheetProtection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        WorksheetProtection.prototype.toJSON = function () {
            return {
                "options": this.m_options,
                "protected": this.m_protected
            };
        };
        return WorksheetProtection;
    }(OfficeExtension.ClientObject));
    Excel.WorksheetProtection = WorksheetProtection;
    var WorksheetFreezePanes = (function (_super) {
        __extends(WorksheetFreezePanes, _super);
        function WorksheetFreezePanes() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(WorksheetFreezePanes.prototype, "_className", {
            get: function () {
                return "WorksheetFreezePanes";
            },
            enumerable: true,
            configurable: true
        });
        WorksheetFreezePanes.prototype.freezeAt = function (frozenRange) {
            _createMethodAction(this.context, this, "FreezeAt", 0, [frozenRange]);
        };
        WorksheetFreezePanes.prototype.freezeColumns = function (count) {
            _createMethodAction(this.context, this, "FreezeColumns", 0, [count]);
        };
        WorksheetFreezePanes.prototype.freezeRows = function (count) {
            _createMethodAction(this.context, this, "FreezeRows", 0, [count]);
        };
        WorksheetFreezePanes.prototype.getLocation = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetLocation", 1, [], false, true, null));
        };
        WorksheetFreezePanes.prototype.getLocationOrNullObject = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetLocationOrNullObject", 1, [], false, true, null));
        };
        WorksheetFreezePanes.prototype.unfreeze = function () {
            _createMethodAction(this.context, this, "Unfreeze", 0, []);
        };
        WorksheetFreezePanes.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
        };
        WorksheetFreezePanes.prototype.toJSON = function () {
            return {};
        };
        return WorksheetFreezePanes;
    }(OfficeExtension.ClientObject));
    Excel.WorksheetFreezePanes = WorksheetFreezePanes;
    var Range = (function (_super) {
        __extends(Range, _super);
        function Range() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(Range.prototype, "_className", {
            get: function () {
                return "Range";
            },
            enumerable: true,
            configurable: true
        });
        Range.prototype._ensureInteger = function (num, methodName) {
            if (!(typeof num === "number" && isFinite(num) && Math.floor(num) === num)) {
                throw new OfficeExtension.Utility.throwError(Excel.ErrorCodes.invalidArgument, num, methodName);
            }
        };
        Range.prototype._getAdjacentRange = function (functionName, count, referenceRange, rowDirection, columnDirection) {
            if (count == null) {
                count = 1;
            }
            this._ensureInteger(count, functionName);
            var startRange;
            var rowOffset = 0;
            var columnOffset = 0;
            if (count > 0) {
                startRange = referenceRange.getOffsetRange(rowDirection, columnDirection);
            }
            else {
                startRange = referenceRange;
                rowOffset = rowDirection;
                columnOffset = columnDirection;
            }
            if (Math.abs(count) == 1) {
                return startRange;
            }
            return startRange.getBoundingRect(referenceRange.getOffsetRange(rowDirection * count + rowOffset, columnDirection * count + columnOffset));
        };
        Object.defineProperty(Range.prototype, "conditionalFormats", {
            get: function () {
                _throwIfApiNotSupported("Range.conditionalFormats", _defaultApiSetName, "1.6", _hostName);
                if (!this.m_conditionalFormats) {
                    this.m_conditionalFormats = new Excel.ConditionalFormatCollection(this.context, _createPropertyObjectPath(this.context, this, "ConditionalFormats", true, false));
                }
                return this.m_conditionalFormats;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "format", {
            get: function () {
                if (!this.m_format) {
                    this.m_format = new Excel.RangeFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
                }
                return this.m_format;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "sort", {
            get: function () {
                _throwIfApiNotSupported("Range.sort", _defaultApiSetName, "1.2", _hostName);
                if (!this.m_sort) {
                    this.m_sort = new Excel.RangeSort(this.context, _createPropertyObjectPath(this.context, this, "Sort", false, false));
                }
                return this.m_sort;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "worksheet", {
            get: function () {
                if (!this.m_worksheet) {
                    this.m_worksheet = new Excel.Worksheet(this.context, _createPropertyObjectPath(this.context, this, "Worksheet", false, false));
                }
                return this.m_worksheet;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "address", {
            get: function () {
                _throwIfNotLoaded("address", this.m_address, "Range", this._isNull);
                return this.m_address;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "addressLocal", {
            get: function () {
                _throwIfNotLoaded("addressLocal", this.m_addressLocal, "Range", this._isNull);
                return this.m_addressLocal;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "cellCount", {
            get: function () {
                _throwIfNotLoaded("cellCount", this.m_cellCount, "Range", this._isNull);
                return this.m_cellCount;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "columnCount", {
            get: function () {
                _throwIfNotLoaded("columnCount", this.m_columnCount, "Range", this._isNull);
                return this.m_columnCount;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "columnHidden", {
            get: function () {
                _throwIfNotLoaded("columnHidden", this.m_columnHidden, "Range", this._isNull);
                _throwIfApiNotSupported("Range.columnHidden", _defaultApiSetName, "1.2", _hostName);
                return this.m_columnHidden;
            },
            set: function (value) {
                this.m_columnHidden = value;
                _createSetPropertyAction(this.context, this, "ColumnHidden", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "columnIndex", {
            get: function () {
                _throwIfNotLoaded("columnIndex", this.m_columnIndex, "Range", this._isNull);
                return this.m_columnIndex;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "formulas", {
            get: function () {
                _throwIfNotLoaded("formulas", this.m_formulas, "Range", this._isNull);
                return this.m_formulas;
            },
            set: function (value) {
                this.m_formulas = value;
                if (setRangePropertiesInBulk(this, "Formulas", value)) {
                    return;
                }
                this.m_formulas = value;
                _createSetPropertyAction(this.context, this, "Formulas", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "formulasLocal", {
            get: function () {
                _throwIfNotLoaded("formulasLocal", this.m_formulasLocal, "Range", this._isNull);
                return this.m_formulasLocal;
            },
            set: function (value) {
                this.m_formulasLocal = value;
                if (setRangePropertiesInBulk(this, "FormulasLocal", value)) {
                    return;
                }
                this.m_formulasLocal = value;
                _createSetPropertyAction(this.context, this, "FormulasLocal", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "formulasR1C1", {
            get: function () {
                _throwIfNotLoaded("formulasR1C1", this.m_formulasR1C1, "Range", this._isNull);
                _throwIfApiNotSupported("Range.formulasR1C1", _defaultApiSetName, "1.2", _hostName);
                return this.m_formulasR1C1;
            },
            set: function (value) {
                this.m_formulasR1C1 = value;
                if (setRangePropertiesInBulk(this, "FormulasR1C1", value)) {
                    return;
                }
                this.m_formulasR1C1 = value;
                _createSetPropertyAction(this.context, this, "FormulasR1C1", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "hidden", {
            get: function () {
                _throwIfNotLoaded("hidden", this.m_hidden, "Range", this._isNull);
                _throwIfApiNotSupported("Range.hidden", _defaultApiSetName, "1.2", _hostName);
                return this.m_hidden;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "hyperlink", {
            get: function () {
                _throwIfNotLoaded("hyperlink", this.m_hyperlink, "Range", this._isNull);
                _throwIfApiNotSupported("Range.hyperlink", _defaultApiSetName, "1.7", _hostName);
                return this.m_hyperlink;
            },
            set: function (value) {
                this.m_hyperlink = value;
                _createSetPropertyAction(this.context, this, "Hyperlink", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "numberFormat", {
            get: function () {
                _throwIfNotLoaded("numberFormat", this.m_numberFormat, "Range", this._isNull);
                return this.m_numberFormat;
            },
            set: function (value) {
                this.m_numberFormat = value;
                if (setRangePropertiesInBulk(this, "NumberFormat", value)) {
                    return;
                }
                this.m_numberFormat = value;
                _createSetPropertyAction(this.context, this, "NumberFormat", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "rowCount", {
            get: function () {
                _throwIfNotLoaded("rowCount", this.m_rowCount, "Range", this._isNull);
                return this.m_rowCount;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "rowHidden", {
            get: function () {
                _throwIfNotLoaded("rowHidden", this.m_rowHidden, "Range", this._isNull);
                _throwIfApiNotSupported("Range.rowHidden", _defaultApiSetName, "1.2", _hostName);
                return this.m_rowHidden;
            },
            set: function (value) {
                this.m_rowHidden = value;
                _createSetPropertyAction(this.context, this, "RowHidden", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "rowIndex", {
            get: function () {
                _throwIfNotLoaded("rowIndex", this.m_rowIndex, "Range", this._isNull);
                return this.m_rowIndex;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "text", {
            get: function () {
                _throwIfNotLoaded("text", this.m_text, "Range", this._isNull);
                return this.m_text;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "valueTypes", {
            get: function () {
                _throwIfNotLoaded("valueTypes", this.m_valueTypes, "Range", this._isNull);
                return this.m_valueTypes;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "values", {
            get: function () {
                _throwIfNotLoaded("values", this.m_values, "Range", this._isNull);
                return this.m_values;
            },
            set: function (value) {
                this.m_values = value;
                if (setRangePropertiesInBulk(this, "Values", value)) {
                    return;
                }
                this.m_values = value;
                _createSetPropertyAction(this.context, this, "Values", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "_ReferenceId", {
            get: function () {
                _throwIfNotLoaded("_ReferenceId", this.m__ReferenceId, "Range", this._isNull);
                return this.m__ReferenceId;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "isEntireColumn", {
            get: function () {
                _throwIfNotLoaded("isEntireColumn", this.m_isEntireColumn, "Range", this._isNull);
                _throwIfApiNotSupported("Range.isEntireColumn", _defaultApiSetName, "1.7", _hostName);
                return this.m_isEntireColumn;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "isEntireRow", {
            get: function () {
                _throwIfNotLoaded("isEntireRow", this.m_isEntireRow, "Range", this._isNull);
                _throwIfApiNotSupported("Range.isEntireRow", _defaultApiSetName, "1.7", _hostName);
                return this.m_isEntireRow;
            },
            enumerable: true,
            configurable: true
        });
        Range.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["numberFormat", "values", "formulas", "formulasLocal", "formulasR1C1", "rowHidden", "columnHidden", "hyperlink"], ["format"], [
                "conditionalFormats",
                "sort",
                "worksheet",
                "conditionalFormats",
                "sort",
                "worksheet"
            ]);
        };
        Range.prototype.calculate = function () {
            _throwIfApiNotSupported("Range.calculate", _defaultApiSetName, "1.6", _hostName);
            _createMethodAction(this.context, this, "Calculate", 0, []);
        };
        Range.prototype.clear = function (applyTo) {
            _createMethodAction(this.context, this, "Clear", 0, [applyTo]);
        };
        Range.prototype.delete = function (shift) {
            _createMethodAction(this.context, this, "Delete", 0, [shift]);
        };
        Range.prototype.getAbsoluteResizedRange = function (numRows, numColumns) {
            _throwIfApiNotSupported("Range.getAbsoluteResizedRange", _defaultApiSetName, "1.7", _hostName);
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetAbsoluteResizedRange", 1, [numRows, numColumns], false, true, null));
        };
        Range.prototype.getBoundingRect = function (anotherRange) {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetBoundingRect", 1, [anotherRange], false, true, null));
        };
        Range.prototype.getCell = function (row, column) {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetCell", 1, [row, column], false, true, null));
        };
        Range.prototype.getColumn = function (column) {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetColumn", 1, [column], false, true, null));
        };
        Range.prototype.getColumnsAfter = function (count) {
            if (!isExcel1_3OrAbove()) {
                if (count == null) {
                    count = 1;
                }
                this._ensureInteger(count, "RowsAbove");
                if (count == 0) {
                    throw new OfficeExtension.Utility.throwError(Excel.ErrorCodes.invalidArgument, "count", "RowsAbove");
                }
                return this._getAdjacentRange("getColumnsAfter", count, this.getLastColumn(), 0, 1);
            }
            _throwIfApiNotSupported("Range.getColumnsAfter", _defaultApiSetName, "1.3", _hostName);
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetColumnsAfter", 1, [count], false, true, null));
        };
        Range.prototype.getColumnsBefore = function (count) {
            if (!isExcel1_3OrAbove()) {
                if (count == null) {
                    count = 1;
                }
                this._ensureInteger(count, "RowsAbove");
                if (count == 0) {
                    throw new OfficeExtension.Utility.throwError(Excel.ErrorCodes.invalidArgument, "count", "RowsAbove");
                }
                return this._getAdjacentRange("getColumnsBefore", count, this.getColumn(0), 0, -1);
            }
            _throwIfApiNotSupported("Range.getColumnsBefore", _defaultApiSetName, "1.3", _hostName);
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetColumnsBefore", 1, [count], false, true, null));
        };
        Range.prototype.getEntireColumn = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetEntireColumn", 1, [], false, true, null));
        };
        Range.prototype.getEntireRow = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetEntireRow", 1, [], false, true, null));
        };
        Range.prototype.getIntersection = function (anotherRange) {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetIntersection", 1, [anotherRange], false, true, null));
        };
        Range.prototype.getIntersectionOrNullObject = function (anotherRange) {
            _throwIfApiNotSupported("Range.getIntersectionOrNullObject", _defaultApiSetName, "1.4", _hostName);
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetIntersectionOrNullObject", 1, [anotherRange], false, true, null));
        };
        Range.prototype.getLastCell = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetLastCell", 1, [], false, true, null));
        };
        Range.prototype.getLastColumn = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetLastColumn", 1, [], false, true, null));
        };
        Range.prototype.getLastRow = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetLastRow", 1, [], false, true, null));
        };
        Range.prototype.getOffsetRange = function (rowOffset, columnOffset) {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetOffsetRange", 1, [rowOffset, columnOffset], false, true, null));
        };
        Range.prototype.getResizedRange = function (deltaRows, deltaColumns) {
            if (!isExcel1_3OrAbove()) {
                this._ensureInteger(deltaRows, "getResizedRange");
                this._ensureInteger(deltaColumns, "getResizedRange");
                var referenceRange = (deltaRows >= 0 && deltaColumns >= 0) ? this : this.getCell(0, 0);
                return referenceRange.getBoundingRect(this.getLastCell().getOffsetRange(deltaRows, deltaColumns));
            }
            _throwIfApiNotSupported("Range.getResizedRange", _defaultApiSetName, "1.3", _hostName);
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetResizedRange", 1, [deltaRows, deltaColumns], false, true, null));
        };
        Range.prototype.getRow = function (row) {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetRow", 1, [row], false, true, null));
        };
        Range.prototype.getRowsAbove = function (count) {
            if (!isExcel1_3OrAbove()) {
                if (count == null) {
                    count = 1;
                }
                this._ensureInteger(count, "RowsAbove");
                if (count == 0) {
                    throw new OfficeExtension.Utility.throwError(Excel.ErrorCodes.invalidArgument, "count", "RowsAbove");
                }
                return this._getAdjacentRange("getRowsAbove", count, this.getRow(0), -1, 0);
            }
            _throwIfApiNotSupported("Range.getRowsAbove", _defaultApiSetName, "1.3", _hostName);
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetRowsAbove", 1, [count], false, true, null));
        };
        Range.prototype.getRowsBelow = function (count) {
            if (!isExcel1_3OrAbove()) {
                if (count == null) {
                    count = 1;
                }
                this._ensureInteger(count, "RowsAbove");
                if (count == 0) {
                    throw new OfficeExtension.Utility.throwError(Excel.ErrorCodes.invalidArgument, "count", "RowsAbove");
                }
                return this._getAdjacentRange("getRowsBelow", count, this.getLastRow(), 1, 0);
            }
            _throwIfApiNotSupported("Range.getRowsBelow", _defaultApiSetName, "1.3", _hostName);
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetRowsBelow", 1, [count], false, true, null));
        };
        Range.prototype.getSurroundingRegion = function () {
            _throwIfApiNotSupported("Range.getSurroundingRegion", _defaultApiSetName, "1.7", _hostName);
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetSurroundingRegion", 1, [], false, true, null));
        };
        Range.prototype.getUsedRange = function (valuesOnly) {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetUsedRange", 1, [valuesOnly], false, true, null));
        };
        Range.prototype.getUsedRangeOrNullObject = function (valuesOnly) {
            _throwIfApiNotSupported("Range.getUsedRangeOrNullObject", _defaultApiSetName, "1.4", _hostName);
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetUsedRangeOrNullObject", 1, [valuesOnly], false, true, null));
        };
        Range.prototype.getVisibleView = function () {
            _throwIfApiNotSupported("Range.getVisibleView", _defaultApiSetName, "1.3", _hostName);
            return new Excel.RangeView(this.context, _createMethodObjectPath(this.context, this, "GetVisibleView", 1, [], false, false, null));
        };
        Range.prototype.insert = function (shift) {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "Insert", 0, [shift], false, true, null));
        };
        Range.prototype.merge = function (across) {
            _throwIfApiNotSupported("Range.merge", _defaultApiSetName, "1.2", _hostName);
            _createMethodAction(this.context, this, "Merge", 0, [across]);
        };
        Range.prototype.select = function () {
            _createMethodAction(this.context, this, "Select", 1, []);
        };
        Range.prototype.showCard = function () {
            _throwIfApiNotSupported("Range.showCard", _defaultApiSetName, "1.8", _hostName);
            _createMethodAction(this.context, this, "ShowCard", 0, []);
        };
        Range.prototype.unmerge = function () {
            _throwIfApiNotSupported("Range.unmerge", _defaultApiSetName, "1.2", _hostName);
            _createMethodAction(this.context, this, "Unmerge", 0, []);
        };
        Range.prototype._KeepReference = function () {
            _createMethodAction(this.context, this, "_KeepReference", 1, []);
        };
        Range.prototype._ValidateArraySize = function (rows, columns) {
            _throwIfApiNotSupported("Range._ValidateArraySize", _defaultApiSetName, "1.3", _hostName);
            _createMethodAction(this.context, this, "_ValidateArraySize", 1, [rows, columns]);
        };
        Range.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Address"])) {
                this.m_address = obj["Address"];
            }
            if (!_isUndefined(obj["AddressLocal"])) {
                this.m_addressLocal = obj["AddressLocal"];
            }
            if (!_isUndefined(obj["CellCount"])) {
                this.m_cellCount = obj["CellCount"];
            }
            if (!_isUndefined(obj["ColumnCount"])) {
                this.m_columnCount = obj["ColumnCount"];
            }
            if (!_isUndefined(obj["ColumnHidden"])) {
                this.m_columnHidden = obj["ColumnHidden"];
            }
            if (!_isUndefined(obj["ColumnIndex"])) {
                this.m_columnIndex = obj["ColumnIndex"];
            }
            if (!_isUndefined(obj["Formulas"])) {
                this.m_formulas = obj["Formulas"];
            }
            if (!_isUndefined(obj["FormulasLocal"])) {
                this.m_formulasLocal = obj["FormulasLocal"];
            }
            if (!_isUndefined(obj["FormulasR1C1"])) {
                this.m_formulasR1C1 = obj["FormulasR1C1"];
            }
            if (!_isUndefined(obj["Hidden"])) {
                this.m_hidden = obj["Hidden"];
            }
            if (!_isUndefined(obj["Hyperlink"])) {
                this.m_hyperlink = obj["Hyperlink"];
            }
            if (!_isUndefined(obj["NumberFormat"])) {
                this.m_numberFormat = obj["NumberFormat"];
            }
            if (!_isUndefined(obj["RowCount"])) {
                this.m_rowCount = obj["RowCount"];
            }
            if (!_isUndefined(obj["RowHidden"])) {
                this.m_rowHidden = obj["RowHidden"];
            }
            if (!_isUndefined(obj["RowIndex"])) {
                this.m_rowIndex = obj["RowIndex"];
            }
            if (!_isUndefined(obj["Text"])) {
                this.m_text = obj["Text"];
            }
            if (!_isUndefined(obj["ValueTypes"])) {
                this.m_valueTypes = obj["ValueTypes"];
            }
            if (!_isUndefined(obj["Values"])) {
                this.m_values = obj["Values"];
            }
            if (!_isUndefined(obj["_ReferenceId"])) {
                this.m__ReferenceId = obj["_ReferenceId"];
            }
            if (!_isUndefined(obj["isEntireColumn"])) {
                this.m_isEntireColumn = obj["isEntireColumn"];
            }
            if (!_isUndefined(obj["isEntireRow"])) {
                this.m_isEntireRow = obj["isEntireRow"];
            }
            _handleNavigationPropertyResults(this, obj, ["conditionalFormats", "ConditionalFormats", "format", "Format", "sort", "Sort", "worksheet", "Worksheet"]);
        };
        Range.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        Range.prototype._handleIdResult = function (value) {
            _super.prototype._handleIdResult.call(this, value);
            if (_isNullOrUndefined(value)) {
                return;
            }
            if (!_isUndefined(value["_ReferenceId"])) {
                this.m__ReferenceId = value["_ReferenceId"];
            }
        };
        Range.prototype.track = function () {
            this.context.trackedObjects.add(this);
            return this;
        };
        Range.prototype.untrack = function () {
            this.context.trackedObjects.remove(this);
            return this;
        };
        Range.prototype.toJSON = function () {
            return {
                "address": this.m_address,
                "addressLocal": this.m_addressLocal,
                "cellCount": this.m_cellCount,
                "columnCount": this.m_columnCount,
                "columnHidden": this.m_columnHidden,
                "columnIndex": this.m_columnIndex,
                "format": this.m_format,
                "formulas": this.m_formulas,
                "formulasLocal": this.m_formulasLocal,
                "formulasR1C1": this.m_formulasR1C1,
                "hidden": this.m_hidden,
                "hyperlink": this.m_hyperlink,
                "isEntireColumn": this.m_isEntireColumn,
                "isEntireRow": this.m_isEntireRow,
                "numberFormat": this.m_numberFormat,
                "rowCount": this.m_rowCount,
                "rowHidden": this.m_rowHidden,
                "rowIndex": this.m_rowIndex,
                "text": this.m_text,
                "values": this.m_values,
                "valueTypes": this.m_valueTypes
            };
        };
        return Range;
    }(OfficeExtension.ClientObject));
    Excel.Range = Range;
    var RangeView = (function (_super) {
        __extends(RangeView, _super);
        function RangeView() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(RangeView.prototype, "_className", {
            get: function () {
                return "RangeView";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeView.prototype, "rows", {
            get: function () {
                if (!this.m_rows) {
                    this.m_rows = new Excel.RangeViewCollection(this.context, _createPropertyObjectPath(this.context, this, "Rows", true, false));
                }
                return this.m_rows;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeView.prototype, "cellAddresses", {
            get: function () {
                _throwIfNotLoaded("cellAddresses", this.m_cellAddresses, "RangeView", this._isNull);
                return this.m_cellAddresses;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeView.prototype, "columnCount", {
            get: function () {
                _throwIfNotLoaded("columnCount", this.m_columnCount, "RangeView", this._isNull);
                return this.m_columnCount;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeView.prototype, "formulas", {
            get: function () {
                _throwIfNotLoaded("formulas", this.m_formulas, "RangeView", this._isNull);
                return this.m_formulas;
            },
            set: function (value) {
                this.m_formulas = value;
                _createSetPropertyAction(this.context, this, "Formulas", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeView.prototype, "formulasLocal", {
            get: function () {
                _throwIfNotLoaded("formulasLocal", this.m_formulasLocal, "RangeView", this._isNull);
                return this.m_formulasLocal;
            },
            set: function (value) {
                this.m_formulasLocal = value;
                _createSetPropertyAction(this.context, this, "FormulasLocal", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeView.prototype, "formulasR1C1", {
            get: function () {
                _throwIfNotLoaded("formulasR1C1", this.m_formulasR1C1, "RangeView", this._isNull);
                return this.m_formulasR1C1;
            },
            set: function (value) {
                this.m_formulasR1C1 = value;
                _createSetPropertyAction(this.context, this, "FormulasR1C1", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeView.prototype, "index", {
            get: function () {
                _throwIfNotLoaded("index", this.m_index, "RangeView", this._isNull);
                return this.m_index;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeView.prototype, "numberFormat", {
            get: function () {
                _throwIfNotLoaded("numberFormat", this.m_numberFormat, "RangeView", this._isNull);
                return this.m_numberFormat;
            },
            set: function (value) {
                this.m_numberFormat = value;
                _createSetPropertyAction(this.context, this, "NumberFormat", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeView.prototype, "rowCount", {
            get: function () {
                _throwIfNotLoaded("rowCount", this.m_rowCount, "RangeView", this._isNull);
                return this.m_rowCount;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeView.prototype, "text", {
            get: function () {
                _throwIfNotLoaded("text", this.m_text, "RangeView", this._isNull);
                return this.m_text;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeView.prototype, "valueTypes", {
            get: function () {
                _throwIfNotLoaded("valueTypes", this.m_valueTypes, "RangeView", this._isNull);
                return this.m_valueTypes;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeView.prototype, "values", {
            get: function () {
                _throwIfNotLoaded("values", this.m_values, "RangeView", this._isNull);
                return this.m_values;
            },
            set: function (value) {
                this.m_values = value;
                _createSetPropertyAction(this.context, this, "Values", value);
            },
            enumerable: true,
            configurable: true
        });
        RangeView.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["numberFormat", "values", "formulas", "formulasLocal", "formulasR1C1"], [], [
                "rows",
                "rows"
            ]);
        };
        RangeView.prototype.getRange = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetRange", 1, [], false, true, null));
        };
        RangeView.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["CellAddresses"])) {
                this.m_cellAddresses = obj["CellAddresses"];
            }
            if (!_isUndefined(obj["ColumnCount"])) {
                this.m_columnCount = obj["ColumnCount"];
            }
            if (!_isUndefined(obj["Formulas"])) {
                this.m_formulas = obj["Formulas"];
            }
            if (!_isUndefined(obj["FormulasLocal"])) {
                this.m_formulasLocal = obj["FormulasLocal"];
            }
            if (!_isUndefined(obj["FormulasR1C1"])) {
                this.m_formulasR1C1 = obj["FormulasR1C1"];
            }
            if (!_isUndefined(obj["Index"])) {
                this.m_index = obj["Index"];
            }
            if (!_isUndefined(obj["NumberFormat"])) {
                this.m_numberFormat = obj["NumberFormat"];
            }
            if (!_isUndefined(obj["RowCount"])) {
                this.m_rowCount = obj["RowCount"];
            }
            if (!_isUndefined(obj["Text"])) {
                this.m_text = obj["Text"];
            }
            if (!_isUndefined(obj["ValueTypes"])) {
                this.m_valueTypes = obj["ValueTypes"];
            }
            if (!_isUndefined(obj["Values"])) {
                this.m_values = obj["Values"];
            }
            _handleNavigationPropertyResults(this, obj, ["rows", "Rows"]);
        };
        RangeView.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        RangeView.prototype.toJSON = function () {
            return {
                "cellAddresses": this.m_cellAddresses,
                "columnCount": this.m_columnCount,
                "formulas": this.m_formulas,
                "formulasLocal": this.m_formulasLocal,
                "formulasR1C1": this.m_formulasR1C1,
                "index": this.m_index,
                "numberFormat": this.m_numberFormat,
                "rowCount": this.m_rowCount,
                "text": this.m_text,
                "values": this.m_values,
                "valueTypes": this.m_valueTypes
            };
        };
        return RangeView;
    }(OfficeExtension.ClientObject));
    Excel.RangeView = RangeView;
    var RangeViewCollection = (function (_super) {
        __extends(RangeViewCollection, _super);
        function RangeViewCollection() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(RangeViewCollection.prototype, "_className", {
            get: function () {
                return "RangeViewCollection";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeViewCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, "RangeViewCollection", this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        RangeViewCollection.prototype.getCount = function () {
            _throwIfApiNotSupported("RangeViewCollection.getCount", _defaultApiSetName, "1.4", _hostName);
            var action = _createMethodAction(this.context, this, "GetCount", 1, []);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        RangeViewCollection.prototype.getItemAt = function (index) {
            return new Excel.RangeView(this.context, _createMethodObjectPath(this.context, this, "GetItemAt", 1, [index], false, false, null));
        };
        RangeViewCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.RangeView(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(false, this.context, this, _data[i], i));
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        RangeViewCollection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        RangeViewCollection.prototype.toJSON = function () {
            return {};
        };
        return RangeViewCollection;
    }(OfficeExtension.ClientObject));
    Excel.RangeViewCollection = RangeViewCollection;
    var SettingCollection = (function (_super) {
        __extends(SettingCollection, _super);
        function SettingCollection() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(SettingCollection.prototype, "_className", {
            get: function () {
                return "SettingCollection";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(SettingCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, "SettingCollection", this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        SettingCollection.prototype.add = function (key, value) {
            value = Setting._replaceDateWithStringDate(value);
            return new Excel.Setting(this.context, _createMethodObjectPath(this.context, this, "Add", 0, [key, value], false, true, null));
        };
        SettingCollection.prototype.getCount = function () {
            var action = _createMethodAction(this.context, this, "GetCount", 1, []);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        SettingCollection.prototype.getItem = function (key) {
            return new Excel.Setting(this.context, _createIndexerObjectPath(this.context, this, [key]));
        };
        SettingCollection.prototype.getItemOrNullObject = function (key) {
            return new Excel.Setting(this.context, _createMethodObjectPath(this.context, this, "GetItemOrNullObject", 1, [key], false, false, null));
        };
        SettingCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.Setting(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        SettingCollection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        Object.defineProperty(SettingCollection.prototype, "onSettingsChanged", {
            get: function () {
                var _this = this;
                if (!this.m_settingsChanged) {
                    this.m_settingsChanged = new OfficeExtension.EventHandlers(this.context, this, "SettingsChanged", {
                        registerFunc: function (handlerCallback) {
                            return _this.context.eventRegistration.register(1, "", handlerCallback);
                        },
                        unregisterFunc: function (handlerCallback) {
                            return _this.context.eventRegistration.unregister(1, "", handlerCallback);
                        },
                        eventArgsTransformFunc: function (args) {
                            return OfficeExtension.Utility._createPromiseFromResult({ settings: _this });
                        }
                    });
                }
                return this.m_settingsChanged;
            },
            enumerable: true,
            configurable: true
        });
        SettingCollection.prototype.toJSON = function () {
            return {};
        };
        return SettingCollection;
    }(OfficeExtension.ClientObject));
    Excel.SettingCollection = SettingCollection;
    var Setting = (function (_super) {
        __extends(Setting, _super);
        function Setting() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(Setting.prototype, "_className", {
            get: function () {
                return "Setting";
            },
            enumerable: true,
            configurable: true
        });
        Setting.replaceStringDateWithDate = function (value) {
            var strValue = JSON.stringify(value);
            value = JSON.parse(strValue, function dateReviver(k, v) {
                var d;
                if (typeof v === 'string' && v && v.length > 6 && v.slice(0, 5) === Setting.DateJSONPrefix && v.slice(-1) === Setting.DateJSONSuffix) {
                    d = new Date(parseInt(v.slice(5, -1)));
                    if (d) {
                        return d;
                    }
                }
                return v;
            });
            return value;
        };
        Setting._replaceDateWithStringDate = function (value) {
            var strValue = JSON.stringify(value, function dateReplacer(k, v) {
                return (this[k] instanceof Date) ? (Setting.DateJSONPrefix + this[k].getTime() + Setting.DateJSONSuffix) : v;
            });
            value = JSON.parse(strValue);
            return value;
        };
        Object.defineProperty(Setting.prototype, "key", {
            get: function () {
                _throwIfNotLoaded("key", this.m_key, "Setting", this._isNull);
                return this.m_key;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Setting.prototype, "value", {
            get: function () {
                _throwIfNotLoaded("value", this.m_value, "Setting", this._isNull);
                return this.m_value;
            },
            set: function (value) {
                if (!_isNullOrUndefined(value)) {
                    this.m_value = value;
                    var newValue = Setting._replaceDateWithStringDate(value);
                    _createSetPropertyAction(this.context, this, "Value", newValue);
                    return;
                }
                this.m_value = value;
                _createSetPropertyAction(this.context, this, "Value", value);
            },
            enumerable: true,
            configurable: true
        });
        Setting.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["value"], [], []);
        };
        Setting.prototype.delete = function () {
            _createMethodAction(this.context, this, "Delete", 0, []);
        };
        Setting.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Key"])) {
                this.m_key = obj["Key"];
            }
            if (!_isUndefined(obj["Value"])) {
                this.m_value = obj["Value"];
                this.m_value = Setting.replaceStringDateWithDate(this.m_value);
            }
        };
        Setting.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        Setting.prototype.toJSON = function () {
            return {
                "key": this.m_key,
                "value": this.m_value
            };
        };
        return Setting;
    }(OfficeExtension.ClientObject));
    Setting.DateJSONPrefix = "Date(";
    Setting.DateJSONSuffix = ")";
    Excel.Setting = Setting;
    var NamedItemCollection = (function (_super) {
        __extends(NamedItemCollection, _super);
        function NamedItemCollection() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(NamedItemCollection.prototype, "_className", {
            get: function () {
                return "NamedItemCollection";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(NamedItemCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, "NamedItemCollection", this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        NamedItemCollection.prototype.add = function (name, reference, comment) {
            _throwIfApiNotSupported("NamedItemCollection.add", _defaultApiSetName, "1.4", _hostName);
            return new Excel.NamedItem(this.context, _createMethodObjectPath(this.context, this, "Add", 0, [name, reference, comment], false, true, null));
        };
        NamedItemCollection.prototype.addFormulaLocal = function (name, formula, comment) {
            _throwIfApiNotSupported("NamedItemCollection.addFormulaLocal", _defaultApiSetName, "1.4", _hostName);
            return new Excel.NamedItem(this.context, _createMethodObjectPath(this.context, this, "AddFormulaLocal", 0, [name, formula, comment], false, false, null));
        };
        NamedItemCollection.prototype.getCount = function () {
            _throwIfApiNotSupported("NamedItemCollection.getCount", _defaultApiSetName, "1.4", _hostName);
            var action = _createMethodAction(this.context, this, "GetCount", 1, []);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        NamedItemCollection.prototype.getItem = function (name) {
            return new Excel.NamedItem(this.context, _createIndexerObjectPath(this.context, this, [name]));
        };
        NamedItemCollection.prototype.getItemOrNullObject = function (name) {
            _throwIfApiNotSupported("NamedItemCollection.getItemOrNullObject", _defaultApiSetName, "1.4", _hostName);
            return new Excel.NamedItem(this.context, _createMethodObjectPath(this.context, this, "GetItemOrNullObject", 1, [name], false, false, null));
        };
        NamedItemCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.NamedItem(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        NamedItemCollection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        NamedItemCollection.prototype.toJSON = function () {
            return {};
        };
        return NamedItemCollection;
    }(OfficeExtension.ClientObject));
    Excel.NamedItemCollection = NamedItemCollection;
    var NamedItem = (function (_super) {
        __extends(NamedItem, _super);
        function NamedItem() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(NamedItem.prototype, "_className", {
            get: function () {
                return "NamedItem";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(NamedItem.prototype, "arrayValues", {
            get: function () {
                _throwIfApiNotSupported("NamedItem.arrayValues", _defaultApiSetName, "1.7", _hostName);
                if (!this.m_arrayValues) {
                    this.m_arrayValues = new Excel.NamedItemArrayValues(this.context, _createPropertyObjectPath(this.context, this, "ArrayValues", false, false));
                }
                return this.m_arrayValues;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(NamedItem.prototype, "worksheet", {
            get: function () {
                _throwIfApiNotSupported("NamedItem.worksheet", _defaultApiSetName, "1.4", _hostName);
                if (!this.m_worksheet) {
                    this.m_worksheet = new Excel.Worksheet(this.context, _createPropertyObjectPath(this.context, this, "Worksheet", false, false));
                }
                return this.m_worksheet;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(NamedItem.prototype, "worksheetOrNullObject", {
            get: function () {
                _throwIfApiNotSupported("NamedItem.worksheetOrNullObject", _defaultApiSetName, "1.4", _hostName);
                if (!this.m_worksheetOrNullObject) {
                    this.m_worksheetOrNullObject = new Excel.Worksheet(this.context, _createPropertyObjectPath(this.context, this, "WorksheetOrNullObject", false, false));
                }
                return this.m_worksheetOrNullObject;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(NamedItem.prototype, "comment", {
            get: function () {
                _throwIfNotLoaded("comment", this.m_comment, "NamedItem", this._isNull);
                _throwIfApiNotSupported("NamedItem.comment", _defaultApiSetName, "1.4", _hostName);
                return this.m_comment;
            },
            set: function (value) {
                this.m_comment = value;
                _createSetPropertyAction(this.context, this, "Comment", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(NamedItem.prototype, "formula", {
            get: function () {
                _throwIfNotLoaded("formula", this.m_formula, "NamedItem", this._isNull);
                _throwIfApiNotSupported("NamedItem.formula", _defaultApiSetName, "1.7", _hostName);
                return this.m_formula;
            },
            set: function (value) {
                this.m_formula = value;
                _createSetPropertyAction(this.context, this, "Formula", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(NamedItem.prototype, "name", {
            get: function () {
                _throwIfNotLoaded("name", this.m_name, "NamedItem", this._isNull);
                return this.m_name;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(NamedItem.prototype, "scope", {
            get: function () {
                _throwIfNotLoaded("scope", this.m_scope, "NamedItem", this._isNull);
                _throwIfApiNotSupported("NamedItem.scope", _defaultApiSetName, "1.4", _hostName);
                return this.m_scope;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(NamedItem.prototype, "type", {
            get: function () {
                _throwIfNotLoaded("type", this.m_type, "NamedItem", this._isNull);
                return this.m_type;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(NamedItem.prototype, "value", {
            get: function () {
                _throwIfNotLoaded("value", this.m_value, "NamedItem", this._isNull);
                return this.m_value;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(NamedItem.prototype, "visible", {
            get: function () {
                _throwIfNotLoaded("visible", this.m_visible, "NamedItem", this._isNull);
                return this.m_visible;
            },
            set: function (value) {
                this.m_visible = value;
                _createSetPropertyAction(this.context, this, "Visible", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(NamedItem.prototype, "_Id", {
            get: function () {
                _throwIfNotLoaded("_Id", this.m__Id, "NamedItem", this._isNull);
                return this.m__Id;
            },
            enumerable: true,
            configurable: true
        });
        NamedItem.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["visible", "comment", "formula"], [], [
                "arrayValues",
                "worksheet",
                "worksheetOrNullObject",
                "arrayValues",
                "worksheet",
                "worksheetOrNullObject"
            ]);
        };
        NamedItem.prototype.delete = function () {
            _throwIfApiNotSupported("NamedItem.delete", _defaultApiSetName, "1.4", _hostName);
            _createMethodAction(this.context, this, "Delete", 0, []);
        };
        NamedItem.prototype.getRange = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetRange", 1, [], false, true, null));
        };
        NamedItem.prototype.getRangeOrNullObject = function () {
            _throwIfApiNotSupported("NamedItem.getRangeOrNullObject", _defaultApiSetName, "1.4", _hostName);
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetRangeOrNullObject", 1, [], false, true, null));
        };
        NamedItem.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Comment"])) {
                this.m_comment = obj["Comment"];
            }
            if (!_isUndefined(obj["Formula"])) {
                this.m_formula = obj["Formula"];
            }
            if (!_isUndefined(obj["Name"])) {
                this.m_name = obj["Name"];
            }
            if (!_isUndefined(obj["Scope"])) {
                this.m_scope = obj["Scope"];
            }
            if (!_isUndefined(obj["Type"])) {
                this.m_type = obj["Type"];
            }
            if (!_isUndefined(obj["Value"])) {
                this.m_value = obj["Value"];
            }
            if (!_isUndefined(obj["Visible"])) {
                this.m_visible = obj["Visible"];
            }
            if (!_isUndefined(obj["_Id"])) {
                this.m__Id = obj["_Id"];
            }
            _handleNavigationPropertyResults(this, obj, ["arrayValues", "ArrayValues", "worksheet", "Worksheet", "worksheetOrNullObject", "WorksheetOrNullObject"]);
        };
        NamedItem.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        NamedItem.prototype._handleIdResult = function (value) {
            _super.prototype._handleIdResult.call(this, value);
            if (_isNullOrUndefined(value)) {
                return;
            }
            if (!_isUndefined(value["_Id"])) {
                this.m__Id = value["_Id"];
            }
        };
        NamedItem.prototype.toJSON = function () {
            return {
                "comment": this.m_comment,
                "formula": this.m_formula,
                "name": this.m_name,
                "scope": this.m_scope,
                "type": this.m_type,
                "value": this.m_value,
                "visible": this.m_visible
            };
        };
        return NamedItem;
    }(OfficeExtension.ClientObject));
    Excel.NamedItem = NamedItem;
    var NamedItemArrayValues = (function (_super) {
        __extends(NamedItemArrayValues, _super);
        function NamedItemArrayValues() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(NamedItemArrayValues.prototype, "_className", {
            get: function () {
                return "NamedItemArrayValues";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(NamedItemArrayValues.prototype, "types", {
            get: function () {
                _throwIfNotLoaded("types", this.m_types, "NamedItemArrayValues", this._isNull);
                return this.m_types;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(NamedItemArrayValues.prototype, "values", {
            get: function () {
                _throwIfNotLoaded("values", this.m_values, "NamedItemArrayValues", this._isNull);
                return this.m_values;
            },
            enumerable: true,
            configurable: true
        });
        NamedItemArrayValues.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Types"])) {
                this.m_types = obj["Types"];
            }
            if (!_isUndefined(obj["Values"])) {
                this.m_values = obj["Values"];
            }
        };
        NamedItemArrayValues.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        NamedItemArrayValues.prototype.toJSON = function () {
            return {
                "types": this.m_types,
                "values": this.m_values
            };
        };
        return NamedItemArrayValues;
    }(OfficeExtension.ClientObject));
    Excel.NamedItemArrayValues = NamedItemArrayValues;
    var Binding = (function (_super) {
        __extends(Binding, _super);
        function Binding() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(Binding.prototype, "_className", {
            get: function () {
                return "Binding";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Binding.prototype, "id", {
            get: function () {
                _throwIfNotLoaded("id", this.m_id, "Binding", this._isNull);
                return this.m_id;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Binding.prototype, "type", {
            get: function () {
                _throwIfNotLoaded("type", this.m_type, "Binding", this._isNull);
                return this.m_type;
            },
            enumerable: true,
            configurable: true
        });
        Binding.prototype.delete = function () {
            _throwIfApiNotSupported("Binding.delete", _defaultApiSetName, "1.3", _hostName);
            _createMethodAction(this.context, this, "Delete", 0, []);
        };
        Binding.prototype.getRange = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetRange", 1, [], false, false, null));
        };
        Binding.prototype.getTable = function () {
            return new Excel.Table(this.context, _createMethodObjectPath(this.context, this, "GetTable", 1, [], false, false, null));
        };
        Binding.prototype.getText = function () {
            var action = _createMethodAction(this.context, this, "GetText", 1, []);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        Binding.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Id"])) {
                this.m_id = obj["Id"];
            }
            if (!_isUndefined(obj["Type"])) {
                this.m_type = obj["Type"];
            }
        };
        Binding.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        Binding.prototype._handleIdResult = function (value) {
            _super.prototype._handleIdResult.call(this, value);
            if (_isNullOrUndefined(value)) {
                return;
            }
            if (!_isUndefined(value["Id"])) {
                this.m_id = value["Id"];
            }
        };
        Object.defineProperty(Binding.prototype, "onDataChanged", {
            get: function () {
                var _this = this;
                _throwIfApiNotSupported("Binding.onDataChanged", _defaultApiSetName, "1.3", _hostName);
                if (!this.m_dataChanged) {
                    this.m_dataChanged = new OfficeExtension.EventHandlers(this.context, this, "DataChanged", {
                        registerFunc: function (handlerCallback) {
                            return _this.context.eventRegistration.register(4, _this.id, handlerCallback);
                        },
                        unregisterFunc: function (handlerCallback) {
                            return _this.context.eventRegistration.unregister(4, _this.id, handlerCallback);
                        },
                        eventArgsTransformFunc: function (args) {
                            var evt = {
                                binding: _this
                            };
                            return OfficeExtension.Utility._createPromiseFromResult(evt);
                        }
                    });
                }
                return this.m_dataChanged;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Binding.prototype, "onSelectionChanged", {
            get: function () {
                var _this = this;
                _throwIfApiNotSupported("Binding.onSelectionChanged", _defaultApiSetName, "1.3", _hostName);
                if (!this.m_selectionChanged) {
                    this.m_selectionChanged = new OfficeExtension.EventHandlers(this.context, this, "SelectionChanged", {
                        registerFunc: function (handlerCallback) {
                            return _this.context.eventRegistration.register(3, _this.id, handlerCallback);
                        },
                        unregisterFunc: function (handlerCallback) {
                            return _this.context.eventRegistration.unregister(3, _this.id, handlerCallback);
                        },
                        eventArgsTransformFunc: function (args) {
                            var evt = {
                                binding: _this,
                                columnCount: args.columnCount,
                                rowCount: args.rowCount,
                                startColumn: args.startColumn,
                                startRow: args.startRow
                            };
                            return OfficeExtension.Utility._createPromiseFromResult(evt);
                        }
                    });
                }
                return this.m_selectionChanged;
            },
            enumerable: true,
            configurable: true
        });
        Binding.prototype.toJSON = function () {
            return {
                "id": this.m_id,
                "type": this.m_type
            };
        };
        return Binding;
    }(OfficeExtension.ClientObject));
    Excel.Binding = Binding;
    var BindingCollection = (function (_super) {
        __extends(BindingCollection, _super);
        function BindingCollection() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(BindingCollection.prototype, "_className", {
            get: function () {
                return "BindingCollection";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(BindingCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, "BindingCollection", this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(BindingCollection.prototype, "count", {
            get: function () {
                _throwIfNotLoaded("count", this.m_count, "BindingCollection", this._isNull);
                return this.m_count;
            },
            enumerable: true,
            configurable: true
        });
        BindingCollection.prototype.add = function (range, bindingType, id) {
            _throwIfApiNotSupported("BindingCollection.add", _defaultApiSetName, "1.3", _hostName);
            return new Excel.Binding(this.context, _createMethodObjectPath(this.context, this, "Add", 0, [range, bindingType, id], false, true, null));
        };
        BindingCollection.prototype.addFromNamedItem = function (name, bindingType, id) {
            _throwIfApiNotSupported("BindingCollection.addFromNamedItem", _defaultApiSetName, "1.3", _hostName);
            return new Excel.Binding(this.context, _createMethodObjectPath(this.context, this, "AddFromNamedItem", 0, [name, bindingType, id], false, false, null));
        };
        BindingCollection.prototype.addFromSelection = function (bindingType, id) {
            _throwIfApiNotSupported("BindingCollection.addFromSelection", _defaultApiSetName, "1.3", _hostName);
            return new Excel.Binding(this.context, _createMethodObjectPath(this.context, this, "AddFromSelection", 0, [bindingType, id], false, false, null));
        };
        BindingCollection.prototype.getCount = function () {
            _throwIfApiNotSupported("BindingCollection.getCount", _defaultApiSetName, "1.4", _hostName);
            var action = _createMethodAction(this.context, this, "GetCount", 1, []);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        BindingCollection.prototype.getItem = function (id) {
            return new Excel.Binding(this.context, _createIndexerObjectPath(this.context, this, [id]));
        };
        BindingCollection.prototype.getItemAt = function (index) {
            return new Excel.Binding(this.context, _createMethodObjectPath(this.context, this, "GetItemAt", 1, [index], false, false, null));
        };
        BindingCollection.prototype.getItemOrNullObject = function (id) {
            _throwIfApiNotSupported("BindingCollection.getItemOrNullObject", _defaultApiSetName, "1.4", _hostName);
            return new Excel.Binding(this.context, _createMethodObjectPath(this.context, this, "GetItemOrNullObject", 1, [id], false, false, null));
        };
        BindingCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Count"])) {
                this.m_count = obj["Count"];
            }
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.Binding(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        BindingCollection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        BindingCollection.prototype.toJSON = function () {
            return {
                "count": this.m_count
            };
        };
        return BindingCollection;
    }(OfficeExtension.ClientObject));
    Excel.BindingCollection = BindingCollection;
    var TableCollection = (function (_super) {
        __extends(TableCollection, _super);
        function TableCollection() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(TableCollection.prototype, "_className", {
            get: function () {
                return "TableCollection";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(TableCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, "TableCollection", this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(TableCollection.prototype, "count", {
            get: function () {
                _throwIfNotLoaded("count", this.m_count, "TableCollection", this._isNull);
                return this.m_count;
            },
            enumerable: true,
            configurable: true
        });
        TableCollection.prototype.add = function (address, hasHeaders) {
            return new Excel.Table(this.context, _createMethodObjectPath(this.context, this, "Add", 0, [address, hasHeaders], false, true, null));
        };
        TableCollection.prototype.getCount = function () {
            _throwIfApiNotSupported("TableCollection.getCount", _defaultApiSetName, "1.4", _hostName);
            var action = _createMethodAction(this.context, this, "GetCount", 1, []);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        TableCollection.prototype.getItem = function (key) {
            return new Excel.Table(this.context, _createIndexerObjectPath(this.context, this, [key]));
        };
        TableCollection.prototype.getItemAt = function (index) {
            return new Excel.Table(this.context, _createMethodObjectPath(this.context, this, "GetItemAt", 1, [index], false, false, null));
        };
        TableCollection.prototype.getItemOrNullObject = function (key) {
            _throwIfApiNotSupported("TableCollection.getItemOrNullObject", _defaultApiSetName, "1.4", _hostName);
            return new Excel.Table(this.context, _createMethodObjectPath(this.context, this, "GetItemOrNullObject", 1, [key], false, false, null));
        };
        TableCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Count"])) {
                this.m_count = obj["Count"];
            }
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.Table(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        TableCollection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        Object.defineProperty(TableCollection.prototype, "onDataChanged", {
            get: function () {
                _throwIfApiNotSupported("TableCollection.onDataChanged", _defaultApiSetName, "1.8", _hostName);
                if (!this.m_dataChanged) {
                    this.m_dataChanged = new OfficeExtension.GenericEventHandlers(this.context, this, "DataChanged", {
                        eventType: 0,
                        registerFunc: function () { return null; },
                        unregisterFunc: function () { return null; },
                        getTargetIdFunc: null,
                        eventArgsTransformFunc: function (value) {
                            return null;
                        }
                    });
                }
                return this.m_dataChanged;
            },
            enumerable: true,
            configurable: true
        });
        TableCollection.prototype.toJSON = function () {
            return {
                "count": this.m_count
            };
        };
        return TableCollection;
    }(OfficeExtension.ClientObject));
    Excel.TableCollection = TableCollection;
    var Table = (function (_super) {
        __extends(Table, _super);
        function Table() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(Table.prototype, "_className", {
            get: function () {
                return "Table";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Table.prototype, "columns", {
            get: function () {
                if (!this.m_columns) {
                    this.m_columns = new Excel.TableColumnCollection(this.context, _createPropertyObjectPath(this.context, this, "Columns", true, false));
                }
                return this.m_columns;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Table.prototype, "rows", {
            get: function () {
                if (!this.m_rows) {
                    this.m_rows = new Excel.TableRowCollection(this.context, _createPropertyObjectPath(this.context, this, "Rows", true, false));
                }
                return this.m_rows;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Table.prototype, "sort", {
            get: function () {
                _throwIfApiNotSupported("Table.sort", _defaultApiSetName, "1.2", _hostName);
                if (!this.m_sort) {
                    this.m_sort = new Excel.TableSort(this.context, _createPropertyObjectPath(this.context, this, "Sort", false, false));
                }
                return this.m_sort;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Table.prototype, "worksheet", {
            get: function () {
                _throwIfApiNotSupported("Table.worksheet", _defaultApiSetName, "1.2", _hostName);
                if (!this.m_worksheet) {
                    this.m_worksheet = new Excel.Worksheet(this.context, _createPropertyObjectPath(this.context, this, "Worksheet", false, false));
                }
                return this.m_worksheet;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Table.prototype, "highlightFirstColumn", {
            get: function () {
                _throwIfNotLoaded("highlightFirstColumn", this.m_highlightFirstColumn, "Table", this._isNull);
                _throwIfApiNotSupported("Table.highlightFirstColumn", _defaultApiSetName, "1.3", _hostName);
                return this.m_highlightFirstColumn;
            },
            set: function (value) {
                this.m_highlightFirstColumn = value;
                _createSetPropertyAction(this.context, this, "HighlightFirstColumn", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Table.prototype, "highlightLastColumn", {
            get: function () {
                _throwIfNotLoaded("highlightLastColumn", this.m_highlightLastColumn, "Table", this._isNull);
                _throwIfApiNotSupported("Table.highlightLastColumn", _defaultApiSetName, "1.3", _hostName);
                return this.m_highlightLastColumn;
            },
            set: function (value) {
                this.m_highlightLastColumn = value;
                _createSetPropertyAction(this.context, this, "HighlightLastColumn", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Table.prototype, "id", {
            get: function () {
                _throwIfNotLoaded("id", this.m_id, "Table", this._isNull);
                return this.m_id;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Table.prototype, "name", {
            get: function () {
                _throwIfNotLoaded("name", this.m_name, "Table", this._isNull);
                return this.m_name;
            },
            set: function (value) {
                this.m_name = value;
                _createSetPropertyAction(this.context, this, "Name", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Table.prototype, "showBandedColumns", {
            get: function () {
                _throwIfNotLoaded("showBandedColumns", this.m_showBandedColumns, "Table", this._isNull);
                _throwIfApiNotSupported("Table.showBandedColumns", _defaultApiSetName, "1.3", _hostName);
                return this.m_showBandedColumns;
            },
            set: function (value) {
                this.m_showBandedColumns = value;
                _createSetPropertyAction(this.context, this, "ShowBandedColumns", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Table.prototype, "showBandedRows", {
            get: function () {
                _throwIfNotLoaded("showBandedRows", this.m_showBandedRows, "Table", this._isNull);
                _throwIfApiNotSupported("Table.showBandedRows", _defaultApiSetName, "1.3", _hostName);
                return this.m_showBandedRows;
            },
            set: function (value) {
                this.m_showBandedRows = value;
                _createSetPropertyAction(this.context, this, "ShowBandedRows", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Table.prototype, "showFilterButton", {
            get: function () {
                _throwIfNotLoaded("showFilterButton", this.m_showFilterButton, "Table", this._isNull);
                _throwIfApiNotSupported("Table.showFilterButton", _defaultApiSetName, "1.3", _hostName);
                return this.m_showFilterButton;
            },
            set: function (value) {
                this.m_showFilterButton = value;
                _createSetPropertyAction(this.context, this, "ShowFilterButton", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Table.prototype, "showHeaders", {
            get: function () {
                _throwIfNotLoaded("showHeaders", this.m_showHeaders, "Table", this._isNull);
                return this.m_showHeaders;
            },
            set: function (value) {
                this.m_showHeaders = value;
                _createSetPropertyAction(this.context, this, "ShowHeaders", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Table.prototype, "showTotals", {
            get: function () {
                _throwIfNotLoaded("showTotals", this.m_showTotals, "Table", this._isNull);
                return this.m_showTotals;
            },
            set: function (value) {
                this.m_showTotals = value;
                _createSetPropertyAction(this.context, this, "ShowTotals", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Table.prototype, "style", {
            get: function () {
                _throwIfNotLoaded("style", this.m_style, "Table", this._isNull);
                return this.m_style;
            },
            set: function (value) {
                this.m_style = value;
                _createSetPropertyAction(this.context, this, "Style", value);
            },
            enumerable: true,
            configurable: true
        });
        Table.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["name", "showHeaders", "showTotals", "style", "highlightFirstColumn", "highlightLastColumn", "showBandedRows", "showBandedColumns", "showFilterButton"], [], [
                "columns",
                "rows",
                "sort",
                "worksheet",
                "columns",
                "rows",
                "sort",
                "worksheet"
            ]);
        };
        Table.prototype.clearFilters = function () {
            _throwIfApiNotSupported("Table.clearFilters", _defaultApiSetName, "1.2", _hostName);
            _createMethodAction(this.context, this, "ClearFilters", 0, []);
        };
        Table.prototype.convertToRange = function () {
            _throwIfApiNotSupported("Table.convertToRange", _defaultApiSetName, "1.2", _hostName);
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "ConvertToRange", 0, [], false, true, null));
        };
        Table.prototype.delete = function () {
            _createMethodAction(this.context, this, "Delete", 0, []);
        };
        Table.prototype.getDataBodyRange = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetDataBodyRange", 1, [], false, true, null));
        };
        Table.prototype.getHeaderRowRange = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetHeaderRowRange", 1, [], false, true, null));
        };
        Table.prototype.getRange = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetRange", 1, [], false, true, null));
        };
        Table.prototype.getTotalRowRange = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetTotalRowRange", 1, [], false, true, null));
        };
        Table.prototype.reapplyFilters = function () {
            _throwIfApiNotSupported("Table.reapplyFilters", _defaultApiSetName, "1.2", _hostName);
            _createMethodAction(this.context, this, "ReapplyFilters", 0, []);
        };
        Table.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["HighlightFirstColumn"])) {
                this.m_highlightFirstColumn = obj["HighlightFirstColumn"];
            }
            if (!_isUndefined(obj["HighlightLastColumn"])) {
                this.m_highlightLastColumn = obj["HighlightLastColumn"];
            }
            if (!_isUndefined(obj["Id"])) {
                this.m_id = obj["Id"];
            }
            if (!_isUndefined(obj["Name"])) {
                this.m_name = obj["Name"];
            }
            if (!_isUndefined(obj["ShowBandedColumns"])) {
                this.m_showBandedColumns = obj["ShowBandedColumns"];
            }
            if (!_isUndefined(obj["ShowBandedRows"])) {
                this.m_showBandedRows = obj["ShowBandedRows"];
            }
            if (!_isUndefined(obj["ShowFilterButton"])) {
                this.m_showFilterButton = obj["ShowFilterButton"];
            }
            if (!_isUndefined(obj["ShowHeaders"])) {
                this.m_showHeaders = obj["ShowHeaders"];
            }
            if (!_isUndefined(obj["ShowTotals"])) {
                this.m_showTotals = obj["ShowTotals"];
            }
            if (!_isUndefined(obj["Style"])) {
                this.m_style = obj["Style"];
            }
            _handleNavigationPropertyResults(this, obj, ["columns", "Columns", "rows", "Rows", "sort", "Sort", "worksheet", "Worksheet"]);
        };
        Table.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        Table.prototype._handleIdResult = function (value) {
            _super.prototype._handleIdResult.call(this, value);
            if (_isNullOrUndefined(value)) {
                return;
            }
            if (!_isUndefined(value["Id"])) {
                this.m_id = value["Id"];
            }
        };
        Object.defineProperty(Table.prototype, "onDataChanged", {
            get: function () {
                _throwIfApiNotSupported("Table.onDataChanged", _defaultApiSetName, "1.8", _hostName);
                if (!this.m_dataChanged) {
                    this.m_dataChanged = new OfficeExtension.GenericEventHandlers(this.context, this, "DataChanged", {
                        eventType: 0,
                        registerFunc: function () { return null; },
                        unregisterFunc: function () { return null; },
                        getTargetIdFunc: null,
                        eventArgsTransformFunc: function (value) {
                            return null;
                        }
                    });
                }
                return this.m_dataChanged;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Table.prototype, "onSelectionChanged", {
            get: function () {
                _throwIfApiNotSupported("Table.onSelectionChanged", _defaultApiSetName, "1.8", _hostName);
                if (!this.m_selectionChanged) {
                    this.m_selectionChanged = new OfficeExtension.GenericEventHandlers(this.context, this, "SelectionChanged", {
                        eventType: 0,
                        registerFunc: function () { return null; },
                        unregisterFunc: function () { return null; },
                        getTargetIdFunc: null,
                        eventArgsTransformFunc: function (value) {
                            return null;
                        }
                    });
                }
                return this.m_selectionChanged;
            },
            enumerable: true,
            configurable: true
        });
        Table.prototype.toJSON = function () {
            return {
                "highlightFirstColumn": this.m_highlightFirstColumn,
                "highlightLastColumn": this.m_highlightLastColumn,
                "id": this.m_id,
                "name": this.m_name,
                "showBandedColumns": this.m_showBandedColumns,
                "showBandedRows": this.m_showBandedRows,
                "showFilterButton": this.m_showFilterButton,
                "showHeaders": this.m_showHeaders,
                "showTotals": this.m_showTotals,
                "style": this.m_style
            };
        };
        return Table;
    }(OfficeExtension.ClientObject));
    Excel.Table = Table;
    var TableColumnCollection = (function (_super) {
        __extends(TableColumnCollection, _super);
        function TableColumnCollection() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(TableColumnCollection.prototype, "_className", {
            get: function () {
                return "TableColumnCollection";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(TableColumnCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, "TableColumnCollection", this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(TableColumnCollection.prototype, "count", {
            get: function () {
                _throwIfNotLoaded("count", this.m_count, "TableColumnCollection", this._isNull);
                return this.m_count;
            },
            enumerable: true,
            configurable: true
        });
        TableColumnCollection.prototype.add = function (index, values, name) {
            return new Excel.TableColumn(this.context, _createMethodObjectPath(this.context, this, "Add", 0, [index, values, name], false, true, null));
        };
        TableColumnCollection.prototype.getCount = function () {
            _throwIfApiNotSupported("TableColumnCollection.getCount", _defaultApiSetName, "1.4", _hostName);
            var action = _createMethodAction(this.context, this, "GetCount", 1, []);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        TableColumnCollection.prototype.getItem = function (key) {
            return new Excel.TableColumn(this.context, _createIndexerObjectPath(this.context, this, [key]));
        };
        TableColumnCollection.prototype.getItemAt = function (index) {
            return new Excel.TableColumn(this.context, _createMethodObjectPath(this.context, this, "GetItemAt", 1, [index], false, false, null));
        };
        TableColumnCollection.prototype.getItemOrNullObject = function (key) {
            _throwIfApiNotSupported("TableColumnCollection.getItemOrNullObject", _defaultApiSetName, "1.4", _hostName);
            return new Excel.TableColumn(this.context, _createMethodObjectPath(this.context, this, "GetItemOrNullObject", 1, [key], false, false, null));
        };
        TableColumnCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Count"])) {
                this.m_count = obj["Count"];
            }
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.TableColumn(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        TableColumnCollection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        TableColumnCollection.prototype.toJSON = function () {
            return {
                "count": this.m_count
            };
        };
        return TableColumnCollection;
    }(OfficeExtension.ClientObject));
    Excel.TableColumnCollection = TableColumnCollection;
    var TableColumn = (function (_super) {
        __extends(TableColumn, _super);
        function TableColumn() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(TableColumn.prototype, "_className", {
            get: function () {
                return "TableColumn";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(TableColumn.prototype, "filter", {
            get: function () {
                _throwIfApiNotSupported("TableColumn.filter", _defaultApiSetName, "1.2", _hostName);
                if (!this.m_filter) {
                    this.m_filter = new Excel.Filter(this.context, _createPropertyObjectPath(this.context, this, "Filter", false, false));
                }
                return this.m_filter;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(TableColumn.prototype, "id", {
            get: function () {
                _throwIfNotLoaded("id", this.m_id, "TableColumn", this._isNull);
                return this.m_id;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(TableColumn.prototype, "index", {
            get: function () {
                _throwIfNotLoaded("index", this.m_index, "TableColumn", this._isNull);
                return this.m_index;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(TableColumn.prototype, "name", {
            get: function () {
                _throwIfNotLoaded("name", this.m_name, "TableColumn", this._isNull);
                return this.m_name;
            },
            set: function (value) {
                this.m_name = value;
                _createSetPropertyAction(this.context, this, "Name", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(TableColumn.prototype, "values", {
            get: function () {
                _throwIfNotLoaded("values", this.m_values, "TableColumn", this._isNull);
                return this.m_values;
            },
            set: function (value) {
                this.m_values = value;
                _createSetPropertyAction(this.context, this, "Values", value);
            },
            enumerable: true,
            configurable: true
        });
        TableColumn.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["values", "name"], [], [
                "filter",
                "filter"
            ]);
        };
        TableColumn.prototype.delete = function () {
            _createMethodAction(this.context, this, "Delete", 0, []);
        };
        TableColumn.prototype.getDataBodyRange = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetDataBodyRange", 1, [], false, true, null));
        };
        TableColumn.prototype.getHeaderRowRange = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetHeaderRowRange", 1, [], false, true, null));
        };
        TableColumn.prototype.getRange = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetRange", 1, [], false, true, null));
        };
        TableColumn.prototype.getTotalRowRange = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetTotalRowRange", 1, [], false, true, null));
        };
        TableColumn.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Id"])) {
                this.m_id = obj["Id"];
            }
            if (!_isUndefined(obj["Index"])) {
                this.m_index = obj["Index"];
            }
            if (!_isUndefined(obj["Name"])) {
                this.m_name = obj["Name"];
            }
            if (!_isUndefined(obj["Values"])) {
                this.m_values = obj["Values"];
            }
            _handleNavigationPropertyResults(this, obj, ["filter", "Filter"]);
        };
        TableColumn.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        TableColumn.prototype._handleIdResult = function (value) {
            _super.prototype._handleIdResult.call(this, value);
            if (_isNullOrUndefined(value)) {
                return;
            }
            if (!_isUndefined(value["Id"])) {
                this.m_id = value["Id"];
            }
        };
        TableColumn.prototype.toJSON = function () {
            return {
                "id": this.m_id,
                "index": this.m_index,
                "name": this.m_name,
                "values": this.m_values
            };
        };
        return TableColumn;
    }(OfficeExtension.ClientObject));
    Excel.TableColumn = TableColumn;
    var TableRowCollection = (function (_super) {
        __extends(TableRowCollection, _super);
        function TableRowCollection() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(TableRowCollection.prototype, "_className", {
            get: function () {
                return "TableRowCollection";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(TableRowCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, "TableRowCollection", this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(TableRowCollection.prototype, "count", {
            get: function () {
                _throwIfNotLoaded("count", this.m_count, "TableRowCollection", this._isNull);
                return this.m_count;
            },
            enumerable: true,
            configurable: true
        });
        TableRowCollection.prototype.add = function (index, values) {
            return new Excel.TableRow(this.context, _createMethodObjectPath(this.context, this, "Add", 0, [index, values], false, true, null));
        };
        TableRowCollection.prototype.getCount = function () {
            _throwIfApiNotSupported("TableRowCollection.getCount", _defaultApiSetName, "1.4", _hostName);
            var action = _createMethodAction(this.context, this, "GetCount", 1, []);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        TableRowCollection.prototype.getItemAt = function (index) {
            return new Excel.TableRow(this.context, _createMethodObjectPath(this.context, this, "GetItemAt", 1, [index], false, false, null));
        };
        TableRowCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Count"])) {
                this.m_count = obj["Count"];
            }
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.TableRow(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(false, this.context, this, _data[i], i));
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        TableRowCollection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        TableRowCollection.prototype.toJSON = function () {
            return {
                "count": this.m_count
            };
        };
        return TableRowCollection;
    }(OfficeExtension.ClientObject));
    Excel.TableRowCollection = TableRowCollection;
    var TableRow = (function (_super) {
        __extends(TableRow, _super);
        function TableRow() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(TableRow.prototype, "_className", {
            get: function () {
                return "TableRow";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(TableRow.prototype, "index", {
            get: function () {
                _throwIfNotLoaded("index", this.m_index, "TableRow", this._isNull);
                return this.m_index;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(TableRow.prototype, "values", {
            get: function () {
                _throwIfNotLoaded("values", this.m_values, "TableRow", this._isNull);
                return this.m_values;
            },
            set: function (value) {
                this.m_values = value;
                _createSetPropertyAction(this.context, this, "Values", value);
            },
            enumerable: true,
            configurable: true
        });
        TableRow.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["values"], [], []);
        };
        TableRow.prototype.delete = function () {
            _createMethodAction(this.context, this, "Delete", 0, []);
        };
        TableRow.prototype.getRange = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetRange", 1, [], false, true, null));
        };
        TableRow.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Index"])) {
                this.m_index = obj["Index"];
            }
            if (!_isUndefined(obj["Values"])) {
                this.m_values = obj["Values"];
            }
        };
        TableRow.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        TableRow.prototype.toJSON = function () {
            return {
                "index": this.m_index,
                "values": this.m_values
            };
        };
        return TableRow;
    }(OfficeExtension.ClientObject));
    Excel.TableRow = TableRow;
    var RangeFormat = (function (_super) {
        __extends(RangeFormat, _super);
        function RangeFormat() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(RangeFormat.prototype, "_className", {
            get: function () {
                return "RangeFormat";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeFormat.prototype, "borders", {
            get: function () {
                if (!this.m_borders) {
                    this.m_borders = new Excel.RangeBorderCollection(this.context, _createPropertyObjectPath(this.context, this, "Borders", true, false));
                }
                return this.m_borders;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeFormat.prototype, "fill", {
            get: function () {
                if (!this.m_fill) {
                    this.m_fill = new Excel.RangeFill(this.context, _createPropertyObjectPath(this.context, this, "Fill", false, false));
                }
                return this.m_fill;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeFormat.prototype, "font", {
            get: function () {
                if (!this.m_font) {
                    this.m_font = new Excel.RangeFont(this.context, _createPropertyObjectPath(this.context, this, "Font", false, false));
                }
                return this.m_font;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeFormat.prototype, "protection", {
            get: function () {
                _throwIfApiNotSupported("RangeFormat.protection", _defaultApiSetName, "1.2", _hostName);
                if (!this.m_protection) {
                    this.m_protection = new Excel.FormatProtection(this.context, _createPropertyObjectPath(this.context, this, "Protection", false, false));
                }
                return this.m_protection;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeFormat.prototype, "columnWidth", {
            get: function () {
                _throwIfNotLoaded("columnWidth", this.m_columnWidth, "RangeFormat", this._isNull);
                _throwIfApiNotSupported("RangeFormat.columnWidth", _defaultApiSetName, "1.2", _hostName);
                return this.m_columnWidth;
            },
            set: function (value) {
                this.m_columnWidth = value;
                _createSetPropertyAction(this.context, this, "ColumnWidth", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeFormat.prototype, "horizontalAlignment", {
            get: function () {
                _throwIfNotLoaded("horizontalAlignment", this.m_horizontalAlignment, "RangeFormat", this._isNull);
                return this.m_horizontalAlignment;
            },
            set: function (value) {
                this.m_horizontalAlignment = value;
                _createSetPropertyAction(this.context, this, "HorizontalAlignment", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeFormat.prototype, "rowHeight", {
            get: function () {
                _throwIfNotLoaded("rowHeight", this.m_rowHeight, "RangeFormat", this._isNull);
                _throwIfApiNotSupported("RangeFormat.rowHeight", _defaultApiSetName, "1.2", _hostName);
                return this.m_rowHeight;
            },
            set: function (value) {
                this.m_rowHeight = value;
                _createSetPropertyAction(this.context, this, "RowHeight", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeFormat.prototype, "textOrientation", {
            get: function () {
                _throwIfNotLoaded("textOrientation", this.m_textOrientation, "RangeFormat", this._isNull);
                _throwIfApiNotSupported("RangeFormat.textOrientation", _defaultApiSetName, "1.7", _hostName);
                return this.m_textOrientation;
            },
            set: function (value) {
                this.m_textOrientation = value;
                _createSetPropertyAction(this.context, this, "TextOrientation", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeFormat.prototype, "verticalAlignment", {
            get: function () {
                _throwIfNotLoaded("verticalAlignment", this.m_verticalAlignment, "RangeFormat", this._isNull);
                return this.m_verticalAlignment;
            },
            set: function (value) {
                this.m_verticalAlignment = value;
                _createSetPropertyAction(this.context, this, "VerticalAlignment", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeFormat.prototype, "wrapText", {
            get: function () {
                _throwIfNotLoaded("wrapText", this.m_wrapText, "RangeFormat", this._isNull);
                return this.m_wrapText;
            },
            set: function (value) {
                this.m_wrapText = value;
                _createSetPropertyAction(this.context, this, "WrapText", value);
            },
            enumerable: true,
            configurable: true
        });
        RangeFormat.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["wrapText", "horizontalAlignment", "verticalAlignment", "columnWidth", "rowHeight", "textOrientation"], ["fill", "font", "protection"], [
                "borders",
                "borders"
            ]);
        };
        RangeFormat.prototype.autofitColumns = function () {
            _throwIfApiNotSupported("RangeFormat.autofitColumns", _defaultApiSetName, "1.2", _hostName);
            _createMethodAction(this.context, this, "AutofitColumns", 0, []);
        };
        RangeFormat.prototype.autofitRows = function () {
            _throwIfApiNotSupported("RangeFormat.autofitRows", _defaultApiSetName, "1.2", _hostName);
            _createMethodAction(this.context, this, "AutofitRows", 0, []);
        };
        RangeFormat.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["ColumnWidth"])) {
                this.m_columnWidth = obj["ColumnWidth"];
            }
            if (!_isUndefined(obj["HorizontalAlignment"])) {
                this.m_horizontalAlignment = obj["HorizontalAlignment"];
            }
            if (!_isUndefined(obj["RowHeight"])) {
                this.m_rowHeight = obj["RowHeight"];
            }
            if (!_isUndefined(obj["TextOrientation"])) {
                this.m_textOrientation = obj["TextOrientation"];
            }
            if (!_isUndefined(obj["VerticalAlignment"])) {
                this.m_verticalAlignment = obj["VerticalAlignment"];
            }
            if (!_isUndefined(obj["WrapText"])) {
                this.m_wrapText = obj["WrapText"];
            }
            _handleNavigationPropertyResults(this, obj, ["borders", "Borders", "fill", "Fill", "font", "Font", "protection", "Protection"]);
        };
        RangeFormat.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        RangeFormat.prototype.toJSON = function () {
            return {
                "columnWidth": this.m_columnWidth,
                "fill": this.m_fill,
                "font": this.m_font,
                "horizontalAlignment": this.m_horizontalAlignment,
                "protection": this.m_protection,
                "rowHeight": this.m_rowHeight,
                "textOrientation": this.m_textOrientation,
                "verticalAlignment": this.m_verticalAlignment,
                "wrapText": this.m_wrapText
            };
        };
        return RangeFormat;
    }(OfficeExtension.ClientObject));
    Excel.RangeFormat = RangeFormat;
    var FormatProtection = (function (_super) {
        __extends(FormatProtection, _super);
        function FormatProtection() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(FormatProtection.prototype, "_className", {
            get: function () {
                return "FormatProtection";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(FormatProtection.prototype, "formulaHidden", {
            get: function () {
                _throwIfNotLoaded("formulaHidden", this.m_formulaHidden, "FormatProtection", this._isNull);
                return this.m_formulaHidden;
            },
            set: function (value) {
                this.m_formulaHidden = value;
                _createSetPropertyAction(this.context, this, "FormulaHidden", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(FormatProtection.prototype, "locked", {
            get: function () {
                _throwIfNotLoaded("locked", this.m_locked, "FormatProtection", this._isNull);
                return this.m_locked;
            },
            set: function (value) {
                this.m_locked = value;
                _createSetPropertyAction(this.context, this, "Locked", value);
            },
            enumerable: true,
            configurable: true
        });
        FormatProtection.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["locked", "formulaHidden"], [], []);
        };
        FormatProtection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["FormulaHidden"])) {
                this.m_formulaHidden = obj["FormulaHidden"];
            }
            if (!_isUndefined(obj["Locked"])) {
                this.m_locked = obj["Locked"];
            }
        };
        FormatProtection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        FormatProtection.prototype.toJSON = function () {
            return {
                "formulaHidden": this.m_formulaHidden,
                "locked": this.m_locked
            };
        };
        return FormatProtection;
    }(OfficeExtension.ClientObject));
    Excel.FormatProtection = FormatProtection;
    var RangeFill = (function (_super) {
        __extends(RangeFill, _super);
        function RangeFill() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(RangeFill.prototype, "_className", {
            get: function () {
                return "RangeFill";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeFill.prototype, "color", {
            get: function () {
                _throwIfNotLoaded("color", this.m_color, "RangeFill", this._isNull);
                return this.m_color;
            },
            set: function (value) {
                this.m_color = value;
                _createSetPropertyAction(this.context, this, "Color", value);
            },
            enumerable: true,
            configurable: true
        });
        RangeFill.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["color"], [], []);
        };
        RangeFill.prototype.clear = function () {
            _createMethodAction(this.context, this, "Clear", 0, []);
        };
        RangeFill.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Color"])) {
                this.m_color = obj["Color"];
            }
        };
        RangeFill.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        RangeFill.prototype.toJSON = function () {
            return {
                "color": this.m_color
            };
        };
        return RangeFill;
    }(OfficeExtension.ClientObject));
    Excel.RangeFill = RangeFill;
    var RangeBorder = (function (_super) {
        __extends(RangeBorder, _super);
        function RangeBorder() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(RangeBorder.prototype, "_className", {
            get: function () {
                return "RangeBorder";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeBorder.prototype, "color", {
            get: function () {
                _throwIfNotLoaded("color", this.m_color, "RangeBorder", this._isNull);
                return this.m_color;
            },
            set: function (value) {
                this.m_color = value;
                _createSetPropertyAction(this.context, this, "Color", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeBorder.prototype, "sideIndex", {
            get: function () {
                _throwIfNotLoaded("sideIndex", this.m_sideIndex, "RangeBorder", this._isNull);
                return this.m_sideIndex;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeBorder.prototype, "style", {
            get: function () {
                _throwIfNotLoaded("style", this.m_style, "RangeBorder", this._isNull);
                return this.m_style;
            },
            set: function (value) {
                this.m_style = value;
                _createSetPropertyAction(this.context, this, "Style", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeBorder.prototype, "weight", {
            get: function () {
                _throwIfNotLoaded("weight", this.m_weight, "RangeBorder", this._isNull);
                return this.m_weight;
            },
            set: function (value) {
                this.m_weight = value;
                _createSetPropertyAction(this.context, this, "Weight", value);
            },
            enumerable: true,
            configurable: true
        });
        RangeBorder.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["style", "weight", "color"], [], []);
        };
        RangeBorder.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Color"])) {
                this.m_color = obj["Color"];
            }
            if (!_isUndefined(obj["SideIndex"])) {
                this.m_sideIndex = obj["SideIndex"];
            }
            if (!_isUndefined(obj["Style"])) {
                this.m_style = obj["Style"];
            }
            if (!_isUndefined(obj["Weight"])) {
                this.m_weight = obj["Weight"];
            }
        };
        RangeBorder.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        RangeBorder.prototype.toJSON = function () {
            return {
                "color": this.m_color,
                "sideIndex": this.m_sideIndex,
                "style": this.m_style,
                "weight": this.m_weight
            };
        };
        return RangeBorder;
    }(OfficeExtension.ClientObject));
    Excel.RangeBorder = RangeBorder;
    var RangeBorderCollection = (function (_super) {
        __extends(RangeBorderCollection, _super);
        function RangeBorderCollection() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(RangeBorderCollection.prototype, "_className", {
            get: function () {
                return "RangeBorderCollection";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeBorderCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, "RangeBorderCollection", this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeBorderCollection.prototype, "count", {
            get: function () {
                _throwIfNotLoaded("count", this.m_count, "RangeBorderCollection", this._isNull);
                return this.m_count;
            },
            enumerable: true,
            configurable: true
        });
        RangeBorderCollection.prototype.getItem = function (index) {
            return new Excel.RangeBorder(this.context, _createIndexerObjectPath(this.context, this, [index]));
        };
        RangeBorderCollection.prototype.getItemAt = function (index) {
            return new Excel.RangeBorder(this.context, _createMethodObjectPath(this.context, this, "GetItemAt", 1, [index], false, false, null));
        };
        RangeBorderCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Count"])) {
                this.m_count = obj["Count"];
            }
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.RangeBorder(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        RangeBorderCollection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        RangeBorderCollection.prototype.toJSON = function () {
            return {
                "count": this.m_count
            };
        };
        return RangeBorderCollection;
    }(OfficeExtension.ClientObject));
    Excel.RangeBorderCollection = RangeBorderCollection;
    var RangeFont = (function (_super) {
        __extends(RangeFont, _super);
        function RangeFont() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(RangeFont.prototype, "_className", {
            get: function () {
                return "RangeFont";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeFont.prototype, "bold", {
            get: function () {
                _throwIfNotLoaded("bold", this.m_bold, "RangeFont", this._isNull);
                return this.m_bold;
            },
            set: function (value) {
                this.m_bold = value;
                _createSetPropertyAction(this.context, this, "Bold", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeFont.prototype, "color", {
            get: function () {
                _throwIfNotLoaded("color", this.m_color, "RangeFont", this._isNull);
                return this.m_color;
            },
            set: function (value) {
                this.m_color = value;
                _createSetPropertyAction(this.context, this, "Color", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeFont.prototype, "italic", {
            get: function () {
                _throwIfNotLoaded("italic", this.m_italic, "RangeFont", this._isNull);
                return this.m_italic;
            },
            set: function (value) {
                this.m_italic = value;
                _createSetPropertyAction(this.context, this, "Italic", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeFont.prototype, "name", {
            get: function () {
                _throwIfNotLoaded("name", this.m_name, "RangeFont", this._isNull);
                return this.m_name;
            },
            set: function (value) {
                this.m_name = value;
                _createSetPropertyAction(this.context, this, "Name", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeFont.prototype, "size", {
            get: function () {
                _throwIfNotLoaded("size", this.m_size, "RangeFont", this._isNull);
                return this.m_size;
            },
            set: function (value) {
                this.m_size = value;
                _createSetPropertyAction(this.context, this, "Size", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeFont.prototype, "underline", {
            get: function () {
                _throwIfNotLoaded("underline", this.m_underline, "RangeFont", this._isNull);
                return this.m_underline;
            },
            set: function (value) {
                this.m_underline = value;
                _createSetPropertyAction(this.context, this, "Underline", value);
            },
            enumerable: true,
            configurable: true
        });
        RangeFont.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["name", "size", "color", "italic", "bold", "underline"], [], []);
        };
        RangeFont.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Bold"])) {
                this.m_bold = obj["Bold"];
            }
            if (!_isUndefined(obj["Color"])) {
                this.m_color = obj["Color"];
            }
            if (!_isUndefined(obj["Italic"])) {
                this.m_italic = obj["Italic"];
            }
            if (!_isUndefined(obj["Name"])) {
                this.m_name = obj["Name"];
            }
            if (!_isUndefined(obj["Size"])) {
                this.m_size = obj["Size"];
            }
            if (!_isUndefined(obj["Underline"])) {
                this.m_underline = obj["Underline"];
            }
        };
        RangeFont.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        RangeFont.prototype.toJSON = function () {
            return {
                "bold": this.m_bold,
                "color": this.m_color,
                "italic": this.m_italic,
                "name": this.m_name,
                "size": this.m_size,
                "underline": this.m_underline
            };
        };
        return RangeFont;
    }(OfficeExtension.ClientObject));
    Excel.RangeFont = RangeFont;
    var ChartCollection = (function (_super) {
        __extends(ChartCollection, _super);
        function ChartCollection() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ChartCollection.prototype, "_className", {
            get: function () {
                return "ChartCollection";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, "ChartCollection", this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartCollection.prototype, "count", {
            get: function () {
                _throwIfNotLoaded("count", this.m_count, "ChartCollection", this._isNull);
                return this.m_count;
            },
            enumerable: true,
            configurable: true
        });
        ChartCollection.prototype.add = function (type, sourceData, seriesBy) {
            if (!(sourceData instanceof Range)) {
                throw OfficeExtension.Utility.createRuntimeError(OfficeExtension.ResourceStrings.invalidArgument, "sourceData", "Charts.Add");
            }
            return new Excel.Chart(this.context, _createMethodObjectPath(this.context, this, "Add", 0, [type, sourceData, seriesBy], false, true, null));
        };
        ChartCollection.prototype.getCount = function () {
            _throwIfApiNotSupported("ChartCollection.getCount", _defaultApiSetName, "1.4", _hostName);
            var action = _createMethodAction(this.context, this, "GetCount", 1, []);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        ChartCollection.prototype.getItem = function (name) {
            return new Excel.Chart(this.context, _createMethodObjectPath(this.context, this, "GetItem", 1, [name], false, false, null));
        };
        ChartCollection.prototype.getItemAt = function (index) {
            return new Excel.Chart(this.context, _createMethodObjectPath(this.context, this, "GetItemAt", 1, [index], false, false, null));
        };
        ChartCollection.prototype.getItemOrNullObject = function (name) {
            _throwIfApiNotSupported("ChartCollection.getItemOrNullObject", _defaultApiSetName, "1.4", _hostName);
            return new Excel.Chart(this.context, _createMethodObjectPath(this.context, this, "GetItemOrNullObject", 1, [name], false, false, null));
        };
        ChartCollection.prototype._GetItem = function (key) {
            return new Excel.Chart(this.context, _createIndexerObjectPath(this.context, this, [key]));
        };
        ChartCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Count"])) {
                this.m_count = obj["Count"];
            }
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.Chart(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        ChartCollection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ChartCollection.prototype.toJSON = function () {
            return {
                "count": this.m_count
            };
        };
        return ChartCollection;
    }(OfficeExtension.ClientObject));
    Excel.ChartCollection = ChartCollection;
    var Chart = (function (_super) {
        __extends(Chart, _super);
        function Chart() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(Chart.prototype, "_className", {
            get: function () {
                return "Chart";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Chart.prototype, "axes", {
            get: function () {
                if (!this.m_axes) {
                    this.m_axes = new Excel.ChartAxes(this.context, _createPropertyObjectPath(this.context, this, "Axes", false, false));
                }
                return this.m_axes;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Chart.prototype, "dataLabels", {
            get: function () {
                if (!this.m_dataLabels) {
                    this.m_dataLabels = new Excel.ChartDataLabels(this.context, _createPropertyObjectPath(this.context, this, "DataLabels", false, false));
                }
                return this.m_dataLabels;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Chart.prototype, "format", {
            get: function () {
                if (!this.m_format) {
                    this.m_format = new Excel.ChartAreaFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
                }
                return this.m_format;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Chart.prototype, "legend", {
            get: function () {
                if (!this.m_legend) {
                    this.m_legend = new Excel.ChartLegend(this.context, _createPropertyObjectPath(this.context, this, "Legend", false, false));
                }
                return this.m_legend;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Chart.prototype, "series", {
            get: function () {
                if (!this.m_series) {
                    this.m_series = new Excel.ChartSeriesCollection(this.context, _createPropertyObjectPath(this.context, this, "Series", true, false));
                }
                return this.m_series;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Chart.prototype, "title", {
            get: function () {
                if (!this.m_title) {
                    this.m_title = new Excel.ChartTitle(this.context, _createPropertyObjectPath(this.context, this, "Title", false, false));
                }
                return this.m_title;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Chart.prototype, "worksheet", {
            get: function () {
                _throwIfApiNotSupported("Chart.worksheet", _defaultApiSetName, "1.2", _hostName);
                if (!this.m_worksheet) {
                    this.m_worksheet = new Excel.Worksheet(this.context, _createPropertyObjectPath(this.context, this, "Worksheet", false, false));
                }
                return this.m_worksheet;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Chart.prototype, "height", {
            get: function () {
                _throwIfNotLoaded("height", this.m_height, "Chart", this._isNull);
                return this.m_height;
            },
            set: function (value) {
                this.m_height = value;
                _createSetPropertyAction(this.context, this, "Height", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Chart.prototype, "id", {
            get: function () {
                _throwIfNotLoaded("id", this.m_id, "Chart", this._isNull);
                _throwIfApiNotSupported("Chart.id", _defaultApiSetName, "1.8", _hostName);
                return this.m_id;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Chart.prototype, "left", {
            get: function () {
                _throwIfNotLoaded("left", this.m_left, "Chart", this._isNull);
                return this.m_left;
            },
            set: function (value) {
                this.m_left = value;
                _createSetPropertyAction(this.context, this, "Left", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Chart.prototype, "name", {
            get: function () {
                _throwIfNotLoaded("name", this.m_name, "Chart", this._isNull);
                return this.m_name;
            },
            set: function (value) {
                this.m_name = value;
                _createSetPropertyAction(this.context, this, "Name", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Chart.prototype, "showAllFieldButtons", {
            get: function () {
                _throwIfNotLoaded("showAllFieldButtons", this.m_showAllFieldButtons, "Chart", this._isNull);
                _throwIfApiNotSupported("Chart.showAllFieldButtons", _defaultApiSetName, "1.8", _hostName);
                return this.m_showAllFieldButtons;
            },
            set: function (value) {
                this.m_showAllFieldButtons = value;
                _createSetPropertyAction(this.context, this, "ShowAllFieldButtons", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Chart.prototype, "top", {
            get: function () {
                _throwIfNotLoaded("top", this.m_top, "Chart", this._isNull);
                return this.m_top;
            },
            set: function (value) {
                this.m_top = value;
                _createSetPropertyAction(this.context, this, "Top", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Chart.prototype, "width", {
            get: function () {
                _throwIfNotLoaded("width", this.m_width, "Chart", this._isNull);
                return this.m_width;
            },
            set: function (value) {
                this.m_width = value;
                _createSetPropertyAction(this.context, this, "Width", value);
            },
            enumerable: true,
            configurable: true
        });
        Chart.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["name", "top", "left", "width", "height", "showAllFieldButtons"], ["title", "dataLabels", "legend", "axes", "format"], [
                "series",
                "worksheet",
                "series",
                "worksheet"
            ]);
        };
        Chart.prototype.delete = function () {
            _createMethodAction(this.context, this, "Delete", 0, []);
        };
        Chart.prototype.getImage = function (width, height, fittingMode) {
            _throwIfApiNotSupported("Chart.getImage", _defaultApiSetName, "1.2", _hostName);
            var action = _createMethodAction(this.context, this, "GetImage", 1, [width, height, fittingMode]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        Chart.prototype.setData = function (sourceData, seriesBy) {
            if (!(sourceData instanceof Range)) {
                throw OfficeExtension.Utility.createRuntimeError(OfficeExtension.ResourceStrings.invalidArgument, "sourceData", "Chart.setData");
            }
            _createMethodAction(this.context, this, "SetData", 0, [sourceData, seriesBy]);
        };
        Chart.prototype.setPosition = function (startCell, endCell) {
            _createMethodAction(this.context, this, "SetPosition", 0, [startCell, endCell]);
        };
        Chart.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Height"])) {
                this.m_height = obj["Height"];
            }
            if (!_isUndefined(obj["Id"])) {
                this.m_id = obj["Id"];
            }
            if (!_isUndefined(obj["Left"])) {
                this.m_left = obj["Left"];
            }
            if (!_isUndefined(obj["Name"])) {
                this.m_name = obj["Name"];
            }
            if (!_isUndefined(obj["ShowAllFieldButtons"])) {
                this.m_showAllFieldButtons = obj["ShowAllFieldButtons"];
            }
            if (!_isUndefined(obj["Top"])) {
                this.m_top = obj["Top"];
            }
            if (!_isUndefined(obj["Width"])) {
                this.m_width = obj["Width"];
            }
            _handleNavigationPropertyResults(this, obj, ["axes", "Axes", "dataLabels", "DataLabels", "format", "Format", "legend", "Legend", "series", "Series", "title", "Title", "worksheet", "Worksheet"]);
        };
        Chart.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        Chart.prototype._handleIdResult = function (value) {
            _super.prototype._handleIdResult.call(this, value);
            if (_isNullOrUndefined(value)) {
                return;
            }
            if (!_isUndefined(value["Id"])) {
                this.m_id = value["Id"];
            }
        };
        Chart.prototype.toJSON = function () {
            return {
                "axes": this.m_axes,
                "dataLabels": this.m_dataLabels,
                "format": this.m_format,
                "height": this.m_height,
                "id": this.m_id,
                "left": this.m_left,
                "legend": this.m_legend,
                "name": this.m_name,
                "showAllFieldButtons": this.m_showAllFieldButtons,
                "title": this.m_title,
                "top": this.m_top,
                "width": this.m_width
            };
        };
        return Chart;
    }(OfficeExtension.ClientObject));
    Excel.Chart = Chart;
    var ChartAreaFormat = (function (_super) {
        __extends(ChartAreaFormat, _super);
        function ChartAreaFormat() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ChartAreaFormat.prototype, "_className", {
            get: function () {
                return "ChartAreaFormat";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartAreaFormat.prototype, "fill", {
            get: function () {
                if (!this.m_fill) {
                    this.m_fill = new Excel.ChartFill(this.context, _createPropertyObjectPath(this.context, this, "Fill", false, false));
                }
                return this.m_fill;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartAreaFormat.prototype, "font", {
            get: function () {
                if (!this.m_font) {
                    this.m_font = new Excel.ChartFont(this.context, _createPropertyObjectPath(this.context, this, "Font", false, false));
                }
                return this.m_font;
            },
            enumerable: true,
            configurable: true
        });
        ChartAreaFormat.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, [], ["font"], [
                "fill"
            ]);
        };
        ChartAreaFormat.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            _handleNavigationPropertyResults(this, obj, ["fill", "Fill", "font", "Font"]);
        };
        ChartAreaFormat.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ChartAreaFormat.prototype.toJSON = function () {
            return {
                "fill": this.m_fill,
                "font": this.m_font
            };
        };
        return ChartAreaFormat;
    }(OfficeExtension.ClientObject));
    Excel.ChartAreaFormat = ChartAreaFormat;
    var ChartSeriesCollection = (function (_super) {
        __extends(ChartSeriesCollection, _super);
        function ChartSeriesCollection() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ChartSeriesCollection.prototype, "_className", {
            get: function () {
                return "ChartSeriesCollection";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartSeriesCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, "ChartSeriesCollection", this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartSeriesCollection.prototype, "count", {
            get: function () {
                _throwIfNotLoaded("count", this.m_count, "ChartSeriesCollection", this._isNull);
                return this.m_count;
            },
            enumerable: true,
            configurable: true
        });
        ChartSeriesCollection.prototype.add = function (name, index) {
            _throwIfApiNotSupported("ChartSeriesCollection.add", _defaultApiSetName, "1.8", _hostName);
            return new Excel.ChartSeries(this.context, _createMethodObjectPath(this.context, this, "Add", 0, [name, index], false, true, null));
        };
        ChartSeriesCollection.prototype.getCount = function () {
            _throwIfApiNotSupported("ChartSeriesCollection.getCount", _defaultApiSetName, "1.4", _hostName);
            var action = _createMethodAction(this.context, this, "GetCount", 1, []);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        ChartSeriesCollection.prototype.getItemAt = function (index) {
            return new Excel.ChartSeries(this.context, _createMethodObjectPath(this.context, this, "GetItemAt", 1, [index], false, false, null));
        };
        ChartSeriesCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Count"])) {
                this.m_count = obj["Count"];
            }
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.ChartSeries(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(false, this.context, this, _data[i], i));
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        ChartSeriesCollection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ChartSeriesCollection.prototype.toJSON = function () {
            return {
                "count": this.m_count
            };
        };
        return ChartSeriesCollection;
    }(OfficeExtension.ClientObject));
    Excel.ChartSeriesCollection = ChartSeriesCollection;
    var ChartSeries = (function (_super) {
        __extends(ChartSeries, _super);
        function ChartSeries() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ChartSeries.prototype, "_className", {
            get: function () {
                return "ChartSeries";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartSeries.prototype, "format", {
            get: function () {
                if (!this.m_format) {
                    this.m_format = new Excel.ChartSeriesFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
                }
                return this.m_format;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartSeries.prototype, "points", {
            get: function () {
                if (!this.m_points) {
                    this.m_points = new Excel.ChartPointsCollection(this.context, _createPropertyObjectPath(this.context, this, "Points", true, false));
                }
                return this.m_points;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartSeries.prototype, "trendlines", {
            get: function () {
                _throwIfApiNotSupported("ChartSeries.trendlines", _defaultApiSetName, "1.8", _hostName);
                if (!this.m_trendlines) {
                    this.m_trendlines = new Excel.ChartTrendlineCollection(this.context, _createPropertyObjectPath(this.context, this, "Trendlines", true, false));
                }
                return this.m_trendlines;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartSeries.prototype, "name", {
            get: function () {
                _throwIfNotLoaded("name", this.m_name, "ChartSeries", this._isNull);
                return this.m_name;
            },
            set: function (value) {
                this.m_name = value;
                _createSetPropertyAction(this.context, this, "Name", value);
            },
            enumerable: true,
            configurable: true
        });
        ChartSeries.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["name"], ["format"], [
                "points",
                "trendlines",
                "points",
                "trendlines"
            ]);
        };
        ChartSeries.prototype.delete = function () {
            _throwIfApiNotSupported("ChartSeries.delete", _defaultApiSetName, "1.8", _hostName);
            _createMethodAction(this.context, this, "Delete", 0, []);
        };
        ChartSeries.prototype.setBubbleSizes = function (sourceData) {
            _throwIfApiNotSupported("ChartSeries.setBubbleSizes", _defaultApiSetName, "1.8", _hostName);
            _createMethodAction(this.context, this, "SetBubbleSizes", 0, [sourceData]);
        };
        ChartSeries.prototype.setValues = function (sourceData) {
            _throwIfApiNotSupported("ChartSeries.setValues", _defaultApiSetName, "1.8", _hostName);
            _createMethodAction(this.context, this, "SetValues", 0, [sourceData]);
        };
        ChartSeries.prototype.setXAxisValues = function (sourceData) {
            _throwIfApiNotSupported("ChartSeries.setXAxisValues", _defaultApiSetName, "1.8", _hostName);
            _createMethodAction(this.context, this, "SetXAxisValues", 0, [sourceData]);
        };
        ChartSeries.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Name"])) {
                this.m_name = obj["Name"];
            }
            _handleNavigationPropertyResults(this, obj, ["format", "Format", "points", "Points", "trendlines", "Trendlines"]);
        };
        ChartSeries.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ChartSeries.prototype.toJSON = function () {
            return {
                "format": this.m_format,
                "name": this.m_name
            };
        };
        return ChartSeries;
    }(OfficeExtension.ClientObject));
    Excel.ChartSeries = ChartSeries;
    var ChartSeriesFormat = (function (_super) {
        __extends(ChartSeriesFormat, _super);
        function ChartSeriesFormat() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ChartSeriesFormat.prototype, "_className", {
            get: function () {
                return "ChartSeriesFormat";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartSeriesFormat.prototype, "fill", {
            get: function () {
                if (!this.m_fill) {
                    this.m_fill = new Excel.ChartFill(this.context, _createPropertyObjectPath(this.context, this, "Fill", false, false));
                }
                return this.m_fill;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartSeriesFormat.prototype, "line", {
            get: function () {
                if (!this.m_line) {
                    this.m_line = new Excel.ChartLineFormat(this.context, _createPropertyObjectPath(this.context, this, "Line", false, false));
                }
                return this.m_line;
            },
            enumerable: true,
            configurable: true
        });
        ChartSeriesFormat.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, [], ["line"], [
                "fill"
            ]);
        };
        ChartSeriesFormat.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            _handleNavigationPropertyResults(this, obj, ["fill", "Fill", "line", "Line"]);
        };
        ChartSeriesFormat.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ChartSeriesFormat.prototype.toJSON = function () {
            return {
                "fill": this.m_fill,
                "line": this.m_line
            };
        };
        return ChartSeriesFormat;
    }(OfficeExtension.ClientObject));
    Excel.ChartSeriesFormat = ChartSeriesFormat;
    var ChartPointsCollection = (function (_super) {
        __extends(ChartPointsCollection, _super);
        function ChartPointsCollection() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ChartPointsCollection.prototype, "_className", {
            get: function () {
                return "ChartPointsCollection";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartPointsCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, "ChartPointsCollection", this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartPointsCollection.prototype, "count", {
            get: function () {
                _throwIfNotLoaded("count", this.m_count, "ChartPointsCollection", this._isNull);
                return this.m_count;
            },
            enumerable: true,
            configurable: true
        });
        ChartPointsCollection.prototype.getCount = function () {
            _throwIfApiNotSupported("ChartPointsCollection.getCount", _defaultApiSetName, "1.4", _hostName);
            var action = _createMethodAction(this.context, this, "GetCount", 1, []);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        ChartPointsCollection.prototype.getItemAt = function (index) {
            return new Excel.ChartPoint(this.context, _createMethodObjectPath(this.context, this, "GetItemAt", 1, [index], false, false, null));
        };
        ChartPointsCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Count"])) {
                this.m_count = obj["Count"];
            }
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.ChartPoint(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(false, this.context, this, _data[i], i));
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        ChartPointsCollection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ChartPointsCollection.prototype.toJSON = function () {
            return {
                "count": this.m_count
            };
        };
        return ChartPointsCollection;
    }(OfficeExtension.ClientObject));
    Excel.ChartPointsCollection = ChartPointsCollection;
    var ChartPoint = (function (_super) {
        __extends(ChartPoint, _super);
        function ChartPoint() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ChartPoint.prototype, "_className", {
            get: function () {
                return "ChartPoint";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartPoint.prototype, "format", {
            get: function () {
                if (!this.m_format) {
                    this.m_format = new Excel.ChartPointFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
                }
                return this.m_format;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartPoint.prototype, "value", {
            get: function () {
                _throwIfNotLoaded("value", this.m_value, "ChartPoint", this._isNull);
                return this.m_value;
            },
            enumerable: true,
            configurable: true
        });
        ChartPoint.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Value"])) {
                this.m_value = obj["Value"];
            }
            _handleNavigationPropertyResults(this, obj, ["format", "Format"]);
        };
        ChartPoint.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ChartPoint.prototype.toJSON = function () {
            return {
                "format": this.m_format,
                "value": this.m_value
            };
        };
        return ChartPoint;
    }(OfficeExtension.ClientObject));
    Excel.ChartPoint = ChartPoint;
    var ChartPointFormat = (function (_super) {
        __extends(ChartPointFormat, _super);
        function ChartPointFormat() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ChartPointFormat.prototype, "_className", {
            get: function () {
                return "ChartPointFormat";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartPointFormat.prototype, "fill", {
            get: function () {
                if (!this.m_fill) {
                    this.m_fill = new Excel.ChartFill(this.context, _createPropertyObjectPath(this.context, this, "Fill", false, false));
                }
                return this.m_fill;
            },
            enumerable: true,
            configurable: true
        });
        ChartPointFormat.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            _handleNavigationPropertyResults(this, obj, ["fill", "Fill"]);
        };
        ChartPointFormat.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ChartPointFormat.prototype.toJSON = function () {
            return {
                "fill": this.m_fill
            };
        };
        return ChartPointFormat;
    }(OfficeExtension.ClientObject));
    Excel.ChartPointFormat = ChartPointFormat;
    var ChartAxes = (function (_super) {
        __extends(ChartAxes, _super);
        function ChartAxes() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ChartAxes.prototype, "_className", {
            get: function () {
                return "ChartAxes";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartAxes.prototype, "categoryAxis", {
            get: function () {
                if (!this.m_categoryAxis) {
                    this.m_categoryAxis = new Excel.ChartAxis(this.context, _createPropertyObjectPath(this.context, this, "CategoryAxis", false, false));
                }
                return this.m_categoryAxis;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartAxes.prototype, "seriesAxis", {
            get: function () {
                if (!this.m_seriesAxis) {
                    this.m_seriesAxis = new Excel.ChartAxis(this.context, _createPropertyObjectPath(this.context, this, "SeriesAxis", false, false));
                }
                return this.m_seriesAxis;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartAxes.prototype, "valueAxis", {
            get: function () {
                if (!this.m_valueAxis) {
                    this.m_valueAxis = new Excel.ChartAxis(this.context, _createPropertyObjectPath(this.context, this, "ValueAxis", false, false));
                }
                return this.m_valueAxis;
            },
            enumerable: true,
            configurable: true
        });
        ChartAxes.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, [], ["categoryAxis", "seriesAxis", "valueAxis"], []);
        };
        ChartAxes.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            _handleNavigationPropertyResults(this, obj, ["categoryAxis", "CategoryAxis", "seriesAxis", "SeriesAxis", "valueAxis", "ValueAxis"]);
        };
        ChartAxes.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ChartAxes.prototype.toJSON = function () {
            return {
                "categoryAxis": this.m_categoryAxis,
                "seriesAxis": this.m_seriesAxis,
                "valueAxis": this.m_valueAxis
            };
        };
        return ChartAxes;
    }(OfficeExtension.ClientObject));
    Excel.ChartAxes = ChartAxes;
    var ChartAxis = (function (_super) {
        __extends(ChartAxis, _super);
        function ChartAxis() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ChartAxis.prototype, "_className", {
            get: function () {
                return "ChartAxis";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartAxis.prototype, "format", {
            get: function () {
                if (!this.m_format) {
                    this.m_format = new Excel.ChartAxisFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
                }
                return this.m_format;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartAxis.prototype, "majorGridlines", {
            get: function () {
                if (!this.m_majorGridlines) {
                    this.m_majorGridlines = new Excel.ChartGridlines(this.context, _createPropertyObjectPath(this.context, this, "MajorGridlines", false, false));
                }
                return this.m_majorGridlines;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartAxis.prototype, "minorGridlines", {
            get: function () {
                if (!this.m_minorGridlines) {
                    this.m_minorGridlines = new Excel.ChartGridlines(this.context, _createPropertyObjectPath(this.context, this, "MinorGridlines", false, false));
                }
                return this.m_minorGridlines;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartAxis.prototype, "title", {
            get: function () {
                if (!this.m_title) {
                    this.m_title = new Excel.ChartAxisTitle(this.context, _createPropertyObjectPath(this.context, this, "Title", false, false));
                }
                return this.m_title;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartAxis.prototype, "baseTimeUnit", {
            get: function () {
                _throwIfNotLoaded("baseTimeUnit", this.m_baseTimeUnit, "ChartAxis", this._isNull);
                _throwIfApiNotSupported("ChartAxis.baseTimeUnit", _defaultApiSetName, "1.8", _hostName);
                return this.m_baseTimeUnit;
            },
            set: function (value) {
                this.m_baseTimeUnit = value;
                _createSetPropertyAction(this.context, this, "BaseTimeUnit", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartAxis.prototype, "categoryType", {
            get: function () {
                _throwIfNotLoaded("categoryType", this.m_categoryType, "ChartAxis", this._isNull);
                _throwIfApiNotSupported("ChartAxis.categoryType", _defaultApiSetName, "1.8", _hostName);
                return this.m_categoryType;
            },
            set: function (value) {
                this.m_categoryType = value;
                _createSetPropertyAction(this.context, this, "CategoryType", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartAxis.prototype, "customDisplayUnit", {
            get: function () {
                _throwIfNotLoaded("customDisplayUnit", this.m_customDisplayUnit, "ChartAxis", this._isNull);
                _throwIfApiNotSupported("ChartAxis.customDisplayUnit", _defaultApiSetName, "1.8", _hostName);
                return this.m_customDisplayUnit;
            },
            set: function (value) {
                this.m_customDisplayUnit = value;
                _createSetPropertyAction(this.context, this, "CustomDisplayUnit", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartAxis.prototype, "displayUnit", {
            get: function () {
                _throwIfNotLoaded("displayUnit", this.m_displayUnit, "ChartAxis", this._isNull);
                _throwIfApiNotSupported("ChartAxis.displayUnit", _defaultApiSetName, "1.8", _hostName);
                return this.m_displayUnit;
            },
            set: function (value) {
                this.m_displayUnit = value;
                _createSetPropertyAction(this.context, this, "DisplayUnit", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartAxis.prototype, "majorTimeUnitScale", {
            get: function () {
                _throwIfNotLoaded("majorTimeUnitScale", this.m_majorTimeUnitScale, "ChartAxis", this._isNull);
                _throwIfApiNotSupported("ChartAxis.majorTimeUnitScale", _defaultApiSetName, "1.8", _hostName);
                return this.m_majorTimeUnitScale;
            },
            set: function (value) {
                this.m_majorTimeUnitScale = value;
                _createSetPropertyAction(this.context, this, "MajorTimeUnitScale", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartAxis.prototype, "majorUnit", {
            get: function () {
                _throwIfNotLoaded("majorUnit", this.m_majorUnit, "ChartAxis", this._isNull);
                return this.m_majorUnit;
            },
            set: function (value) {
                this.m_majorUnit = value;
                _createSetPropertyAction(this.context, this, "MajorUnit", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartAxis.prototype, "maximum", {
            get: function () {
                _throwIfNotLoaded("maximum", this.m_maximum, "ChartAxis", this._isNull);
                return this.m_maximum;
            },
            set: function (value) {
                this.m_maximum = value;
                _createSetPropertyAction(this.context, this, "Maximum", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartAxis.prototype, "minimum", {
            get: function () {
                _throwIfNotLoaded("minimum", this.m_minimum, "ChartAxis", this._isNull);
                return this.m_minimum;
            },
            set: function (value) {
                this.m_minimum = value;
                _createSetPropertyAction(this.context, this, "Minimum", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartAxis.prototype, "minorTimeUnitScale", {
            get: function () {
                _throwIfNotLoaded("minorTimeUnitScale", this.m_minorTimeUnitScale, "ChartAxis", this._isNull);
                _throwIfApiNotSupported("ChartAxis.minorTimeUnitScale", _defaultApiSetName, "1.8", _hostName);
                return this.m_minorTimeUnitScale;
            },
            set: function (value) {
                this.m_minorTimeUnitScale = value;
                _createSetPropertyAction(this.context, this, "MinorTimeUnitScale", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartAxis.prototype, "minorUnit", {
            get: function () {
                _throwIfNotLoaded("minorUnit", this.m_minorUnit, "ChartAxis", this._isNull);
                return this.m_minorUnit;
            },
            set: function (value) {
                this.m_minorUnit = value;
                _createSetPropertyAction(this.context, this, "MinorUnit", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartAxis.prototype, "showDisplayUnitLabel", {
            get: function () {
                _throwIfNotLoaded("showDisplayUnitLabel", this.m_showDisplayUnitLabel, "ChartAxis", this._isNull);
                _throwIfApiNotSupported("ChartAxis.showDisplayUnitLabel", _defaultApiSetName, "1.8", _hostName);
                return this.m_showDisplayUnitLabel;
            },
            set: function (value) {
                this.m_showDisplayUnitLabel = value;
                _createSetPropertyAction(this.context, this, "ShowDisplayUnitLabel", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartAxis.prototype, "type", {
            get: function () {
                _throwIfNotLoaded("type", this.m_type, "ChartAxis", this._isNull);
                _throwIfApiNotSupported("ChartAxis.type", _defaultApiSetName, "1.8", _hostName);
                return this.m_type;
            },
            enumerable: true,
            configurable: true
        });
        ChartAxis.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["majorUnit", "maximum", "minimum", "minorUnit", "displayUnit", "showDisplayUnitLabel", "customDisplayUnit", "minorTimeUnitScale", "majorTimeUnitScale", "baseTimeUnit", "categoryType"], ["majorGridlines", "minorGridlines", "title", "format"], []);
        };
        ChartAxis.prototype.setCategoryNames = function (sourceData) {
            _throwIfApiNotSupported("ChartAxis.setCategoryNames", _defaultApiSetName, "1.8", _hostName);
            _createMethodAction(this.context, this, "SetCategoryNames", 0, [sourceData]);
        };
        ChartAxis.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["BaseTimeUnit"])) {
                this.m_baseTimeUnit = obj["BaseTimeUnit"];
            }
            if (!_isUndefined(obj["CategoryType"])) {
                this.m_categoryType = obj["CategoryType"];
            }
            if (!_isUndefined(obj["CustomDisplayUnit"])) {
                this.m_customDisplayUnit = obj["CustomDisplayUnit"];
            }
            if (!_isUndefined(obj["DisplayUnit"])) {
                this.m_displayUnit = obj["DisplayUnit"];
            }
            if (!_isUndefined(obj["MajorTimeUnitScale"])) {
                this.m_majorTimeUnitScale = obj["MajorTimeUnitScale"];
            }
            if (!_isUndefined(obj["MajorUnit"])) {
                this.m_majorUnit = obj["MajorUnit"];
            }
            if (!_isUndefined(obj["Maximum"])) {
                this.m_maximum = obj["Maximum"];
            }
            if (!_isUndefined(obj["Minimum"])) {
                this.m_minimum = obj["Minimum"];
            }
            if (!_isUndefined(obj["MinorTimeUnitScale"])) {
                this.m_minorTimeUnitScale = obj["MinorTimeUnitScale"];
            }
            if (!_isUndefined(obj["MinorUnit"])) {
                this.m_minorUnit = obj["MinorUnit"];
            }
            if (!_isUndefined(obj["ShowDisplayUnitLabel"])) {
                this.m_showDisplayUnitLabel = obj["ShowDisplayUnitLabel"];
            }
            if (!_isUndefined(obj["Type"])) {
                this.m_type = obj["Type"];
            }
            _handleNavigationPropertyResults(this, obj, ["format", "Format", "majorGridlines", "MajorGridlines", "minorGridlines", "MinorGridlines", "title", "Title"]);
        };
        ChartAxis.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ChartAxis.prototype.toJSON = function () {
            return {
                "baseTimeUnit": this.m_baseTimeUnit,
                "categoryType": this.m_categoryType,
                "customDisplayUnit": this.m_customDisplayUnit,
                "displayUnit": this.m_displayUnit,
                "format": this.m_format,
                "majorGridlines": this.m_majorGridlines,
                "majorTimeUnitScale": this.m_majorTimeUnitScale,
                "majorUnit": this.m_majorUnit,
                "maximum": this.m_maximum,
                "minimum": this.m_minimum,
                "minorGridlines": this.m_minorGridlines,
                "minorTimeUnitScale": this.m_minorTimeUnitScale,
                "minorUnit": this.m_minorUnit,
                "showDisplayUnitLabel": this.m_showDisplayUnitLabel,
                "title": this.m_title,
                "type": this.m_type
            };
        };
        return ChartAxis;
    }(OfficeExtension.ClientObject));
    Excel.ChartAxis = ChartAxis;
    var ChartAxisFormat = (function (_super) {
        __extends(ChartAxisFormat, _super);
        function ChartAxisFormat() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ChartAxisFormat.prototype, "_className", {
            get: function () {
                return "ChartAxisFormat";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartAxisFormat.prototype, "font", {
            get: function () {
                if (!this.m_font) {
                    this.m_font = new Excel.ChartFont(this.context, _createPropertyObjectPath(this.context, this, "Font", false, false));
                }
                return this.m_font;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartAxisFormat.prototype, "line", {
            get: function () {
                if (!this.m_line) {
                    this.m_line = new Excel.ChartLineFormat(this.context, _createPropertyObjectPath(this.context, this, "Line", false, false));
                }
                return this.m_line;
            },
            enumerable: true,
            configurable: true
        });
        ChartAxisFormat.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, [], ["font", "line"], []);
        };
        ChartAxisFormat.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            _handleNavigationPropertyResults(this, obj, ["font", "Font", "line", "Line"]);
        };
        ChartAxisFormat.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ChartAxisFormat.prototype.toJSON = function () {
            return {
                "font": this.m_font,
                "line": this.m_line
            };
        };
        return ChartAxisFormat;
    }(OfficeExtension.ClientObject));
    Excel.ChartAxisFormat = ChartAxisFormat;
    var ChartAxisTitle = (function (_super) {
        __extends(ChartAxisTitle, _super);
        function ChartAxisTitle() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ChartAxisTitle.prototype, "_className", {
            get: function () {
                return "ChartAxisTitle";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartAxisTitle.prototype, "format", {
            get: function () {
                if (!this.m_format) {
                    this.m_format = new Excel.ChartAxisTitleFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
                }
                return this.m_format;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartAxisTitle.prototype, "text", {
            get: function () {
                _throwIfNotLoaded("text", this.m_text, "ChartAxisTitle", this._isNull);
                return this.m_text;
            },
            set: function (value) {
                this.m_text = value;
                _createSetPropertyAction(this.context, this, "Text", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartAxisTitle.prototype, "visible", {
            get: function () {
                _throwIfNotLoaded("visible", this.m_visible, "ChartAxisTitle", this._isNull);
                return this.m_visible;
            },
            set: function (value) {
                this.m_visible = value;
                _createSetPropertyAction(this.context, this, "Visible", value);
            },
            enumerable: true,
            configurable: true
        });
        ChartAxisTitle.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["text", "visible"], ["format"], []);
        };
        ChartAxisTitle.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Text"])) {
                this.m_text = obj["Text"];
            }
            if (!_isUndefined(obj["Visible"])) {
                this.m_visible = obj["Visible"];
            }
            _handleNavigationPropertyResults(this, obj, ["format", "Format"]);
        };
        ChartAxisTitle.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ChartAxisTitle.prototype.toJSON = function () {
            return {
                "format": this.m_format,
                "text": this.m_text,
                "visible": this.m_visible
            };
        };
        return ChartAxisTitle;
    }(OfficeExtension.ClientObject));
    Excel.ChartAxisTitle = ChartAxisTitle;
    var ChartAxisTitleFormat = (function (_super) {
        __extends(ChartAxisTitleFormat, _super);
        function ChartAxisTitleFormat() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ChartAxisTitleFormat.prototype, "_className", {
            get: function () {
                return "ChartAxisTitleFormat";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartAxisTitleFormat.prototype, "font", {
            get: function () {
                if (!this.m_font) {
                    this.m_font = new Excel.ChartFont(this.context, _createPropertyObjectPath(this.context, this, "Font", false, false));
                }
                return this.m_font;
            },
            enumerable: true,
            configurable: true
        });
        ChartAxisTitleFormat.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, [], ["font"], []);
        };
        ChartAxisTitleFormat.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            _handleNavigationPropertyResults(this, obj, ["font", "Font"]);
        };
        ChartAxisTitleFormat.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ChartAxisTitleFormat.prototype.toJSON = function () {
            return {
                "font": this.m_font
            };
        };
        return ChartAxisTitleFormat;
    }(OfficeExtension.ClientObject));
    Excel.ChartAxisTitleFormat = ChartAxisTitleFormat;
    var ChartDataLabels = (function (_super) {
        __extends(ChartDataLabels, _super);
        function ChartDataLabels() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ChartDataLabels.prototype, "_className", {
            get: function () {
                return "ChartDataLabels";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartDataLabels.prototype, "format", {
            get: function () {
                if (!this.m_format) {
                    this.m_format = new Excel.ChartDataLabelFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
                }
                return this.m_format;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartDataLabels.prototype, "position", {
            get: function () {
                _throwIfNotLoaded("position", this.m_position, "ChartDataLabels", this._isNull);
                return this.m_position;
            },
            set: function (value) {
                this.m_position = value;
                _createSetPropertyAction(this.context, this, "Position", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartDataLabels.prototype, "separator", {
            get: function () {
                _throwIfNotLoaded("separator", this.m_separator, "ChartDataLabels", this._isNull);
                return this.m_separator;
            },
            set: function (value) {
                this.m_separator = value;
                _createSetPropertyAction(this.context, this, "Separator", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartDataLabels.prototype, "showBubbleSize", {
            get: function () {
                _throwIfNotLoaded("showBubbleSize", this.m_showBubbleSize, "ChartDataLabels", this._isNull);
                return this.m_showBubbleSize;
            },
            set: function (value) {
                this.m_showBubbleSize = value;
                _createSetPropertyAction(this.context, this, "ShowBubbleSize", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartDataLabels.prototype, "showCategoryName", {
            get: function () {
                _throwIfNotLoaded("showCategoryName", this.m_showCategoryName, "ChartDataLabels", this._isNull);
                return this.m_showCategoryName;
            },
            set: function (value) {
                this.m_showCategoryName = value;
                _createSetPropertyAction(this.context, this, "ShowCategoryName", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartDataLabels.prototype, "showLegendKey", {
            get: function () {
                _throwIfNotLoaded("showLegendKey", this.m_showLegendKey, "ChartDataLabels", this._isNull);
                return this.m_showLegendKey;
            },
            set: function (value) {
                this.m_showLegendKey = value;
                _createSetPropertyAction(this.context, this, "ShowLegendKey", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartDataLabels.prototype, "showPercentage", {
            get: function () {
                _throwIfNotLoaded("showPercentage", this.m_showPercentage, "ChartDataLabels", this._isNull);
                return this.m_showPercentage;
            },
            set: function (value) {
                this.m_showPercentage = value;
                _createSetPropertyAction(this.context, this, "ShowPercentage", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartDataLabels.prototype, "showSeriesName", {
            get: function () {
                _throwIfNotLoaded("showSeriesName", this.m_showSeriesName, "ChartDataLabels", this._isNull);
                return this.m_showSeriesName;
            },
            set: function (value) {
                this.m_showSeriesName = value;
                _createSetPropertyAction(this.context, this, "ShowSeriesName", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartDataLabels.prototype, "showValue", {
            get: function () {
                _throwIfNotLoaded("showValue", this.m_showValue, "ChartDataLabels", this._isNull);
                return this.m_showValue;
            },
            set: function (value) {
                this.m_showValue = value;
                _createSetPropertyAction(this.context, this, "ShowValue", value);
            },
            enumerable: true,
            configurable: true
        });
        ChartDataLabels.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["position", "showValue", "showSeriesName", "showCategoryName", "showLegendKey", "showPercentage", "showBubbleSize", "separator"], ["format"], []);
        };
        ChartDataLabels.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Position"])) {
                this.m_position = obj["Position"];
            }
            if (!_isUndefined(obj["Separator"])) {
                this.m_separator = obj["Separator"];
            }
            if (!_isUndefined(obj["ShowBubbleSize"])) {
                this.m_showBubbleSize = obj["ShowBubbleSize"];
            }
            if (!_isUndefined(obj["ShowCategoryName"])) {
                this.m_showCategoryName = obj["ShowCategoryName"];
            }
            if (!_isUndefined(obj["ShowLegendKey"])) {
                this.m_showLegendKey = obj["ShowLegendKey"];
            }
            if (!_isUndefined(obj["ShowPercentage"])) {
                this.m_showPercentage = obj["ShowPercentage"];
            }
            if (!_isUndefined(obj["ShowSeriesName"])) {
                this.m_showSeriesName = obj["ShowSeriesName"];
            }
            if (!_isUndefined(obj["ShowValue"])) {
                this.m_showValue = obj["ShowValue"];
            }
            _handleNavigationPropertyResults(this, obj, ["format", "Format"]);
        };
        ChartDataLabels.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ChartDataLabels.prototype.toJSON = function () {
            return {
                "format": this.m_format,
                "position": this.m_position,
                "separator": this.m_separator,
                "showBubbleSize": this.m_showBubbleSize,
                "showCategoryName": this.m_showCategoryName,
                "showLegendKey": this.m_showLegendKey,
                "showPercentage": this.m_showPercentage,
                "showSeriesName": this.m_showSeriesName,
                "showValue": this.m_showValue
            };
        };
        return ChartDataLabels;
    }(OfficeExtension.ClientObject));
    Excel.ChartDataLabels = ChartDataLabels;
    var ChartDataLabelFormat = (function (_super) {
        __extends(ChartDataLabelFormat, _super);
        function ChartDataLabelFormat() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ChartDataLabelFormat.prototype, "_className", {
            get: function () {
                return "ChartDataLabelFormat";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartDataLabelFormat.prototype, "fill", {
            get: function () {
                if (!this.m_fill) {
                    this.m_fill = new Excel.ChartFill(this.context, _createPropertyObjectPath(this.context, this, "Fill", false, false));
                }
                return this.m_fill;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartDataLabelFormat.prototype, "font", {
            get: function () {
                if (!this.m_font) {
                    this.m_font = new Excel.ChartFont(this.context, _createPropertyObjectPath(this.context, this, "Font", false, false));
                }
                return this.m_font;
            },
            enumerable: true,
            configurable: true
        });
        ChartDataLabelFormat.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, [], ["font"], [
                "fill"
            ]);
        };
        ChartDataLabelFormat.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            _handleNavigationPropertyResults(this, obj, ["fill", "Fill", "font", "Font"]);
        };
        ChartDataLabelFormat.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ChartDataLabelFormat.prototype.toJSON = function () {
            return {
                "fill": this.m_fill,
                "font": this.m_font
            };
        };
        return ChartDataLabelFormat;
    }(OfficeExtension.ClientObject));
    Excel.ChartDataLabelFormat = ChartDataLabelFormat;
    var ChartGridlines = (function (_super) {
        __extends(ChartGridlines, _super);
        function ChartGridlines() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ChartGridlines.prototype, "_className", {
            get: function () {
                return "ChartGridlines";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartGridlines.prototype, "format", {
            get: function () {
                if (!this.m_format) {
                    this.m_format = new Excel.ChartGridlinesFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
                }
                return this.m_format;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartGridlines.prototype, "visible", {
            get: function () {
                _throwIfNotLoaded("visible", this.m_visible, "ChartGridlines", this._isNull);
                return this.m_visible;
            },
            set: function (value) {
                this.m_visible = value;
                _createSetPropertyAction(this.context, this, "Visible", value);
            },
            enumerable: true,
            configurable: true
        });
        ChartGridlines.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["visible"], ["format"], []);
        };
        ChartGridlines.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Visible"])) {
                this.m_visible = obj["Visible"];
            }
            _handleNavigationPropertyResults(this, obj, ["format", "Format"]);
        };
        ChartGridlines.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ChartGridlines.prototype.toJSON = function () {
            return {
                "format": this.m_format,
                "visible": this.m_visible
            };
        };
        return ChartGridlines;
    }(OfficeExtension.ClientObject));
    Excel.ChartGridlines = ChartGridlines;
    var ChartGridlinesFormat = (function (_super) {
        __extends(ChartGridlinesFormat, _super);
        function ChartGridlinesFormat() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ChartGridlinesFormat.prototype, "_className", {
            get: function () {
                return "ChartGridlinesFormat";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartGridlinesFormat.prototype, "line", {
            get: function () {
                if (!this.m_line) {
                    this.m_line = new Excel.ChartLineFormat(this.context, _createPropertyObjectPath(this.context, this, "Line", false, false));
                }
                return this.m_line;
            },
            enumerable: true,
            configurable: true
        });
        ChartGridlinesFormat.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, [], ["line"], []);
        };
        ChartGridlinesFormat.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            _handleNavigationPropertyResults(this, obj, ["line", "Line"]);
        };
        ChartGridlinesFormat.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ChartGridlinesFormat.prototype.toJSON = function () {
            return {
                "line": this.m_line
            };
        };
        return ChartGridlinesFormat;
    }(OfficeExtension.ClientObject));
    Excel.ChartGridlinesFormat = ChartGridlinesFormat;
    var ChartLegend = (function (_super) {
        __extends(ChartLegend, _super);
        function ChartLegend() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ChartLegend.prototype, "_className", {
            get: function () {
                return "ChartLegend";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartLegend.prototype, "format", {
            get: function () {
                if (!this.m_format) {
                    this.m_format = new Excel.ChartLegendFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
                }
                return this.m_format;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartLegend.prototype, "overlay", {
            get: function () {
                _throwIfNotLoaded("overlay", this.m_overlay, "ChartLegend", this._isNull);
                return this.m_overlay;
            },
            set: function (value) {
                this.m_overlay = value;
                _createSetPropertyAction(this.context, this, "Overlay", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartLegend.prototype, "position", {
            get: function () {
                _throwIfNotLoaded("position", this.m_position, "ChartLegend", this._isNull);
                return this.m_position;
            },
            set: function (value) {
                this.m_position = value;
                _createSetPropertyAction(this.context, this, "Position", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartLegend.prototype, "visible", {
            get: function () {
                _throwIfNotLoaded("visible", this.m_visible, "ChartLegend", this._isNull);
                return this.m_visible;
            },
            set: function (value) {
                this.m_visible = value;
                _createSetPropertyAction(this.context, this, "Visible", value);
            },
            enumerable: true,
            configurable: true
        });
        ChartLegend.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["visible", "position", "overlay"], ["format"], []);
        };
        ChartLegend.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Overlay"])) {
                this.m_overlay = obj["Overlay"];
            }
            if (!_isUndefined(obj["Position"])) {
                this.m_position = obj["Position"];
            }
            if (!_isUndefined(obj["Visible"])) {
                this.m_visible = obj["Visible"];
            }
            _handleNavigationPropertyResults(this, obj, ["format", "Format"]);
        };
        ChartLegend.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ChartLegend.prototype.toJSON = function () {
            return {
                "format": this.m_format,
                "overlay": this.m_overlay,
                "position": this.m_position,
                "visible": this.m_visible
            };
        };
        return ChartLegend;
    }(OfficeExtension.ClientObject));
    Excel.ChartLegend = ChartLegend;
    var ChartLegendFormat = (function (_super) {
        __extends(ChartLegendFormat, _super);
        function ChartLegendFormat() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ChartLegendFormat.prototype, "_className", {
            get: function () {
                return "ChartLegendFormat";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartLegendFormat.prototype, "fill", {
            get: function () {
                if (!this.m_fill) {
                    this.m_fill = new Excel.ChartFill(this.context, _createPropertyObjectPath(this.context, this, "Fill", false, false));
                }
                return this.m_fill;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartLegendFormat.prototype, "font", {
            get: function () {
                if (!this.m_font) {
                    this.m_font = new Excel.ChartFont(this.context, _createPropertyObjectPath(this.context, this, "Font", false, false));
                }
                return this.m_font;
            },
            enumerable: true,
            configurable: true
        });
        ChartLegendFormat.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, [], ["font"], [
                "fill"
            ]);
        };
        ChartLegendFormat.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            _handleNavigationPropertyResults(this, obj, ["fill", "Fill", "font", "Font"]);
        };
        ChartLegendFormat.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ChartLegendFormat.prototype.toJSON = function () {
            return {
                "fill": this.m_fill,
                "font": this.m_font
            };
        };
        return ChartLegendFormat;
    }(OfficeExtension.ClientObject));
    Excel.ChartLegendFormat = ChartLegendFormat;
    var ChartTitle = (function (_super) {
        __extends(ChartTitle, _super);
        function ChartTitle() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ChartTitle.prototype, "_className", {
            get: function () {
                return "ChartTitle";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartTitle.prototype, "format", {
            get: function () {
                if (!this.m_format) {
                    this.m_format = new Excel.ChartTitleFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
                }
                return this.m_format;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartTitle.prototype, "horizontalAlignment", {
            get: function () {
                _throwIfNotLoaded("horizontalAlignment", this.m_horizontalAlignment, "ChartTitle", this._isNull);
                _throwIfApiNotSupported("ChartTitle.horizontalAlignment", _defaultApiSetName, "1.8", _hostName);
                return this.m_horizontalAlignment;
            },
            set: function (value) {
                this.m_horizontalAlignment = value;
                _createSetPropertyAction(this.context, this, "HorizontalAlignment", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartTitle.prototype, "overlay", {
            get: function () {
                _throwIfNotLoaded("overlay", this.m_overlay, "ChartTitle", this._isNull);
                return this.m_overlay;
            },
            set: function (value) {
                this.m_overlay = value;
                _createSetPropertyAction(this.context, this, "Overlay", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartTitle.prototype, "text", {
            get: function () {
                _throwIfNotLoaded("text", this.m_text, "ChartTitle", this._isNull);
                return this.m_text;
            },
            set: function (value) {
                this.m_text = value;
                _createSetPropertyAction(this.context, this, "Text", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartTitle.prototype, "visible", {
            get: function () {
                _throwIfNotLoaded("visible", this.m_visible, "ChartTitle", this._isNull);
                return this.m_visible;
            },
            set: function (value) {
                this.m_visible = value;
                _createSetPropertyAction(this.context, this, "Visible", value);
            },
            enumerable: true,
            configurable: true
        });
        ChartTitle.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["visible", "text", "overlay", "horizontalAlignment"], ["format"], []);
        };
        ChartTitle.prototype.getSubstring = function (start, length) {
            _throwIfApiNotSupported("ChartTitle.getSubstring", _defaultApiSetName, "1.8", _hostName);
            return new Excel.ChartFormatString(this.context, _createMethodObjectPath(this.context, this, "GetSubstring", 1, [start, length], false, false, null));
        };
        ChartTitle.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["HorizontalAlignment"])) {
                this.m_horizontalAlignment = obj["HorizontalAlignment"];
            }
            if (!_isUndefined(obj["Overlay"])) {
                this.m_overlay = obj["Overlay"];
            }
            if (!_isUndefined(obj["Text"])) {
                this.m_text = obj["Text"];
            }
            if (!_isUndefined(obj["Visible"])) {
                this.m_visible = obj["Visible"];
            }
            _handleNavigationPropertyResults(this, obj, ["format", "Format"]);
        };
        ChartTitle.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ChartTitle.prototype.toJSON = function () {
            return {
                "format": this.m_format,
                "horizontalAlignment": this.m_horizontalAlignment,
                "overlay": this.m_overlay,
                "text": this.m_text,
                "visible": this.m_visible
            };
        };
        return ChartTitle;
    }(OfficeExtension.ClientObject));
    Excel.ChartTitle = ChartTitle;
    var ChartFormatString = (function (_super) {
        __extends(ChartFormatString, _super);
        function ChartFormatString() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ChartFormatString.prototype, "_className", {
            get: function () {
                return "ChartFormatString";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartFormatString.prototype, "font", {
            get: function () {
                if (!this.m_font) {
                    this.m_font = new Excel.ChartFont(this.context, _createPropertyObjectPath(this.context, this, "Font", false, false));
                }
                return this.m_font;
            },
            enumerable: true,
            configurable: true
        });
        ChartFormatString.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            _handleNavigationPropertyResults(this, obj, ["font", "Font"]);
        };
        ChartFormatString.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ChartFormatString.prototype.toJSON = function () {
            return {};
        };
        return ChartFormatString;
    }(OfficeExtension.ClientObject));
    Excel.ChartFormatString = ChartFormatString;
    var ChartTitleFormat = (function (_super) {
        __extends(ChartTitleFormat, _super);
        function ChartTitleFormat() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ChartTitleFormat.prototype, "_className", {
            get: function () {
                return "ChartTitleFormat";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartTitleFormat.prototype, "fill", {
            get: function () {
                if (!this.m_fill) {
                    this.m_fill = new Excel.ChartFill(this.context, _createPropertyObjectPath(this.context, this, "Fill", false, false));
                }
                return this.m_fill;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartTitleFormat.prototype, "font", {
            get: function () {
                if (!this.m_font) {
                    this.m_font = new Excel.ChartFont(this.context, _createPropertyObjectPath(this.context, this, "Font", false, false));
                }
                return this.m_font;
            },
            enumerable: true,
            configurable: true
        });
        ChartTitleFormat.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, [], ["font"], [
                "fill"
            ]);
        };
        ChartTitleFormat.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            _handleNavigationPropertyResults(this, obj, ["fill", "Fill", "font", "Font"]);
        };
        ChartTitleFormat.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ChartTitleFormat.prototype.toJSON = function () {
            return {
                "fill": this.m_fill,
                "font": this.m_font
            };
        };
        return ChartTitleFormat;
    }(OfficeExtension.ClientObject));
    Excel.ChartTitleFormat = ChartTitleFormat;
    var ChartFill = (function (_super) {
        __extends(ChartFill, _super);
        function ChartFill() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ChartFill.prototype, "_className", {
            get: function () {
                return "ChartFill";
            },
            enumerable: true,
            configurable: true
        });
        ChartFill.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ChartFill.prototype.clear = function () {
            _createMethodAction(this.context, this, "Clear", 0, []);
        };
        ChartFill.prototype.setSolidColor = function (color) {
            _createMethodAction(this.context, this, "SetSolidColor", 0, [color]);
        };
        ChartFill.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
        };
        ChartFill.prototype.toJSON = function () {
            return {};
        };
        return ChartFill;
    }(OfficeExtension.ClientObject));
    Excel.ChartFill = ChartFill;
    var ChartLineFormat = (function (_super) {
        __extends(ChartLineFormat, _super);
        function ChartLineFormat() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ChartLineFormat.prototype, "_className", {
            get: function () {
                return "ChartLineFormat";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartLineFormat.prototype, "color", {
            get: function () {
                _throwIfNotLoaded("color", this.m_color, "ChartLineFormat", this._isNull);
                return this.m_color;
            },
            set: function (value) {
                this.m_color = value;
                _createSetPropertyAction(this.context, this, "Color", value);
            },
            enumerable: true,
            configurable: true
        });
        ChartLineFormat.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["color"], [], []);
        };
        ChartLineFormat.prototype.clear = function () {
            _createMethodAction(this.context, this, "Clear", 0, []);
        };
        ChartLineFormat.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Color"])) {
                this.m_color = obj["Color"];
            }
        };
        ChartLineFormat.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ChartLineFormat.prototype.toJSON = function () {
            return {
                "color": this.m_color
            };
        };
        return ChartLineFormat;
    }(OfficeExtension.ClientObject));
    Excel.ChartLineFormat = ChartLineFormat;
    var ChartFont = (function (_super) {
        __extends(ChartFont, _super);
        function ChartFont() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ChartFont.prototype, "_className", {
            get: function () {
                return "ChartFont";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartFont.prototype, "bold", {
            get: function () {
                _throwIfNotLoaded("bold", this.m_bold, "ChartFont", this._isNull);
                return this.m_bold;
            },
            set: function (value) {
                this.m_bold = value;
                _createSetPropertyAction(this.context, this, "Bold", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartFont.prototype, "color", {
            get: function () {
                _throwIfNotLoaded("color", this.m_color, "ChartFont", this._isNull);
                return this.m_color;
            },
            set: function (value) {
                this.m_color = value;
                _createSetPropertyAction(this.context, this, "Color", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartFont.prototype, "italic", {
            get: function () {
                _throwIfNotLoaded("italic", this.m_italic, "ChartFont", this._isNull);
                return this.m_italic;
            },
            set: function (value) {
                this.m_italic = value;
                _createSetPropertyAction(this.context, this, "Italic", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartFont.prototype, "name", {
            get: function () {
                _throwIfNotLoaded("name", this.m_name, "ChartFont", this._isNull);
                return this.m_name;
            },
            set: function (value) {
                this.m_name = value;
                _createSetPropertyAction(this.context, this, "Name", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartFont.prototype, "size", {
            get: function () {
                _throwIfNotLoaded("size", this.m_size, "ChartFont", this._isNull);
                return this.m_size;
            },
            set: function (value) {
                this.m_size = value;
                _createSetPropertyAction(this.context, this, "Size", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartFont.prototype, "underline", {
            get: function () {
                _throwIfNotLoaded("underline", this.m_underline, "ChartFont", this._isNull);
                return this.m_underline;
            },
            set: function (value) {
                this.m_underline = value;
                _createSetPropertyAction(this.context, this, "Underline", value);
            },
            enumerable: true,
            configurable: true
        });
        ChartFont.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["bold", "color", "italic", "name", "size", "underline"], [], []);
        };
        ChartFont.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Bold"])) {
                this.m_bold = obj["Bold"];
            }
            if (!_isUndefined(obj["Color"])) {
                this.m_color = obj["Color"];
            }
            if (!_isUndefined(obj["Italic"])) {
                this.m_italic = obj["Italic"];
            }
            if (!_isUndefined(obj["Name"])) {
                this.m_name = obj["Name"];
            }
            if (!_isUndefined(obj["Size"])) {
                this.m_size = obj["Size"];
            }
            if (!_isUndefined(obj["Underline"])) {
                this.m_underline = obj["Underline"];
            }
        };
        ChartFont.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ChartFont.prototype.toJSON = function () {
            return {
                "bold": this.m_bold,
                "color": this.m_color,
                "italic": this.m_italic,
                "name": this.m_name,
                "size": this.m_size,
                "underline": this.m_underline
            };
        };
        return ChartFont;
    }(OfficeExtension.ClientObject));
    Excel.ChartFont = ChartFont;
    var ChartTrendline = (function (_super) {
        __extends(ChartTrendline, _super);
        function ChartTrendline() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ChartTrendline.prototype, "_className", {
            get: function () {
                return "ChartTrendline";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartTrendline.prototype, "format", {
            get: function () {
                if (!this.m_format) {
                    this.m_format = new Excel.ChartTrendlineFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
                }
                return this.m_format;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartTrendline.prototype, "movingAveragePeriod", {
            get: function () {
                _throwIfNotLoaded("movingAveragePeriod", this.m_movingAveragePeriod, "ChartTrendline", this._isNull);
                return this.m_movingAveragePeriod;
            },
            set: function (value) {
                this.m_movingAveragePeriod = value;
                _createSetPropertyAction(this.context, this, "MovingAveragePeriod", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartTrendline.prototype, "polynomialOrder", {
            get: function () {
                _throwIfNotLoaded("polynomialOrder", this.m_polynomialOrder, "ChartTrendline", this._isNull);
                return this.m_polynomialOrder;
            },
            set: function (value) {
                this.m_polynomialOrder = value;
                _createSetPropertyAction(this.context, this, "PolynomialOrder", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartTrendline.prototype, "type", {
            get: function () {
                _throwIfNotLoaded("type", this.m_type, "ChartTrendline", this._isNull);
                return this.m_type;
            },
            set: function (value) {
                this.m_type = value;
                _createSetPropertyAction(this.context, this, "Type", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartTrendline.prototype, "_Id", {
            get: function () {
                _throwIfNotLoaded("_Id", this.m__Id, "ChartTrendline", this._isNull);
                return this.m__Id;
            },
            enumerable: true,
            configurable: true
        });
        ChartTrendline.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["type", "polynomialOrder", "movingAveragePeriod"], ["format"], []);
        };
        ChartTrendline.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["MovingAveragePeriod"])) {
                this.m_movingAveragePeriod = obj["MovingAveragePeriod"];
            }
            if (!_isUndefined(obj["PolynomialOrder"])) {
                this.m_polynomialOrder = obj["PolynomialOrder"];
            }
            if (!_isUndefined(obj["Type"])) {
                this.m_type = obj["Type"];
            }
            if (!_isUndefined(obj["_Id"])) {
                this.m__Id = obj["_Id"];
            }
            _handleNavigationPropertyResults(this, obj, ["format", "Format"]);
        };
        ChartTrendline.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ChartTrendline.prototype._handleIdResult = function (value) {
            _super.prototype._handleIdResult.call(this, value);
            if (_isNullOrUndefined(value)) {
                return;
            }
            if (!_isUndefined(value["_Id"])) {
                this.m__Id = value["_Id"];
            }
        };
        ChartTrendline.prototype.toJSON = function () {
            return {
                "format": this.m_format,
                "movingAveragePeriod": this.m_movingAveragePeriod,
                "polynomialOrder": this.m_polynomialOrder,
                "type": this.m_type
            };
        };
        return ChartTrendline;
    }(OfficeExtension.ClientObject));
    Excel.ChartTrendline = ChartTrendline;
    var ChartTrendlineCollection = (function (_super) {
        __extends(ChartTrendlineCollection, _super);
        function ChartTrendlineCollection() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ChartTrendlineCollection.prototype, "_className", {
            get: function () {
                return "ChartTrendlineCollection";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartTrendlineCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, "ChartTrendlineCollection", this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        ChartTrendlineCollection.prototype.add = function (type) {
            return new Excel.ChartTrendline(this.context, _createMethodObjectPath(this.context, this, "Add", 0, [type], false, true, null));
        };
        ChartTrendlineCollection.prototype.getCount = function () {
            var action = _createMethodAction(this.context, this, "GetCount", 1, []);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        ChartTrendlineCollection.prototype.getItem = function (index) {
            return new Excel.ChartTrendline(this.context, _createIndexerObjectPath(this.context, this, [index]));
        };
        ChartTrendlineCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.ChartTrendline(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        ChartTrendlineCollection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ChartTrendlineCollection.prototype.toJSON = function () {
            return {};
        };
        return ChartTrendlineCollection;
    }(OfficeExtension.ClientObject));
    Excel.ChartTrendlineCollection = ChartTrendlineCollection;
    var ChartTrendlineFormat = (function (_super) {
        __extends(ChartTrendlineFormat, _super);
        function ChartTrendlineFormat() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ChartTrendlineFormat.prototype, "_className", {
            get: function () {
                return "ChartTrendlineFormat";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartTrendlineFormat.prototype, "line", {
            get: function () {
                if (!this.m_line) {
                    this.m_line = new Excel.ChartLineFormat(this.context, _createPropertyObjectPath(this.context, this, "Line", false, false));
                }
                return this.m_line;
            },
            enumerable: true,
            configurable: true
        });
        ChartTrendlineFormat.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, [], ["line"], []);
        };
        ChartTrendlineFormat.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            _handleNavigationPropertyResults(this, obj, ["line", "Line"]);
        };
        ChartTrendlineFormat.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ChartTrendlineFormat.prototype.toJSON = function () {
            return {
                "line": this.m_line
            };
        };
        return ChartTrendlineFormat;
    }(OfficeExtension.ClientObject));
    Excel.ChartTrendlineFormat = ChartTrendlineFormat;
    var RangeSort = (function (_super) {
        __extends(RangeSort, _super);
        function RangeSort() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(RangeSort.prototype, "_className", {
            get: function () {
                return "RangeSort";
            },
            enumerable: true,
            configurable: true
        });
        RangeSort.prototype.apply = function (fields, matchCase, hasHeaders, orientation, method) {
            _createMethodAction(this.context, this, "Apply", 0, [fields, matchCase, hasHeaders, orientation, method]);
        };
        RangeSort.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
        };
        RangeSort.prototype.toJSON = function () {
            return {};
        };
        return RangeSort;
    }(OfficeExtension.ClientObject));
    Excel.RangeSort = RangeSort;
    var TableSort = (function (_super) {
        __extends(TableSort, _super);
        function TableSort() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(TableSort.prototype, "_className", {
            get: function () {
                return "TableSort";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(TableSort.prototype, "fields", {
            get: function () {
                _throwIfNotLoaded("fields", this.m_fields, "TableSort", this._isNull);
                return this.m_fields;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(TableSort.prototype, "matchCase", {
            get: function () {
                _throwIfNotLoaded("matchCase", this.m_matchCase, "TableSort", this._isNull);
                return this.m_matchCase;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(TableSort.prototype, "method", {
            get: function () {
                _throwIfNotLoaded("method", this.m_method, "TableSort", this._isNull);
                return this.m_method;
            },
            enumerable: true,
            configurable: true
        });
        TableSort.prototype.apply = function (fields, matchCase, method) {
            _createMethodAction(this.context, this, "Apply", 0, [fields, matchCase, method]);
        };
        TableSort.prototype.clear = function () {
            _createMethodAction(this.context, this, "Clear", 0, []);
        };
        TableSort.prototype.reapply = function () {
            _createMethodAction(this.context, this, "Reapply", 0, []);
        };
        TableSort.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Fields"])) {
                this.m_fields = obj["Fields"];
            }
            if (!_isUndefined(obj["MatchCase"])) {
                this.m_matchCase = obj["MatchCase"];
            }
            if (!_isUndefined(obj["Method"])) {
                this.m_method = obj["Method"];
            }
        };
        TableSort.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        TableSort.prototype.toJSON = function () {
            return {
                "fields": this.m_fields,
                "matchCase": this.m_matchCase,
                "method": this.m_method
            };
        };
        return TableSort;
    }(OfficeExtension.ClientObject));
    Excel.TableSort = TableSort;
    var Filter = (function (_super) {
        __extends(Filter, _super);
        function Filter() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(Filter.prototype, "_className", {
            get: function () {
                return "Filter";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Filter.prototype, "criteria", {
            get: function () {
                _throwIfNotLoaded("criteria", this.m_criteria, "Filter", this._isNull);
                return this.m_criteria;
            },
            enumerable: true,
            configurable: true
        });
        Filter.prototype.apply = function (criteria) {
            _createMethodAction(this.context, this, "Apply", 0, [criteria]);
        };
        Filter.prototype.applyBottomItemsFilter = function (count) {
            _createMethodAction(this.context, this, "ApplyBottomItemsFilter", 0, [count]);
        };
        Filter.prototype.applyBottomPercentFilter = function (percent) {
            _createMethodAction(this.context, this, "ApplyBottomPercentFilter", 0, [percent]);
        };
        Filter.prototype.applyCellColorFilter = function (color) {
            _createMethodAction(this.context, this, "ApplyCellColorFilter", 0, [color]);
        };
        Filter.prototype.applyCustomFilter = function (criteria1, criteria2, oper) {
            _createMethodAction(this.context, this, "ApplyCustomFilter", 0, [criteria1, criteria2, oper]);
        };
        Filter.prototype.applyDynamicFilter = function (criteria) {
            _createMethodAction(this.context, this, "ApplyDynamicFilter", 0, [criteria]);
        };
        Filter.prototype.applyFontColorFilter = function (color) {
            _createMethodAction(this.context, this, "ApplyFontColorFilter", 0, [color]);
        };
        Filter.prototype.applyIconFilter = function (icon) {
            _createMethodAction(this.context, this, "ApplyIconFilter", 0, [icon]);
        };
        Filter.prototype.applyTopItemsFilter = function (count) {
            _createMethodAction(this.context, this, "ApplyTopItemsFilter", 0, [count]);
        };
        Filter.prototype.applyTopPercentFilter = function (percent) {
            _createMethodAction(this.context, this, "ApplyTopPercentFilter", 0, [percent]);
        };
        Filter.prototype.applyValuesFilter = function (values) {
            _createMethodAction(this.context, this, "ApplyValuesFilter", 0, [values]);
        };
        Filter.prototype.clear = function () {
            _createMethodAction(this.context, this, "Clear", 0, []);
        };
        Filter.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Criteria"])) {
                this.m_criteria = obj["Criteria"];
            }
        };
        Filter.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        Filter.prototype.toJSON = function () {
            return {
                "criteria": this.m_criteria
            };
        };
        return Filter;
    }(OfficeExtension.ClientObject));
    Excel.Filter = Filter;
    var CustomXmlPartScopedCollection = (function (_super) {
        __extends(CustomXmlPartScopedCollection, _super);
        function CustomXmlPartScopedCollection() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(CustomXmlPartScopedCollection.prototype, "_className", {
            get: function () {
                return "CustomXmlPartScopedCollection";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(CustomXmlPartScopedCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, "CustomXmlPartScopedCollection", this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        CustomXmlPartScopedCollection.prototype.getCount = function () {
            var action = _createMethodAction(this.context, this, "GetCount", 1, []);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        CustomXmlPartScopedCollection.prototype.getItem = function (id) {
            return new Excel.CustomXmlPart(this.context, _createIndexerObjectPath(this.context, this, [id]));
        };
        CustomXmlPartScopedCollection.prototype.getItemOrNullObject = function (id) {
            return new Excel.CustomXmlPart(this.context, _createMethodObjectPath(this.context, this, "GetItemOrNullObject", 1, [id], false, false, null));
        };
        CustomXmlPartScopedCollection.prototype.getOnlyItem = function () {
            return new Excel.CustomXmlPart(this.context, _createMethodObjectPath(this.context, this, "GetOnlyItem", 1, [], false, false, null));
        };
        CustomXmlPartScopedCollection.prototype.getOnlyItemOrNullObject = function () {
            return new Excel.CustomXmlPart(this.context, _createMethodObjectPath(this.context, this, "GetOnlyItemOrNullObject", 1, [], false, false, null));
        };
        CustomXmlPartScopedCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.CustomXmlPart(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        CustomXmlPartScopedCollection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        CustomXmlPartScopedCollection.prototype.toJSON = function () {
            return {};
        };
        return CustomXmlPartScopedCollection;
    }(OfficeExtension.ClientObject));
    Excel.CustomXmlPartScopedCollection = CustomXmlPartScopedCollection;
    var CustomXmlPartCollection = (function (_super) {
        __extends(CustomXmlPartCollection, _super);
        function CustomXmlPartCollection() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(CustomXmlPartCollection.prototype, "_className", {
            get: function () {
                return "CustomXmlPartCollection";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(CustomXmlPartCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, "CustomXmlPartCollection", this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        CustomXmlPartCollection.prototype.add = function (xml) {
            return new Excel.CustomXmlPart(this.context, _createMethodObjectPath(this.context, this, "Add", 0, [xml], false, true, null));
        };
        CustomXmlPartCollection.prototype.getByNamespace = function (namespaceUri) {
            return new Excel.CustomXmlPartScopedCollection(this.context, _createMethodObjectPath(this.context, this, "GetByNamespace", 1, [namespaceUri], true, false, null));
        };
        CustomXmlPartCollection.prototype.getCount = function () {
            var action = _createMethodAction(this.context, this, "GetCount", 1, []);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        CustomXmlPartCollection.prototype.getItem = function (id) {
            return new Excel.CustomXmlPart(this.context, _createIndexerObjectPath(this.context, this, [id]));
        };
        CustomXmlPartCollection.prototype.getItemOrNullObject = function (id) {
            return new Excel.CustomXmlPart(this.context, _createMethodObjectPath(this.context, this, "GetItemOrNullObject", 1, [id], false, false, null));
        };
        CustomXmlPartCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.CustomXmlPart(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        CustomXmlPartCollection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        CustomXmlPartCollection.prototype.toJSON = function () {
            return {};
        };
        return CustomXmlPartCollection;
    }(OfficeExtension.ClientObject));
    Excel.CustomXmlPartCollection = CustomXmlPartCollection;
    var CustomXmlPart = (function (_super) {
        __extends(CustomXmlPart, _super);
        function CustomXmlPart() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(CustomXmlPart.prototype, "_className", {
            get: function () {
                return "CustomXmlPart";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(CustomXmlPart.prototype, "id", {
            get: function () {
                _throwIfNotLoaded("id", this.m_id, "CustomXmlPart", this._isNull);
                return this.m_id;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(CustomXmlPart.prototype, "namespaceUri", {
            get: function () {
                _throwIfNotLoaded("namespaceUri", this.m_namespaceUri, "CustomXmlPart", this._isNull);
                return this.m_namespaceUri;
            },
            enumerable: true,
            configurable: true
        });
        CustomXmlPart.prototype.delete = function () {
            _createMethodAction(this.context, this, "Delete", 0, []);
        };
        CustomXmlPart.prototype.getXml = function () {
            var action = _createMethodAction(this.context, this, "GetXml", 1, []);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        CustomXmlPart.prototype.setXml = function (xml) {
            _createMethodAction(this.context, this, "SetXml", 0, [xml]);
        };
        CustomXmlPart.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Id"])) {
                this.m_id = obj["Id"];
            }
            if (!_isUndefined(obj["NamespaceUri"])) {
                this.m_namespaceUri = obj["NamespaceUri"];
            }
        };
        CustomXmlPart.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        CustomXmlPart.prototype._handleIdResult = function (value) {
            _super.prototype._handleIdResult.call(this, value);
            if (_isNullOrUndefined(value)) {
                return;
            }
            if (!_isUndefined(value["Id"])) {
                this.m_id = value["Id"];
            }
        };
        CustomXmlPart.prototype.toJSON = function () {
            return {
                "id": this.m_id,
                "namespaceUri": this.m_namespaceUri
            };
        };
        return CustomXmlPart;
    }(OfficeExtension.ClientObject));
    Excel.CustomXmlPart = CustomXmlPart;
    var _V1Api = (function (_super) {
        __extends(_V1Api, _super);
        function _V1Api() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(_V1Api.prototype, "_className", {
            get: function () {
                return "_V1Api";
            },
            enumerable: true,
            configurable: true
        });
        _V1Api.prototype.bindingAddColumns = function (input) {
            var action = _createMethodAction(this.context, this, "BindingAddColumns", 0, [input]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        _V1Api.prototype.bindingAddFromNamedItem = function (input) {
            var action = _createMethodAction(this.context, this, "BindingAddFromNamedItem", 1, [input]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        _V1Api.prototype.bindingAddFromPrompt = function (input) {
            var action = _createMethodAction(this.context, this, "BindingAddFromPrompt", 1, [input]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        _V1Api.prototype.bindingAddFromSelection = function (input) {
            var action = _createMethodAction(this.context, this, "BindingAddFromSelection", 1, [input]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        _V1Api.prototype.bindingAddRows = function (input) {
            var action = _createMethodAction(this.context, this, "BindingAddRows", 0, [input]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        _V1Api.prototype.bindingClearFormats = function (input) {
            var action = _createMethodAction(this.context, this, "BindingClearFormats", 0, [input]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        _V1Api.prototype.bindingDeleteAllDataValues = function (input) {
            var action = _createMethodAction(this.context, this, "BindingDeleteAllDataValues", 0, [input]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        _V1Api.prototype.bindingGetAll = function () {
            var action = _createMethodAction(this.context, this, "BindingGetAll", 1, []);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        _V1Api.prototype.bindingGetById = function (input) {
            var action = _createMethodAction(this.context, this, "BindingGetById", 1, [input]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        _V1Api.prototype.bindingGetData = function (input) {
            var action = _createMethodAction(this.context, this, "BindingGetData", 1, [input]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        _V1Api.prototype.bindingReleaseById = function (input) {
            var action = _createMethodAction(this.context, this, "BindingReleaseById", 1, [input]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        _V1Api.prototype.bindingSetData = function (input) {
            var action = _createMethodAction(this.context, this, "BindingSetData", 0, [input]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        _V1Api.prototype.bindingSetFormats = function (input) {
            var action = _createMethodAction(this.context, this, "BindingSetFormats", 0, [input]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        _V1Api.prototype.bindingSetTableOptions = function (input) {
            var action = _createMethodAction(this.context, this, "BindingSetTableOptions", 0, [input]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        _V1Api.prototype.getFilePropertiesAsync = function () {
            _throwIfApiNotSupported("_V1Api.getFilePropertiesAsync", _defaultApiSetName, "1.6", _hostName);
            var action = _createMethodAction(this.context, this, "GetFilePropertiesAsync", 1, []);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        _V1Api.prototype.getSelectedData = function (input) {
            var action = _createMethodAction(this.context, this, "GetSelectedData", 1, [input]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        _V1Api.prototype.gotoById = function (input) {
            var action = _createMethodAction(this.context, this, "GotoById", 1, [input]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        _V1Api.prototype.setSelectedData = function (input) {
            var action = _createMethodAction(this.context, this, "SetSelectedData", 0, [input]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        _V1Api.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
        };
        _V1Api.prototype.toJSON = function () {
            return {};
        };
        return _V1Api;
    }(OfficeExtension.ClientObject));
    Excel._V1Api = _V1Api;
    var PivotTableCollection = (function (_super) {
        __extends(PivotTableCollection, _super);
        function PivotTableCollection() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(PivotTableCollection.prototype, "_className", {
            get: function () {
                return "PivotTableCollection";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTableCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, "PivotTableCollection", this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        PivotTableCollection.prototype.getCount = function () {
            _throwIfApiNotSupported("PivotTableCollection.getCount", _defaultApiSetName, "1.4", _hostName);
            var action = _createMethodAction(this.context, this, "GetCount", 1, []);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        PivotTableCollection.prototype.getItem = function (name) {
            return new Excel.PivotTable(this.context, _createIndexerObjectPath(this.context, this, [name]));
        };
        PivotTableCollection.prototype.getItemOrNullObject = function (name) {
            _throwIfApiNotSupported("PivotTableCollection.getItemOrNullObject", _defaultApiSetName, "1.4", _hostName);
            return new Excel.PivotTable(this.context, _createMethodObjectPath(this.context, this, "GetItemOrNullObject", 1, [name], false, false, null));
        };
        PivotTableCollection.prototype.refreshAll = function () {
            _createMethodAction(this.context, this, "RefreshAll", 0, []);
        };
        PivotTableCollection.prototype.add = function (name, address, pivotCache) {
            _throwIfApiNotSupported("PivotTableCollection.add", "Pivot", "1.1", _hostName);
            return new Excel.PivotTable(this.context, _createMethodObjectPath(this.context, this, "Add", 0, [name, address, pivotCache], false, true, null));
        };
        PivotTableCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.PivotTable(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        PivotTableCollection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        PivotTableCollection.prototype.toJSON = function () {
            return {};
        };
        return PivotTableCollection;
    }(OfficeExtension.ClientObject));
    Excel.PivotTableCollection = PivotTableCollection;
    var PivotTable = (function (_super) {
        __extends(PivotTable, _super);
        function PivotTable() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(PivotTable.prototype, "_className", {
            get: function () {
                return "PivotTable";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "worksheet", {
            get: function () {
                if (!this.m_worksheet) {
                    this.m_worksheet = new Excel.Worksheet(this.context, _createPropertyObjectPath(this.context, this, "Worksheet", false, false));
                }
                return this.m_worksheet;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "calculatedFields", {
            get: function () {
                _throwIfApiNotSupported("PivotTable.calculatedFields", "Pivot", "1.1", _hostName);
                if (!this.m_calculatedFields) {
                    this.m_calculatedFields = new Excel.CalculatedFieldCollection(this.context, _createPropertyObjectPath(this.context, this, "CalculatedFields", true, false));
                }
                return this.m_calculatedFields;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "dataBodyRange", {
            get: function () {
                _throwIfApiNotSupported("PivotTable.dataBodyRange", "Pivot", "1.1", _hostName);
                if (!this.m_dataBodyRange) {
                    this.m_dataBodyRange = new Excel.Range(this.context, _createPropertyObjectPath(this.context, this, "DataBodyRange", false, false));
                }
                return this.m_dataBodyRange;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "dataLabelRange", {
            get: function () {
                _throwIfApiNotSupported("PivotTable.dataLabelRange", "Pivot", "1.1", _hostName);
                if (!this.m_dataLabelRange) {
                    this.m_dataLabelRange = new Excel.Range(this.context, _createPropertyObjectPath(this.context, this, "DataLabelRange", false, false));
                }
                return this.m_dataLabelRange;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "pivotFields", {
            get: function () {
                _throwIfApiNotSupported("PivotTable.pivotFields", "Pivot", "1.1", _hostName);
                if (!this.m_pivotFields) {
                    this.m_pivotFields = new Excel.PivotFieldCollection(this.context, _createPropertyObjectPath(this.context, this, "PivotFields", true, false));
                }
                return this.m_pivotFields;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "columnGrandTotals", {
            get: function () {
                _throwIfNotLoaded("columnGrandTotals", this.m_columnGrandTotals, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.columnGrandTotals", _defaultApiSetName, "1.7", _hostName);
                return this.m_columnGrandTotals;
            },
            set: function (value) {
                this.m_columnGrandTotals = value;
                _createSetPropertyAction(this.context, this, "ColumnGrandTotals", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "id", {
            get: function () {
                _throwIfNotLoaded("id", this.m_id, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.id", _defaultApiSetName, "1.5", _hostName);
                return this.m_id;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "name", {
            get: function () {
                _throwIfNotLoaded("name", this.m_name, "PivotTable", this._isNull);
                return this.m_name;
            },
            set: function (value) {
                this.m_name = value;
                _createSetPropertyAction(this.context, this, "Name", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "rowGrandTotals", {
            get: function () {
                _throwIfNotLoaded("rowGrandTotals", this.m_rowGrandTotals, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.rowGrandTotals", _defaultApiSetName, "1.7", _hostName);
                return this.m_rowGrandTotals;
            },
            set: function (value) {
                this.m_rowGrandTotals = value;
                _createSetPropertyAction(this.context, this, "RowGrandTotals", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "allowMultipleFilters", {
            get: function () {
                _throwIfNotLoaded("allowMultipleFilters", this.m_allowMultipleFilters, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.allowMultipleFilters", "Pivot", "1.1", _hostName);
                return this.m_allowMultipleFilters;
            },
            set: function (value) {
                this.m_allowMultipleFilters = value;
                _createSetPropertyAction(this.context, this, "AllowMultipleFilters", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "alternativeText", {
            get: function () {
                _throwIfNotLoaded("alternativeText", this.m_alternativeText, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.alternativeText", "Pivot", "1.1", _hostName);
                return this.m_alternativeText;
            },
            set: function (value) {
                this.m_alternativeText = value;
                _createSetPropertyAction(this.context, this, "AlternativeText", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "compactLayoutColumnHeader", {
            get: function () {
                _throwIfNotLoaded("compactLayoutColumnHeader", this.m_compactLayoutColumnHeader, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.compactLayoutColumnHeader", "Pivot", "1.1", _hostName);
                return this.m_compactLayoutColumnHeader;
            },
            set: function (value) {
                this.m_compactLayoutColumnHeader = value;
                _createSetPropertyAction(this.context, this, "CompactLayoutColumnHeader", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "compactLayoutRowHeader", {
            get: function () {
                _throwIfNotLoaded("compactLayoutRowHeader", this.m_compactLayoutRowHeader, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.compactLayoutRowHeader", "Pivot", "1.1", _hostName);
                return this.m_compactLayoutRowHeader;
            },
            set: function (value) {
                this.m_compactLayoutRowHeader = value;
                _createSetPropertyAction(this.context, this, "CompactLayoutRowHeader", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "compactRowIndent", {
            get: function () {
                _throwIfNotLoaded("compactRowIndent", this.m_compactRowIndent, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.compactRowIndent", "Pivot", "1.1", _hostName);
                return this.m_compactRowIndent;
            },
            set: function (value) {
                this.m_compactRowIndent = value;
                _createSetPropertyAction(this.context, this, "CompactRowIndent", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "displayContextTooltips", {
            get: function () {
                _throwIfNotLoaded("displayContextTooltips", this.m_displayContextTooltips, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.displayContextTooltips", "Pivot", "1.1", _hostName);
                return this.m_displayContextTooltips;
            },
            set: function (value) {
                this.m_displayContextTooltips = value;
                _createSetPropertyAction(this.context, this, "DisplayContextTooltips", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "displayEmptyColumn", {
            get: function () {
                _throwIfNotLoaded("displayEmptyColumn", this.m_displayEmptyColumn, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.displayEmptyColumn", "Pivot", "1.1", _hostName);
                return this.m_displayEmptyColumn;
            },
            set: function (value) {
                this.m_displayEmptyColumn = value;
                _createSetPropertyAction(this.context, this, "DisplayEmptyColumn", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "displayEmptyRow", {
            get: function () {
                _throwIfNotLoaded("displayEmptyRow", this.m_displayEmptyRow, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.displayEmptyRow", "Pivot", "1.1", _hostName);
                return this.m_displayEmptyRow;
            },
            set: function (value) {
                this.m_displayEmptyRow = value;
                _createSetPropertyAction(this.context, this, "DisplayEmptyRow", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "displayErrorString", {
            get: function () {
                _throwIfNotLoaded("displayErrorString", this.m_displayErrorString, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.displayErrorString", "Pivot", "1.1", _hostName);
                return this.m_displayErrorString;
            },
            set: function (value) {
                this.m_displayErrorString = value;
                _createSetPropertyAction(this.context, this, "DisplayErrorString", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "displayFieldCaptions", {
            get: function () {
                _throwIfNotLoaded("displayFieldCaptions", this.m_displayFieldCaptions, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.displayFieldCaptions", "Pivot", "1.1", _hostName);
                return this.m_displayFieldCaptions;
            },
            set: function (value) {
                this.m_displayFieldCaptions = value;
                _createSetPropertyAction(this.context, this, "DisplayFieldCaptions", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "displayNullString", {
            get: function () {
                _throwIfNotLoaded("displayNullString", this.m_displayNullString, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.displayNullString", "Pivot", "1.1", _hostName);
                return this.m_displayNullString;
            },
            set: function (value) {
                this.m_displayNullString = value;
                _createSetPropertyAction(this.context, this, "DisplayNullString", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "enableDataValueEditing", {
            get: function () {
                _throwIfNotLoaded("enableDataValueEditing", this.m_enableDataValueEditing, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.enableDataValueEditing", "Pivot", "1.1", _hostName);
                return this.m_enableDataValueEditing;
            },
            set: function (value) {
                this.m_enableDataValueEditing = value;
                _createSetPropertyAction(this.context, this, "EnableDataValueEditing", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "enableDrilldown", {
            get: function () {
                _throwIfNotLoaded("enableDrilldown", this.m_enableDrilldown, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.enableDrilldown", "Pivot", "1.1", _hostName);
                return this.m_enableDrilldown;
            },
            set: function (value) {
                this.m_enableDrilldown = value;
                _createSetPropertyAction(this.context, this, "EnableDrilldown", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "enableFieldDialog", {
            get: function () {
                _throwIfNotLoaded("enableFieldDialog", this.m_enableFieldDialog, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.enableFieldDialog", "Pivot", "1.1", _hostName);
                return this.m_enableFieldDialog;
            },
            set: function (value) {
                this.m_enableFieldDialog = value;
                _createSetPropertyAction(this.context, this, "EnableFieldDialog", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "enableFieldList", {
            get: function () {
                _throwIfNotLoaded("enableFieldList", this.m_enableFieldList, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.enableFieldList", "Pivot", "1.1", _hostName);
                return this.m_enableFieldList;
            },
            set: function (value) {
                this.m_enableFieldList = value;
                _createSetPropertyAction(this.context, this, "EnableFieldList", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "enableWizard", {
            get: function () {
                _throwIfNotLoaded("enableWizard", this.m_enableWizard, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.enableWizard", "Pivot", "1.1", _hostName);
                return this.m_enableWizard;
            },
            set: function (value) {
                this.m_enableWizard = value;
                _createSetPropertyAction(this.context, this, "EnableWizard", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "errorString", {
            get: function () {
                _throwIfNotLoaded("errorString", this.m_errorString, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.errorString", "Pivot", "1.1", _hostName);
                return this.m_errorString;
            },
            set: function (value) {
                this.m_errorString = value;
                _createSetPropertyAction(this.context, this, "ErrorString", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "fieldListSortAscending", {
            get: function () {
                _throwIfNotLoaded("fieldListSortAscending", this.m_fieldListSortAscending, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.fieldListSortAscending", "Pivot", "1.1", _hostName);
                return this.m_fieldListSortAscending;
            },
            set: function (value) {
                this.m_fieldListSortAscending = value;
                _createSetPropertyAction(this.context, this, "FieldListSortAscending", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "grandTotalName", {
            get: function () {
                _throwIfNotLoaded("grandTotalName", this.m_grandTotalName, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.grandTotalName", "Pivot", "1.1", _hostName);
                return this.m_grandTotalName;
            },
            set: function (value) {
                this.m_grandTotalName = value;
                _createSetPropertyAction(this.context, this, "GrandTotalName", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "hasAutoFormat", {
            get: function () {
                _throwIfNotLoaded("hasAutoFormat", this.m_hasAutoFormat, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.hasAutoFormat", "Pivot", "1.1", _hostName);
                return this.m_hasAutoFormat;
            },
            set: function (value) {
                this.m_hasAutoFormat = value;
                _createSetPropertyAction(this.context, this, "HasAutoFormat", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "hidden", {
            get: function () {
                _throwIfNotLoaded("hidden", this.m_hidden, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.hidden", "Pivot", "1.1", _hostName);
                return this.m_hidden;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "inGridDropZones", {
            get: function () {
                _throwIfNotLoaded("inGridDropZones", this.m_inGridDropZones, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.inGridDropZones", "Pivot", "1.1", _hostName);
                return this.m_inGridDropZones;
            },
            set: function (value) {
                this.m_inGridDropZones = value;
                _createSetPropertyAction(this.context, this, "InGridDropZones", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "innerDetail", {
            get: function () {
                _throwIfNotLoaded("innerDetail", this.m_innerDetail, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.innerDetail", "Pivot", "1.1", _hostName);
                return this.m_innerDetail;
            },
            set: function (value) {
                this.m_innerDetail = value;
                _createSetPropertyAction(this.context, this, "InnerDetail", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "manualUpdate", {
            get: function () {
                _throwIfNotLoaded("manualUpdate", this.m_manualUpdate, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.manualUpdate", "Pivot", "1.1", _hostName);
                return this.m_manualUpdate;
            },
            set: function (value) {
                this.m_manualUpdate = value;
                _createSetPropertyAction(this.context, this, "ManualUpdate", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "nullString", {
            get: function () {
                _throwIfNotLoaded("nullString", this.m_nullString, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.nullString", "Pivot", "1.1", _hostName);
                return this.m_nullString;
            },
            set: function (value) {
                this.m_nullString = value;
                _createSetPropertyAction(this.context, this, "NullString", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "preserveFormatting", {
            get: function () {
                _throwIfNotLoaded("preserveFormatting", this.m_preserveFormatting, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.preserveFormatting", "Pivot", "1.1", _hostName);
                return this.m_preserveFormatting;
            },
            set: function (value) {
                this.m_preserveFormatting = value;
                _createSetPropertyAction(this.context, this, "PreserveFormatting", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "printDrillIndicators", {
            get: function () {
                _throwIfNotLoaded("printDrillIndicators", this.m_printDrillIndicators, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.printDrillIndicators", "Pivot", "1.1", _hostName);
                return this.m_printDrillIndicators;
            },
            set: function (value) {
                this.m_printDrillIndicators = value;
                _createSetPropertyAction(this.context, this, "PrintDrillIndicators", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "printTitles", {
            get: function () {
                _throwIfNotLoaded("printTitles", this.m_printTitles, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.printTitles", "Pivot", "1.1", _hostName);
                return this.m_printTitles;
            },
            set: function (value) {
                this.m_printTitles = value;
                _createSetPropertyAction(this.context, this, "PrintTitles", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "refreshDate", {
            get: function () {
                _throwIfNotLoaded("refreshDate", this.m_refreshDate, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.refreshDate", "Pivot", "1.1", _hostName);
                return this.m_refreshDate;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "refreshName", {
            get: function () {
                _throwIfNotLoaded("refreshName", this.m_refreshName, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.refreshName", "Pivot", "1.1", _hostName);
                return this.m_refreshName;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "repeatItemsOnEachPrintedPage", {
            get: function () {
                _throwIfNotLoaded("repeatItemsOnEachPrintedPage", this.m_repeatItemsOnEachPrintedPage, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.repeatItemsOnEachPrintedPage", "Pivot", "1.1", _hostName);
                return this.m_repeatItemsOnEachPrintedPage;
            },
            set: function (value) {
                this.m_repeatItemsOnEachPrintedPage = value;
                _createSetPropertyAction(this.context, this, "RepeatItemsOnEachPrintedPage", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "saveData", {
            get: function () {
                _throwIfNotLoaded("saveData", this.m_saveData, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.saveData", "Pivot", "1.1", _hostName);
                return this.m_saveData;
            },
            set: function (value) {
                this.m_saveData = value;
                _createSetPropertyAction(this.context, this, "SaveData", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "showDrillIndicators", {
            get: function () {
                _throwIfNotLoaded("showDrillIndicators", this.m_showDrillIndicators, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.showDrillIndicators", "Pivot", "1.1", _hostName);
                return this.m_showDrillIndicators;
            },
            set: function (value) {
                this.m_showDrillIndicators = value;
                _createSetPropertyAction(this.context, this, "ShowDrillIndicators", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "showPageMultipleItemLabel", {
            get: function () {
                _throwIfNotLoaded("showPageMultipleItemLabel", this.m_showPageMultipleItemLabel, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.showPageMultipleItemLabel", "Pivot", "1.1", _hostName);
                return this.m_showPageMultipleItemLabel;
            },
            set: function (value) {
                this.m_showPageMultipleItemLabel = value;
                _createSetPropertyAction(this.context, this, "ShowPageMultipleItemLabel", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "showTableStyleColumnHeaders", {
            get: function () {
                _throwIfNotLoaded("showTableStyleColumnHeaders", this.m_showTableStyleColumnHeaders, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.showTableStyleColumnHeaders", "Pivot", "1.1", _hostName);
                return this.m_showTableStyleColumnHeaders;
            },
            set: function (value) {
                this.m_showTableStyleColumnHeaders = value;
                _createSetPropertyAction(this.context, this, "ShowTableStyleColumnHeaders", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "showTableStyleColumnStripes", {
            get: function () {
                _throwIfNotLoaded("showTableStyleColumnStripes", this.m_showTableStyleColumnStripes, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.showTableStyleColumnStripes", "Pivot", "1.1", _hostName);
                return this.m_showTableStyleColumnStripes;
            },
            set: function (value) {
                this.m_showTableStyleColumnStripes = value;
                _createSetPropertyAction(this.context, this, "ShowTableStyleColumnStripes", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "showTableStyleLastColumn", {
            get: function () {
                _throwIfNotLoaded("showTableStyleLastColumn", this.m_showTableStyleLastColumn, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.showTableStyleLastColumn", "Pivot", "1.1", _hostName);
                return this.m_showTableStyleLastColumn;
            },
            set: function (value) {
                this.m_showTableStyleLastColumn = value;
                _createSetPropertyAction(this.context, this, "ShowTableStyleLastColumn", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "showTableStyleRowHeaders", {
            get: function () {
                _throwIfNotLoaded("showTableStyleRowHeaders", this.m_showTableStyleRowHeaders, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.showTableStyleRowHeaders", "Pivot", "1.1", _hostName);
                return this.m_showTableStyleRowHeaders;
            },
            set: function (value) {
                this.m_showTableStyleRowHeaders = value;
                _createSetPropertyAction(this.context, this, "ShowTableStyleRowHeaders", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "showTableStyleRowStripes", {
            get: function () {
                _throwIfNotLoaded("showTableStyleRowStripes", this.m_showTableStyleRowStripes, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.showTableStyleRowStripes", "Pivot", "1.1", _hostName);
                return this.m_showTableStyleRowStripes;
            },
            set: function (value) {
                this.m_showTableStyleRowStripes = value;
                _createSetPropertyAction(this.context, this, "ShowTableStyleRowStripes", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "showValuesRow", {
            get: function () {
                _throwIfNotLoaded("showValuesRow", this.m_showValuesRow, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.showValuesRow", "Pivot", "1.1", _hostName);
                return this.m_showValuesRow;
            },
            set: function (value) {
                this.m_showValuesRow = value;
                _createSetPropertyAction(this.context, this, "ShowValuesRow", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "smallGrid", {
            get: function () {
                _throwIfNotLoaded("smallGrid", this.m_smallGrid, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.smallGrid", "Pivot", "1.1", _hostName);
                return this.m_smallGrid;
            },
            set: function (value) {
                this.m_smallGrid = value;
                _createSetPropertyAction(this.context, this, "SmallGrid", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "sortUsingCustomLists", {
            get: function () {
                _throwIfNotLoaded("sortUsingCustomLists", this.m_sortUsingCustomLists, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.sortUsingCustomLists", "Pivot", "1.1", _hostName);
                return this.m_sortUsingCustomLists;
            },
            set: function (value) {
                this.m_sortUsingCustomLists = value;
                _createSetPropertyAction(this.context, this, "SortUsingCustomLists", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "subtotalHiddenPageItems", {
            get: function () {
                _throwIfNotLoaded("subtotalHiddenPageItems", this.m_subtotalHiddenPageItems, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.subtotalHiddenPageItems", "Pivot", "1.1", _hostName);
                return this.m_subtotalHiddenPageItems;
            },
            set: function (value) {
                this.m_subtotalHiddenPageItems = value;
                _createSetPropertyAction(this.context, this, "SubtotalHiddenPageItems", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "summary", {
            get: function () {
                _throwIfNotLoaded("summary", this.m_summary, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.summary", "Pivot", "1.1", _hostName);
                return this.m_summary;
            },
            set: function (value) {
                this.m_summary = value;
                _createSetPropertyAction(this.context, this, "Summary", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "tag", {
            get: function () {
                _throwIfNotLoaded("tag", this.m_tag, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.tag", "Pivot", "1.1", _hostName);
                return this.m_tag;
            },
            set: function (value) {
                this.m_tag = value;
                _createSetPropertyAction(this.context, this, "Tag", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "totalsAnnotation", {
            get: function () {
                _throwIfNotLoaded("totalsAnnotation", this.m_totalsAnnotation, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.totalsAnnotation", "Pivot", "1.1", _hostName);
                return this.m_totalsAnnotation;
            },
            set: function (value) {
                this.m_totalsAnnotation = value;
                _createSetPropertyAction(this.context, this, "TotalsAnnotation", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "vacatedStyle", {
            get: function () {
                _throwIfNotLoaded("vacatedStyle", this.m_vacatedStyle, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.vacatedStyle", "Pivot", "1.1", _hostName);
                return this.m_vacatedStyle;
            },
            set: function (value) {
                this.m_vacatedStyle = value;
                _createSetPropertyAction(this.context, this, "VacatedStyle", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "value", {
            get: function () {
                _throwIfNotLoaded("value", this.m_value, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.value", "Pivot", "1.1", _hostName);
                return this.m_value;
            },
            set: function (value) {
                this.m_value = value;
                _createSetPropertyAction(this.context, this, "Value", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "version", {
            get: function () {
                _throwIfNotLoaded("version", this.m_version, "PivotTable", this._isNull);
                _throwIfApiNotSupported("PivotTable.version", "Pivot", "1.1", _hostName);
                return this.m_version;
            },
            enumerable: true,
            configurable: true
        });
        PivotTable.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["name", "columnGrandTotals", "value", "rowGrandTotals", "summary", "saveData", "hasAutoFormat", "innerDetail", "displayErrorString", "displayNullString", "enableDrilldown", "enableFieldDialog", "enableWizard", "errorString", "manualUpdate", "nullString", "subtotalHiddenPageItems", "preserveFormatting", "tag", "vacatedStyle", "printTitles", "grandTotalName", "smallGrid", "repeatItemsOnEachPrintedPage", "totalsAnnotation", "alternativeText", "enableDataValueEditing", "enableFieldList", "showPageMultipleItemLabel", "displayEmptyRow", "displayEmptyColumn", "showDrillIndicators", "printDrillIndicators", "displayContextTooltips", "compactRowIndent", "displayFieldCaptions", "inGridDropZones", "showTableStyleLastColumn", "showTableStyleRowStripes", "showTableStyleColumnStripes", "showTableStyleRowHeaders", "showTableStyleColumnHeaders", "allowMultipleFilters", "compactLayoutRowHeader", "compactLayoutColumnHeader", "fieldListSortAscending", "sortUsingCustomLists", "showValuesRow"], [], [
                "worksheet",
                "calculatedFields",
                "dataBodyRange",
                "dataLabelRange",
                "pivotFields",
                "worksheet",
                "calculatedFields",
                "dataBodyRange",
                "dataLabelRange",
                "pivotFields"
            ]);
        };
        PivotTable.prototype.refresh = function () {
            _createMethodAction(this.context, this, "Refresh", 0, []);
        };
        PivotTable.prototype.rowAxisLayout = function (RowLayout) {
            _throwIfApiNotSupported("PivotTable.rowAxisLayout", _defaultApiSetName, "1.7", _hostName);
            _createMethodAction(this.context, this, "RowAxisLayout", 0, [RowLayout]);
        };
        PivotTable.prototype.subtotalLocation = function (Location) {
            _throwIfApiNotSupported("PivotTable.subtotalLocation", _defaultApiSetName, "1.7", _hostName);
            _createMethodAction(this.context, this, "SubtotalLocation", 0, [Location]);
        };
        PivotTable.prototype.addChart = function (chartType, seriesBy) {
            _throwIfApiNotSupported("PivotTable.addChart", "Pivot", "1.1", _hostName);
            return new Excel.Chart(this.context, _createMethodObjectPath(this.context, this, "AddChart", 0, [chartType, seriesBy], false, false, null));
        };
        PivotTable.prototype.addDataField = function (field, caption, func) {
            _throwIfApiNotSupported("PivotTable.addDataField", "Pivot", "1.1", _hostName);
            return new Excel.PivotField(this.context, _createMethodObjectPath(this.context, this, "AddDataField", 0, [field, caption, func], false, false, null));
        };
        PivotTable.prototype.clearTable = function () {
            _throwIfApiNotSupported("PivotTable.clearTable", "Pivot", "1.1", _hostName);
            _createMethodAction(this.context, this, "ClearTable", 0, []);
        };
        PivotTable.prototype.getColumnField = function (Index) {
            _throwIfApiNotSupported("PivotTable.getColumnField", "Pivot", "1.1", _hostName);
            return new Excel.PivotField(this.context, _createMethodObjectPath(this.context, this, "GetColumnField", 1, [Index], false, false, null));
        };
        PivotTable.prototype.getColumnRange = function () {
            _throwIfApiNotSupported("PivotTable.getColumnRange", "Pivot", "1.1", _hostName);
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetColumnRange", 1, [], false, true, null));
        };
        PivotTable.prototype.getDataField = function (Index) {
            _throwIfApiNotSupported("PivotTable.getDataField", "Pivot", "1.1", _hostName);
            return new Excel.PivotField(this.context, _createMethodObjectPath(this.context, this, "GetDataField", 1, [Index], false, false, null));
        };
        PivotTable.prototype.getDataPivotField = function () {
            _throwIfApiNotSupported("PivotTable.getDataPivotField", "Pivot", "1.1", _hostName);
            return new Excel.PivotField(this.context, _createMethodObjectPath(this.context, this, "GetDataPivotField", 1, [], false, false, null));
        };
        PivotTable.prototype.getEntireRange = function () {
            _throwIfApiNotSupported("PivotTable.getEntireRange", "Pivot", "1.1", _hostName);
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetEntireRange", 1, [], false, true, null));
        };
        PivotTable.prototype.getHiddenField = function (Index) {
            _throwIfApiNotSupported("PivotTable.getHiddenField", "Pivot", "1.1", _hostName);
            return new Excel.PivotField(this.context, _createMethodObjectPath(this.context, this, "GetHiddenField", 1, [Index], false, false, null));
        };
        PivotTable.prototype.getPageRange = function () {
            _throwIfApiNotSupported("PivotTable.getPageRange", "Pivot", "1.1", _hostName);
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetPageRange", 1, [], false, true, null));
        };
        PivotTable.prototype.getRowField = function (Index) {
            _throwIfApiNotSupported("PivotTable.getRowField", "Pivot", "1.1", _hostName);
            return new Excel.PivotField(this.context, _createMethodObjectPath(this.context, this, "GetRowField", 1, [Index], false, false, null));
        };
        PivotTable.prototype.getRowRange = function () {
            _throwIfApiNotSupported("PivotTable.getRowRange", "Pivot", "1.1", _hostName);
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetRowRange", 1, [], false, true, null));
        };
        PivotTable.prototype.getVisibleFields = function () {
            _throwIfApiNotSupported("PivotTable.getVisibleFields", "Pivot", "1.1", _hostName);
            return new Excel.PivotFieldCollection(this.context, _createMethodObjectPath(this.context, this, "GetVisibleFields", 1, [], true, false, null));
        };
        PivotTable.prototype.listFormulas = function () {
            _throwIfApiNotSupported("PivotTable.listFormulas", "Pivot", "1.1", _hostName);
            _createMethodAction(this.context, this, "ListFormulas", 0, []);
        };
        PivotTable.prototype.pivotSelect = function (Name, Mode, UseStandardName) {
            _throwIfApiNotSupported("PivotTable.pivotSelect", "Pivot", "1.1", _hostName);
            _createMethodAction(this.context, this, "PivotSelect", 0, [Name, Mode, UseStandardName]);
        };
        PivotTable.prototype.refreshTable = function () {
            _throwIfApiNotSupported("PivotTable.refreshTable", "Pivot", "1.1", _hostName);
            var action = _createMethodAction(this.context, this, "RefreshTable", 0, []);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        PivotTable.prototype.repeatAllLabels = function (Repeat) {
            _throwIfApiNotSupported("PivotTable.repeatAllLabels", "Pivot", "1.1", _hostName);
            _createMethodAction(this.context, this, "RepeatAllLabels", 0, [Repeat]);
        };
        PivotTable.prototype.update = function () {
            _throwIfApiNotSupported("PivotTable.update", "Pivot", "1.1", _hostName);
            _createMethodAction(this.context, this, "Update", 0, []);
        };
        PivotTable.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["ColumnGrandTotals"])) {
                this.m_columnGrandTotals = obj["ColumnGrandTotals"];
            }
            if (!_isUndefined(obj["Id"])) {
                this.m_id = obj["Id"];
            }
            if (!_isUndefined(obj["Name"])) {
                this.m_name = obj["Name"];
            }
            if (!_isUndefined(obj["RowGrandTotals"])) {
                this.m_rowGrandTotals = obj["RowGrandTotals"];
            }
            if (!_isUndefined(obj["AllowMultipleFilters"])) {
                this.m_allowMultipleFilters = obj["AllowMultipleFilters"];
            }
            if (!_isUndefined(obj["AlternativeText"])) {
                this.m_alternativeText = obj["AlternativeText"];
            }
            if (!_isUndefined(obj["CompactLayoutColumnHeader"])) {
                this.m_compactLayoutColumnHeader = obj["CompactLayoutColumnHeader"];
            }
            if (!_isUndefined(obj["CompactLayoutRowHeader"])) {
                this.m_compactLayoutRowHeader = obj["CompactLayoutRowHeader"];
            }
            if (!_isUndefined(obj["CompactRowIndent"])) {
                this.m_compactRowIndent = obj["CompactRowIndent"];
            }
            if (!_isUndefined(obj["DisplayContextTooltips"])) {
                this.m_displayContextTooltips = obj["DisplayContextTooltips"];
            }
            if (!_isUndefined(obj["DisplayEmptyColumn"])) {
                this.m_displayEmptyColumn = obj["DisplayEmptyColumn"];
            }
            if (!_isUndefined(obj["DisplayEmptyRow"])) {
                this.m_displayEmptyRow = obj["DisplayEmptyRow"];
            }
            if (!_isUndefined(obj["DisplayErrorString"])) {
                this.m_displayErrorString = obj["DisplayErrorString"];
            }
            if (!_isUndefined(obj["DisplayFieldCaptions"])) {
                this.m_displayFieldCaptions = obj["DisplayFieldCaptions"];
            }
            if (!_isUndefined(obj["DisplayNullString"])) {
                this.m_displayNullString = obj["DisplayNullString"];
            }
            if (!_isUndefined(obj["EnableDataValueEditing"])) {
                this.m_enableDataValueEditing = obj["EnableDataValueEditing"];
            }
            if (!_isUndefined(obj["EnableDrilldown"])) {
                this.m_enableDrilldown = obj["EnableDrilldown"];
            }
            if (!_isUndefined(obj["EnableFieldDialog"])) {
                this.m_enableFieldDialog = obj["EnableFieldDialog"];
            }
            if (!_isUndefined(obj["EnableFieldList"])) {
                this.m_enableFieldList = obj["EnableFieldList"];
            }
            if (!_isUndefined(obj["EnableWizard"])) {
                this.m_enableWizard = obj["EnableWizard"];
            }
            if (!_isUndefined(obj["ErrorString"])) {
                this.m_errorString = obj["ErrorString"];
            }
            if (!_isUndefined(obj["FieldListSortAscending"])) {
                this.m_fieldListSortAscending = obj["FieldListSortAscending"];
            }
            if (!_isUndefined(obj["GrandTotalName"])) {
                this.m_grandTotalName = obj["GrandTotalName"];
            }
            if (!_isUndefined(obj["HasAutoFormat"])) {
                this.m_hasAutoFormat = obj["HasAutoFormat"];
            }
            if (!_isUndefined(obj["Hidden"])) {
                this.m_hidden = obj["Hidden"];
            }
            if (!_isUndefined(obj["InGridDropZones"])) {
                this.m_inGridDropZones = obj["InGridDropZones"];
            }
            if (!_isUndefined(obj["InnerDetail"])) {
                this.m_innerDetail = obj["InnerDetail"];
            }
            if (!_isUndefined(obj["ManualUpdate"])) {
                this.m_manualUpdate = obj["ManualUpdate"];
            }
            if (!_isUndefined(obj["NullString"])) {
                this.m_nullString = obj["NullString"];
            }
            if (!_isUndefined(obj["PreserveFormatting"])) {
                this.m_preserveFormatting = obj["PreserveFormatting"];
            }
            if (!_isUndefined(obj["PrintDrillIndicators"])) {
                this.m_printDrillIndicators = obj["PrintDrillIndicators"];
            }
            if (!_isUndefined(obj["PrintTitles"])) {
                this.m_printTitles = obj["PrintTitles"];
            }
            if (!_isUndefined(obj["RefreshDate"])) {
                this.m_refreshDate = _adjustToDateTime(obj["RefreshDate"]);
            }
            if (!_isUndefined(obj["RefreshName"])) {
                this.m_refreshName = obj["RefreshName"];
            }
            if (!_isUndefined(obj["RepeatItemsOnEachPrintedPage"])) {
                this.m_repeatItemsOnEachPrintedPage = obj["RepeatItemsOnEachPrintedPage"];
            }
            if (!_isUndefined(obj["SaveData"])) {
                this.m_saveData = obj["SaveData"];
            }
            if (!_isUndefined(obj["ShowDrillIndicators"])) {
                this.m_showDrillIndicators = obj["ShowDrillIndicators"];
            }
            if (!_isUndefined(obj["ShowPageMultipleItemLabel"])) {
                this.m_showPageMultipleItemLabel = obj["ShowPageMultipleItemLabel"];
            }
            if (!_isUndefined(obj["ShowTableStyleColumnHeaders"])) {
                this.m_showTableStyleColumnHeaders = obj["ShowTableStyleColumnHeaders"];
            }
            if (!_isUndefined(obj["ShowTableStyleColumnStripes"])) {
                this.m_showTableStyleColumnStripes = obj["ShowTableStyleColumnStripes"];
            }
            if (!_isUndefined(obj["ShowTableStyleLastColumn"])) {
                this.m_showTableStyleLastColumn = obj["ShowTableStyleLastColumn"];
            }
            if (!_isUndefined(obj["ShowTableStyleRowHeaders"])) {
                this.m_showTableStyleRowHeaders = obj["ShowTableStyleRowHeaders"];
            }
            if (!_isUndefined(obj["ShowTableStyleRowStripes"])) {
                this.m_showTableStyleRowStripes = obj["ShowTableStyleRowStripes"];
            }
            if (!_isUndefined(obj["ShowValuesRow"])) {
                this.m_showValuesRow = obj["ShowValuesRow"];
            }
            if (!_isUndefined(obj["SmallGrid"])) {
                this.m_smallGrid = obj["SmallGrid"];
            }
            if (!_isUndefined(obj["SortUsingCustomLists"])) {
                this.m_sortUsingCustomLists = obj["SortUsingCustomLists"];
            }
            if (!_isUndefined(obj["SubtotalHiddenPageItems"])) {
                this.m_subtotalHiddenPageItems = obj["SubtotalHiddenPageItems"];
            }
            if (!_isUndefined(obj["Summary"])) {
                this.m_summary = obj["Summary"];
            }
            if (!_isUndefined(obj["Tag"])) {
                this.m_tag = obj["Tag"];
            }
            if (!_isUndefined(obj["TotalsAnnotation"])) {
                this.m_totalsAnnotation = obj["TotalsAnnotation"];
            }
            if (!_isUndefined(obj["VacatedStyle"])) {
                this.m_vacatedStyle = obj["VacatedStyle"];
            }
            if (!_isUndefined(obj["Value"])) {
                this.m_value = obj["Value"];
            }
            if (!_isUndefined(obj["Version"])) {
                this.m_version = obj["Version"];
            }
            _handleNavigationPropertyResults(this, obj, ["worksheet", "Worksheet", "calculatedFields", "CalculatedFields", "dataBodyRange", "DataBodyRange", "dataLabelRange", "DataLabelRange", "pivotFields", "PivotFields"]);
        };
        PivotTable.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        PivotTable.prototype._handleIdResult = function (value) {
            _super.prototype._handleIdResult.call(this, value);
            if (_isNullOrUndefined(value)) {
                return;
            }
            if (!_isUndefined(value["Id"])) {
                this.m_id = value["Id"];
            }
        };
        PivotTable.prototype.toJSON = function () {
            return {
                "allowMultipleFilters": this.m_allowMultipleFilters,
                "alternativeText": this.m_alternativeText,
                "columnGrandTotals": this.m_columnGrandTotals,
                "compactLayoutColumnHeader": this.m_compactLayoutColumnHeader,
                "compactLayoutRowHeader": this.m_compactLayoutRowHeader,
                "compactRowIndent": this.m_compactRowIndent,
                "displayContextTooltips": this.m_displayContextTooltips,
                "displayEmptyColumn": this.m_displayEmptyColumn,
                "displayEmptyRow": this.m_displayEmptyRow,
                "displayErrorString": this.m_displayErrorString,
                "displayFieldCaptions": this.m_displayFieldCaptions,
                "displayNullString": this.m_displayNullString,
                "enableDataValueEditing": this.m_enableDataValueEditing,
                "enableDrilldown": this.m_enableDrilldown,
                "enableFieldDialog": this.m_enableFieldDialog,
                "enableFieldList": this.m_enableFieldList,
                "enableWizard": this.m_enableWizard,
                "errorString": this.m_errorString,
                "fieldListSortAscending": this.m_fieldListSortAscending,
                "grandTotalName": this.m_grandTotalName,
                "hasAutoFormat": this.m_hasAutoFormat,
                "hidden": this.m_hidden,
                "id": this.m_id,
                "inGridDropZones": this.m_inGridDropZones,
                "innerDetail": this.m_innerDetail,
                "manualUpdate": this.m_manualUpdate,
                "name": this.m_name,
                "nullString": this.m_nullString,
                "preserveFormatting": this.m_preserveFormatting,
                "printDrillIndicators": this.m_printDrillIndicators,
                "printTitles": this.m_printTitles,
                "refreshDate": this.m_refreshDate,
                "refreshName": this.m_refreshName,
                "repeatItemsOnEachPrintedPage": this.m_repeatItemsOnEachPrintedPage,
                "rowGrandTotals": this.m_rowGrandTotals,
                "saveData": this.m_saveData,
                "showDrillIndicators": this.m_showDrillIndicators,
                "showPageMultipleItemLabel": this.m_showPageMultipleItemLabel,
                "showTableStyleColumnHeaders": this.m_showTableStyleColumnHeaders,
                "showTableStyleColumnStripes": this.m_showTableStyleColumnStripes,
                "showTableStyleLastColumn": this.m_showTableStyleLastColumn,
                "showTableStyleRowHeaders": this.m_showTableStyleRowHeaders,
                "showTableStyleRowStripes": this.m_showTableStyleRowStripes,
                "showValuesRow": this.m_showValuesRow,
                "smallGrid": this.m_smallGrid,
                "sortUsingCustomLists": this.m_sortUsingCustomLists,
                "subtotalHiddenPageItems": this.m_subtotalHiddenPageItems,
                "summary": this.m_summary,
                "tag": this.m_tag,
                "totalsAnnotation": this.m_totalsAnnotation,
                "vacatedStyle": this.m_vacatedStyle,
                "value": this.m_value,
                "version": this.m_version
            };
        };
        return PivotTable;
    }(OfficeExtension.ClientObject));
    Excel.PivotTable = PivotTable;
    var DocumentProperties = (function (_super) {
        __extends(DocumentProperties, _super);
        function DocumentProperties() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(DocumentProperties.prototype, "_className", {
            get: function () {
                return "DocumentProperties";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DocumentProperties.prototype, "author", {
            get: function () {
                _throwIfNotLoaded("author", this.m_author, "DocumentProperties", this._isNull);
                return this.m_author;
            },
            set: function (value) {
                this.m_author = value;
                _createSetPropertyAction(this.context, this, "Author", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DocumentProperties.prototype, "category", {
            get: function () {
                _throwIfNotLoaded("category", this.m_category, "DocumentProperties", this._isNull);
                return this.m_category;
            },
            set: function (value) {
                this.m_category = value;
                _createSetPropertyAction(this.context, this, "Category", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DocumentProperties.prototype, "comments", {
            get: function () {
                _throwIfNotLoaded("comments", this.m_comments, "DocumentProperties", this._isNull);
                return this.m_comments;
            },
            set: function (value) {
                this.m_comments = value;
                _createSetPropertyAction(this.context, this, "Comments", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DocumentProperties.prototype, "company", {
            get: function () {
                _throwIfNotLoaded("company", this.m_company, "DocumentProperties", this._isNull);
                return this.m_company;
            },
            set: function (value) {
                this.m_company = value;
                _createSetPropertyAction(this.context, this, "Company", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DocumentProperties.prototype, "creationDate", {
            get: function () {
                _throwIfNotLoaded("creationDate", this.m_creationDate, "DocumentProperties", this._isNull);
                return this.m_creationDate;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DocumentProperties.prototype, "keywords", {
            get: function () {
                _throwIfNotLoaded("keywords", this.m_keywords, "DocumentProperties", this._isNull);
                return this.m_keywords;
            },
            set: function (value) {
                this.m_keywords = value;
                _createSetPropertyAction(this.context, this, "Keywords", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DocumentProperties.prototype, "lastAuthor", {
            get: function () {
                _throwIfNotLoaded("lastAuthor", this.m_lastAuthor, "DocumentProperties", this._isNull);
                return this.m_lastAuthor;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DocumentProperties.prototype, "manager", {
            get: function () {
                _throwIfNotLoaded("manager", this.m_manager, "DocumentProperties", this._isNull);
                return this.m_manager;
            },
            set: function (value) {
                this.m_manager = value;
                _createSetPropertyAction(this.context, this, "Manager", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DocumentProperties.prototype, "revisionNumber", {
            get: function () {
                _throwIfNotLoaded("revisionNumber", this.m_revisionNumber, "DocumentProperties", this._isNull);
                return this.m_revisionNumber;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DocumentProperties.prototype, "subject", {
            get: function () {
                _throwIfNotLoaded("subject", this.m_subject, "DocumentProperties", this._isNull);
                return this.m_subject;
            },
            set: function (value) {
                this.m_subject = value;
                _createSetPropertyAction(this.context, this, "Subject", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DocumentProperties.prototype, "title", {
            get: function () {
                _throwIfNotLoaded("title", this.m_title, "DocumentProperties", this._isNull);
                return this.m_title;
            },
            set: function (value) {
                this.m_title = value;
                _createSetPropertyAction(this.context, this, "Title", value);
            },
            enumerable: true,
            configurable: true
        });
        DocumentProperties.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["title", "subject", "author", "keywords", "comments", "category", "manager", "company"], [], []);
        };
        DocumentProperties.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Author"])) {
                this.m_author = obj["Author"];
            }
            if (!_isUndefined(obj["Category"])) {
                this.m_category = obj["Category"];
            }
            if (!_isUndefined(obj["Comments"])) {
                this.m_comments = obj["Comments"];
            }
            if (!_isUndefined(obj["Company"])) {
                this.m_company = obj["Company"];
            }
            if (!_isUndefined(obj["CreationDate"])) {
                this.m_creationDate = _adjustToDateTime(obj["CreationDate"]);
            }
            if (!_isUndefined(obj["Keywords"])) {
                this.m_keywords = obj["Keywords"];
            }
            if (!_isUndefined(obj["LastAuthor"])) {
                this.m_lastAuthor = obj["LastAuthor"];
            }
            if (!_isUndefined(obj["Manager"])) {
                this.m_manager = obj["Manager"];
            }
            if (!_isUndefined(obj["RevisionNumber"])) {
                this.m_revisionNumber = obj["RevisionNumber"];
            }
            if (!_isUndefined(obj["Subject"])) {
                this.m_subject = obj["Subject"];
            }
            if (!_isUndefined(obj["Title"])) {
                this.m_title = obj["Title"];
            }
        };
        DocumentProperties.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        DocumentProperties.prototype.toJSON = function () {
            return {
                "author": this.m_author,
                "category": this.m_category,
                "comments": this.m_comments,
                "company": this.m_company,
                "creationDate": this.m_creationDate,
                "keywords": this.m_keywords,
                "lastAuthor": this.m_lastAuthor,
                "manager": this.m_manager,
                "revisionNumber": this.m_revisionNumber,
                "subject": this.m_subject,
                "title": this.m_title
            };
        };
        return DocumentProperties;
    }(OfficeExtension.ClientObject));
    Excel.DocumentProperties = DocumentProperties;
    var ConditionalFormatCollection = (function (_super) {
        __extends(ConditionalFormatCollection, _super);
        function ConditionalFormatCollection() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ConditionalFormatCollection.prototype, "_className", {
            get: function () {
                return "ConditionalFormatCollection";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormatCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, "ConditionalFormatCollection", this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        ConditionalFormatCollection.prototype.add = function (type) {
            return new Excel.ConditionalFormat(this.context, _createMethodObjectPath(this.context, this, "Add", 0, [type], false, true, null));
        };
        ConditionalFormatCollection.prototype.clearAll = function () {
            _createMethodAction(this.context, this, "ClearAll", 0, []);
        };
        ConditionalFormatCollection.prototype.getCount = function () {
            var action = _createMethodAction(this.context, this, "GetCount", 1, []);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        ConditionalFormatCollection.prototype.getItem = function (id) {
            return new Excel.ConditionalFormat(this.context, _createIndexerObjectPath(this.context, this, [id]));
        };
        ConditionalFormatCollection.prototype.getItemAt = function (index) {
            return new Excel.ConditionalFormat(this.context, _createMethodObjectPath(this.context, this, "GetItemAt", 1, [index], false, false, null));
        };
        ConditionalFormatCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.ConditionalFormat(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        ConditionalFormatCollection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ConditionalFormatCollection.prototype.toJSON = function () {
            return {};
        };
        return ConditionalFormatCollection;
    }(OfficeExtension.ClientObject));
    Excel.ConditionalFormatCollection = ConditionalFormatCollection;
    var ConditionalFormat = (function (_super) {
        __extends(ConditionalFormat, _super);
        function ConditionalFormat() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ConditionalFormat.prototype, "_className", {
            get: function () {
                return "ConditionalFormat";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormat.prototype, "cellValue", {
            get: function () {
                if (!this.m_cellValue) {
                    this.m_cellValue = new Excel.CellValueConditionalFormat(this.context, _createPropertyObjectPath(this.context, this, "CellValue", false, false));
                }
                return this.m_cellValue;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormat.prototype, "cellValueOrNullObject", {
            get: function () {
                if (!this.m_cellValueOrNullObject) {
                    this.m_cellValueOrNullObject = new Excel.CellValueConditionalFormat(this.context, _createPropertyObjectPath(this.context, this, "CellValueOrNullObject", false, false));
                }
                return this.m_cellValueOrNullObject;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormat.prototype, "colorScale", {
            get: function () {
                if (!this.m_colorScale) {
                    this.m_colorScale = new Excel.ColorScaleConditionalFormat(this.context, _createPropertyObjectPath(this.context, this, "ColorScale", false, false));
                }
                return this.m_colorScale;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormat.prototype, "colorScaleOrNullObject", {
            get: function () {
                if (!this.m_colorScaleOrNullObject) {
                    this.m_colorScaleOrNullObject = new Excel.ColorScaleConditionalFormat(this.context, _createPropertyObjectPath(this.context, this, "ColorScaleOrNullObject", false, false));
                }
                return this.m_colorScaleOrNullObject;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormat.prototype, "custom", {
            get: function () {
                if (!this.m_custom) {
                    this.m_custom = new Excel.CustomConditionalFormat(this.context, _createPropertyObjectPath(this.context, this, "Custom", false, false));
                }
                return this.m_custom;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormat.prototype, "customOrNullObject", {
            get: function () {
                if (!this.m_customOrNullObject) {
                    this.m_customOrNullObject = new Excel.CustomConditionalFormat(this.context, _createPropertyObjectPath(this.context, this, "CustomOrNullObject", false, false));
                }
                return this.m_customOrNullObject;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormat.prototype, "dataBar", {
            get: function () {
                if (!this.m_dataBar) {
                    this.m_dataBar = new Excel.DataBarConditionalFormat(this.context, _createPropertyObjectPath(this.context, this, "DataBar", false, false));
                }
                return this.m_dataBar;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormat.prototype, "dataBarOrNullObject", {
            get: function () {
                if (!this.m_dataBarOrNullObject) {
                    this.m_dataBarOrNullObject = new Excel.DataBarConditionalFormat(this.context, _createPropertyObjectPath(this.context, this, "DataBarOrNullObject", false, false));
                }
                return this.m_dataBarOrNullObject;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormat.prototype, "iconSet", {
            get: function () {
                if (!this.m_iconSet) {
                    this.m_iconSet = new Excel.IconSetConditionalFormat(this.context, _createPropertyObjectPath(this.context, this, "IconSet", false, false));
                }
                return this.m_iconSet;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormat.prototype, "iconSetOrNullObject", {
            get: function () {
                if (!this.m_iconSetOrNullObject) {
                    this.m_iconSetOrNullObject = new Excel.IconSetConditionalFormat(this.context, _createPropertyObjectPath(this.context, this, "IconSetOrNullObject", false, false));
                }
                return this.m_iconSetOrNullObject;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormat.prototype, "preset", {
            get: function () {
                if (!this.m_preset) {
                    this.m_preset = new Excel.PresetCriteriaConditionalFormat(this.context, _createPropertyObjectPath(this.context, this, "Preset", false, false));
                }
                return this.m_preset;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormat.prototype, "presetOrNullObject", {
            get: function () {
                if (!this.m_presetOrNullObject) {
                    this.m_presetOrNullObject = new Excel.PresetCriteriaConditionalFormat(this.context, _createPropertyObjectPath(this.context, this, "PresetOrNullObject", false, false));
                }
                return this.m_presetOrNullObject;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormat.prototype, "textComparison", {
            get: function () {
                if (!this.m_textComparison) {
                    this.m_textComparison = new Excel.TextConditionalFormat(this.context, _createPropertyObjectPath(this.context, this, "TextComparison", false, false));
                }
                return this.m_textComparison;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormat.prototype, "textComparisonOrNullObject", {
            get: function () {
                if (!this.m_textComparisonOrNullObject) {
                    this.m_textComparisonOrNullObject = new Excel.TextConditionalFormat(this.context, _createPropertyObjectPath(this.context, this, "TextComparisonOrNullObject", false, false));
                }
                return this.m_textComparisonOrNullObject;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormat.prototype, "topBottom", {
            get: function () {
                if (!this.m_topBottom) {
                    this.m_topBottom = new Excel.TopBottomConditionalFormat(this.context, _createPropertyObjectPath(this.context, this, "TopBottom", false, false));
                }
                return this.m_topBottom;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormat.prototype, "topBottomOrNullObject", {
            get: function () {
                if (!this.m_topBottomOrNullObject) {
                    this.m_topBottomOrNullObject = new Excel.TopBottomConditionalFormat(this.context, _createPropertyObjectPath(this.context, this, "TopBottomOrNullObject", false, false));
                }
                return this.m_topBottomOrNullObject;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormat.prototype, "id", {
            get: function () {
                _throwIfNotLoaded("id", this.m_id, "ConditionalFormat", this._isNull);
                return this.m_id;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormat.prototype, "priority", {
            get: function () {
                _throwIfNotLoaded("priority", this.m_priority, "ConditionalFormat", this._isNull);
                return this.m_priority;
            },
            set: function (value) {
                this.m_priority = value;
                _createSetPropertyAction(this.context, this, "Priority", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormat.prototype, "stopIfTrue", {
            get: function () {
                _throwIfNotLoaded("stopIfTrue", this.m_stopIfTrue, "ConditionalFormat", this._isNull);
                return this.m_stopIfTrue;
            },
            set: function (value) {
                this.m_stopIfTrue = value;
                _createSetPropertyAction(this.context, this, "StopIfTrue", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormat.prototype, "type", {
            get: function () {
                _throwIfNotLoaded("type", this.m_type, "ConditionalFormat", this._isNull);
                return this.m_type;
            },
            enumerable: true,
            configurable: true
        });
        ConditionalFormat.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["stopIfTrue", "priority"], ["dataBarOrNullObject", "dataBar", "customOrNullObject", "custom", "iconSet", "iconSetOrNullObject", "colorScale", "colorScaleOrNullObject", "topBottom", "topBottomOrNullObject", "preset", "presetOrNullObject", "textComparison", "textComparisonOrNullObject", "cellValue", "cellValueOrNullObject"], []);
        };
        ConditionalFormat.prototype.delete = function () {
            _createMethodAction(this.context, this, "Delete", 0, []);
        };
        ConditionalFormat.prototype.getRange = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetRange", 1, [], false, true, null));
        };
        ConditionalFormat.prototype.getRangeOrNullObject = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetRangeOrNullObject", 1, [], false, true, null));
        };
        ConditionalFormat.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Id"])) {
                this.m_id = obj["Id"];
            }
            if (!_isUndefined(obj["Priority"])) {
                this.m_priority = obj["Priority"];
            }
            if (!_isUndefined(obj["StopIfTrue"])) {
                this.m_stopIfTrue = obj["StopIfTrue"];
            }
            if (!_isUndefined(obj["Type"])) {
                this.m_type = obj["Type"];
            }
            _handleNavigationPropertyResults(this, obj, ["cellValue", "CellValue", "cellValueOrNullObject", "CellValueOrNullObject", "colorScale", "ColorScale", "colorScaleOrNullObject", "ColorScaleOrNullObject", "custom", "Custom", "customOrNullObject", "CustomOrNullObject", "dataBar", "DataBar", "dataBarOrNullObject", "DataBarOrNullObject", "iconSet", "IconSet", "iconSetOrNullObject", "IconSetOrNullObject", "preset", "Preset", "presetOrNullObject", "PresetOrNullObject", "textComparison", "TextComparison", "textComparisonOrNullObject", "TextComparisonOrNullObject", "topBottom", "TopBottom", "topBottomOrNullObject", "TopBottomOrNullObject"]);
        };
        ConditionalFormat.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ConditionalFormat.prototype._handleIdResult = function (value) {
            _super.prototype._handleIdResult.call(this, value);
            if (_isNullOrUndefined(value)) {
                return;
            }
            if (!_isUndefined(value["Id"])) {
                this.m_id = value["Id"];
            }
        };
        ConditionalFormat.prototype.toJSON = function () {
            return {
                "cellValue": this.m_cellValue,
                "cellValueOrNullObject": this.m_cellValueOrNullObject,
                "colorScale": this.m_colorScale,
                "colorScaleOrNullObject": this.m_colorScaleOrNullObject,
                "custom": this.m_custom,
                "customOrNullObject": this.m_customOrNullObject,
                "dataBar": this.m_dataBar,
                "dataBarOrNullObject": this.m_dataBarOrNullObject,
                "iconSet": this.m_iconSet,
                "iconSetOrNullObject": this.m_iconSetOrNullObject,
                "id": this.m_id,
                "preset": this.m_preset,
                "presetOrNullObject": this.m_presetOrNullObject,
                "priority": this.m_priority,
                "stopIfTrue": this.m_stopIfTrue,
                "textComparison": this.m_textComparison,
                "textComparisonOrNullObject": this.m_textComparisonOrNullObject,
                "topBottom": this.m_topBottom,
                "topBottomOrNullObject": this.m_topBottomOrNullObject,
                "type": this.m_type
            };
        };
        return ConditionalFormat;
    }(OfficeExtension.ClientObject));
    Excel.ConditionalFormat = ConditionalFormat;
    var DataBarConditionalFormat = (function (_super) {
        __extends(DataBarConditionalFormat, _super);
        function DataBarConditionalFormat() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(DataBarConditionalFormat.prototype, "_className", {
            get: function () {
                return "DataBarConditionalFormat";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DataBarConditionalFormat.prototype, "negativeFormat", {
            get: function () {
                if (!this.m_negativeFormat) {
                    this.m_negativeFormat = new Excel.ConditionalDataBarNegativeFormat(this.context, _createPropertyObjectPath(this.context, this, "NegativeFormat", false, false));
                }
                return this.m_negativeFormat;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DataBarConditionalFormat.prototype, "positiveFormat", {
            get: function () {
                if (!this.m_positiveFormat) {
                    this.m_positiveFormat = new Excel.ConditionalDataBarPositiveFormat(this.context, _createPropertyObjectPath(this.context, this, "PositiveFormat", false, false));
                }
                return this.m_positiveFormat;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DataBarConditionalFormat.prototype, "axisColor", {
            get: function () {
                _throwIfNotLoaded("axisColor", this.m_axisColor, "DataBarConditionalFormat", this._isNull);
                return this.m_axisColor;
            },
            set: function (value) {
                this.m_axisColor = value;
                _createSetPropertyAction(this.context, this, "AxisColor", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DataBarConditionalFormat.prototype, "axisFormat", {
            get: function () {
                _throwIfNotLoaded("axisFormat", this.m_axisFormat, "DataBarConditionalFormat", this._isNull);
                return this.m_axisFormat;
            },
            set: function (value) {
                this.m_axisFormat = value;
                _createSetPropertyAction(this.context, this, "AxisFormat", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DataBarConditionalFormat.prototype, "barDirection", {
            get: function () {
                _throwIfNotLoaded("barDirection", this.m_barDirection, "DataBarConditionalFormat", this._isNull);
                return this.m_barDirection;
            },
            set: function (value) {
                this.m_barDirection = value;
                _createSetPropertyAction(this.context, this, "BarDirection", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DataBarConditionalFormat.prototype, "lowerBoundRule", {
            get: function () {
                _throwIfNotLoaded("lowerBoundRule", this.m_lowerBoundRule, "DataBarConditionalFormat", this._isNull);
                return this.m_lowerBoundRule;
            },
            set: function (value) {
                this.m_lowerBoundRule = value;
                _createSetPropertyAction(this.context, this, "LowerBoundRule", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DataBarConditionalFormat.prototype, "showDataBarOnly", {
            get: function () {
                _throwIfNotLoaded("showDataBarOnly", this.m_showDataBarOnly, "DataBarConditionalFormat", this._isNull);
                return this.m_showDataBarOnly;
            },
            set: function (value) {
                this.m_showDataBarOnly = value;
                _createSetPropertyAction(this.context, this, "ShowDataBarOnly", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DataBarConditionalFormat.prototype, "upperBoundRule", {
            get: function () {
                _throwIfNotLoaded("upperBoundRule", this.m_upperBoundRule, "DataBarConditionalFormat", this._isNull);
                return this.m_upperBoundRule;
            },
            set: function (value) {
                this.m_upperBoundRule = value;
                _createSetPropertyAction(this.context, this, "UpperBoundRule", value);
            },
            enumerable: true,
            configurable: true
        });
        DataBarConditionalFormat.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["showDataBarOnly", "barDirection", "axisFormat", "axisColor", "lowerBoundRule", "upperBoundRule"], ["positiveFormat", "negativeFormat"], []);
        };
        DataBarConditionalFormat.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["AxisColor"])) {
                this.m_axisColor = obj["AxisColor"];
            }
            if (!_isUndefined(obj["AxisFormat"])) {
                this.m_axisFormat = obj["AxisFormat"];
            }
            if (!_isUndefined(obj["BarDirection"])) {
                this.m_barDirection = obj["BarDirection"];
            }
            if (!_isUndefined(obj["LowerBoundRule"])) {
                this.m_lowerBoundRule = obj["LowerBoundRule"];
            }
            if (!_isUndefined(obj["ShowDataBarOnly"])) {
                this.m_showDataBarOnly = obj["ShowDataBarOnly"];
            }
            if (!_isUndefined(obj["UpperBoundRule"])) {
                this.m_upperBoundRule = obj["UpperBoundRule"];
            }
            _handleNavigationPropertyResults(this, obj, ["negativeFormat", "NegativeFormat", "positiveFormat", "PositiveFormat"]);
        };
        DataBarConditionalFormat.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        DataBarConditionalFormat.prototype.toJSON = function () {
            return {
                "axisColor": this.m_axisColor,
                "axisFormat": this.m_axisFormat,
                "barDirection": this.m_barDirection,
                "lowerBoundRule": this.m_lowerBoundRule,
                "negativeFormat": this.m_negativeFormat,
                "positiveFormat": this.m_positiveFormat,
                "showDataBarOnly": this.m_showDataBarOnly,
                "upperBoundRule": this.m_upperBoundRule
            };
        };
        return DataBarConditionalFormat;
    }(OfficeExtension.ClientObject));
    Excel.DataBarConditionalFormat = DataBarConditionalFormat;
    var ConditionalDataBarPositiveFormat = (function (_super) {
        __extends(ConditionalDataBarPositiveFormat, _super);
        function ConditionalDataBarPositiveFormat() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ConditionalDataBarPositiveFormat.prototype, "_className", {
            get: function () {
                return "ConditionalDataBarPositiveFormat";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalDataBarPositiveFormat.prototype, "borderColor", {
            get: function () {
                _throwIfNotLoaded("borderColor", this.m_borderColor, "ConditionalDataBarPositiveFormat", this._isNull);
                return this.m_borderColor;
            },
            set: function (value) {
                this.m_borderColor = value;
                _createSetPropertyAction(this.context, this, "BorderColor", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalDataBarPositiveFormat.prototype, "fillColor", {
            get: function () {
                _throwIfNotLoaded("fillColor", this.m_fillColor, "ConditionalDataBarPositiveFormat", this._isNull);
                return this.m_fillColor;
            },
            set: function (value) {
                this.m_fillColor = value;
                _createSetPropertyAction(this.context, this, "FillColor", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalDataBarPositiveFormat.prototype, "gradientFill", {
            get: function () {
                _throwIfNotLoaded("gradientFill", this.m_gradientFill, "ConditionalDataBarPositiveFormat", this._isNull);
                return this.m_gradientFill;
            },
            set: function (value) {
                this.m_gradientFill = value;
                _createSetPropertyAction(this.context, this, "GradientFill", value);
            },
            enumerable: true,
            configurable: true
        });
        ConditionalDataBarPositiveFormat.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["fillColor", "gradientFill", "borderColor"], [], []);
        };
        ConditionalDataBarPositiveFormat.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["BorderColor"])) {
                this.m_borderColor = obj["BorderColor"];
            }
            if (!_isUndefined(obj["FillColor"])) {
                this.m_fillColor = obj["FillColor"];
            }
            if (!_isUndefined(obj["GradientFill"])) {
                this.m_gradientFill = obj["GradientFill"];
            }
        };
        ConditionalDataBarPositiveFormat.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ConditionalDataBarPositiveFormat.prototype.toJSON = function () {
            return {
                "borderColor": this.m_borderColor,
                "fillColor": this.m_fillColor,
                "gradientFill": this.m_gradientFill
            };
        };
        return ConditionalDataBarPositiveFormat;
    }(OfficeExtension.ClientObject));
    Excel.ConditionalDataBarPositiveFormat = ConditionalDataBarPositiveFormat;
    var ConditionalDataBarNegativeFormat = (function (_super) {
        __extends(ConditionalDataBarNegativeFormat, _super);
        function ConditionalDataBarNegativeFormat() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ConditionalDataBarNegativeFormat.prototype, "_className", {
            get: function () {
                return "ConditionalDataBarNegativeFormat";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalDataBarNegativeFormat.prototype, "borderColor", {
            get: function () {
                _throwIfNotLoaded("borderColor", this.m_borderColor, "ConditionalDataBarNegativeFormat", this._isNull);
                return this.m_borderColor;
            },
            set: function (value) {
                this.m_borderColor = value;
                _createSetPropertyAction(this.context, this, "BorderColor", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalDataBarNegativeFormat.prototype, "fillColor", {
            get: function () {
                _throwIfNotLoaded("fillColor", this.m_fillColor, "ConditionalDataBarNegativeFormat", this._isNull);
                return this.m_fillColor;
            },
            set: function (value) {
                this.m_fillColor = value;
                _createSetPropertyAction(this.context, this, "FillColor", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalDataBarNegativeFormat.prototype, "matchPositiveBorderColor", {
            get: function () {
                _throwIfNotLoaded("matchPositiveBorderColor", this.m_matchPositiveBorderColor, "ConditionalDataBarNegativeFormat", this._isNull);
                return this.m_matchPositiveBorderColor;
            },
            set: function (value) {
                this.m_matchPositiveBorderColor = value;
                _createSetPropertyAction(this.context, this, "MatchPositiveBorderColor", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalDataBarNegativeFormat.prototype, "matchPositiveFillColor", {
            get: function () {
                _throwIfNotLoaded("matchPositiveFillColor", this.m_matchPositiveFillColor, "ConditionalDataBarNegativeFormat", this._isNull);
                return this.m_matchPositiveFillColor;
            },
            set: function (value) {
                this.m_matchPositiveFillColor = value;
                _createSetPropertyAction(this.context, this, "MatchPositiveFillColor", value);
            },
            enumerable: true,
            configurable: true
        });
        ConditionalDataBarNegativeFormat.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["fillColor", "matchPositiveFillColor", "borderColor", "matchPositiveBorderColor"], [], []);
        };
        ConditionalDataBarNegativeFormat.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["BorderColor"])) {
                this.m_borderColor = obj["BorderColor"];
            }
            if (!_isUndefined(obj["FillColor"])) {
                this.m_fillColor = obj["FillColor"];
            }
            if (!_isUndefined(obj["MatchPositiveBorderColor"])) {
                this.m_matchPositiveBorderColor = obj["MatchPositiveBorderColor"];
            }
            if (!_isUndefined(obj["MatchPositiveFillColor"])) {
                this.m_matchPositiveFillColor = obj["MatchPositiveFillColor"];
            }
        };
        ConditionalDataBarNegativeFormat.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ConditionalDataBarNegativeFormat.prototype.toJSON = function () {
            return {
                "borderColor": this.m_borderColor,
                "fillColor": this.m_fillColor,
                "matchPositiveBorderColor": this.m_matchPositiveBorderColor,
                "matchPositiveFillColor": this.m_matchPositiveFillColor
            };
        };
        return ConditionalDataBarNegativeFormat;
    }(OfficeExtension.ClientObject));
    Excel.ConditionalDataBarNegativeFormat = ConditionalDataBarNegativeFormat;
    var CustomConditionalFormat = (function (_super) {
        __extends(CustomConditionalFormat, _super);
        function CustomConditionalFormat() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(CustomConditionalFormat.prototype, "_className", {
            get: function () {
                return "CustomConditionalFormat";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(CustomConditionalFormat.prototype, "format", {
            get: function () {
                if (!this.m_format) {
                    this.m_format = new Excel.ConditionalRangeFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
                }
                return this.m_format;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(CustomConditionalFormat.prototype, "rule", {
            get: function () {
                if (!this.m_rule) {
                    this.m_rule = new Excel.ConditionalFormatRule(this.context, _createPropertyObjectPath(this.context, this, "Rule", false, false));
                }
                return this.m_rule;
            },
            enumerable: true,
            configurable: true
        });
        CustomConditionalFormat.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, [], ["rule", "format"], []);
        };
        CustomConditionalFormat.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            _handleNavigationPropertyResults(this, obj, ["format", "Format", "rule", "Rule"]);
        };
        CustomConditionalFormat.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        CustomConditionalFormat.prototype.toJSON = function () {
            return {
                "format": this.m_format,
                "rule": this.m_rule
            };
        };
        return CustomConditionalFormat;
    }(OfficeExtension.ClientObject));
    Excel.CustomConditionalFormat = CustomConditionalFormat;
    var ConditionalFormatRule = (function (_super) {
        __extends(ConditionalFormatRule, _super);
        function ConditionalFormatRule() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ConditionalFormatRule.prototype, "_className", {
            get: function () {
                return "ConditionalFormatRule";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormatRule.prototype, "formula", {
            get: function () {
                _throwIfNotLoaded("formula", this.m_formula, "ConditionalFormatRule", this._isNull);
                return this.m_formula;
            },
            set: function (value) {
                this.m_formula = value;
                _createSetPropertyAction(this.context, this, "Formula", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormatRule.prototype, "formulaLocal", {
            get: function () {
                _throwIfNotLoaded("formulaLocal", this.m_formulaLocal, "ConditionalFormatRule", this._isNull);
                return this.m_formulaLocal;
            },
            set: function (value) {
                this.m_formulaLocal = value;
                _createSetPropertyAction(this.context, this, "FormulaLocal", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormatRule.prototype, "formulaR1C1", {
            get: function () {
                _throwIfNotLoaded("formulaR1C1", this.m_formulaR1C1, "ConditionalFormatRule", this._isNull);
                return this.m_formulaR1C1;
            },
            set: function (value) {
                this.m_formulaR1C1 = value;
                _createSetPropertyAction(this.context, this, "FormulaR1C1", value);
            },
            enumerable: true,
            configurable: true
        });
        ConditionalFormatRule.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["formula", "formulaLocal", "formulaR1C1"], [], []);
        };
        ConditionalFormatRule.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Formula"])) {
                this.m_formula = obj["Formula"];
            }
            if (!_isUndefined(obj["FormulaLocal"])) {
                this.m_formulaLocal = obj["FormulaLocal"];
            }
            if (!_isUndefined(obj["FormulaR1C1"])) {
                this.m_formulaR1C1 = obj["FormulaR1C1"];
            }
        };
        ConditionalFormatRule.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ConditionalFormatRule.prototype.toJSON = function () {
            return {
                "formula": this.m_formula,
                "formulaLocal": this.m_formulaLocal,
                "formulaR1C1": this.m_formulaR1C1
            };
        };
        return ConditionalFormatRule;
    }(OfficeExtension.ClientObject));
    Excel.ConditionalFormatRule = ConditionalFormatRule;
    var IconSetConditionalFormat = (function (_super) {
        __extends(IconSetConditionalFormat, _super);
        function IconSetConditionalFormat() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(IconSetConditionalFormat.prototype, "_className", {
            get: function () {
                return "IconSetConditionalFormat";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(IconSetConditionalFormat.prototype, "criteria", {
            get: function () {
                _throwIfNotLoaded("criteria", this.m_criteria, "IconSetConditionalFormat", this._isNull);
                return this.m_criteria;
            },
            set: function (value) {
                this.m_criteria = value;
                _createSetPropertyAction(this.context, this, "Criteria", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(IconSetConditionalFormat.prototype, "reverseIconOrder", {
            get: function () {
                _throwIfNotLoaded("reverseIconOrder", this.m_reverseIconOrder, "IconSetConditionalFormat", this._isNull);
                return this.m_reverseIconOrder;
            },
            set: function (value) {
                this.m_reverseIconOrder = value;
                _createSetPropertyAction(this.context, this, "ReverseIconOrder", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(IconSetConditionalFormat.prototype, "showIconOnly", {
            get: function () {
                _throwIfNotLoaded("showIconOnly", this.m_showIconOnly, "IconSetConditionalFormat", this._isNull);
                return this.m_showIconOnly;
            },
            set: function (value) {
                this.m_showIconOnly = value;
                _createSetPropertyAction(this.context, this, "ShowIconOnly", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(IconSetConditionalFormat.prototype, "style", {
            get: function () {
                _throwIfNotLoaded("style", this.m_style, "IconSetConditionalFormat", this._isNull);
                return this.m_style;
            },
            set: function (value) {
                this.m_style = value;
                _createSetPropertyAction(this.context, this, "Style", value);
            },
            enumerable: true,
            configurable: true
        });
        IconSetConditionalFormat.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["reverseIconOrder", "showIconOnly", "style", "criteria"], [], []);
        };
        IconSetConditionalFormat.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Criteria"])) {
                this.m_criteria = obj["Criteria"];
            }
            if (!_isUndefined(obj["ReverseIconOrder"])) {
                this.m_reverseIconOrder = obj["ReverseIconOrder"];
            }
            if (!_isUndefined(obj["ShowIconOnly"])) {
                this.m_showIconOnly = obj["ShowIconOnly"];
            }
            if (!_isUndefined(obj["Style"])) {
                this.m_style = obj["Style"];
            }
        };
        IconSetConditionalFormat.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        IconSetConditionalFormat.prototype.toJSON = function () {
            return {
                "criteria": this.m_criteria,
                "reverseIconOrder": this.m_reverseIconOrder,
                "showIconOnly": this.m_showIconOnly,
                "style": this.m_style
            };
        };
        return IconSetConditionalFormat;
    }(OfficeExtension.ClientObject));
    Excel.IconSetConditionalFormat = IconSetConditionalFormat;
    var ColorScaleConditionalFormat = (function (_super) {
        __extends(ColorScaleConditionalFormat, _super);
        function ColorScaleConditionalFormat() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ColorScaleConditionalFormat.prototype, "_className", {
            get: function () {
                return "ColorScaleConditionalFormat";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ColorScaleConditionalFormat.prototype, "criteria", {
            get: function () {
                _throwIfNotLoaded("criteria", this.m_criteria, "ColorScaleConditionalFormat", this._isNull);
                return this.m_criteria;
            },
            set: function (value) {
                this.m_criteria = value;
                _createSetPropertyAction(this.context, this, "Criteria", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ColorScaleConditionalFormat.prototype, "threeColorScale", {
            get: function () {
                _throwIfNotLoaded("threeColorScale", this.m_threeColorScale, "ColorScaleConditionalFormat", this._isNull);
                return this.m_threeColorScale;
            },
            enumerable: true,
            configurable: true
        });
        ColorScaleConditionalFormat.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["criteria"], [], []);
        };
        ColorScaleConditionalFormat.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Criteria"])) {
                this.m_criteria = obj["Criteria"];
            }
            if (!_isUndefined(obj["ThreeColorScale"])) {
                this.m_threeColorScale = obj["ThreeColorScale"];
            }
        };
        ColorScaleConditionalFormat.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ColorScaleConditionalFormat.prototype.toJSON = function () {
            return {
                "criteria": this.m_criteria,
                "threeColorScale": this.m_threeColorScale
            };
        };
        return ColorScaleConditionalFormat;
    }(OfficeExtension.ClientObject));
    Excel.ColorScaleConditionalFormat = ColorScaleConditionalFormat;
    var TopBottomConditionalFormat = (function (_super) {
        __extends(TopBottomConditionalFormat, _super);
        function TopBottomConditionalFormat() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(TopBottomConditionalFormat.prototype, "_className", {
            get: function () {
                return "TopBottomConditionalFormat";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(TopBottomConditionalFormat.prototype, "format", {
            get: function () {
                if (!this.m_format) {
                    this.m_format = new Excel.ConditionalRangeFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
                }
                return this.m_format;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(TopBottomConditionalFormat.prototype, "rule", {
            get: function () {
                _throwIfNotLoaded("rule", this.m_rule, "TopBottomConditionalFormat", this._isNull);
                return this.m_rule;
            },
            set: function (value) {
                this.m_rule = value;
                _createSetPropertyAction(this.context, this, "Rule", value);
            },
            enumerable: true,
            configurable: true
        });
        TopBottomConditionalFormat.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["rule"], ["format"], []);
        };
        TopBottomConditionalFormat.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Rule"])) {
                this.m_rule = obj["Rule"];
            }
            _handleNavigationPropertyResults(this, obj, ["format", "Format"]);
        };
        TopBottomConditionalFormat.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        TopBottomConditionalFormat.prototype.toJSON = function () {
            return {
                "format": this.m_format,
                "rule": this.m_rule
            };
        };
        return TopBottomConditionalFormat;
    }(OfficeExtension.ClientObject));
    Excel.TopBottomConditionalFormat = TopBottomConditionalFormat;
    var PresetCriteriaConditionalFormat = (function (_super) {
        __extends(PresetCriteriaConditionalFormat, _super);
        function PresetCriteriaConditionalFormat() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(PresetCriteriaConditionalFormat.prototype, "_className", {
            get: function () {
                return "PresetCriteriaConditionalFormat";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PresetCriteriaConditionalFormat.prototype, "format", {
            get: function () {
                if (!this.m_format) {
                    this.m_format = new Excel.ConditionalRangeFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
                }
                return this.m_format;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PresetCriteriaConditionalFormat.prototype, "rule", {
            get: function () {
                _throwIfNotLoaded("rule", this.m_rule, "PresetCriteriaConditionalFormat", this._isNull);
                return this.m_rule;
            },
            set: function (value) {
                this.m_rule = value;
                _createSetPropertyAction(this.context, this, "Rule", value);
            },
            enumerable: true,
            configurable: true
        });
        PresetCriteriaConditionalFormat.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["rule"], ["format"], []);
        };
        PresetCriteriaConditionalFormat.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Rule"])) {
                this.m_rule = obj["Rule"];
            }
            _handleNavigationPropertyResults(this, obj, ["format", "Format"]);
        };
        PresetCriteriaConditionalFormat.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        PresetCriteriaConditionalFormat.prototype.toJSON = function () {
            return {
                "format": this.m_format,
                "rule": this.m_rule
            };
        };
        return PresetCriteriaConditionalFormat;
    }(OfficeExtension.ClientObject));
    Excel.PresetCriteriaConditionalFormat = PresetCriteriaConditionalFormat;
    var TextConditionalFormat = (function (_super) {
        __extends(TextConditionalFormat, _super);
        function TextConditionalFormat() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(TextConditionalFormat.prototype, "_className", {
            get: function () {
                return "TextConditionalFormat";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(TextConditionalFormat.prototype, "format", {
            get: function () {
                if (!this.m_format) {
                    this.m_format = new Excel.ConditionalRangeFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
                }
                return this.m_format;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(TextConditionalFormat.prototype, "rule", {
            get: function () {
                _throwIfNotLoaded("rule", this.m_rule, "TextConditionalFormat", this._isNull);
                return this.m_rule;
            },
            set: function (value) {
                this.m_rule = value;
                _createSetPropertyAction(this.context, this, "Rule", value);
            },
            enumerable: true,
            configurable: true
        });
        TextConditionalFormat.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["rule"], ["format"], []);
        };
        TextConditionalFormat.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Rule"])) {
                this.m_rule = obj["Rule"];
            }
            _handleNavigationPropertyResults(this, obj, ["format", "Format"]);
        };
        TextConditionalFormat.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        TextConditionalFormat.prototype.toJSON = function () {
            return {
                "format": this.m_format,
                "rule": this.m_rule
            };
        };
        return TextConditionalFormat;
    }(OfficeExtension.ClientObject));
    Excel.TextConditionalFormat = TextConditionalFormat;
    var CellValueConditionalFormat = (function (_super) {
        __extends(CellValueConditionalFormat, _super);
        function CellValueConditionalFormat() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(CellValueConditionalFormat.prototype, "_className", {
            get: function () {
                return "CellValueConditionalFormat";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(CellValueConditionalFormat.prototype, "format", {
            get: function () {
                if (!this.m_format) {
                    this.m_format = new Excel.ConditionalRangeFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
                }
                return this.m_format;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(CellValueConditionalFormat.prototype, "rule", {
            get: function () {
                _throwIfNotLoaded("rule", this.m_rule, "CellValueConditionalFormat", this._isNull);
                return this.m_rule;
            },
            set: function (value) {
                this.m_rule = value;
                _createSetPropertyAction(this.context, this, "Rule", value);
            },
            enumerable: true,
            configurable: true
        });
        CellValueConditionalFormat.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["rule"], ["format"], []);
        };
        CellValueConditionalFormat.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Rule"])) {
                this.m_rule = obj["Rule"];
            }
            _handleNavigationPropertyResults(this, obj, ["format", "Format"]);
        };
        CellValueConditionalFormat.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        CellValueConditionalFormat.prototype.toJSON = function () {
            return {
                "format": this.m_format,
                "rule": this.m_rule
            };
        };
        return CellValueConditionalFormat;
    }(OfficeExtension.ClientObject));
    Excel.CellValueConditionalFormat = CellValueConditionalFormat;
    var ConditionalRangeFormat = (function (_super) {
        __extends(ConditionalRangeFormat, _super);
        function ConditionalRangeFormat() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ConditionalRangeFormat.prototype, "_className", {
            get: function () {
                return "ConditionalRangeFormat";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalRangeFormat.prototype, "borders", {
            get: function () {
                if (!this.m_borders) {
                    this.m_borders = new Excel.ConditionalRangeBorderCollection(this.context, _createPropertyObjectPath(this.context, this, "Borders", true, false));
                }
                return this.m_borders;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalRangeFormat.prototype, "fill", {
            get: function () {
                if (!this.m_fill) {
                    this.m_fill = new Excel.ConditionalRangeFill(this.context, _createPropertyObjectPath(this.context, this, "Fill", false, false));
                }
                return this.m_fill;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalRangeFormat.prototype, "font", {
            get: function () {
                if (!this.m_font) {
                    this.m_font = new Excel.ConditionalRangeFont(this.context, _createPropertyObjectPath(this.context, this, "Font", false, false));
                }
                return this.m_font;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalRangeFormat.prototype, "numberFormat", {
            get: function () {
                _throwIfNotLoaded("numberFormat", this.m_numberFormat, "ConditionalRangeFormat", this._isNull);
                return this.m_numberFormat;
            },
            set: function (value) {
                this.m_numberFormat = value;
                _createSetPropertyAction(this.context, this, "NumberFormat", value);
            },
            enumerable: true,
            configurable: true
        });
        ConditionalRangeFormat.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["numberFormat"], [], [
                "borders",
                "fill",
                "font",
                "borders",
                "fill",
                "font"
            ]);
        };
        ConditionalRangeFormat.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["NumberFormat"])) {
                this.m_numberFormat = obj["NumberFormat"];
            }
            _handleNavigationPropertyResults(this, obj, ["borders", "Borders", "fill", "Fill", "font", "Font"]);
        };
        ConditionalRangeFormat.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ConditionalRangeFormat.prototype.toJSON = function () {
            return {
                "numberFormat": this.m_numberFormat
            };
        };
        return ConditionalRangeFormat;
    }(OfficeExtension.ClientObject));
    Excel.ConditionalRangeFormat = ConditionalRangeFormat;
    var ConditionalRangeFont = (function (_super) {
        __extends(ConditionalRangeFont, _super);
        function ConditionalRangeFont() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ConditionalRangeFont.prototype, "_className", {
            get: function () {
                return "ConditionalRangeFont";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalRangeFont.prototype, "bold", {
            get: function () {
                _throwIfNotLoaded("bold", this.m_bold, "ConditionalRangeFont", this._isNull);
                return this.m_bold;
            },
            set: function (value) {
                this.m_bold = value;
                _createSetPropertyAction(this.context, this, "Bold", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalRangeFont.prototype, "color", {
            get: function () {
                _throwIfNotLoaded("color", this.m_color, "ConditionalRangeFont", this._isNull);
                return this.m_color;
            },
            set: function (value) {
                this.m_color = value;
                _createSetPropertyAction(this.context, this, "Color", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalRangeFont.prototype, "italic", {
            get: function () {
                _throwIfNotLoaded("italic", this.m_italic, "ConditionalRangeFont", this._isNull);
                return this.m_italic;
            },
            set: function (value) {
                this.m_italic = value;
                _createSetPropertyAction(this.context, this, "Italic", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalRangeFont.prototype, "strikethrough", {
            get: function () {
                _throwIfNotLoaded("strikethrough", this.m_strikethrough, "ConditionalRangeFont", this._isNull);
                return this.m_strikethrough;
            },
            set: function (value) {
                this.m_strikethrough = value;
                _createSetPropertyAction(this.context, this, "Strikethrough", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalRangeFont.prototype, "underline", {
            get: function () {
                _throwIfNotLoaded("underline", this.m_underline, "ConditionalRangeFont", this._isNull);
                return this.m_underline;
            },
            set: function (value) {
                this.m_underline = value;
                _createSetPropertyAction(this.context, this, "Underline", value);
            },
            enumerable: true,
            configurable: true
        });
        ConditionalRangeFont.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["color", "italic", "bold", "underline", "strikethrough"], [], []);
        };
        ConditionalRangeFont.prototype.clear = function () {
            _createMethodAction(this.context, this, "Clear", 0, []);
        };
        ConditionalRangeFont.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Bold"])) {
                this.m_bold = obj["Bold"];
            }
            if (!_isUndefined(obj["Color"])) {
                this.m_color = obj["Color"];
            }
            if (!_isUndefined(obj["Italic"])) {
                this.m_italic = obj["Italic"];
            }
            if (!_isUndefined(obj["Strikethrough"])) {
                this.m_strikethrough = obj["Strikethrough"];
            }
            if (!_isUndefined(obj["Underline"])) {
                this.m_underline = obj["Underline"];
            }
        };
        ConditionalRangeFont.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ConditionalRangeFont.prototype.toJSON = function () {
            return {
                "bold": this.m_bold,
                "color": this.m_color,
                "italic": this.m_italic,
                "strikethrough": this.m_strikethrough,
                "underline": this.m_underline
            };
        };
        return ConditionalRangeFont;
    }(OfficeExtension.ClientObject));
    Excel.ConditionalRangeFont = ConditionalRangeFont;
    var ConditionalRangeFill = (function (_super) {
        __extends(ConditionalRangeFill, _super);
        function ConditionalRangeFill() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ConditionalRangeFill.prototype, "_className", {
            get: function () {
                return "ConditionalRangeFill";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalRangeFill.prototype, "color", {
            get: function () {
                _throwIfNotLoaded("color", this.m_color, "ConditionalRangeFill", this._isNull);
                return this.m_color;
            },
            set: function (value) {
                this.m_color = value;
                _createSetPropertyAction(this.context, this, "Color", value);
            },
            enumerable: true,
            configurable: true
        });
        ConditionalRangeFill.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["color"], [], []);
        };
        ConditionalRangeFill.prototype.clear = function () {
            _createMethodAction(this.context, this, "Clear", 0, []);
        };
        ConditionalRangeFill.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Color"])) {
                this.m_color = obj["Color"];
            }
        };
        ConditionalRangeFill.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ConditionalRangeFill.prototype.toJSON = function () {
            return {
                "color": this.m_color
            };
        };
        return ConditionalRangeFill;
    }(OfficeExtension.ClientObject));
    Excel.ConditionalRangeFill = ConditionalRangeFill;
    var ConditionalRangeBorder = (function (_super) {
        __extends(ConditionalRangeBorder, _super);
        function ConditionalRangeBorder() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ConditionalRangeBorder.prototype, "_className", {
            get: function () {
                return "ConditionalRangeBorder";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalRangeBorder.prototype, "color", {
            get: function () {
                _throwIfNotLoaded("color", this.m_color, "ConditionalRangeBorder", this._isNull);
                return this.m_color;
            },
            set: function (value) {
                this.m_color = value;
                _createSetPropertyAction(this.context, this, "Color", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalRangeBorder.prototype, "sideIndex", {
            get: function () {
                _throwIfNotLoaded("sideIndex", this.m_sideIndex, "ConditionalRangeBorder", this._isNull);
                return this.m_sideIndex;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalRangeBorder.prototype, "style", {
            get: function () {
                _throwIfNotLoaded("style", this.m_style, "ConditionalRangeBorder", this._isNull);
                return this.m_style;
            },
            set: function (value) {
                this.m_style = value;
                _createSetPropertyAction(this.context, this, "Style", value);
            },
            enumerable: true,
            configurable: true
        });
        ConditionalRangeBorder.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["style", "color"], [], []);
        };
        ConditionalRangeBorder.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Color"])) {
                this.m_color = obj["Color"];
            }
            if (!_isUndefined(obj["SideIndex"])) {
                this.m_sideIndex = obj["SideIndex"];
            }
            if (!_isUndefined(obj["Style"])) {
                this.m_style = obj["Style"];
            }
        };
        ConditionalRangeBorder.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ConditionalRangeBorder.prototype.toJSON = function () {
            return {
                "color": this.m_color,
                "sideIndex": this.m_sideIndex,
                "style": this.m_style
            };
        };
        return ConditionalRangeBorder;
    }(OfficeExtension.ClientObject));
    Excel.ConditionalRangeBorder = ConditionalRangeBorder;
    var ConditionalRangeBorderCollection = (function (_super) {
        __extends(ConditionalRangeBorderCollection, _super);
        function ConditionalRangeBorderCollection() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ConditionalRangeBorderCollection.prototype, "_className", {
            get: function () {
                return "ConditionalRangeBorderCollection";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalRangeBorderCollection.prototype, "bottom", {
            get: function () {
                if (!this.m_bottom) {
                    this.m_bottom = new Excel.ConditionalRangeBorder(this.context, _createPropertyObjectPath(this.context, this, "Bottom", false, false));
                }
                return this.m_bottom;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalRangeBorderCollection.prototype, "left", {
            get: function () {
                if (!this.m_left) {
                    this.m_left = new Excel.ConditionalRangeBorder(this.context, _createPropertyObjectPath(this.context, this, "Left", false, false));
                }
                return this.m_left;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalRangeBorderCollection.prototype, "right", {
            get: function () {
                if (!this.m_right) {
                    this.m_right = new Excel.ConditionalRangeBorder(this.context, _createPropertyObjectPath(this.context, this, "Right", false, false));
                }
                return this.m_right;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalRangeBorderCollection.prototype, "top", {
            get: function () {
                if (!this.m_top) {
                    this.m_top = new Excel.ConditionalRangeBorder(this.context, _createPropertyObjectPath(this.context, this, "Top", false, false));
                }
                return this.m_top;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalRangeBorderCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, "ConditionalRangeBorderCollection", this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalRangeBorderCollection.prototype, "count", {
            get: function () {
                _throwIfNotLoaded("count", this.m_count, "ConditionalRangeBorderCollection", this._isNull);
                return this.m_count;
            },
            enumerable: true,
            configurable: true
        });
        ConditionalRangeBorderCollection.prototype.getItem = function (index) {
            return new Excel.ConditionalRangeBorder(this.context, _createIndexerObjectPath(this.context, this, [index]));
        };
        ConditionalRangeBorderCollection.prototype.getItemAt = function (index) {
            return new Excel.ConditionalRangeBorder(this.context, _createMethodObjectPath(this.context, this, "GetItemAt", 1, [index], false, false, null));
        };
        ConditionalRangeBorderCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Count"])) {
                this.m_count = obj["Count"];
            }
            _handleNavigationPropertyResults(this, obj, ["bottom", "Bottom", "left", "Left", "right", "Right", "top", "Top"]);
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.ConditionalRangeBorder(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        ConditionalRangeBorderCollection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ConditionalRangeBorderCollection.prototype.toJSON = function () {
            return {
                "count": this.m_count
            };
        };
        return ConditionalRangeBorderCollection;
    }(OfficeExtension.ClientObject));
    Excel.ConditionalRangeBorderCollection = ConditionalRangeBorderCollection;
    var CustomFunction = (function (_super) {
        __extends(CustomFunction, _super);
        function CustomFunction() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(CustomFunction.prototype, "_className", {
            get: function () {
                return "CustomFunction";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(CustomFunction.prototype, "description", {
            get: function () {
                _throwIfNotLoaded("description", this.m_description, "CustomFunction", this._isNull);
                return this.m_description;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(CustomFunction.prototype, "id", {
            get: function () {
                _throwIfNotLoaded("id", this.m_id, "CustomFunction", this._isNull);
                return this.m_id;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(CustomFunction.prototype, "name", {
            get: function () {
                _throwIfNotLoaded("name", this.m_name, "CustomFunction", this._isNull);
                return this.m_name;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(CustomFunction.prototype, "parameters", {
            get: function () {
                _throwIfNotLoaded("parameters", this.m_parameters, "CustomFunction", this._isNull);
                return this.m_parameters;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(CustomFunction.prototype, "resultDimensionality", {
            get: function () {
                _throwIfNotLoaded("resultDimensionality", this.m_resultDimensionality, "CustomFunction", this._isNull);
                return this.m_resultDimensionality;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(CustomFunction.prototype, "resultType", {
            get: function () {
                _throwIfNotLoaded("resultType", this.m_resultType, "CustomFunction", this._isNull);
                return this.m_resultType;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(CustomFunction.prototype, "streaming", {
            get: function () {
                _throwIfNotLoaded("streaming", this.m_streaming, "CustomFunction", this._isNull);
                return this.m_streaming;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(CustomFunction.prototype, "type", {
            get: function () {
                _throwIfNotLoaded("type", this.m_type, "CustomFunction", this._isNull);
                return this.m_type;
            },
            enumerable: true,
            configurable: true
        });
        CustomFunction.prototype.delete = function () {
            _createMethodAction(this.context, this, "Delete", 0, []);
        };
        CustomFunction.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Description"])) {
                this.m_description = obj["Description"];
            }
            if (!_isUndefined(obj["Id"])) {
                this.m_id = obj["Id"];
            }
            if (!_isUndefined(obj["Name"])) {
                this.m_name = obj["Name"];
            }
            if (!_isUndefined(obj["Parameters"])) {
                this.m_parameters = obj["Parameters"];
            }
            if (!_isUndefined(obj["ResultDimensionality"])) {
                this.m_resultDimensionality = obj["ResultDimensionality"];
            }
            if (!_isUndefined(obj["ResultType"])) {
                this.m_resultType = obj["ResultType"];
            }
            if (!_isUndefined(obj["Streaming"])) {
                this.m_streaming = obj["Streaming"];
            }
            if (!_isUndefined(obj["Type"])) {
                this.m_type = obj["Type"];
            }
        };
        CustomFunction.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        CustomFunction.prototype._handleIdResult = function (value) {
            _super.prototype._handleIdResult.call(this, value);
            if (_isNullOrUndefined(value)) {
                return;
            }
            if (!_isUndefined(value["Id"])) {
                this.m_id = value["Id"];
            }
        };
        CustomFunction.prototype.toJSON = function () {
            return {
                "description": this.m_description,
                "id": this.m_id,
                "name": this.m_name,
                "parameters": this.m_parameters,
                "resultDimensionality": this.m_resultDimensionality,
                "resultType": this.m_resultType,
                "streaming": this.m_streaming,
                "type": this.m_type
            };
        };
        return CustomFunction;
    }(OfficeExtension.ClientObject));
    Excel.CustomFunction = CustomFunction;
    var CustomFunctionProxy = (function () {
        function CustomFunctionProxy() {
            this._isInit = false;
        }
        CustomFunctionProxy.prototype.addAll = function (context) {
            if (!Excel.Script || !Excel.Script.CustomFunctions) {
                return;
            }
            for (var namespace in Excel.Script.CustomFunctions) {
                for (var name_1 in Excel.Script.CustomFunctions[namespace]) {
                    this.add(context, namespace + "." + name_1);
                }
            }
        };
        CustomFunctionProxy.prototype.add = function (context, name) {
            if (OfficeExtension.Utility.isNullOrEmptyString(name)) {
                throw OfficeExtension._Internal.RuntimeError._createInvalidArgError("name");
            }
            if (_isNullOrUndefined(Excel.Script) || _isNullOrUndefined(Excel.Script.CustomFunctions)) {
                throw OfficeExtension.Utility.createRuntimeError(ErrorCodes.invalidOperation, OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.customFunctionDefintionMissing), "CustomFunctionProxy.add");
            }
            var nameSplit = CustomFunctionProxy.splitName(name);
            var definitionCollection = Excel.Script.CustomFunctions[nameSplit.namespace];
            if (_isNullOrUndefined(definitionCollection)) {
                throw OfficeExtension.Utility.createRuntimeError(ErrorCodes.invalidOperation, OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.customFunctionDefintionMissing), "CustomFunctionProxy.add");
            }
            var definition = definitionCollection[nameSplit.name];
            if (_isNullOrUndefined(definition)) {
                throw OfficeExtension.Utility.createRuntimeError(ErrorCodes.invalidOperation, OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.customFunctionDefintionMissing), "CustomFunctionProxy.add");
            }
            this._ensureInit(context);
            var apiCustomFunction = context.workbook.customFunctions._Add(Excel.CustomFunctionType.script, name, definition.description, definition.result.resultType, definition.result.resultDimensionality ? definition.result.resultDimensionality : Excel.CustomFunctionDimensionality.scalar, definition.options && definition.options.stream ? definition.options.stream : false, definition.parameters);
            return apiCustomFunction;
        };
        CustomFunctionProxy.splitName = function (name) {
            var matches = name.match(/[a-z_][a-z_0-9\.]+/gi);
            if (matches === null || matches.length !== 1 || matches[0] !== name) {
                throw OfficeExtension.Utility.createRuntimeError(ErrorCodes.invalidOperation, "The function name may only contain letters, digits, underscores, and dots.", "CustomFunctionProxy.splitName");
            }
            var splitIndex = name.lastIndexOf(".");
            if (splitIndex < 1 || splitIndex === name.length - 1) {
                throw OfficeExtension.Utility.createRuntimeError(ErrorCodes.invalidOperation, "The function name must contain a non-empty namespace and a non-empty short name.", "CustomFunctionProxy.splitName");
            }
            var nameSplit = {
                namespace: name.substring(0, splitIndex),
                name: name.substr(splitIndex + 1)
            };
            return nameSplit;
        };
        CustomFunctionProxy.prototype._ensureInit = function (context) {
            if (!this._isInit) {
                context.workbook._onMessage.add(this._handleMessage.bind(this));
                this._isInit = true;
            }
        };
        CustomFunctionProxy.prototype._handleMessage = function (args) {
            var _this = this;
            OfficeExtension.Utility.checkArgumentNull(args, "args");
            var entryArray = args.entries;
            var resultArray = [];
            var _loop_1 = function () {
                if (entryArray[i].messageCategory === 1) {
                    OfficeExtension.Utility.checkArgumentNull(args.workbook, "workbook");
                    var messageJson = entryArray[i].message;
                    if (OfficeExtension.Utility.isNullOrEmptyString(messageJson)) {
                        throw OfficeExtension.Utility.createRuntimeError(ErrorCodes.generalException, "messageJson", "CustomFunctionProxy._handleMessage");
                    }
                    var message_1 = JSON.parse(messageJson);
                    if (_isNullOrUndefined(message_1)) {
                        throw OfficeExtension.Utility.createRuntimeError(ErrorCodes.generalException, "message", "CustomFunctionProxy._handleMessage");
                    }
                    if (_isNullOrUndefined(message_1.invocationId) || message_1.invocationId < 0) {
                        throw OfficeExtension.Utility.createRuntimeError(ErrorCodes.generalException, "invocationId", "CustomFunctionProxy._handleMessage");
                    }
                    if (_isNullOrUndefined(message_1.functionName)) {
                        throw OfficeExtension.Utility.createRuntimeError(ErrorCodes.generalException, "functionName", "CustomFunctionProxy._handleMessage");
                    }
                    var nameSplit = CustomFunctionProxy.splitName(message_1.functionName);
                    var definitionCollection = Excel.Script.CustomFunctions[nameSplit.namespace];
                    if (_isNullOrUndefined(definitionCollection)) {
                        throw OfficeExtension.Utility.createRuntimeError(ErrorCodes.invalidOperation, OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.customFunctionDefintionMissing), "CustomFunctionProxy._handleMessage");
                    }
                    var definition = definitionCollection[nameSplit.name];
                    if (_isNullOrUndefined(definition)) {
                        throw OfficeExtension.Utility.createRuntimeError(ErrorCodes.invalidOperation, OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.customFunctionDefintionMissing), "CustomFunctionProxy._handleMessage");
                    }
                    if (_isNullOrUndefined(definition.call)) {
                        throw OfficeExtension.Utility.createRuntimeError(ErrorCodes.invalidOperation, OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.customFunctionImplementationMissing), "CustomFunctionProxy._handleMessage");
                    }
                    var parameterValues = message_1.parameterValues;
                    if (!_isNullOrUndefined(definition.options) && definition.options.stream) {
                        parameterValues.push(function (result) {
                            return _this._setResult(args.workbook, message_1.invocationId, result);
                        });
                    }
                    var result = definition.call.apply(null, parameterValues);
                    if (!_isNullOrUndefined(definition.options) && definition.options.stream) {
                        resultArray.push(null);
                    }
                    else {
                        if (typeof result === "object" && typeof result.then === "function") {
                            resultArray.push(result.then(function (value) {
                                return _this._setResult(args.workbook, message_1.invocationId, value);
                            }, function (reason) {
                                return _this._setResult(args.workbook, message_1.invocationId, "Error: " + reason);
                            }));
                        }
                        else {
                            resultArray.push(this_1._setResult(args.workbook, message_1.invocationId, result));
                        }
                    }
                }
            };
            var this_1 = this;
            for (var i = 0; i < entryArray.length; i++) {
                _loop_1();
            }
            return OfficeExtension.Promise.all(resultArray);
        };
        CustomFunctionProxy.prototype._setResult = function (workbook, invocationId, result) {
            workbook.customFunctions._SetInvocationResult(invocationId, result);
            return workbook.context.sync();
        };
        return CustomFunctionProxy;
    }());
    Excel.CustomFunctionProxy = CustomFunctionProxy;
    Excel.customFunctionProxy = new CustomFunctionProxy();
    var CustomFunctionCollection = (function (_super) {
        __extends(CustomFunctionCollection, _super);
        function CustomFunctionCollection() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(CustomFunctionCollection.prototype, "_className", {
            get: function () {
                return "CustomFunctionCollection";
            },
            enumerable: true,
            configurable: true
        });
        CustomFunctionCollection.prototype.addAll = function () {
            this.deleteAll();
            Excel.customFunctionProxy.addAll(this.context);
        };
        CustomFunctionCollection.prototype.add = function (name) {
            return Excel.customFunctionProxy.add(this.context, name);
        };
        Object.defineProperty(CustomFunctionCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, "CustomFunctionCollection", this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        CustomFunctionCollection.prototype.deleteAll = function () {
            _createMethodAction(this.context, this, "DeleteAll", 0, []);
        };
        CustomFunctionCollection.prototype.getCount = function () {
            var action = _createMethodAction(this.context, this, "GetCount", 1, []);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        CustomFunctionCollection.prototype.getItem = function (name) {
            return new Excel.CustomFunction(this.context, _createIndexerObjectPath(this.context, this, [name]));
        };
        CustomFunctionCollection.prototype.getItemOrNullObject = function (name) {
            return new Excel.CustomFunction(this.context, _createMethodObjectPath(this.context, this, "GetItemOrNullObject", 1, [name], false, false, null));
        };
        CustomFunctionCollection.prototype.importFromWeb = function (metadataFormat, metadataUrl, name) {
            return new Excel.CustomFunction(this.context, _createMethodObjectPath(this.context, this, "ImportFromWeb", 0, [metadataFormat, metadataUrl, name], false, false, null));
        };
        CustomFunctionCollection.prototype._Add = function (type, name, description, resultType, resultDimensionality, streaming, parameters) {
            return new Excel.CustomFunction(this.context, _createMethodObjectPath(this.context, this, "_Add", 0, [type, name, description, resultType, resultDimensionality, streaming, parameters], false, false, null));
        };
        CustomFunctionCollection.prototype._SetInvocationResult = function (invocationId, result) {
            _createMethodAction(this.context, this, "_SetInvocationResult", 0, [invocationId, result]);
        };
        CustomFunctionCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.CustomFunction(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        CustomFunctionCollection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        CustomFunctionCollection.prototype.toJSON = function () {
            return {};
        };
        return CustomFunctionCollection;
    }(OfficeExtension.ClientObject));
    Excel.CustomFunctionCollection = CustomFunctionCollection;
    var InternalTest = (function (_super) {
        __extends(InternalTest, _super);
        function InternalTest() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(InternalTest.prototype, "_className", {
            get: function () {
                return "InternalTest";
            },
            enumerable: true,
            configurable: true
        });
        InternalTest.prototype.delay = function (seconds) {
            var action = _createMethodAction(this.context, this, "Delay", 0, [seconds]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        InternalTest.prototype.deserializeCustomFunctions = function (serializedString) {
            _throwIfApiNotSupported("InternalTest.deserializeCustomFunctions", _defaultApiSetName, "1.7", _hostName);
            _createMethodAction(this.context, this, "DeserializeCustomFunctions", 0, [serializedString]);
        };
        InternalTest.prototype.serializeCustomFunctions = function () {
            _throwIfApiNotSupported("InternalTest.serializeCustomFunctions", _defaultApiSetName, "1.7", _hostName);
            var action = _createMethodAction(this.context, this, "SerializeCustomFunctions", 0, []);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        InternalTest.prototype.triggerMessage = function (messageCategory, messageType, targetId, message) {
            _throwIfApiNotSupported("InternalTest.triggerMessage", _defaultApiSetName, "1.8", _hostName);
            _createMethodAction(this.context, this, "TriggerMessage", 0, [messageCategory, messageType, targetId, message]);
        };
        InternalTest.prototype.triggerPostProcess = function () {
            _throwIfApiNotSupported("InternalTest.triggerPostProcess", _defaultApiSetName, "1.8", _hostName);
            _createMethodAction(this.context, this, "TriggerPostProcess", 0, []);
        };
        InternalTest.prototype.triggerTestEvent = function (prop1, worksheet) {
            _throwIfApiNotSupported("InternalTest.triggerTestEvent", _defaultApiSetName, "1.8", _hostName);
            _createMethodAction(this.context, this, "TriggerTestEvent", 0, [prop1, worksheet]);
        };
        InternalTest.prototype._RegisterTestEvent = function () {
            _throwIfApiNotSupported("InternalTest._RegisterTestEvent", _defaultApiSetName, "1.8", _hostName);
            _createMethodAction(this.context, this, "_RegisterTestEvent", 0, []);
        };
        InternalTest.prototype._UnregisterTestEvent = function () {
            _throwIfApiNotSupported("InternalTest._UnregisterTestEvent", _defaultApiSetName, "1.8", _hostName);
            _createMethodAction(this.context, this, "_UnregisterTestEvent", 0, []);
        };
        InternalTest.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
        };
        Object.defineProperty(InternalTest.prototype, "onTestEvent", {
            get: function () {
                var _this = this;
                _throwIfApiNotSupported("InternalTest.onTestEvent", _defaultApiSetName, "1.8", _hostName);
                if (!this.m_testEvent) {
                    this.m_testEvent = new OfficeExtension.GenericEventHandlers(this.context, this, "TestEvent", {
                        eventType: 1,
                        registerFunc: function () { return _this._RegisterTestEvent(); },
                        unregisterFunc: function () { return _this._UnregisterTestEvent(); },
                        getTargetIdFunc: function () { return ""; },
                        eventArgsTransformFunc: function (value) {
                            var newArgs = {
                                prop1: value.prop1,
                                worksheet: _this.context.workbook.worksheets.getItem(value.worksheetId)
                            };
                            return OfficeExtension.Utility._createPromiseFromResult(newArgs);
                        }
                    });
                }
                return this.m_testEvent;
            },
            enumerable: true,
            configurable: true
        });
        InternalTest.prototype.toJSON = function () {
            return {};
        };
        return InternalTest;
    }(OfficeExtension.ClientObject));
    Excel.InternalTest = InternalTest;
    var AxisType;
    (function (AxisType) {
        AxisType.invalid = "Invalid";
        AxisType.category = "Category";
        AxisType.value = "Value";
        AxisType.series = "Series";
    })(AxisType = Excel.AxisType || (Excel.AxisType = {}));
    var BindingType;
    (function (BindingType) {
        BindingType.range = "Range";
        BindingType.table = "Table";
        BindingType.text = "Text";
    })(BindingType = Excel.BindingType || (Excel.BindingType = {}));
    var BorderIndex;
    (function (BorderIndex) {
        BorderIndex.edgeTop = "EdgeTop";
        BorderIndex.edgeBottom = "EdgeBottom";
        BorderIndex.edgeLeft = "EdgeLeft";
        BorderIndex.edgeRight = "EdgeRight";
        BorderIndex.insideVertical = "InsideVertical";
        BorderIndex.insideHorizontal = "InsideHorizontal";
        BorderIndex.diagonalDown = "DiagonalDown";
        BorderIndex.diagonalUp = "DiagonalUp";
    })(BorderIndex = Excel.BorderIndex || (Excel.BorderIndex = {}));
    var BorderLineStyle;
    (function (BorderLineStyle) {
        BorderLineStyle.none = "None";
        BorderLineStyle.continuous = "Continuous";
        BorderLineStyle.dash = "Dash";
        BorderLineStyle.dashDot = "DashDot";
        BorderLineStyle.dashDotDot = "DashDotDot";
        BorderLineStyle.dot = "Dot";
        BorderLineStyle.double = "Double";
        BorderLineStyle.slantDashDot = "SlantDashDot";
    })(BorderLineStyle = Excel.BorderLineStyle || (Excel.BorderLineStyle = {}));
    var BorderWeight;
    (function (BorderWeight) {
        BorderWeight.hairline = "Hairline";
        BorderWeight.thin = "Thin";
        BorderWeight.medium = "Medium";
        BorderWeight.thick = "Thick";
    })(BorderWeight = Excel.BorderWeight || (Excel.BorderWeight = {}));
    var CalculationMode;
    (function (CalculationMode) {
        CalculationMode.automatic = "Automatic";
        CalculationMode.automaticExceptTables = "AutomaticExceptTables";
        CalculationMode.manual = "Manual";
    })(CalculationMode = Excel.CalculationMode || (Excel.CalculationMode = {}));
    var CalculationType;
    (function (CalculationType) {
        CalculationType.recalculate = "Recalculate";
        CalculationType.full = "Full";
        CalculationType.fullRebuild = "FullRebuild";
    })(CalculationType = Excel.CalculationType || (Excel.CalculationType = {}));
    var ClearApplyTo;
    (function (ClearApplyTo) {
        ClearApplyTo.all = "All";
        ClearApplyTo.formats = "Formats";
        ClearApplyTo.contents = "Contents";
        ClearApplyTo.hyperlinks = "Hyperlinks";
    })(ClearApplyTo = Excel.ClearApplyTo || (Excel.ClearApplyTo = {}));
    var ChartAxisDisplayUnit;
    (function (ChartAxisDisplayUnit) {
        ChartAxisDisplayUnit.none = "None";
        ChartAxisDisplayUnit.hundreds = "Hundreds";
        ChartAxisDisplayUnit.thousands = "Thousands";
        ChartAxisDisplayUnit.tenThousands = "TenThousands";
        ChartAxisDisplayUnit.hundredThousands = "HundredThousands";
        ChartAxisDisplayUnit.millions = "Millions";
        ChartAxisDisplayUnit.tenMillions = "TenMillions";
        ChartAxisDisplayUnit.hundredMillions = "HundredMillions";
        ChartAxisDisplayUnit.billions = "Billions";
        ChartAxisDisplayUnit.trillions = "Trillions";
        ChartAxisDisplayUnit.custom = "Custom";
    })(ChartAxisDisplayUnit = Excel.ChartAxisDisplayUnit || (Excel.ChartAxisDisplayUnit = {}));
    var ChartAxisTimeUnit;
    (function (ChartAxisTimeUnit) {
        ChartAxisTimeUnit.days = "Days";
        ChartAxisTimeUnit.months = "Months";
        ChartAxisTimeUnit.years = "Years";
    })(ChartAxisTimeUnit = Excel.ChartAxisTimeUnit || (Excel.ChartAxisTimeUnit = {}));
    var ChartAxisCategoryType;
    (function (ChartAxisCategoryType) {
        ChartAxisCategoryType.automatic = "Automatic";
        ChartAxisCategoryType.textAxis = "TextAxis";
        ChartAxisCategoryType.dateAxis = "DateAxis";
    })(ChartAxisCategoryType = Excel.ChartAxisCategoryType || (Excel.ChartAxisCategoryType = {}));
    var ChartDataLabelPosition;
    (function (ChartDataLabelPosition) {
        ChartDataLabelPosition.invalid = "Invalid";
        ChartDataLabelPosition.none = "None";
        ChartDataLabelPosition.center = "Center";
        ChartDataLabelPosition.insideEnd = "InsideEnd";
        ChartDataLabelPosition.insideBase = "InsideBase";
        ChartDataLabelPosition.outsideEnd = "OutsideEnd";
        ChartDataLabelPosition.left = "Left";
        ChartDataLabelPosition.right = "Right";
        ChartDataLabelPosition.top = "Top";
        ChartDataLabelPosition.bottom = "Bottom";
        ChartDataLabelPosition.bestFit = "BestFit";
        ChartDataLabelPosition.callout = "Callout";
    })(ChartDataLabelPosition = Excel.ChartDataLabelPosition || (Excel.ChartDataLabelPosition = {}));
    var ChartLegendPosition;
    (function (ChartLegendPosition) {
        ChartLegendPosition.invalid = "Invalid";
        ChartLegendPosition.top = "Top";
        ChartLegendPosition.bottom = "Bottom";
        ChartLegendPosition.left = "Left";
        ChartLegendPosition.right = "Right";
        ChartLegendPosition.corner = "Corner";
        ChartLegendPosition.custom = "Custom";
    })(ChartLegendPosition = Excel.ChartLegendPosition || (Excel.ChartLegendPosition = {}));
    var ChartSeriesBy;
    (function (ChartSeriesBy) {
        ChartSeriesBy.auto = "Auto";
        ChartSeriesBy.columns = "Columns";
        ChartSeriesBy.rows = "Rows";
    })(ChartSeriesBy = Excel.ChartSeriesBy || (Excel.ChartSeriesBy = {}));
    var ChartTextHorizontalAlignment;
    (function (ChartTextHorizontalAlignment) {
        ChartTextHorizontalAlignment.center = "Center";
        ChartTextHorizontalAlignment.left = "Left";
        ChartTextHorizontalAlignment.right = "Right";
        ChartTextHorizontalAlignment.justify = "Justify";
        ChartTextHorizontalAlignment.distributed = "Distributed";
    })(ChartTextHorizontalAlignment = Excel.ChartTextHorizontalAlignment || (Excel.ChartTextHorizontalAlignment = {}));
    var ChartType;
    (function (ChartType) {
        ChartType.invalid = "Invalid";
        ChartType.columnClustered = "ColumnClustered";
        ChartType.columnStacked = "ColumnStacked";
        ChartType.columnStacked100 = "ColumnStacked100";
        ChartType._3DColumnClustered = "3DColumnClustered";
        ChartType._3DColumnStacked = "3DColumnStacked";
        ChartType._3DColumnStacked100 = "3DColumnStacked100";
        ChartType.barClustered = "BarClustered";
        ChartType.barStacked = "BarStacked";
        ChartType.barStacked100 = "BarStacked100";
        ChartType._3DBarClustered = "3DBarClustered";
        ChartType._3DBarStacked = "3DBarStacked";
        ChartType._3DBarStacked100 = "3DBarStacked100";
        ChartType.lineStacked = "LineStacked";
        ChartType.lineStacked100 = "LineStacked100";
        ChartType.lineMarkers = "LineMarkers";
        ChartType.lineMarkersStacked = "LineMarkersStacked";
        ChartType.lineMarkersStacked100 = "LineMarkersStacked100";
        ChartType.pieOfPie = "PieOfPie";
        ChartType.pieExploded = "PieExploded";
        ChartType._3DPieExploded = "3DPieExploded";
        ChartType.barOfPie = "BarOfPie";
        ChartType.xyscatterSmooth = "XYScatterSmooth";
        ChartType.xyscatterSmoothNoMarkers = "XYScatterSmoothNoMarkers";
        ChartType.xyscatterLines = "XYScatterLines";
        ChartType.xyscatterLinesNoMarkers = "XYScatterLinesNoMarkers";
        ChartType.areaStacked = "AreaStacked";
        ChartType.areaStacked100 = "AreaStacked100";
        ChartType._3DAreaStacked = "3DAreaStacked";
        ChartType._3DAreaStacked100 = "3DAreaStacked100";
        ChartType.doughnutExploded = "DoughnutExploded";
        ChartType.radarMarkers = "RadarMarkers";
        ChartType.radarFilled = "RadarFilled";
        ChartType.surface = "Surface";
        ChartType.surfaceWireframe = "SurfaceWireframe";
        ChartType.surfaceTopView = "SurfaceTopView";
        ChartType.surfaceTopViewWireframe = "SurfaceTopViewWireframe";
        ChartType.bubble = "Bubble";
        ChartType.bubble3DEffect = "Bubble3DEffect";
        ChartType.stockHLC = "StockHLC";
        ChartType.stockOHLC = "StockOHLC";
        ChartType.stockVHLC = "StockVHLC";
        ChartType.stockVOHLC = "StockVOHLC";
        ChartType.cylinderColClustered = "CylinderColClustered";
        ChartType.cylinderColStacked = "CylinderColStacked";
        ChartType.cylinderColStacked100 = "CylinderColStacked100";
        ChartType.cylinderBarClustered = "CylinderBarClustered";
        ChartType.cylinderBarStacked = "CylinderBarStacked";
        ChartType.cylinderBarStacked100 = "CylinderBarStacked100";
        ChartType.cylinderCol = "CylinderCol";
        ChartType.coneColClustered = "ConeColClustered";
        ChartType.coneColStacked = "ConeColStacked";
        ChartType.coneColStacked100 = "ConeColStacked100";
        ChartType.coneBarClustered = "ConeBarClustered";
        ChartType.coneBarStacked = "ConeBarStacked";
        ChartType.coneBarStacked100 = "ConeBarStacked100";
        ChartType.coneCol = "ConeCol";
        ChartType.pyramidColClustered = "PyramidColClustered";
        ChartType.pyramidColStacked = "PyramidColStacked";
        ChartType.pyramidColStacked100 = "PyramidColStacked100";
        ChartType.pyramidBarClustered = "PyramidBarClustered";
        ChartType.pyramidBarStacked = "PyramidBarStacked";
        ChartType.pyramidBarStacked100 = "PyramidBarStacked100";
        ChartType.pyramidCol = "PyramidCol";
        ChartType._3DColumn = "3DColumn";
        ChartType.line = "Line";
        ChartType._3DLine = "3DLine";
        ChartType._3DPie = "3DPie";
        ChartType.pie = "Pie";
        ChartType.xyscatter = "XYScatter";
        ChartType._3DArea = "3DArea";
        ChartType.area = "Area";
        ChartType.doughnut = "Doughnut";
        ChartType.radar = "Radar";
    })(ChartType = Excel.ChartType || (Excel.ChartType = {}));
    var ChartUnderlineStyle;
    (function (ChartUnderlineStyle) {
        ChartUnderlineStyle.none = "None";
        ChartUnderlineStyle.single = "Single";
    })(ChartUnderlineStyle = Excel.ChartUnderlineStyle || (Excel.ChartUnderlineStyle = {}));
    var ConditionalDataBarAxisFormat;
    (function (ConditionalDataBarAxisFormat) {
        ConditionalDataBarAxisFormat.automatic = "Automatic";
        ConditionalDataBarAxisFormat.none = "None";
        ConditionalDataBarAxisFormat.cellMidPoint = "CellMidPoint";
    })(ConditionalDataBarAxisFormat = Excel.ConditionalDataBarAxisFormat || (Excel.ConditionalDataBarAxisFormat = {}));
    var ConditionalDataBarDirection;
    (function (ConditionalDataBarDirection) {
        ConditionalDataBarDirection.context = "Context";
        ConditionalDataBarDirection.leftToRight = "LeftToRight";
        ConditionalDataBarDirection.rightToLeft = "RightToLeft";
    })(ConditionalDataBarDirection = Excel.ConditionalDataBarDirection || (Excel.ConditionalDataBarDirection = {}));
    var ConditionalFormatDirection;
    (function (ConditionalFormatDirection) {
        ConditionalFormatDirection.top = "Top";
        ConditionalFormatDirection.bottom = "Bottom";
    })(ConditionalFormatDirection = Excel.ConditionalFormatDirection || (Excel.ConditionalFormatDirection = {}));
    var ConditionalFormatType;
    (function (ConditionalFormatType) {
        ConditionalFormatType.custom = "Custom";
        ConditionalFormatType.dataBar = "DataBar";
        ConditionalFormatType.colorScale = "ColorScale";
        ConditionalFormatType.iconSet = "IconSet";
        ConditionalFormatType.topBottom = "TopBottom";
        ConditionalFormatType.presetCriteria = "PresetCriteria";
        ConditionalFormatType.containsText = "ContainsText";
        ConditionalFormatType.cellValue = "CellValue";
    })(ConditionalFormatType = Excel.ConditionalFormatType || (Excel.ConditionalFormatType = {}));
    var ConditionalFormatRuleType;
    (function (ConditionalFormatRuleType) {
        ConditionalFormatRuleType.invalid = "Invalid";
        ConditionalFormatRuleType.automatic = "Automatic";
        ConditionalFormatRuleType.lowestValue = "LowestValue";
        ConditionalFormatRuleType.highestValue = "HighestValue";
        ConditionalFormatRuleType.number = "Number";
        ConditionalFormatRuleType.percent = "Percent";
        ConditionalFormatRuleType.formula = "Formula";
        ConditionalFormatRuleType.percentile = "Percentile";
    })(ConditionalFormatRuleType = Excel.ConditionalFormatRuleType || (Excel.ConditionalFormatRuleType = {}));
    var ConditionalFormatIconRuleType;
    (function (ConditionalFormatIconRuleType) {
        ConditionalFormatIconRuleType.invalid = "Invalid";
        ConditionalFormatIconRuleType.number = "Number";
        ConditionalFormatIconRuleType.percent = "Percent";
        ConditionalFormatIconRuleType.formula = "Formula";
        ConditionalFormatIconRuleType.percentile = "Percentile";
    })(ConditionalFormatIconRuleType = Excel.ConditionalFormatIconRuleType || (Excel.ConditionalFormatIconRuleType = {}));
    var ConditionalFormatColorCriterionType;
    (function (ConditionalFormatColorCriterionType) {
        ConditionalFormatColorCriterionType.invalid = "Invalid";
        ConditionalFormatColorCriterionType.lowestValue = "LowestValue";
        ConditionalFormatColorCriterionType.highestValue = "HighestValue";
        ConditionalFormatColorCriterionType.number = "Number";
        ConditionalFormatColorCriterionType.percent = "Percent";
        ConditionalFormatColorCriterionType.formula = "Formula";
        ConditionalFormatColorCriterionType.percentile = "Percentile";
    })(ConditionalFormatColorCriterionType = Excel.ConditionalFormatColorCriterionType || (Excel.ConditionalFormatColorCriterionType = {}));
    var ConditionalTopBottomCriterionType;
    (function (ConditionalTopBottomCriterionType) {
        ConditionalTopBottomCriterionType.invalid = "Invalid";
        ConditionalTopBottomCriterionType.topItems = "TopItems";
        ConditionalTopBottomCriterionType.topPercent = "TopPercent";
        ConditionalTopBottomCriterionType.bottomItems = "BottomItems";
        ConditionalTopBottomCriterionType.bottomPercent = "BottomPercent";
    })(ConditionalTopBottomCriterionType = Excel.ConditionalTopBottomCriterionType || (Excel.ConditionalTopBottomCriterionType = {}));
    var ConditionalFormatPresetCriterion;
    (function (ConditionalFormatPresetCriterion) {
        ConditionalFormatPresetCriterion.invalid = "Invalid";
        ConditionalFormatPresetCriterion.blanks = "Blanks";
        ConditionalFormatPresetCriterion.nonBlanks = "NonBlanks";
        ConditionalFormatPresetCriterion.errors = "Errors";
        ConditionalFormatPresetCriterion.nonErrors = "NonErrors";
        ConditionalFormatPresetCriterion.yesterday = "Yesterday";
        ConditionalFormatPresetCriterion.today = "Today";
        ConditionalFormatPresetCriterion.tomorrow = "Tomorrow";
        ConditionalFormatPresetCriterion.lastSevenDays = "LastSevenDays";
        ConditionalFormatPresetCriterion.lastWeek = "LastWeek";
        ConditionalFormatPresetCriterion.thisWeek = "ThisWeek";
        ConditionalFormatPresetCriterion.nextWeek = "NextWeek";
        ConditionalFormatPresetCriterion.lastMonth = "LastMonth";
        ConditionalFormatPresetCriterion.thisMonth = "ThisMonth";
        ConditionalFormatPresetCriterion.nextMonth = "NextMonth";
        ConditionalFormatPresetCriterion.aboveAverage = "AboveAverage";
        ConditionalFormatPresetCriterion.belowAverage = "BelowAverage";
        ConditionalFormatPresetCriterion.equalOrAboveAverage = "EqualOrAboveAverage";
        ConditionalFormatPresetCriterion.equalOrBelowAverage = "EqualOrBelowAverage";
        ConditionalFormatPresetCriterion.oneStdDevAboveAverage = "OneStdDevAboveAverage";
        ConditionalFormatPresetCriterion.oneStdDevBelowAverage = "OneStdDevBelowAverage";
        ConditionalFormatPresetCriterion.twoStdDevAboveAverage = "TwoStdDevAboveAverage";
        ConditionalFormatPresetCriterion.twoStdDevBelowAverage = "TwoStdDevBelowAverage";
        ConditionalFormatPresetCriterion.threeStdDevAboveAverage = "ThreeStdDevAboveAverage";
        ConditionalFormatPresetCriterion.threeStdDevBelowAverage = "ThreeStdDevBelowAverage";
        ConditionalFormatPresetCriterion.uniqueValues = "UniqueValues";
        ConditionalFormatPresetCriterion.duplicateValues = "DuplicateValues";
    })(ConditionalFormatPresetCriterion = Excel.ConditionalFormatPresetCriterion || (Excel.ConditionalFormatPresetCriterion = {}));
    var ConditionalTextOperator;
    (function (ConditionalTextOperator) {
        ConditionalTextOperator.invalid = "Invalid";
        ConditionalTextOperator.contains = "Contains";
        ConditionalTextOperator.notContains = "NotContains";
        ConditionalTextOperator.beginsWith = "BeginsWith";
        ConditionalTextOperator.endsWith = "EndsWith";
    })(ConditionalTextOperator = Excel.ConditionalTextOperator || (Excel.ConditionalTextOperator = {}));
    var ConditionalCellValueOperator;
    (function (ConditionalCellValueOperator) {
        ConditionalCellValueOperator.invalid = "Invalid";
        ConditionalCellValueOperator.between = "Between";
        ConditionalCellValueOperator.notBetween = "NotBetween";
        ConditionalCellValueOperator.equalTo = "EqualTo";
        ConditionalCellValueOperator.notEqualTo = "NotEqualTo";
        ConditionalCellValueOperator.greaterThan = "GreaterThan";
        ConditionalCellValueOperator.lessThan = "LessThan";
        ConditionalCellValueOperator.greaterThanOrEqual = "GreaterThanOrEqual";
        ConditionalCellValueOperator.lessThanOrEqual = "LessThanOrEqual";
    })(ConditionalCellValueOperator = Excel.ConditionalCellValueOperator || (Excel.ConditionalCellValueOperator = {}));
    var ConditionalIconCriterionOperator;
    (function (ConditionalIconCriterionOperator) {
        ConditionalIconCriterionOperator.invalid = "Invalid";
        ConditionalIconCriterionOperator.greaterThan = "GreaterThan";
        ConditionalIconCriterionOperator.greaterThanOrEqual = "GreaterThanOrEqual";
    })(ConditionalIconCriterionOperator = Excel.ConditionalIconCriterionOperator || (Excel.ConditionalIconCriterionOperator = {}));
    var ConditionalRangeBorderIndex;
    (function (ConditionalRangeBorderIndex) {
        ConditionalRangeBorderIndex.edgeTop = "EdgeTop";
        ConditionalRangeBorderIndex.edgeBottom = "EdgeBottom";
        ConditionalRangeBorderIndex.edgeLeft = "EdgeLeft";
        ConditionalRangeBorderIndex.edgeRight = "EdgeRight";
    })(ConditionalRangeBorderIndex = Excel.ConditionalRangeBorderIndex || (Excel.ConditionalRangeBorderIndex = {}));
    var ConditionalRangeBorderLineStyle;
    (function (ConditionalRangeBorderLineStyle) {
        ConditionalRangeBorderLineStyle.none = "None";
        ConditionalRangeBorderLineStyle.continuous = "Continuous";
        ConditionalRangeBorderLineStyle.dash = "Dash";
        ConditionalRangeBorderLineStyle.dashDot = "DashDot";
        ConditionalRangeBorderLineStyle.dashDotDot = "DashDotDot";
        ConditionalRangeBorderLineStyle.dot = "Dot";
    })(ConditionalRangeBorderLineStyle = Excel.ConditionalRangeBorderLineStyle || (Excel.ConditionalRangeBorderLineStyle = {}));
    var ConditionalRangeFontUnderlineStyle;
    (function (ConditionalRangeFontUnderlineStyle) {
        ConditionalRangeFontUnderlineStyle.none = "None";
        ConditionalRangeFontUnderlineStyle.single = "Single";
        ConditionalRangeFontUnderlineStyle.double = "Double";
    })(ConditionalRangeFontUnderlineStyle = Excel.ConditionalRangeFontUnderlineStyle || (Excel.ConditionalRangeFontUnderlineStyle = {}));
    var CustomFunctionType;
    (function (CustomFunctionType) {
        CustomFunctionType.invalid = "Invalid";
        CustomFunctionType.script = "Script";
        CustomFunctionType.webService = "WebService";
    })(CustomFunctionType = Excel.CustomFunctionType || (Excel.CustomFunctionType = {}));
    var CustomFunctionMetadataFormat;
    (function (CustomFunctionMetadataFormat) {
        CustomFunctionMetadataFormat.invalid = "Invalid";
        CustomFunctionMetadataFormat.openApi = "OpenApi";
    })(CustomFunctionMetadataFormat = Excel.CustomFunctionMetadataFormat || (Excel.CustomFunctionMetadataFormat = {}));
    var CustomFunctionValueType;
    (function (CustomFunctionValueType) {
        CustomFunctionValueType.invalid = "Invalid";
        CustomFunctionValueType.boolean = "Boolean";
        CustomFunctionValueType.number = "Number";
        CustomFunctionValueType.string = "String";
        CustomFunctionValueType.isodate = "ISODate";
    })(CustomFunctionValueType = Excel.CustomFunctionValueType || (Excel.CustomFunctionValueType = {}));
    var CustomFunctionDimensionality;
    (function (CustomFunctionDimensionality) {
        CustomFunctionDimensionality.invalid = "Invalid";
        CustomFunctionDimensionality.scalar = "Scalar";
        CustomFunctionDimensionality.matrix = "Matrix";
    })(CustomFunctionDimensionality = Excel.CustomFunctionDimensionality || (Excel.CustomFunctionDimensionality = {}));
    var DeleteShiftDirection;
    (function (DeleteShiftDirection) {
        DeleteShiftDirection.up = "Up";
        DeleteShiftDirection.left = "Left";
    })(DeleteShiftDirection = Excel.DeleteShiftDirection || (Excel.DeleteShiftDirection = {}));
    var DynamicFilterCriteria;
    (function (DynamicFilterCriteria) {
        DynamicFilterCriteria.unknown = "Unknown";
        DynamicFilterCriteria.aboveAverage = "AboveAverage";
        DynamicFilterCriteria.allDatesInPeriodApril = "AllDatesInPeriodApril";
        DynamicFilterCriteria.allDatesInPeriodAugust = "AllDatesInPeriodAugust";
        DynamicFilterCriteria.allDatesInPeriodDecember = "AllDatesInPeriodDecember";
        DynamicFilterCriteria.allDatesInPeriodFebruray = "AllDatesInPeriodFebruray";
        DynamicFilterCriteria.allDatesInPeriodJanuary = "AllDatesInPeriodJanuary";
        DynamicFilterCriteria.allDatesInPeriodJuly = "AllDatesInPeriodJuly";
        DynamicFilterCriteria.allDatesInPeriodJune = "AllDatesInPeriodJune";
        DynamicFilterCriteria.allDatesInPeriodMarch = "AllDatesInPeriodMarch";
        DynamicFilterCriteria.allDatesInPeriodMay = "AllDatesInPeriodMay";
        DynamicFilterCriteria.allDatesInPeriodNovember = "AllDatesInPeriodNovember";
        DynamicFilterCriteria.allDatesInPeriodOctober = "AllDatesInPeriodOctober";
        DynamicFilterCriteria.allDatesInPeriodQuarter1 = "AllDatesInPeriodQuarter1";
        DynamicFilterCriteria.allDatesInPeriodQuarter2 = "AllDatesInPeriodQuarter2";
        DynamicFilterCriteria.allDatesInPeriodQuarter3 = "AllDatesInPeriodQuarter3";
        DynamicFilterCriteria.allDatesInPeriodQuarter4 = "AllDatesInPeriodQuarter4";
        DynamicFilterCriteria.allDatesInPeriodSeptember = "AllDatesInPeriodSeptember";
        DynamicFilterCriteria.belowAverage = "BelowAverage";
        DynamicFilterCriteria.lastMonth = "LastMonth";
        DynamicFilterCriteria.lastQuarter = "LastQuarter";
        DynamicFilterCriteria.lastWeek = "LastWeek";
        DynamicFilterCriteria.lastYear = "LastYear";
        DynamicFilterCriteria.nextMonth = "NextMonth";
        DynamicFilterCriteria.nextQuarter = "NextQuarter";
        DynamicFilterCriteria.nextWeek = "NextWeek";
        DynamicFilterCriteria.nextYear = "NextYear";
        DynamicFilterCriteria.thisMonth = "ThisMonth";
        DynamicFilterCriteria.thisQuarter = "ThisQuarter";
        DynamicFilterCriteria.thisWeek = "ThisWeek";
        DynamicFilterCriteria.thisYear = "ThisYear";
        DynamicFilterCriteria.today = "Today";
        DynamicFilterCriteria.tomorrow = "Tomorrow";
        DynamicFilterCriteria.yearToDate = "YearToDate";
        DynamicFilterCriteria.yesterday = "Yesterday";
    })(DynamicFilterCriteria = Excel.DynamicFilterCriteria || (Excel.DynamicFilterCriteria = {}));
    var FilterDatetimeSpecificity;
    (function (FilterDatetimeSpecificity) {
        FilterDatetimeSpecificity.year = "Year";
        FilterDatetimeSpecificity.month = "Month";
        FilterDatetimeSpecificity.day = "Day";
        FilterDatetimeSpecificity.hour = "Hour";
        FilterDatetimeSpecificity.minute = "Minute";
        FilterDatetimeSpecificity.second = "Second";
    })(FilterDatetimeSpecificity = Excel.FilterDatetimeSpecificity || (Excel.FilterDatetimeSpecificity = {}));
    var FilterOn;
    (function (FilterOn) {
        FilterOn.bottomItems = "BottomItems";
        FilterOn.bottomPercent = "BottomPercent";
        FilterOn.cellColor = "CellColor";
        FilterOn.dynamic = "Dynamic";
        FilterOn.fontColor = "FontColor";
        FilterOn.values = "Values";
        FilterOn.topItems = "TopItems";
        FilterOn.topPercent = "TopPercent";
        FilterOn.icon = "Icon";
        FilterOn.custom = "Custom";
    })(FilterOn = Excel.FilterOn || (Excel.FilterOn = {}));
    var FilterOperator;
    (function (FilterOperator) {
        FilterOperator.and = "And";
        FilterOperator.or = "Or";
    })(FilterOperator = Excel.FilterOperator || (Excel.FilterOperator = {}));
    var HorizontalAlignment;
    (function (HorizontalAlignment) {
        HorizontalAlignment.general = "General";
        HorizontalAlignment.left = "Left";
        HorizontalAlignment.center = "Center";
        HorizontalAlignment.right = "Right";
        HorizontalAlignment.fill = "Fill";
        HorizontalAlignment.justify = "Justify";
        HorizontalAlignment.centerAcrossSelection = "CenterAcrossSelection";
        HorizontalAlignment.distributed = "Distributed";
    })(HorizontalAlignment = Excel.HorizontalAlignment || (Excel.HorizontalAlignment = {}));
    var IconSet;
    (function (IconSet) {
        IconSet.invalid = "Invalid";
        IconSet.threeArrows = "ThreeArrows";
        IconSet.threeArrowsGray = "ThreeArrowsGray";
        IconSet.threeFlags = "ThreeFlags";
        IconSet.threeTrafficLights1 = "ThreeTrafficLights1";
        IconSet.threeTrafficLights2 = "ThreeTrafficLights2";
        IconSet.threeSigns = "ThreeSigns";
        IconSet.threeSymbols = "ThreeSymbols";
        IconSet.threeSymbols2 = "ThreeSymbols2";
        IconSet.fourArrows = "FourArrows";
        IconSet.fourArrowsGray = "FourArrowsGray";
        IconSet.fourRedToBlack = "FourRedToBlack";
        IconSet.fourRating = "FourRating";
        IconSet.fourTrafficLights = "FourTrafficLights";
        IconSet.fiveArrows = "FiveArrows";
        IconSet.fiveArrowsGray = "FiveArrowsGray";
        IconSet.fiveRating = "FiveRating";
        IconSet.fiveQuarters = "FiveQuarters";
        IconSet.threeStars = "ThreeStars";
        IconSet.threeTriangles = "ThreeTriangles";
        IconSet.fiveBoxes = "FiveBoxes";
    })(IconSet = Excel.IconSet || (Excel.IconSet = {}));
    var ImageFittingMode;
    (function (ImageFittingMode) {
        ImageFittingMode.fit = "Fit";
        ImageFittingMode.fitAndCenter = "FitAndCenter";
        ImageFittingMode.fill = "Fill";
    })(ImageFittingMode = Excel.ImageFittingMode || (Excel.ImageFittingMode = {}));
    var InsertShiftDirection;
    (function (InsertShiftDirection) {
        InsertShiftDirection.down = "Down";
        InsertShiftDirection.right = "Right";
    })(InsertShiftDirection = Excel.InsertShiftDirection || (Excel.InsertShiftDirection = {}));
    var NamedItemScope;
    (function (NamedItemScope) {
        NamedItemScope.worksheet = "Worksheet";
        NamedItemScope.workbook = "Workbook";
    })(NamedItemScope = Excel.NamedItemScope || (Excel.NamedItemScope = {}));
    var NamedItemType;
    (function (NamedItemType) {
        NamedItemType.string = "String";
        NamedItemType.integer = "Integer";
        NamedItemType.double = "Double";
        NamedItemType.boolean = "Boolean";
        NamedItemType.range = "Range";
        NamedItemType.error = "Error";
        NamedItemType.array = "Array";
    })(NamedItemType = Excel.NamedItemType || (Excel.NamedItemType = {}));
    var RangeUnderlineStyle;
    (function (RangeUnderlineStyle) {
        RangeUnderlineStyle.none = "None";
        RangeUnderlineStyle.single = "Single";
        RangeUnderlineStyle.double = "Double";
        RangeUnderlineStyle.singleAccountant = "SingleAccountant";
        RangeUnderlineStyle.doubleAccountant = "DoubleAccountant";
    })(RangeUnderlineStyle = Excel.RangeUnderlineStyle || (Excel.RangeUnderlineStyle = {}));
    var SheetVisibility;
    (function (SheetVisibility) {
        SheetVisibility.visible = "Visible";
        SheetVisibility.hidden = "Hidden";
        SheetVisibility.veryHidden = "VeryHidden";
    })(SheetVisibility = Excel.SheetVisibility || (Excel.SheetVisibility = {}));
    var RangeValueType;
    (function (RangeValueType) {
        RangeValueType.unknown = "Unknown";
        RangeValueType.empty = "Empty";
        RangeValueType.string = "String";
        RangeValueType.integer = "Integer";
        RangeValueType.double = "Double";
        RangeValueType.boolean = "Boolean";
        RangeValueType.error = "Error";
    })(RangeValueType = Excel.RangeValueType || (Excel.RangeValueType = {}));
    var SortOrientation;
    (function (SortOrientation) {
        SortOrientation.rows = "Rows";
        SortOrientation.columns = "Columns";
    })(SortOrientation = Excel.SortOrientation || (Excel.SortOrientation = {}));
    var SortOn;
    (function (SortOn) {
        SortOn.value = "Value";
        SortOn.cellColor = "CellColor";
        SortOn.fontColor = "FontColor";
        SortOn.icon = "Icon";
    })(SortOn = Excel.SortOn || (Excel.SortOn = {}));
    var SortDataOption;
    (function (SortDataOption) {
        SortDataOption.normal = "Normal";
        SortDataOption.textAsNumber = "TextAsNumber";
    })(SortDataOption = Excel.SortDataOption || (Excel.SortDataOption = {}));
    var SortMethod;
    (function (SortMethod) {
        SortMethod.pinYin = "PinYin";
        SortMethod.strokeCount = "StrokeCount";
    })(SortMethod = Excel.SortMethod || (Excel.SortMethod = {}));
    var VerticalAlignment;
    (function (VerticalAlignment) {
        VerticalAlignment.top = "Top";
        VerticalAlignment.center = "Center";
        VerticalAlignment.bottom = "Bottom";
        VerticalAlignment.justify = "Justify";
        VerticalAlignment.distributed = "Distributed";
    })(VerticalAlignment = Excel.VerticalAlignment || (Excel.VerticalAlignment = {}));
    var EventSource;
    (function (EventSource) {
        EventSource.local = "Local";
        EventSource.remote = "Remote";
    })(EventSource = Excel.EventSource || (Excel.EventSource = {}));
    var DataChangeType;
    (function (DataChangeType) {
        DataChangeType.others = "Others";
        DataChangeType.rangeEdited = "RangeEdited";
        DataChangeType.rowInserted = "RowInserted";
        DataChangeType.rowDeleted = "RowDeleted";
        DataChangeType.columnInserted = "ColumnInserted";
        DataChangeType.columnDeleted = "ColumnDeleted";
        DataChangeType.cellInserted = "CellInserted";
        DataChangeType.cellDeleted = "CellDeleted";
    })(DataChangeType = Excel.DataChangeType || (Excel.DataChangeType = {}));
    var EventType;
    (function (EventType) {
        EventType.worksheetDataChanged = "WorksheetDataChanged";
        EventType.worksheetSelectionChanged = "WorksheetSelectionChanged";
        EventType.worksheetAdded = "WorksheetAdded";
        EventType.worksheetActivated = "WorksheetActivated";
        EventType.worksheetDeactivated = "WorksheetDeactivated";
        EventType.tableDataChanged = "TableDataChanged";
        EventType.tableSelectionChanged = "TableSelectionChanged";
    })(EventType = Excel.EventType || (Excel.EventType = {}));
    var DocumentPropertyItem;
    (function (DocumentPropertyItem) {
        DocumentPropertyItem.title = "Title";
        DocumentPropertyItem.subject = "Subject";
        DocumentPropertyItem.author = "Author";
        DocumentPropertyItem.keywords = "Keywords";
        DocumentPropertyItem.comments = "Comments";
        DocumentPropertyItem.template = "Template";
        DocumentPropertyItem.lastAuth = "LastAuth";
        DocumentPropertyItem.revision = "Revision";
        DocumentPropertyItem.appName = "AppName";
        DocumentPropertyItem.lastPrint = "LastPrint";
        DocumentPropertyItem.creation = "Creation";
        DocumentPropertyItem.lastSave = "LastSave";
        DocumentPropertyItem.category = "Category";
        DocumentPropertyItem.format = "Format";
        DocumentPropertyItem.manager = "Manager";
        DocumentPropertyItem.company = "Company";
    })(DocumentPropertyItem = Excel.DocumentPropertyItem || (Excel.DocumentPropertyItem = {}));
    var TrendlineType;
    (function (TrendlineType) {
        TrendlineType.linear = "Linear";
        TrendlineType.exponential = "Exponential";
        TrendlineType.logarithmic = "Logarithmic";
        TrendlineType.movingAverage = "MovingAverage";
        TrendlineType.polynomial = "Polynomial";
        TrendlineType.power = "Power";
    })(TrendlineType = Excel.TrendlineType || (Excel.TrendlineType = {}));
    var SubtotalLocationType;
    (function (SubtotalLocationType) {
        SubtotalLocationType.atTop = "AtTop";
        SubtotalLocationType.atBottom = "AtBottom";
    })(SubtotalLocationType = Excel.SubtotalLocationType || (Excel.SubtotalLocationType = {}));
    var LayoutRowType;
    (function (LayoutRowType) {
        LayoutRowType.compactRow = "CompactRow";
        LayoutRowType.tabularRow = "TabularRow";
        LayoutRowType.outlineRow = "OutlineRow";
    })(LayoutRowType = Excel.LayoutRowType || (Excel.LayoutRowType = {}));
    var FunctionResult = (function (_super) {
        __extends(FunctionResult, _super);
        function FunctionResult() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(FunctionResult.prototype, "_className", {
            get: function () {
                return "FunctionResult<T>";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(FunctionResult.prototype, "error", {
            get: function () {
                _throwIfNotLoaded("error", this.m_error, "FunctionResult", this._isNull);
                return this.m_error;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(FunctionResult.prototype, "value", {
            get: function () {
                _throwIfNotLoaded("value", this.m_value, "FunctionResult", this._isNull);
                return this.m_value;
            },
            enumerable: true,
            configurable: true
        });
        FunctionResult.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Error"])) {
                this.m_error = obj["Error"];
            }
            if (!_isUndefined(obj["Value"])) {
                this.m_value = obj["Value"];
            }
        };
        FunctionResult.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        FunctionResult.prototype.toJSON = function () {
            return {
                "error": this.m_error,
                "value": this.m_value
            };
        };
        return FunctionResult;
    }(OfficeExtension.ClientObject));
    Excel.FunctionResult = FunctionResult;
    var Functions = (function (_super) {
        __extends(Functions, _super);
        function Functions() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(Functions.prototype, "_className", {
            get: function () {
                return "Functions";
            },
            enumerable: true,
            configurable: true
        });
        Functions.prototype.abs = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Abs", 0, [number], false, true, null));
        };
        Functions.prototype.accrInt = function (issue, firstInterest, settlement, rate, par, frequency, basis, calcMethod) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "AccrInt", 0, [issue, firstInterest, settlement, rate, par, frequency, basis, calcMethod], false, true, null));
        };
        Functions.prototype.accrIntM = function (issue, settlement, rate, par, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "AccrIntM", 0, [issue, settlement, rate, par, basis], false, true, null));
        };
        Functions.prototype.acos = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Acos", 0, [number], false, true, null));
        };
        Functions.prototype.acosh = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Acosh", 0, [number], false, true, null));
        };
        Functions.prototype.acot = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Acot", 0, [number], false, true, null));
        };
        Functions.prototype.acoth = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Acoth", 0, [number], false, true, null));
        };
        Functions.prototype.amorDegrc = function (cost, datePurchased, firstPeriod, salvage, period, rate, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "AmorDegrc", 0, [cost, datePurchased, firstPeriod, salvage, period, rate, basis], false, true, null));
        };
        Functions.prototype.amorLinc = function (cost, datePurchased, firstPeriod, salvage, period, rate, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "AmorLinc", 0, [cost, datePurchased, firstPeriod, salvage, period, rate, basis], false, true, null));
        };
        Functions.prototype.and = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "And", 0, [values], false, true, null));
        };
        Functions.prototype.arabic = function (text) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Arabic", 0, [text], false, true, null));
        };
        Functions.prototype.areas = function (reference) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Areas", 0, [reference], false, true, null));
        };
        Functions.prototype.asc = function (text) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Asc", 0, [text], false, true, null));
        };
        Functions.prototype.asin = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Asin", 0, [number], false, true, null));
        };
        Functions.prototype.asinh = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Asinh", 0, [number], false, true, null));
        };
        Functions.prototype.atan = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Atan", 0, [number], false, true, null));
        };
        Functions.prototype.atan2 = function (xNum, yNum) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Atan2", 0, [xNum, yNum], false, true, null));
        };
        Functions.prototype.atanh = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Atanh", 0, [number], false, true, null));
        };
        Functions.prototype.aveDev = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "AveDev", 0, [values], false, true, null));
        };
        Functions.prototype.average = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Average", 0, [values], false, true, null));
        };
        Functions.prototype.averageA = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "AverageA", 0, [values], false, true, null));
        };
        Functions.prototype.averageIf = function (range, criteria, averageRange) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "AverageIf", 0, [range, criteria, averageRange], false, true, null));
        };
        Functions.prototype.averageIfs = function (averageRange) {
            var values = [];
            for (var _i = 1; _i < arguments.length; _i++) {
                values[_i - 1] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "AverageIfs", 0, [averageRange, values], false, true, null));
        };
        Functions.prototype.bahtText = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "BahtText", 0, [number], false, true, null));
        };
        Functions.prototype.base = function (number, radix, minLength) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Base", 0, [number, radix, minLength], false, true, null));
        };
        Functions.prototype.besselI = function (x, n) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "BesselI", 0, [x, n], false, true, null));
        };
        Functions.prototype.besselJ = function (x, n) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "BesselJ", 0, [x, n], false, true, null));
        };
        Functions.prototype.besselK = function (x, n) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "BesselK", 0, [x, n], false, true, null));
        };
        Functions.prototype.besselY = function (x, n) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "BesselY", 0, [x, n], false, true, null));
        };
        Functions.prototype.beta_Dist = function (x, alpha, beta, cumulative, A, B) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Beta_Dist", 0, [x, alpha, beta, cumulative, A, B], false, true, null));
        };
        Functions.prototype.beta_Inv = function (probability, alpha, beta, A, B) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Beta_Inv", 0, [probability, alpha, beta, A, B], false, true, null));
        };
        Functions.prototype.bin2Dec = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Bin2Dec", 0, [number], false, true, null));
        };
        Functions.prototype.bin2Hex = function (number, places) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Bin2Hex", 0, [number, places], false, true, null));
        };
        Functions.prototype.bin2Oct = function (number, places) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Bin2Oct", 0, [number, places], false, true, null));
        };
        Functions.prototype.binom_Dist = function (numberS, trials, probabilityS, cumulative) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Binom_Dist", 0, [numberS, trials, probabilityS, cumulative], false, true, null));
        };
        Functions.prototype.binom_Dist_Range = function (trials, probabilityS, numberS, numberS2) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Binom_Dist_Range", 0, [trials, probabilityS, numberS, numberS2], false, true, null));
        };
        Functions.prototype.binom_Inv = function (trials, probabilityS, alpha) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Binom_Inv", 0, [trials, probabilityS, alpha], false, true, null));
        };
        Functions.prototype.bitand = function (number1, number2) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Bitand", 0, [number1, number2], false, true, null));
        };
        Functions.prototype.bitlshift = function (number, shiftAmount) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Bitlshift", 0, [number, shiftAmount], false, true, null));
        };
        Functions.prototype.bitor = function (number1, number2) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Bitor", 0, [number1, number2], false, true, null));
        };
        Functions.prototype.bitrshift = function (number, shiftAmount) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Bitrshift", 0, [number, shiftAmount], false, true, null));
        };
        Functions.prototype.bitxor = function (number1, number2) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Bitxor", 0, [number1, number2], false, true, null));
        };
        Functions.prototype.ceiling_Math = function (number, significance, mode) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Ceiling_Math", 0, [number, significance, mode], false, true, null));
        };
        Functions.prototype.ceiling_Precise = function (number, significance) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Ceiling_Precise", 0, [number, significance], false, true, null));
        };
        Functions.prototype.char = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Char", 0, [number], false, true, null));
        };
        Functions.prototype.chiSq_Dist = function (x, degFreedom, cumulative) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ChiSq_Dist", 0, [x, degFreedom, cumulative], false, true, null));
        };
        Functions.prototype.chiSq_Dist_RT = function (x, degFreedom) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ChiSq_Dist_RT", 0, [x, degFreedom], false, true, null));
        };
        Functions.prototype.chiSq_Inv = function (probability, degFreedom) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ChiSq_Inv", 0, [probability, degFreedom], false, true, null));
        };
        Functions.prototype.chiSq_Inv_RT = function (probability, degFreedom) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ChiSq_Inv_RT", 0, [probability, degFreedom], false, true, null));
        };
        Functions.prototype.choose = function (indexNum) {
            var values = [];
            for (var _i = 1; _i < arguments.length; _i++) {
                values[_i - 1] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Choose", 0, [indexNum, values], false, true, null));
        };
        Functions.prototype.clean = function (text) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Clean", 0, [text], false, true, null));
        };
        Functions.prototype.code = function (text) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Code", 0, [text], false, true, null));
        };
        Functions.prototype.columns = function (array) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Columns", 0, [array], false, true, null));
        };
        Functions.prototype.combin = function (number, numberChosen) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Combin", 0, [number, numberChosen], false, true, null));
        };
        Functions.prototype.combina = function (number, numberChosen) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Combina", 0, [number, numberChosen], false, true, null));
        };
        Functions.prototype.complex = function (realNum, iNum, suffix) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Complex", 0, [realNum, iNum, suffix], false, true, null));
        };
        Functions.prototype.concatenate = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Concatenate", 0, [values], false, true, null));
        };
        Functions.prototype.confidence_Norm = function (alpha, standardDev, size) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Confidence_Norm", 0, [alpha, standardDev, size], false, true, null));
        };
        Functions.prototype.confidence_T = function (alpha, standardDev, size) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Confidence_T", 0, [alpha, standardDev, size], false, true, null));
        };
        Functions.prototype.convert = function (number, fromUnit, toUnit) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Convert", 0, [number, fromUnit, toUnit], false, true, null));
        };
        Functions.prototype.cos = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Cos", 0, [number], false, true, null));
        };
        Functions.prototype.cosh = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Cosh", 0, [number], false, true, null));
        };
        Functions.prototype.cot = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Cot", 0, [number], false, true, null));
        };
        Functions.prototype.coth = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Coth", 0, [number], false, true, null));
        };
        Functions.prototype.count = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Count", 0, [values], false, true, null));
        };
        Functions.prototype.countA = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "CountA", 0, [values], false, true, null));
        };
        Functions.prototype.countBlank = function (range) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "CountBlank", 0, [range], false, true, null));
        };
        Functions.prototype.countIf = function (range, criteria) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "CountIf", 0, [range, criteria], false, true, null));
        };
        Functions.prototype.countIfs = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "CountIfs", 0, [values], false, true, null));
        };
        Functions.prototype.coupDayBs = function (settlement, maturity, frequency, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "CoupDayBs", 0, [settlement, maturity, frequency, basis], false, true, null));
        };
        Functions.prototype.coupDays = function (settlement, maturity, frequency, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "CoupDays", 0, [settlement, maturity, frequency, basis], false, true, null));
        };
        Functions.prototype.coupDaysNc = function (settlement, maturity, frequency, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "CoupDaysNc", 0, [settlement, maturity, frequency, basis], false, true, null));
        };
        Functions.prototype.coupNcd = function (settlement, maturity, frequency, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "CoupNcd", 0, [settlement, maturity, frequency, basis], false, true, null));
        };
        Functions.prototype.coupNum = function (settlement, maturity, frequency, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "CoupNum", 0, [settlement, maturity, frequency, basis], false, true, null));
        };
        Functions.prototype.coupPcd = function (settlement, maturity, frequency, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "CoupPcd", 0, [settlement, maturity, frequency, basis], false, true, null));
        };
        Functions.prototype.csc = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Csc", 0, [number], false, true, null));
        };
        Functions.prototype.csch = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Csch", 0, [number], false, true, null));
        };
        Functions.prototype.cumIPmt = function (rate, nper, pv, startPeriod, endPeriod, type) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "CumIPmt", 0, [rate, nper, pv, startPeriod, endPeriod, type], false, true, null));
        };
        Functions.prototype.cumPrinc = function (rate, nper, pv, startPeriod, endPeriod, type) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "CumPrinc", 0, [rate, nper, pv, startPeriod, endPeriod, type], false, true, null));
        };
        Functions.prototype.daverage = function (database, field, criteria) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DAverage", 0, [database, field, criteria], false, true, null));
        };
        Functions.prototype.dcount = function (database, field, criteria) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DCount", 0, [database, field, criteria], false, true, null));
        };
        Functions.prototype.dcountA = function (database, field, criteria) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DCountA", 0, [database, field, criteria], false, true, null));
        };
        Functions.prototype.dget = function (database, field, criteria) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DGet", 0, [database, field, criteria], false, true, null));
        };
        Functions.prototype.dmax = function (database, field, criteria) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DMax", 0, [database, field, criteria], false, true, null));
        };
        Functions.prototype.dmin = function (database, field, criteria) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DMin", 0, [database, field, criteria], false, true, null));
        };
        Functions.prototype.dproduct = function (database, field, criteria) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DProduct", 0, [database, field, criteria], false, true, null));
        };
        Functions.prototype.dstDev = function (database, field, criteria) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DStDev", 0, [database, field, criteria], false, true, null));
        };
        Functions.prototype.dstDevP = function (database, field, criteria) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DStDevP", 0, [database, field, criteria], false, true, null));
        };
        Functions.prototype.dsum = function (database, field, criteria) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DSum", 0, [database, field, criteria], false, true, null));
        };
        Functions.prototype.dvar = function (database, field, criteria) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DVar", 0, [database, field, criteria], false, true, null));
        };
        Functions.prototype.dvarP = function (database, field, criteria) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DVarP", 0, [database, field, criteria], false, true, null));
        };
        Functions.prototype.date = function (year, month, day) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Date", 0, [year, month, day], false, true, null));
        };
        Functions.prototype.datevalue = function (dateText) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Datevalue", 0, [dateText], false, true, null));
        };
        Functions.prototype.day = function (serialNumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Day", 0, [serialNumber], false, true, null));
        };
        Functions.prototype.days = function (endDate, startDate) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Days", 0, [endDate, startDate], false, true, null));
        };
        Functions.prototype.days360 = function (startDate, endDate, method) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Days360", 0, [startDate, endDate, method], false, true, null));
        };
        Functions.prototype.db = function (cost, salvage, life, period, month) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Db", 0, [cost, salvage, life, period, month], false, true, null));
        };
        Functions.prototype.dbcs = function (text) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Dbcs", 0, [text], false, true, null));
        };
        Functions.prototype.ddb = function (cost, salvage, life, period, factor) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Ddb", 0, [cost, salvage, life, period, factor], false, true, null));
        };
        Functions.prototype.dec2Bin = function (number, places) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Dec2Bin", 0, [number, places], false, true, null));
        };
        Functions.prototype.dec2Hex = function (number, places) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Dec2Hex", 0, [number, places], false, true, null));
        };
        Functions.prototype.dec2Oct = function (number, places) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Dec2Oct", 0, [number, places], false, true, null));
        };
        Functions.prototype.decimal = function (number, radix) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Decimal", 0, [number, radix], false, true, null));
        };
        Functions.prototype.degrees = function (angle) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Degrees", 0, [angle], false, true, null));
        };
        Functions.prototype.delta = function (number1, number2) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Delta", 0, [number1, number2], false, true, null));
        };
        Functions.prototype.devSq = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DevSq", 0, [values], false, true, null));
        };
        Functions.prototype.disc = function (settlement, maturity, pr, redemption, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Disc", 0, [settlement, maturity, pr, redemption, basis], false, true, null));
        };
        Functions.prototype.dollar = function (number, decimals) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Dollar", 0, [number, decimals], false, true, null));
        };
        Functions.prototype.dollarDe = function (fractionalDollar, fraction) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DollarDe", 0, [fractionalDollar, fraction], false, true, null));
        };
        Functions.prototype.dollarFr = function (decimalDollar, fraction) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DollarFr", 0, [decimalDollar, fraction], false, true, null));
        };
        Functions.prototype.duration = function (settlement, maturity, coupon, yld, frequency, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Duration", 0, [settlement, maturity, coupon, yld, frequency, basis], false, true, null));
        };
        Functions.prototype.ecma_Ceiling = function (number, significance) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ECMA_Ceiling", 0, [number, significance], false, true, null));
        };
        Functions.prototype.edate = function (startDate, months) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "EDate", 0, [startDate, months], false, true, null));
        };
        Functions.prototype.effect = function (nominalRate, npery) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Effect", 0, [nominalRate, npery], false, true, null));
        };
        Functions.prototype.eoMonth = function (startDate, months) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "EoMonth", 0, [startDate, months], false, true, null));
        };
        Functions.prototype.erf = function (lowerLimit, upperLimit) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Erf", 0, [lowerLimit, upperLimit], false, true, null));
        };
        Functions.prototype.erfC = function (x) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ErfC", 0, [x], false, true, null));
        };
        Functions.prototype.erfC_Precise = function (X) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ErfC_Precise", 0, [X], false, true, null));
        };
        Functions.prototype.erf_Precise = function (X) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Erf_Precise", 0, [X], false, true, null));
        };
        Functions.prototype.error_Type = function (errorVal) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Error_Type", 0, [errorVal], false, true, null));
        };
        Functions.prototype.even = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Even", 0, [number], false, true, null));
        };
        Functions.prototype.exact = function (text1, text2) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Exact", 0, [text1, text2], false, true, null));
        };
        Functions.prototype.exp = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Exp", 0, [number], false, true, null));
        };
        Functions.prototype.expon_Dist = function (x, lambda, cumulative) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Expon_Dist", 0, [x, lambda, cumulative], false, true, null));
        };
        Functions.prototype.fvschedule = function (principal, schedule) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "FVSchedule", 0, [principal, schedule], false, true, null));
        };
        Functions.prototype.f_Dist = function (x, degFreedom1, degFreedom2, cumulative) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "F_Dist", 0, [x, degFreedom1, degFreedom2, cumulative], false, true, null));
        };
        Functions.prototype.f_Dist_RT = function (x, degFreedom1, degFreedom2) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "F_Dist_RT", 0, [x, degFreedom1, degFreedom2], false, true, null));
        };
        Functions.prototype.f_Inv = function (probability, degFreedom1, degFreedom2) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "F_Inv", 0, [probability, degFreedom1, degFreedom2], false, true, null));
        };
        Functions.prototype.f_Inv_RT = function (probability, degFreedom1, degFreedom2) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "F_Inv_RT", 0, [probability, degFreedom1, degFreedom2], false, true, null));
        };
        Functions.prototype.fact = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Fact", 0, [number], false, true, null));
        };
        Functions.prototype.factDouble = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "FactDouble", 0, [number], false, true, null));
        };
        Functions.prototype.false = function () {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "False", 0, [], false, true, null));
        };
        Functions.prototype.find = function (findText, withinText, startNum) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Find", 0, [findText, withinText, startNum], false, true, null));
        };
        Functions.prototype.findB = function (findText, withinText, startNum) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "FindB", 0, [findText, withinText, startNum], false, true, null));
        };
        Functions.prototype.fisher = function (x) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Fisher", 0, [x], false, true, null));
        };
        Functions.prototype.fisherInv = function (y) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "FisherInv", 0, [y], false, true, null));
        };
        Functions.prototype.fixed = function (number, decimals, noCommas) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Fixed", 0, [number, decimals, noCommas], false, true, null));
        };
        Functions.prototype.floor_Math = function (number, significance, mode) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Floor_Math", 0, [number, significance, mode], false, true, null));
        };
        Functions.prototype.floor_Precise = function (number, significance) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Floor_Precise", 0, [number, significance], false, true, null));
        };
        Functions.prototype.fv = function (rate, nper, pmt, pv, type) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Fv", 0, [rate, nper, pmt, pv, type], false, true, null));
        };
        Functions.prototype.gamma = function (x) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Gamma", 0, [x], false, true, null));
        };
        Functions.prototype.gammaLn = function (x) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "GammaLn", 0, [x], false, true, null));
        };
        Functions.prototype.gammaLn_Precise = function (x) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "GammaLn_Precise", 0, [x], false, true, null));
        };
        Functions.prototype.gamma_Dist = function (x, alpha, beta, cumulative) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Gamma_Dist", 0, [x, alpha, beta, cumulative], false, true, null));
        };
        Functions.prototype.gamma_Inv = function (probability, alpha, beta) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Gamma_Inv", 0, [probability, alpha, beta], false, true, null));
        };
        Functions.prototype.gauss = function (x) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Gauss", 0, [x], false, true, null));
        };
        Functions.prototype.gcd = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Gcd", 0, [values], false, true, null));
        };
        Functions.prototype.geStep = function (number, step) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "GeStep", 0, [number, step], false, true, null));
        };
        Functions.prototype.geoMean = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "GeoMean", 0, [values], false, true, null));
        };
        Functions.prototype.hlookup = function (lookupValue, tableArray, rowIndexNum, rangeLookup) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "HLookup", 0, [lookupValue, tableArray, rowIndexNum, rangeLookup], false, true, null));
        };
        Functions.prototype.harMean = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "HarMean", 0, [values], false, true, null));
        };
        Functions.prototype.hex2Bin = function (number, places) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Hex2Bin", 0, [number, places], false, true, null));
        };
        Functions.prototype.hex2Dec = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Hex2Dec", 0, [number], false, true, null));
        };
        Functions.prototype.hex2Oct = function (number, places) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Hex2Oct", 0, [number, places], false, true, null));
        };
        Functions.prototype.hour = function (serialNumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Hour", 0, [serialNumber], false, true, null));
        };
        Functions.prototype.hypGeom_Dist = function (sampleS, numberSample, populationS, numberPop, cumulative) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "HypGeom_Dist", 0, [sampleS, numberSample, populationS, numberPop, cumulative], false, true, null));
        };
        Functions.prototype.hyperlink = function (linkLocation, friendlyName) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Hyperlink", 0, [linkLocation, friendlyName], false, true, null));
        };
        Functions.prototype.iso_Ceiling = function (number, significance) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ISO_Ceiling", 0, [number, significance], false, true, null));
        };
        Functions.prototype.if = function (logicalTest, valueIfTrue, valueIfFalse) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "If", 0, [logicalTest, valueIfTrue, valueIfFalse], false, true, null));
        };
        Functions.prototype.imAbs = function (inumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImAbs", 0, [inumber], false, true, null));
        };
        Functions.prototype.imArgument = function (inumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImArgument", 0, [inumber], false, true, null));
        };
        Functions.prototype.imConjugate = function (inumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImConjugate", 0, [inumber], false, true, null));
        };
        Functions.prototype.imCos = function (inumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImCos", 0, [inumber], false, true, null));
        };
        Functions.prototype.imCosh = function (inumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImCosh", 0, [inumber], false, true, null));
        };
        Functions.prototype.imCot = function (inumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImCot", 0, [inumber], false, true, null));
        };
        Functions.prototype.imCsc = function (inumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImCsc", 0, [inumber], false, true, null));
        };
        Functions.prototype.imCsch = function (inumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImCsch", 0, [inumber], false, true, null));
        };
        Functions.prototype.imDiv = function (inumber1, inumber2) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImDiv", 0, [inumber1, inumber2], false, true, null));
        };
        Functions.prototype.imExp = function (inumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImExp", 0, [inumber], false, true, null));
        };
        Functions.prototype.imLn = function (inumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImLn", 0, [inumber], false, true, null));
        };
        Functions.prototype.imLog10 = function (inumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImLog10", 0, [inumber], false, true, null));
        };
        Functions.prototype.imLog2 = function (inumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImLog2", 0, [inumber], false, true, null));
        };
        Functions.prototype.imPower = function (inumber, number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImPower", 0, [inumber, number], false, true, null));
        };
        Functions.prototype.imProduct = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImProduct", 0, [values], false, true, null));
        };
        Functions.prototype.imReal = function (inumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImReal", 0, [inumber], false, true, null));
        };
        Functions.prototype.imSec = function (inumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImSec", 0, [inumber], false, true, null));
        };
        Functions.prototype.imSech = function (inumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImSech", 0, [inumber], false, true, null));
        };
        Functions.prototype.imSin = function (inumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImSin", 0, [inumber], false, true, null));
        };
        Functions.prototype.imSinh = function (inumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImSinh", 0, [inumber], false, true, null));
        };
        Functions.prototype.imSqrt = function (inumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImSqrt", 0, [inumber], false, true, null));
        };
        Functions.prototype.imSub = function (inumber1, inumber2) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImSub", 0, [inumber1, inumber2], false, true, null));
        };
        Functions.prototype.imSum = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImSum", 0, [values], false, true, null));
        };
        Functions.prototype.imTan = function (inumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImTan", 0, [inumber], false, true, null));
        };
        Functions.prototype.imaginary = function (inumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Imaginary", 0, [inumber], false, true, null));
        };
        Functions.prototype.int = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Int", 0, [number], false, true, null));
        };
        Functions.prototype.intRate = function (settlement, maturity, investment, redemption, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "IntRate", 0, [settlement, maturity, investment, redemption, basis], false, true, null));
        };
        Functions.prototype.ipmt = function (rate, per, nper, pv, fv, type) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Ipmt", 0, [rate, per, nper, pv, fv, type], false, true, null));
        };
        Functions.prototype.irr = function (values, guess) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Irr", 0, [values, guess], false, true, null));
        };
        Functions.prototype.isErr = function (value) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "IsErr", 0, [value], false, true, null));
        };
        Functions.prototype.isError = function (value) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "IsError", 0, [value], false, true, null));
        };
        Functions.prototype.isEven = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "IsEven", 0, [number], false, true, null));
        };
        Functions.prototype.isFormula = function (reference) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "IsFormula", 0, [reference], false, true, null));
        };
        Functions.prototype.isLogical = function (value) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "IsLogical", 0, [value], false, true, null));
        };
        Functions.prototype.isNA = function (value) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "IsNA", 0, [value], false, true, null));
        };
        Functions.prototype.isNonText = function (value) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "IsNonText", 0, [value], false, true, null));
        };
        Functions.prototype.isNumber = function (value) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "IsNumber", 0, [value], false, true, null));
        };
        Functions.prototype.isOdd = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "IsOdd", 0, [number], false, true, null));
        };
        Functions.prototype.isText = function (value) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "IsText", 0, [value], false, true, null));
        };
        Functions.prototype.isoWeekNum = function (date) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "IsoWeekNum", 0, [date], false, true, null));
        };
        Functions.prototype.ispmt = function (rate, per, nper, pv) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Ispmt", 0, [rate, per, nper, pv], false, true, null));
        };
        Functions.prototype.isref = function (value) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Isref", 0, [value], false, true, null));
        };
        Functions.prototype.kurt = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Kurt", 0, [values], false, true, null));
        };
        Functions.prototype.large = function (array, k) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Large", 0, [array, k], false, true, null));
        };
        Functions.prototype.lcm = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Lcm", 0, [values], false, true, null));
        };
        Functions.prototype.left = function (text, numChars) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Left", 0, [text, numChars], false, true, null));
        };
        Functions.prototype.leftb = function (text, numBytes) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Leftb", 0, [text, numBytes], false, true, null));
        };
        Functions.prototype.len = function (text) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Len", 0, [text], false, true, null));
        };
        Functions.prototype.lenb = function (text) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Lenb", 0, [text], false, true, null));
        };
        Functions.prototype.ln = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Ln", 0, [number], false, true, null));
        };
        Functions.prototype.log = function (number, base) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Log", 0, [number, base], false, true, null));
        };
        Functions.prototype.log10 = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Log10", 0, [number], false, true, null));
        };
        Functions.prototype.logNorm_Dist = function (x, mean, standardDev, cumulative) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "LogNorm_Dist", 0, [x, mean, standardDev, cumulative], false, true, null));
        };
        Functions.prototype.logNorm_Inv = function (probability, mean, standardDev) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "LogNorm_Inv", 0, [probability, mean, standardDev], false, true, null));
        };
        Functions.prototype.lookup = function (lookupValue, lookupVector, resultVector) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Lookup", 0, [lookupValue, lookupVector, resultVector], false, true, null));
        };
        Functions.prototype.lower = function (text) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Lower", 0, [text], false, true, null));
        };
        Functions.prototype.mduration = function (settlement, maturity, coupon, yld, frequency, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "MDuration", 0, [settlement, maturity, coupon, yld, frequency, basis], false, true, null));
        };
        Functions.prototype.mirr = function (values, financeRate, reinvestRate) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "MIrr", 0, [values, financeRate, reinvestRate], false, true, null));
        };
        Functions.prototype.mround = function (number, multiple) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "MRound", 0, [number, multiple], false, true, null));
        };
        Functions.prototype.match = function (lookupValue, lookupArray, matchType) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Match", 0, [lookupValue, lookupArray, matchType], false, true, null));
        };
        Functions.prototype.max = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Max", 0, [values], false, true, null));
        };
        Functions.prototype.maxA = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "MaxA", 0, [values], false, true, null));
        };
        Functions.prototype.median = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Median", 0, [values], false, true, null));
        };
        Functions.prototype.mid = function (text, startNum, numChars) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Mid", 0, [text, startNum, numChars], false, true, null));
        };
        Functions.prototype.midb = function (text, startNum, numBytes) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Midb", 0, [text, startNum, numBytes], false, true, null));
        };
        Functions.prototype.min = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Min", 0, [values], false, true, null));
        };
        Functions.prototype.minA = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "MinA", 0, [values], false, true, null));
        };
        Functions.prototype.minute = function (serialNumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Minute", 0, [serialNumber], false, true, null));
        };
        Functions.prototype.mod = function (number, divisor) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Mod", 0, [number, divisor], false, true, null));
        };
        Functions.prototype.month = function (serialNumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Month", 0, [serialNumber], false, true, null));
        };
        Functions.prototype.multiNomial = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "MultiNomial", 0, [values], false, true, null));
        };
        Functions.prototype.n = function (value) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "N", 0, [value], false, true, null));
        };
        Functions.prototype.nper = function (rate, pmt, pv, fv, type) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "NPer", 0, [rate, pmt, pv, fv, type], false, true, null));
        };
        Functions.prototype.na = function () {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Na", 0, [], false, true, null));
        };
        Functions.prototype.negBinom_Dist = function (numberF, numberS, probabilityS, cumulative) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "NegBinom_Dist", 0, [numberF, numberS, probabilityS, cumulative], false, true, null));
        };
        Functions.prototype.networkDays = function (startDate, endDate, holidays) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "NetworkDays", 0, [startDate, endDate, holidays], false, true, null));
        };
        Functions.prototype.networkDays_Intl = function (startDate, endDate, weekend, holidays) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "NetworkDays_Intl", 0, [startDate, endDate, weekend, holidays], false, true, null));
        };
        Functions.prototype.nominal = function (effectRate, npery) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Nominal", 0, [effectRate, npery], false, true, null));
        };
        Functions.prototype.norm_Dist = function (x, mean, standardDev, cumulative) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Norm_Dist", 0, [x, mean, standardDev, cumulative], false, true, null));
        };
        Functions.prototype.norm_Inv = function (probability, mean, standardDev) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Norm_Inv", 0, [probability, mean, standardDev], false, true, null));
        };
        Functions.prototype.norm_S_Dist = function (z, cumulative) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Norm_S_Dist", 0, [z, cumulative], false, true, null));
        };
        Functions.prototype.norm_S_Inv = function (probability) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Norm_S_Inv", 0, [probability], false, true, null));
        };
        Functions.prototype.not = function (logical) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Not", 0, [logical], false, true, null));
        };
        Functions.prototype.now = function () {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Now", 0, [], false, true, null));
        };
        Functions.prototype.npv = function (rate) {
            var values = [];
            for (var _i = 1; _i < arguments.length; _i++) {
                values[_i - 1] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Npv", 0, [rate, values], false, true, null));
        };
        Functions.prototype.numberValue = function (text, decimalSeparator, groupSeparator) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "NumberValue", 0, [text, decimalSeparator, groupSeparator], false, true, null));
        };
        Functions.prototype.oct2Bin = function (number, places) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Oct2Bin", 0, [number, places], false, true, null));
        };
        Functions.prototype.oct2Dec = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Oct2Dec", 0, [number], false, true, null));
        };
        Functions.prototype.oct2Hex = function (number, places) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Oct2Hex", 0, [number, places], false, true, null));
        };
        Functions.prototype.odd = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Odd", 0, [number], false, true, null));
        };
        Functions.prototype.oddFPrice = function (settlement, maturity, issue, firstCoupon, rate, yld, redemption, frequency, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "OddFPrice", 0, [settlement, maturity, issue, firstCoupon, rate, yld, redemption, frequency, basis], false, true, null));
        };
        Functions.prototype.oddFYield = function (settlement, maturity, issue, firstCoupon, rate, pr, redemption, frequency, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "OddFYield", 0, [settlement, maturity, issue, firstCoupon, rate, pr, redemption, frequency, basis], false, true, null));
        };
        Functions.prototype.oddLPrice = function (settlement, maturity, lastInterest, rate, yld, redemption, frequency, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "OddLPrice", 0, [settlement, maturity, lastInterest, rate, yld, redemption, frequency, basis], false, true, null));
        };
        Functions.prototype.oddLYield = function (settlement, maturity, lastInterest, rate, pr, redemption, frequency, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "OddLYield", 0, [settlement, maturity, lastInterest, rate, pr, redemption, frequency, basis], false, true, null));
        };
        Functions.prototype.or = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Or", 0, [values], false, true, null));
        };
        Functions.prototype.pduration = function (rate, pv, fv) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "PDuration", 0, [rate, pv, fv], false, true, null));
        };
        Functions.prototype.percentRank_Exc = function (array, x, significance) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "PercentRank_Exc", 0, [array, x, significance], false, true, null));
        };
        Functions.prototype.percentRank_Inc = function (array, x, significance) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "PercentRank_Inc", 0, [array, x, significance], false, true, null));
        };
        Functions.prototype.percentile_Exc = function (array, k) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Percentile_Exc", 0, [array, k], false, true, null));
        };
        Functions.prototype.percentile_Inc = function (array, k) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Percentile_Inc", 0, [array, k], false, true, null));
        };
        Functions.prototype.permut = function (number, numberChosen) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Permut", 0, [number, numberChosen], false, true, null));
        };
        Functions.prototype.permutationa = function (number, numberChosen) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Permutationa", 0, [number, numberChosen], false, true, null));
        };
        Functions.prototype.phi = function (x) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Phi", 0, [x], false, true, null));
        };
        Functions.prototype.pi = function () {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Pi", 0, [], false, true, null));
        };
        Functions.prototype.pmt = function (rate, nper, pv, fv, type) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Pmt", 0, [rate, nper, pv, fv, type], false, true, null));
        };
        Functions.prototype.poisson_Dist = function (x, mean, cumulative) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Poisson_Dist", 0, [x, mean, cumulative], false, true, null));
        };
        Functions.prototype.power = function (number, power) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Power", 0, [number, power], false, true, null));
        };
        Functions.prototype.ppmt = function (rate, per, nper, pv, fv, type) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Ppmt", 0, [rate, per, nper, pv, fv, type], false, true, null));
        };
        Functions.prototype.price = function (settlement, maturity, rate, yld, redemption, frequency, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Price", 0, [settlement, maturity, rate, yld, redemption, frequency, basis], false, true, null));
        };
        Functions.prototype.priceDisc = function (settlement, maturity, discount, redemption, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "PriceDisc", 0, [settlement, maturity, discount, redemption, basis], false, true, null));
        };
        Functions.prototype.priceMat = function (settlement, maturity, issue, rate, yld, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "PriceMat", 0, [settlement, maturity, issue, rate, yld, basis], false, true, null));
        };
        Functions.prototype.product = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Product", 0, [values], false, true, null));
        };
        Functions.prototype.proper = function (text) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Proper", 0, [text], false, true, null));
        };
        Functions.prototype.pv = function (rate, nper, pmt, fv, type) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Pv", 0, [rate, nper, pmt, fv, type], false, true, null));
        };
        Functions.prototype.quartile_Exc = function (array, quart) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Quartile_Exc", 0, [array, quart], false, true, null));
        };
        Functions.prototype.quartile_Inc = function (array, quart) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Quartile_Inc", 0, [array, quart], false, true, null));
        };
        Functions.prototype.quotient = function (numerator, denominator) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Quotient", 0, [numerator, denominator], false, true, null));
        };
        Functions.prototype.radians = function (angle) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Radians", 0, [angle], false, true, null));
        };
        Functions.prototype.rand = function () {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Rand", 0, [], false, true, null));
        };
        Functions.prototype.randBetween = function (bottom, top) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "RandBetween", 0, [bottom, top], false, true, null));
        };
        Functions.prototype.rank_Avg = function (number, ref, order) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Rank_Avg", 0, [number, ref, order], false, true, null));
        };
        Functions.prototype.rank_Eq = function (number, ref, order) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Rank_Eq", 0, [number, ref, order], false, true, null));
        };
        Functions.prototype.rate = function (nper, pmt, pv, fv, type, guess) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Rate", 0, [nper, pmt, pv, fv, type, guess], false, true, null));
        };
        Functions.prototype.received = function (settlement, maturity, investment, discount, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Received", 0, [settlement, maturity, investment, discount, basis], false, true, null));
        };
        Functions.prototype.replace = function (oldText, startNum, numChars, newText) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Replace", 0, [oldText, startNum, numChars, newText], false, true, null));
        };
        Functions.prototype.replaceB = function (oldText, startNum, numBytes, newText) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ReplaceB", 0, [oldText, startNum, numBytes, newText], false, true, null));
        };
        Functions.prototype.rept = function (text, numberTimes) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Rept", 0, [text, numberTimes], false, true, null));
        };
        Functions.prototype.right = function (text, numChars) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Right", 0, [text, numChars], false, true, null));
        };
        Functions.prototype.rightb = function (text, numBytes) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Rightb", 0, [text, numBytes], false, true, null));
        };
        Functions.prototype.roman = function (number, form) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Roman", 0, [number, form], false, true, null));
        };
        Functions.prototype.round = function (number, numDigits) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Round", 0, [number, numDigits], false, true, null));
        };
        Functions.prototype.roundDown = function (number, numDigits) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "RoundDown", 0, [number, numDigits], false, true, null));
        };
        Functions.prototype.roundUp = function (number, numDigits) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "RoundUp", 0, [number, numDigits], false, true, null));
        };
        Functions.prototype.rows = function (array) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Rows", 0, [array], false, true, null));
        };
        Functions.prototype.rri = function (nper, pv, fv) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Rri", 0, [nper, pv, fv], false, true, null));
        };
        Functions.prototype.sec = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Sec", 0, [number], false, true, null));
        };
        Functions.prototype.sech = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Sech", 0, [number], false, true, null));
        };
        Functions.prototype.second = function (serialNumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Second", 0, [serialNumber], false, true, null));
        };
        Functions.prototype.seriesSum = function (x, n, m, coefficients) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "SeriesSum", 0, [x, n, m, coefficients], false, true, null));
        };
        Functions.prototype.sheet = function (value) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Sheet", 0, [value], false, true, null));
        };
        Functions.prototype.sheets = function (reference) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Sheets", 0, [reference], false, true, null));
        };
        Functions.prototype.sign = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Sign", 0, [number], false, true, null));
        };
        Functions.prototype.sin = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Sin", 0, [number], false, true, null));
        };
        Functions.prototype.sinh = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Sinh", 0, [number], false, true, null));
        };
        Functions.prototype.skew = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Skew", 0, [values], false, true, null));
        };
        Functions.prototype.skew_p = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Skew_p", 0, [values], false, true, null));
        };
        Functions.prototype.sln = function (cost, salvage, life) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Sln", 0, [cost, salvage, life], false, true, null));
        };
        Functions.prototype.small = function (array, k) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Small", 0, [array, k], false, true, null));
        };
        Functions.prototype.sqrt = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Sqrt", 0, [number], false, true, null));
        };
        Functions.prototype.sqrtPi = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "SqrtPi", 0, [number], false, true, null));
        };
        Functions.prototype.stDevA = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "StDevA", 0, [values], false, true, null));
        };
        Functions.prototype.stDevPA = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "StDevPA", 0, [values], false, true, null));
        };
        Functions.prototype.stDev_P = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "StDev_P", 0, [values], false, true, null));
        };
        Functions.prototype.stDev_S = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "StDev_S", 0, [values], false, true, null));
        };
        Functions.prototype.standardize = function (x, mean, standardDev) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Standardize", 0, [x, mean, standardDev], false, true, null));
        };
        Functions.prototype.substitute = function (text, oldText, newText, instanceNum) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Substitute", 0, [text, oldText, newText, instanceNum], false, true, null));
        };
        Functions.prototype.subtotal = function (functionNum) {
            var values = [];
            for (var _i = 1; _i < arguments.length; _i++) {
                values[_i - 1] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Subtotal", 0, [functionNum, values], false, true, null));
        };
        Functions.prototype.sum = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Sum", 0, [values], false, true, null));
        };
        Functions.prototype.sumIf = function (range, criteria, sumRange) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "SumIf", 0, [range, criteria, sumRange], false, true, null));
        };
        Functions.prototype.sumIfs = function (sumRange) {
            var values = [];
            for (var _i = 1; _i < arguments.length; _i++) {
                values[_i - 1] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "SumIfs", 0, [sumRange, values], false, true, null));
        };
        Functions.prototype.sumSq = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "SumSq", 0, [values], false, true, null));
        };
        Functions.prototype.syd = function (cost, salvage, life, per) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Syd", 0, [cost, salvage, life, per], false, true, null));
        };
        Functions.prototype.t = function (value) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "T", 0, [value], false, true, null));
        };
        Functions.prototype.tbillEq = function (settlement, maturity, discount) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "TBillEq", 0, [settlement, maturity, discount], false, true, null));
        };
        Functions.prototype.tbillPrice = function (settlement, maturity, discount) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "TBillPrice", 0, [settlement, maturity, discount], false, true, null));
        };
        Functions.prototype.tbillYield = function (settlement, maturity, pr) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "TBillYield", 0, [settlement, maturity, pr], false, true, null));
        };
        Functions.prototype.t_Dist = function (x, degFreedom, cumulative) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "T_Dist", 0, [x, degFreedom, cumulative], false, true, null));
        };
        Functions.prototype.t_Dist_2T = function (x, degFreedom) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "T_Dist_2T", 0, [x, degFreedom], false, true, null));
        };
        Functions.prototype.t_Dist_RT = function (x, degFreedom) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "T_Dist_RT", 0, [x, degFreedom], false, true, null));
        };
        Functions.prototype.t_Inv = function (probability, degFreedom) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "T_Inv", 0, [probability, degFreedom], false, true, null));
        };
        Functions.prototype.t_Inv_2T = function (probability, degFreedom) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "T_Inv_2T", 0, [probability, degFreedom], false, true, null));
        };
        Functions.prototype.tan = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Tan", 0, [number], false, true, null));
        };
        Functions.prototype.tanh = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Tanh", 0, [number], false, true, null));
        };
        Functions.prototype.text = function (value, formatText) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Text", 0, [value, formatText], false, true, null));
        };
        Functions.prototype.time = function (hour, minute, second) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Time", 0, [hour, minute, second], false, true, null));
        };
        Functions.prototype.timevalue = function (timeText) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Timevalue", 0, [timeText], false, true, null));
        };
        Functions.prototype.today = function () {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Today", 0, [], false, true, null));
        };
        Functions.prototype.trim = function (text) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Trim", 0, [text], false, true, null));
        };
        Functions.prototype.trimMean = function (array, percent) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "TrimMean", 0, [array, percent], false, true, null));
        };
        Functions.prototype.true = function () {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "True", 0, [], false, true, null));
        };
        Functions.prototype.trunc = function (number, numDigits) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Trunc", 0, [number, numDigits], false, true, null));
        };
        Functions.prototype.type = function (value) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Type", 0, [value], false, true, null));
        };
        Functions.prototype.usdollar = function (number, decimals) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "USDollar", 0, [number, decimals], false, true, null));
        };
        Functions.prototype.unichar = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Unichar", 0, [number], false, true, null));
        };
        Functions.prototype.unicode = function (text) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Unicode", 0, [text], false, true, null));
        };
        Functions.prototype.upper = function (text) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Upper", 0, [text], false, true, null));
        };
        Functions.prototype.vlookup = function (lookupValue, tableArray, colIndexNum, rangeLookup) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "VLookup", 0, [lookupValue, tableArray, colIndexNum, rangeLookup], false, true, null));
        };
        Functions.prototype.value = function (text) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Value", 0, [text], false, true, null));
        };
        Functions.prototype.varA = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "VarA", 0, [values], false, true, null));
        };
        Functions.prototype.varPA = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "VarPA", 0, [values], false, true, null));
        };
        Functions.prototype.var_P = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Var_P", 0, [values], false, true, null));
        };
        Functions.prototype.var_S = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Var_S", 0, [values], false, true, null));
        };
        Functions.prototype.vdb = function (cost, salvage, life, startPeriod, endPeriod, factor, noSwitch) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Vdb", 0, [cost, salvage, life, startPeriod, endPeriod, factor, noSwitch], false, true, null));
        };
        Functions.prototype.weekNum = function (serialNumber, returnType) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "WeekNum", 0, [serialNumber, returnType], false, true, null));
        };
        Functions.prototype.weekday = function (serialNumber, returnType) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Weekday", 0, [serialNumber, returnType], false, true, null));
        };
        Functions.prototype.weibull_Dist = function (x, alpha, beta, cumulative) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Weibull_Dist", 0, [x, alpha, beta, cumulative], false, true, null));
        };
        Functions.prototype.workDay = function (startDate, days, holidays) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "WorkDay", 0, [startDate, days, holidays], false, true, null));
        };
        Functions.prototype.workDay_Intl = function (startDate, days, weekend, holidays) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "WorkDay_Intl", 0, [startDate, days, weekend, holidays], false, true, null));
        };
        Functions.prototype.xirr = function (values, dates, guess) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Xirr", 0, [values, dates, guess], false, true, null));
        };
        Functions.prototype.xnpv = function (rate, values, dates) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Xnpv", 0, [rate, values, dates], false, true, null));
        };
        Functions.prototype.xor = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Xor", 0, [values], false, true, null));
        };
        Functions.prototype.year = function (serialNumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Year", 0, [serialNumber], false, true, null));
        };
        Functions.prototype.yearFrac = function (startDate, endDate, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "YearFrac", 0, [startDate, endDate, basis], false, true, null));
        };
        Functions.prototype.yield = function (settlement, maturity, rate, pr, redemption, frequency, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Yield", 0, [settlement, maturity, rate, pr, redemption, frequency, basis], false, true, null));
        };
        Functions.prototype.yieldDisc = function (settlement, maturity, pr, redemption, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "YieldDisc", 0, [settlement, maturity, pr, redemption, basis], false, true, null));
        };
        Functions.prototype.yieldMat = function (settlement, maturity, issue, rate, pr, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "YieldMat", 0, [settlement, maturity, issue, rate, pr, basis], false, true, null));
        };
        Functions.prototype.z_Test = function (array, x, sigma) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Z_Test", 0, [array, x, sigma], false, true, null));
        };
        Functions.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
        };
        Functions.prototype.toJSON = function () {
            return {};
        };
        return Functions;
    }(OfficeExtension.ClientObject));
    Excel.Functions = Functions;
    var CalculatedFieldCollection = (function (_super) {
        __extends(CalculatedFieldCollection, _super);
        function CalculatedFieldCollection() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(CalculatedFieldCollection.prototype, "_className", {
            get: function () {
                return "CalculatedFieldCollection";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(CalculatedFieldCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, "CalculatedFieldCollection", this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        CalculatedFieldCollection.prototype.add = function (Name, Formula, UseStandardFormula) {
            return new Excel.PivotField(this.context, _createMethodObjectPath(this.context, this, "Add", 0, [Name, Formula, UseStandardFormula], false, true, null));
        };
        CalculatedFieldCollection.prototype.getCount = function () {
            var action = _createMethodAction(this.context, this, "GetCount", 1, []);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        CalculatedFieldCollection.prototype.getItem = function (nameOrIndex) {
            return new Excel.PivotField(this.context, _createIndexerObjectPath(this.context, this, [nameOrIndex]));
        };
        CalculatedFieldCollection.prototype.getItemAt = function (index) {
            return new Excel.PivotField(this.context, _createMethodObjectPath(this.context, this, "GetItemAt", 1, [index], false, false, null));
        };
        CalculatedFieldCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.PivotField(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        CalculatedFieldCollection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        CalculatedFieldCollection.prototype.toJSON = function () {
            return {};
        };
        return CalculatedFieldCollection;
    }(OfficeExtension.ClientObject));
    Excel.CalculatedFieldCollection = CalculatedFieldCollection;
    var PivotCache = (function (_super) {
        __extends(PivotCache, _super);
        function PivotCache() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(PivotCache.prototype, "_className", {
            get: function () {
                return "PivotCache";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotCache.prototype, "id", {
            get: function () {
                _throwIfNotLoaded("id", this.m_id, "PivotCache", this._isNull);
                return this.m_id;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotCache.prototype, "index", {
            get: function () {
                _throwIfNotLoaded("index", this.m_index, "PivotCache", this._isNull);
                return this.m_index;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotCache.prototype, "version", {
            get: function () {
                _throwIfNotLoaded("version", this.m_version, "PivotCache", this._isNull);
                return this.m_version;
            },
            enumerable: true,
            configurable: true
        });
        PivotCache.prototype.refresh = function () {
            _createMethodAction(this.context, this, "Refresh", 0, []);
        };
        PivotCache.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Id"])) {
                this.m_id = obj["Id"];
            }
            if (!_isUndefined(obj["Index"])) {
                this.m_index = obj["Index"];
            }
            if (!_isUndefined(obj["Version"])) {
                this.m_version = obj["Version"];
            }
        };
        PivotCache.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        PivotCache.prototype._handleIdResult = function (value) {
            _super.prototype._handleIdResult.call(this, value);
            if (_isNullOrUndefined(value)) {
                return;
            }
            if (!_isUndefined(value["Id"])) {
                this.m_id = value["Id"];
            }
        };
        PivotCache.prototype.toJSON = function () {
            return {
                "id": this.m_id,
                "index": this.m_index,
                "version": this.m_version
            };
        };
        return PivotCache;
    }(OfficeExtension.ClientObject));
    Excel.PivotCache = PivotCache;
    var PivotCacheCollection = (function (_super) {
        __extends(PivotCacheCollection, _super);
        function PivotCacheCollection() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(PivotCacheCollection.prototype, "_className", {
            get: function () {
                return "PivotCacheCollection";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotCacheCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, "PivotCacheCollection", this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        PivotCacheCollection.prototype.add = function (sourceType, address) {
            return new Excel.PivotCache(this.context, _createMethodObjectPath(this.context, this, "Add", 0, [sourceType, address], false, true, null));
        };
        PivotCacheCollection.prototype.getCount = function () {
            var action = _createMethodAction(this.context, this, "GetCount", 1, []);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        PivotCacheCollection.prototype.getItem = function (index) {
            return new Excel.PivotCache(this.context, _createIndexerObjectPath(this.context, this, [index]));
        };
        PivotCacheCollection.prototype.getItemAt = function (index) {
            return new Excel.PivotCache(this.context, _createMethodObjectPath(this.context, this, "GetItemAt", 1, [index], false, false, null));
        };
        PivotCacheCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.PivotCache(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        PivotCacheCollection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        PivotCacheCollection.prototype.toJSON = function () {
            return {};
        };
        return PivotCacheCollection;
    }(OfficeExtension.ClientObject));
    Excel.PivotCacheCollection = PivotCacheCollection;
    var PivotField = (function (_super) {
        __extends(PivotField, _super);
        function PivotField() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(PivotField.prototype, "_className", {
            get: function () {
                return "PivotField";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotField.prototype, "currentPage", {
            get: function () {
                if (!this.m_currentPage) {
                    this.m_currentPage = new Excel.PivotItem(this.context, _createPropertyObjectPath(this.context, this, "CurrentPage", false, false));
                }
                return this.m_currentPage;
            },
            set: function (value) {
                this.m_currentPage = value;
                _createSetPropertyAction(this.context, this, "CurrentPage", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotField.prototype, "hiddenItems", {
            get: function () {
                if (!this.m_hiddenItems) {
                    this.m_hiddenItems = new Excel.PivotItemCollection(this.context, _createPropertyObjectPath(this.context, this, "HiddenItems", true, false));
                }
                return this.m_hiddenItems;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotField.prototype, "pivotItems", {
            get: function () {
                if (!this.m_pivotItems) {
                    this.m_pivotItems = new Excel.PivotItemCollection(this.context, _createPropertyObjectPath(this.context, this, "PivotItems", true, false));
                }
                return this.m_pivotItems;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotField.prototype, "visiblePivotItems", {
            get: function () {
                if (!this.m_visiblePivotItems) {
                    this.m_visiblePivotItems = new Excel.PivotItemCollection(this.context, _createPropertyObjectPath(this.context, this, "VisiblePivotItems", true, false));
                }
                return this.m_visiblePivotItems;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotField.prototype, "aggregationFunction", {
            get: function () {
                _throwIfNotLoaded("aggregationFunction", this.m_aggregationFunction, "PivotField", this._isNull);
                return this.m_aggregationFunction;
            },
            set: function (value) {
                this.m_aggregationFunction = value;
                _createSetPropertyAction(this.context, this, "AggregationFunction", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotField.prototype, "allItemsVisible", {
            get: function () {
                _throwIfNotLoaded("allItemsVisible", this.m_allItemsVisible, "PivotField", this._isNull);
                return this.m_allItemsVisible;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotField.prototype, "autoSortField", {
            get: function () {
                _throwIfNotLoaded("autoSortField", this.m_autoSortField, "PivotField", this._isNull);
                return this.m_autoSortField;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotField.prototype, "autoSortOrder", {
            get: function () {
                _throwIfNotLoaded("autoSortOrder", this.m_autoSortOrder, "PivotField", this._isNull);
                return this.m_autoSortOrder;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotField.prototype, "calculated", {
            get: function () {
                _throwIfNotLoaded("calculated", this.m_calculated, "PivotField", this._isNull);
                return this.m_calculated;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotField.prototype, "calculation", {
            get: function () {
                _throwIfNotLoaded("calculation", this.m_calculation, "PivotField", this._isNull);
                return this.m_calculation;
            },
            set: function (value) {
                this.m_calculation = value;
                _createSetPropertyAction(this.context, this, "Calculation", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotField.prototype, "caption", {
            get: function () {
                _throwIfNotLoaded("caption", this.m_caption, "PivotField", this._isNull);
                return this.m_caption;
            },
            set: function (value) {
                this.m_caption = value;
                _createSetPropertyAction(this.context, this, "Caption", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotField.prototype, "dataType", {
            get: function () {
                _throwIfNotLoaded("dataType", this.m_dataType, "PivotField", this._isNull);
                return this.m_dataType;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotField.prototype, "drilledDown", {
            get: function () {
                _throwIfNotLoaded("drilledDown", this.m_drilledDown, "PivotField", this._isNull);
                return this.m_drilledDown;
            },
            set: function (value) {
                this.m_drilledDown = value;
                _createSetPropertyAction(this.context, this, "DrilledDown", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotField.prototype, "enableMultiplePageItems", {
            get: function () {
                _throwIfNotLoaded("enableMultiplePageItems", this.m_enableMultiplePageItems, "PivotField", this._isNull);
                return this.m_enableMultiplePageItems;
            },
            set: function (value) {
                this.m_enableMultiplePageItems = value;
                _createSetPropertyAction(this.context, this, "EnableMultiplePageItems", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotField.prototype, "formula", {
            get: function () {
                _throwIfNotLoaded("formula", this.m_formula, "PivotField", this._isNull);
                return this.m_formula;
            },
            set: function (value) {
                this.m_formula = value;
                _createSetPropertyAction(this.context, this, "Formula", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotField.prototype, "name", {
            get: function () {
                _throwIfNotLoaded("name", this.m_name, "PivotField", this._isNull);
                return this.m_name;
            },
            set: function (value) {
                this.m_name = value;
                _createSetPropertyAction(this.context, this, "Name", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotField.prototype, "numberFormat", {
            get: function () {
                _throwIfNotLoaded("numberFormat", this.m_numberFormat, "PivotField", this._isNull);
                return this.m_numberFormat;
            },
            set: function (value) {
                this.m_numberFormat = value;
                _createSetPropertyAction(this.context, this, "NumberFormat", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotField.prototype, "orientation", {
            get: function () {
                _throwIfNotLoaded("orientation", this.m_orientation, "PivotField", this._isNull);
                return this.m_orientation;
            },
            set: function (value) {
                this.m_orientation = value;
                _createSetPropertyAction(this.context, this, "Orientation", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotField.prototype, "position", {
            get: function () {
                _throwIfNotLoaded("position", this.m_position, "PivotField", this._isNull);
                return this.m_position;
            },
            set: function (value) {
                this.m_position = value;
                _createSetPropertyAction(this.context, this, "Position", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotField.prototype, "showDetail", {
            get: function () {
                _throwIfNotLoaded("showDetail", this.m_showDetail, "PivotField", this._isNull);
                return this.m_showDetail;
            },
            set: function (value) {
                this.m_showDetail = value;
                _createSetPropertyAction(this.context, this, "ShowDetail", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotField.prototype, "sourceName", {
            get: function () {
                _throwIfNotLoaded("sourceName", this.m_sourceName, "PivotField", this._isNull);
                return this.m_sourceName;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotField.prototype, "subtotals", {
            get: function () {
                _throwIfNotLoaded("subtotals", this.m_subtotals, "PivotField", this._isNull);
                return this.m_subtotals;
            },
            set: function (value) {
                this.m_subtotals = value;
                _createSetPropertyAction(this.context, this, "Subtotals", value);
            },
            enumerable: true,
            configurable: true
        });
        PivotField.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["name", "position", "orientation", "caption", "numberFormat", "formula", "calculation", "showDetail", "aggregationFunction", "drilledDown", "enableMultiplePageItems", "subtotals"], [], [
                "currentPage",
                "hiddenItems",
                "pivotItems",
                "visiblePivotItems",
                "currentPage",
                "hiddenItems",
                "pivotItems",
                "visiblePivotItems"
            ]);
        };
        PivotField.prototype.autoGroup = function () {
            _createMethodAction(this.context, this, "AutoGroup", 0, []);
        };
        PivotField.prototype.autoSort = function (sortOrder, Field) {
            _createMethodAction(this.context, this, "AutoSort", 0, [sortOrder, Field]);
        };
        PivotField.prototype.clearAllFilters = function () {
            _createMethodAction(this.context, this, "ClearAllFilters", 0, []);
        };
        PivotField.prototype.getChildField = function () {
            return new Excel.PivotField(this.context, _createMethodObjectPath(this.context, this, "GetChildField", 1, [], false, false, null));
        };
        PivotField.prototype.getChildItems = function () {
            return new Excel.PivotItemCollection(this.context, _createMethodObjectPath(this.context, this, "GetChildItems", 1, [], true, false, null));
        };
        PivotField.prototype.getDataRange = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetDataRange", 1, [], false, true, null));
        };
        PivotField.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["AggregationFunction"])) {
                this.m_aggregationFunction = obj["AggregationFunction"];
            }
            if (!_isUndefined(obj["AllItemsVisible"])) {
                this.m_allItemsVisible = obj["AllItemsVisible"];
            }
            if (!_isUndefined(obj["AutoSortField"])) {
                this.m_autoSortField = obj["AutoSortField"];
            }
            if (!_isUndefined(obj["AutoSortOrder"])) {
                this.m_autoSortOrder = obj["AutoSortOrder"];
            }
            if (!_isUndefined(obj["Calculated"])) {
                this.m_calculated = obj["Calculated"];
            }
            if (!_isUndefined(obj["Calculation"])) {
                this.m_calculation = obj["Calculation"];
            }
            if (!_isUndefined(obj["Caption"])) {
                this.m_caption = obj["Caption"];
            }
            if (!_isUndefined(obj["DataType"])) {
                this.m_dataType = obj["DataType"];
            }
            if (!_isUndefined(obj["DrilledDown"])) {
                this.m_drilledDown = obj["DrilledDown"];
            }
            if (!_isUndefined(obj["EnableMultiplePageItems"])) {
                this.m_enableMultiplePageItems = obj["EnableMultiplePageItems"];
            }
            if (!_isUndefined(obj["Formula"])) {
                this.m_formula = obj["Formula"];
            }
            if (!_isUndefined(obj["Name"])) {
                this.m_name = obj["Name"];
            }
            if (!_isUndefined(obj["NumberFormat"])) {
                this.m_numberFormat = obj["NumberFormat"];
            }
            if (!_isUndefined(obj["Orientation"])) {
                this.m_orientation = obj["Orientation"];
            }
            if (!_isUndefined(obj["Position"])) {
                this.m_position = obj["Position"];
            }
            if (!_isUndefined(obj["ShowDetail"])) {
                this.m_showDetail = obj["ShowDetail"];
            }
            if (!_isUndefined(obj["SourceName"])) {
                this.m_sourceName = obj["SourceName"];
            }
            if (!_isUndefined(obj["Subtotals"])) {
                this.m_subtotals = obj["Subtotals"];
            }
            _handleNavigationPropertyResults(this, obj, ["currentPage", "CurrentPage", "hiddenItems", "HiddenItems", "pivotItems", "PivotItems", "visiblePivotItems", "VisiblePivotItems"]);
        };
        PivotField.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        PivotField.prototype.toJSON = function () {
            return {
                "aggregationFunction": this.m_aggregationFunction,
                "allItemsVisible": this.m_allItemsVisible,
                "autoSortField": this.m_autoSortField,
                "autoSortOrder": this.m_autoSortOrder,
                "calculated": this.m_calculated,
                "calculation": this.m_calculation,
                "caption": this.m_caption,
                "dataType": this.m_dataType,
                "drilledDown": this.m_drilledDown,
                "enableMultiplePageItems": this.m_enableMultiplePageItems,
                "formula": this.m_formula,
                "name": this.m_name,
                "numberFormat": this.m_numberFormat,
                "orientation": this.m_orientation,
                "position": this.m_position,
                "showDetail": this.m_showDetail,
                "sourceName": this.m_sourceName,
                "subtotals": this.m_subtotals
            };
        };
        return PivotField;
    }(OfficeExtension.ClientObject));
    Excel.PivotField = PivotField;
    var PivotFieldCollection = (function (_super) {
        __extends(PivotFieldCollection, _super);
        function PivotFieldCollection() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(PivotFieldCollection.prototype, "_className", {
            get: function () {
                return "PivotFieldCollection";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotFieldCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, "PivotFieldCollection", this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        PivotFieldCollection.prototype.getCount = function () {
            var action = _createMethodAction(this.context, this, "GetCount", 1, []);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        PivotFieldCollection.prototype.getItem = function (nameOrIndex) {
            return new Excel.PivotField(this.context, _createIndexerObjectPath(this.context, this, [nameOrIndex]));
        };
        PivotFieldCollection.prototype.getItemAt = function (index) {
            return new Excel.PivotField(this.context, _createMethodObjectPath(this.context, this, "GetItemAt", 1, [index], false, false, null));
        };
        PivotFieldCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.PivotField(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        PivotFieldCollection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        PivotFieldCollection.prototype.toJSON = function () {
            return {};
        };
        return PivotFieldCollection;
    }(OfficeExtension.ClientObject));
    Excel.PivotFieldCollection = PivotFieldCollection;
    var PivotItem = (function (_super) {
        __extends(PivotItem, _super);
        function PivotItem() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(PivotItem.prototype, "_className", {
            get: function () {
                return "PivotItem";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotItem.prototype, "pivotField", {
            get: function () {
                if (!this.m_pivotField) {
                    this.m_pivotField = new Excel.PivotField(this.context, _createPropertyObjectPath(this.context, this, "PivotField", false, false));
                }
                return this.m_pivotField;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotItem.prototype, "calculated", {
            get: function () {
                _throwIfNotLoaded("calculated", this.m_calculated, "PivotItem", this._isNull);
                return this.m_calculated;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotItem.prototype, "drilledDown", {
            get: function () {
                _throwIfNotLoaded("drilledDown", this.m_drilledDown, "PivotItem", this._isNull);
                return this.m_drilledDown;
            },
            set: function (value) {
                this.m_drilledDown = value;
                _createSetPropertyAction(this.context, this, "DrilledDown", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotItem.prototype, "name", {
            get: function () {
                _throwIfNotLoaded("name", this.m_name, "PivotItem", this._isNull);
                return this.m_name;
            },
            set: function (value) {
                this.m_name = value;
                _createSetPropertyAction(this.context, this, "Name", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotItem.prototype, "position", {
            get: function () {
                _throwIfNotLoaded("position", this.m_position, "PivotItem", this._isNull);
                return this.m_position;
            },
            set: function (value) {
                this.m_position = value;
                _createSetPropertyAction(this.context, this, "Position", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotItem.prototype, "recordCount", {
            get: function () {
                _throwIfNotLoaded("recordCount", this.m_recordCount, "PivotItem", this._isNull);
                return this.m_recordCount;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotItem.prototype, "showDetail", {
            get: function () {
                _throwIfNotLoaded("showDetail", this.m_showDetail, "PivotItem", this._isNull);
                return this.m_showDetail;
            },
            set: function (value) {
                this.m_showDetail = value;
                _createSetPropertyAction(this.context, this, "ShowDetail", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotItem.prototype, "sourceName", {
            get: function () {
                _throwIfNotLoaded("sourceName", this.m_sourceName, "PivotItem", this._isNull);
                return this.m_sourceName;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotItem.prototype, "value", {
            get: function () {
                _throwIfNotLoaded("value", this.m_value, "PivotItem", this._isNull);
                return this.m_value;
            },
            set: function (value) {
                this.m_value = value;
                _createSetPropertyAction(this.context, this, "Value", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotItem.prototype, "visible", {
            get: function () {
                _throwIfNotLoaded("visible", this.m_visible, "PivotItem", this._isNull);
                return this.m_visible;
            },
            set: function (value) {
                this.m_visible = value;
                _createSetPropertyAction(this.context, this, "Visible", value);
            },
            enumerable: true,
            configurable: true
        });
        PivotItem.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["value", "name", "position", "visible", "showDetail", "drilledDown"], [], [
                "pivotField",
                "pivotField"
            ]);
        };
        PivotItem.prototype.getDataRange = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetDataRange", 1, [], false, true, null));
        };
        PivotItem.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Calculated"])) {
                this.m_calculated = obj["Calculated"];
            }
            if (!_isUndefined(obj["DrilledDown"])) {
                this.m_drilledDown = obj["DrilledDown"];
            }
            if (!_isUndefined(obj["Name"])) {
                this.m_name = obj["Name"];
            }
            if (!_isUndefined(obj["Position"])) {
                this.m_position = obj["Position"];
            }
            if (!_isUndefined(obj["RecordCount"])) {
                this.m_recordCount = obj["RecordCount"];
            }
            if (!_isUndefined(obj["ShowDetail"])) {
                this.m_showDetail = obj["ShowDetail"];
            }
            if (!_isUndefined(obj["SourceName"])) {
                this.m_sourceName = obj["SourceName"];
            }
            if (!_isUndefined(obj["Value"])) {
                this.m_value = obj["Value"];
            }
            if (!_isUndefined(obj["Visible"])) {
                this.m_visible = obj["Visible"];
            }
            _handleNavigationPropertyResults(this, obj, ["pivotField", "PivotField"]);
        };
        PivotItem.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        PivotItem.prototype.toJSON = function () {
            return {
                "calculated": this.m_calculated,
                "drilledDown": this.m_drilledDown,
                "name": this.m_name,
                "position": this.m_position,
                "recordCount": this.m_recordCount,
                "showDetail": this.m_showDetail,
                "sourceName": this.m_sourceName,
                "value": this.m_value,
                "visible": this.m_visible
            };
        };
        return PivotItem;
    }(OfficeExtension.ClientObject));
    Excel.PivotItem = PivotItem;
    var PivotItemCollection = (function (_super) {
        __extends(PivotItemCollection, _super);
        function PivotItemCollection() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(PivotItemCollection.prototype, "_className", {
            get: function () {
                return "PivotItemCollection";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotItemCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, "PivotItemCollection", this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        PivotItemCollection.prototype.getCount = function () {
            var action = _createMethodAction(this.context, this, "GetCount", 1, []);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        PivotItemCollection.prototype.getItem = function (nameOrIndex) {
            return new Excel.PivotItem(this.context, _createIndexerObjectPath(this.context, this, [nameOrIndex]));
        };
        PivotItemCollection.prototype.getItemAt = function (index) {
            return new Excel.PivotItem(this.context, _createMethodObjectPath(this.context, this, "GetItemAt", 1, [index], false, false, null));
        };
        PivotItemCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.PivotItem(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        PivotItemCollection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        PivotItemCollection.prototype.toJSON = function () {
            return {};
        };
        return PivotItemCollection;
    }(OfficeExtension.ClientObject));
    Excel.PivotItemCollection = PivotItemCollection;
    var ConsolidationFunction;
    (function (ConsolidationFunction) {
        ConsolidationFunction.varP = "VarP";
        ConsolidationFunction._Var = "Var";
        ConsolidationFunction.sum = "Sum";
        ConsolidationFunction.stDevP = "StDevP";
        ConsolidationFunction.stDev = "StDev";
        ConsolidationFunction.product = "Product";
        ConsolidationFunction.min = "Min";
        ConsolidationFunction.max = "Max";
        ConsolidationFunction.countNums = "CountNums";
        ConsolidationFunction.count = "Count";
        ConsolidationFunction.average = "Average";
        ConsolidationFunction.distinctCount = "DistinctCount";
        ConsolidationFunction.unknown = "Unknown";
    })(ConsolidationFunction = Excel.ConsolidationFunction || (Excel.ConsolidationFunction = {}));
    var PivotFieldCalculation;
    (function (PivotFieldCalculation) {
        PivotFieldCalculation.noAdditionalCalculation = "NoAdditionalCalculation";
        PivotFieldCalculation.differenceFrom = "DifferenceFrom";
        PivotFieldCalculation.percentOf = "PercentOf";
        PivotFieldCalculation.percentDifferenceFrom = "PercentDifferenceFrom";
        PivotFieldCalculation.runningTotal = "RunningTotal";
        PivotFieldCalculation.percentOfRow = "PercentOfRow";
        PivotFieldCalculation.percentOfColumn = "PercentOfColumn";
        PivotFieldCalculation.percentOfTotal = "PercentOfTotal";
        PivotFieldCalculation.index = "Index";
        PivotFieldCalculation.percentOfParentRow = "PercentOfParentRow";
        PivotFieldCalculation.percentOfParentColumn = "PercentOfParentColumn";
        PivotFieldCalculation.percentOfParent = "PercentOfParent";
        PivotFieldCalculation.percentRunningTotal = "PercentRunningTotal";
        PivotFieldCalculation.rankAscending = "RankAscending";
        PivotFieldCalculation.rankDecending = "RankDecending";
    })(PivotFieldCalculation = Excel.PivotFieldCalculation || (Excel.PivotFieldCalculation = {}));
    var PivotFieldDataType;
    (function (PivotFieldDataType) {
        PivotFieldDataType.text = "Text";
        PivotFieldDataType.number = "Number";
        PivotFieldDataType.date = "Date";
    })(PivotFieldDataType = Excel.PivotFieldDataType || (Excel.PivotFieldDataType = {}));
    var PivotFieldOrientation;
    (function (PivotFieldOrientation) {
        PivotFieldOrientation.hidden = "Hidden";
        PivotFieldOrientation.rowField = "RowField";
        PivotFieldOrientation.columnField = "ColumnField";
        PivotFieldOrientation.pageField = "PageField";
        PivotFieldOrientation.dataField = "DataField";
    })(PivotFieldOrientation = Excel.PivotFieldOrientation || (Excel.PivotFieldOrientation = {}));
    var PivotFieldRepeatLabels;
    (function (PivotFieldRepeatLabels) {
        PivotFieldRepeatLabels.doNotRepeatLabels = "DoNotRepeatLabels";
        PivotFieldRepeatLabels.repeatLabels = "RepeatLabels";
    })(PivotFieldRepeatLabels = Excel.PivotFieldRepeatLabels || (Excel.PivotFieldRepeatLabels = {}));
    var PivotTableSourceType;
    (function (PivotTableSourceType) {
        PivotTableSourceType.database = "Database";
        PivotTableSourceType.external = "External";
        PivotTableSourceType.consolidation = "Consolidation";
        PivotTableSourceType.scenario = "Scenario";
    })(PivotTableSourceType = Excel.PivotTableSourceType || (Excel.PivotTableSourceType = {}));
    var PivotTableVersion;
    (function (PivotTableVersion) {
        PivotTableVersion.pivotTableVersionCurrent = "PivotTableVersionCurrent";
        PivotTableVersion.pivotTableVersion2000 = "PivotTableVersion2000";
        PivotTableVersion.pivotTableVersion10 = "PivotTableVersion10";
        PivotTableVersion.pivotTableVersion11 = "PivotTableVersion11";
        PivotTableVersion.pivotTableVersion12 = "PivotTableVersion12";
        PivotTableVersion.pivotTableVersion14 = "PivotTableVersion14";
        PivotTableVersion.pivotTableVersion15 = "PivotTableVersion15";
    })(PivotTableVersion = Excel.PivotTableVersion || (Excel.PivotTableVersion = {}));
    var PivotTableSelectionMode;
    (function (PivotTableSelectionMode) {
        PivotTableSelectionMode.dataAndLabel = "DataAndLabel";
        PivotTableSelectionMode.labelOnly = "LabelOnly";
        PivotTableSelectionMode.dataOnly = "DataOnly";
        PivotTableSelectionMode.origin = "Origin";
        PivotTableSelectionMode.blanks = "Blanks";
        PivotTableSelectionMode.button = "Button";
        PivotTableSelectionMode.firstRow = "FirstRow";
    })(PivotTableSelectionMode = Excel.PivotTableSelectionMode || (Excel.PivotTableSelectionMode = {}));
    var SortOrder;
    (function (SortOrder) {
        SortOrder.ascending = "Ascending";
        SortOrder.descending = "Descending";
    })(SortOrder = Excel.SortOrder || (Excel.SortOrder = {}));
    var ErrorCodes;
    (function (ErrorCodes) {
        ErrorCodes.accessDenied = "AccessDenied";
        ErrorCodes.apiNotFound = "ApiNotFound";
        ErrorCodes.generalException = "GeneralException";
        ErrorCodes.insertDeleteConflict = "InsertDeleteConflict";
        ErrorCodes.invalidArgument = "InvalidArgument";
        ErrorCodes.invalidBinding = "InvalidBinding";
        ErrorCodes.invalidOperation = "InvalidOperation";
        ErrorCodes.invalidReference = "InvalidReference";
        ErrorCodes.invalidSelection = "InvalidSelection";
        ErrorCodes.itemAlreadyExists = "ItemAlreadyExists";
        ErrorCodes.itemNotFound = "ItemNotFound";
        ErrorCodes.notImplemented = "NotImplemented";
        ErrorCodes.unsupportedOperation = "UnsupportedOperation";
        ErrorCodes.invalidOperationInCellEditMode = "InvalidOperationInCellEditMode";
    })(ErrorCodes = Excel.ErrorCodes || (Excel.ErrorCodes = {}));
})(Excel || (Excel = {}));
