var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var OfficeExtension;
(function (OfficeExtension) {
    var Action = (function () {
        function Action(actionInfo, isWriteOperation) {
            this.m_actionInfo = actionInfo;
            this.m_isWriteOperation = isWriteOperation;
        }
        Object.defineProperty(Action.prototype, "actionInfo", {
            get: function () {
                return this.m_actionInfo;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Action.prototype, "isWriteOperation", {
            get: function () {
                return this.m_isWriteOperation;
            },
            enumerable: true,
            configurable: true
        });
        return Action;
    }());
    OfficeExtension.Action = Action;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var TraceMarkerActionResultHandler = (function () {
        function TraceMarkerActionResultHandler(callback) {
            this.m_callback = callback;
        }
        TraceMarkerActionResultHandler.prototype._handleResult = function (value) {
            if (this.m_callback) {
                this.m_callback();
            }
        };
        return TraceMarkerActionResultHandler;
    }());
    var ActionFactory = (function () {
        function ActionFactory() {
        }
        ActionFactory.createSetPropertyAction = function (context, parent, propertyName, value) {
            OfficeExtension.Utility.validateObjectPath(parent);
            var actionInfo = {
                Id: context._nextId(),
                ActionType: 4,
                Name: propertyName,
                ObjectPathId: parent._objectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            };
            var args = [value];
            var referencedArgumentObjectPaths = OfficeExtension.Utility.setMethodArguments(context, actionInfo.ArgumentInfo, args);
            OfficeExtension.Utility.validateReferencedObjectPaths(referencedArgumentObjectPaths);
            var ret = new OfficeExtension.Action(actionInfo, true);
            context._pendingRequest.addAction(ret);
            context._pendingRequest.addReferencedObjectPath(parent._objectPath);
            context._pendingRequest.addReferencedObjectPaths(referencedArgumentObjectPaths);
            return ret;
        };
        ActionFactory.createMethodAction = function (context, parent, methodName, operationType, args) {
            OfficeExtension.Utility.validateObjectPath(parent);
            var actionInfo = {
                Id: context._nextId(),
                ActionType: 3,
                Name: methodName,
                ObjectPathId: parent._objectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            };
            var referencedArgumentObjectPaths = OfficeExtension.Utility.setMethodArguments(context, actionInfo.ArgumentInfo, args);
            OfficeExtension.Utility.validateReferencedObjectPaths(referencedArgumentObjectPaths);
            var isWriteOperation = operationType != 1;
            var ret = new OfficeExtension.Action(actionInfo, isWriteOperation);
            context._pendingRequest.addAction(ret);
            context._pendingRequest.addReferencedObjectPath(parent._objectPath);
            context._pendingRequest.addReferencedObjectPaths(referencedArgumentObjectPaths);
            return ret;
        };
        ActionFactory.createQueryAction = function (context, parent, queryOption) {
            OfficeExtension.Utility.validateObjectPath(parent);
            var actionInfo = {
                Id: context._nextId(),
                ActionType: 2,
                Name: "",
                ObjectPathId: parent._objectPath.objectPathInfo.Id,
            };
            actionInfo.QueryInfo = queryOption;
            var ret = new OfficeExtension.Action(actionInfo, false);
            context._pendingRequest.addAction(ret);
            context._pendingRequest.addReferencedObjectPath(parent._objectPath);
            return ret;
        };
        ActionFactory.createRecursiveQueryAction = function (context, parent, query) {
            OfficeExtension.Utility.validateObjectPath(parent);
            var actionInfo = {
                Id: context._nextId(),
                ActionType: 6,
                Name: "",
                ObjectPathId: parent._objectPath.objectPathInfo.Id,
                RecursiveQueryInfo: query
            };
            var ret = new OfficeExtension.Action(actionInfo, false);
            context._pendingRequest.addAction(ret);
            context._pendingRequest.addReferencedObjectPath(parent._objectPath);
            return ret;
        };
        ActionFactory.createInstantiateAction = function (context, obj) {
            OfficeExtension.Utility.validateObjectPath(obj);
            var actionInfo = {
                Id: context._nextId(),
                ActionType: 1,
                Name: "",
                ObjectPathId: obj._objectPath.objectPathInfo.Id
            };
            var ret = new OfficeExtension.Action(actionInfo, false);
            context._pendingRequest.addAction(ret);
            context._pendingRequest.addReferencedObjectPath(obj._objectPath);
            context._pendingRequest.addActionResultHandler(ret, new OfficeExtension.InstantiateActionResultHandler(obj));
            return ret;
        };
        ActionFactory.createTraceAction = function (context, message, addTraceMessage) {
            var actionInfo = {
                Id: context._nextId(),
                ActionType: 5,
                Name: "Trace",
                ObjectPathId: 0
            };
            var ret = new OfficeExtension.Action(actionInfo, false);
            context._pendingRequest.addAction(ret);
            if (addTraceMessage) {
                context._pendingRequest.addTrace(actionInfo.Id, message);
            }
            return ret;
        };
        ActionFactory.createTraceMarkerForCallback = function (context, callback) {
            var action = ActionFactory.createTraceAction(context, null, false);
            context._pendingRequest.addActionResultHandler(action, new TraceMarkerActionResultHandler(callback));
        };
        return ActionFactory;
    }());
    OfficeExtension.ActionFactory = ActionFactory;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var ClientObject = (function () {
        function ClientObject(context, objectPath) {
            OfficeExtension.Utility.checkArgumentNull(context, "context");
            this.m_context = context;
            this.m_objectPath = objectPath;
            if (this.m_objectPath) {
                if (!context._processingResult) {
                    OfficeExtension.ActionFactory.createInstantiateAction(context, this);
                    if ((context._autoCleanup) && (this._KeepReference)) {
                        context.trackedObjects._autoAdd(this);
                    }
                }
            }
        }
        Object.defineProperty(ClientObject.prototype, "context", {
            get: function () {
                return this.m_context;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ClientObject.prototype, "_objectPath", {
            get: function () {
                return this.m_objectPath;
            },
            set: function (value) {
                this.m_objectPath = value;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ClientObject.prototype, "isNull", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("isNull", this._isNull, null, this._isNull);
                return this._isNull;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ClientObject.prototype, "isNullObject", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("isNullObject", this._isNull, null, this._isNull);
                return this._isNull;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ClientObject.prototype, "_isNull", {
            get: function () {
                return this.m_isNull;
            },
            set: function (value) {
                this.m_isNull = value;
                if (value && this.m_objectPath) {
                    this.m_objectPath._updateAsNullObject();
                }
            },
            enumerable: true,
            configurable: true
        });
        ClientObject.prototype._handleResult = function (value) {
            this._isNull = OfficeExtension.Utility.isNullOrUndefined(value);
            this.context.trackedObjects._autoTrackIfNecessaryWhenHandleObjectResultValue(this, value);
        };
        ClientObject.prototype._handleIdResult = function (value) {
            this._isNull = OfficeExtension.Utility.isNullOrUndefined(value);
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, value);
            this.context.trackedObjects._autoTrackIfNecessaryWhenHandleObjectResultValue(this, value);
        };
        ClientObject.prototype._recursivelySet = function (input, options, scalarWriteablePropertyNames, objectPropertyNames, notAllowedToBeSetPropertyNames) {
            var isClientObject = (input instanceof ClientObject);
            if (isClientObject) {
                if (Object.getPrototypeOf(this) === Object.getPrototypeOf(input)) {
                    input = JSON.parse(JSON.stringify(input));
                }
                else {
                    throw OfficeExtension._Internal.RuntimeError._createInvalidArgError({
                        argumentName: 'properties',
                        errorLocation: this._className + ".set"
                    });
                }
            }
            try {
                var prop;
                for (var i = 0; i < scalarWriteablePropertyNames.length; i++) {
                    prop = scalarWriteablePropertyNames[i];
                    if (input.hasOwnProperty(prop)) {
                        this[prop] = input[prop];
                    }
                }
                for (var i = 0; i < objectPropertyNames.length; i++) {
                    prop = objectPropertyNames[i];
                    if (input.hasOwnProperty(prop)) {
                        this[prop].set(input[prop], options);
                    }
                }
                for (var i = 0; i < notAllowedToBeSetPropertyNames.length; i++) {
                    prop = notAllowedToBeSetPropertyNames[i];
                    if (input.hasOwnProperty(prop)) {
                        throw new OfficeExtension._Internal.RuntimeError({
                            code: OfficeExtension.ErrorCodes.invalidArgument,
                            message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.cannotApplyPropertyThroughSetMethod, prop),
                            debugInfo: {
                                errorLocation: prop
                            }
                        });
                    }
                }
                var throwOnReadOnly = !isClientObject;
                if (options && !OfficeExtension.Utility.isNullOrUndefined(throwOnReadOnly)) {
                    throwOnReadOnly = options.throwOnReadOnly;
                }
                for (prop in input) {
                    if (scalarWriteablePropertyNames.indexOf(prop) < 0 && objectPropertyNames.indexOf(prop) < 0) {
                        var propertyDescriptor = Object.getOwnPropertyDescriptor(Object.getPrototypeOf(this), prop);
                        if (!propertyDescriptor) {
                            throw new OfficeExtension._Internal.RuntimeError({
                                code: OfficeExtension.ErrorCodes.invalidArgument,
                                message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.propertyDoesNotExist, prop),
                                debugInfo: {
                                    errorLocation: prop
                                }
                            });
                        }
                        if (throwOnReadOnly && !propertyDescriptor.set) {
                            throw new OfficeExtension._Internal.RuntimeError({
                                code: OfficeExtension.ErrorCodes.invalidArgument,
                                message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.attemptingToSetReadOnlyProperty, prop),
                                debugInfo: {
                                    errorLocation: prop
                                }
                            });
                        }
                    }
                }
            }
            catch (innerError) {
                throw new OfficeExtension._Internal.RuntimeError({
                    code: OfficeExtension.ErrorCodes.invalidArgument,
                    message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.invalidArgument, 'properties'),
                    debugInfo: {
                        errorLocation: this._className + ".set"
                    },
                    innerError: innerError
                });
            }
        };
        return ClientObject;
    }());
    OfficeExtension.ClientObject = ClientObject;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var ClientRequest = (function () {
        function ClientRequest(context) {
            this.m_context = context;
            this.m_actions = [];
            this.m_actionResultHandler = {};
            this.m_referencedObjectPaths = {};
            this.m_flags = 0;
            this.m_traceInfos = {};
            this.m_pendingProcessEventHandlers = [];
            this.m_pendingEventHandlerActions = {};
            this.m_responseTraceIds = {};
            this.m_responseTraceMessages = [];
            this.m_preSyncPromises = [];
        }
        Object.defineProperty(ClientRequest.prototype, "flags", {
            get: function () {
                return this.m_flags;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ClientRequest.prototype, "traceInfos", {
            get: function () {
                return this.m_traceInfos;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ClientRequest.prototype, "_responseTraceMessages", {
            get: function () {
                return this.m_responseTraceMessages;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ClientRequest.prototype, "_responseTraceIds", {
            get: function () {
                return this.m_responseTraceIds;
            },
            enumerable: true,
            configurable: true
        });
        ClientRequest.prototype._setResponseTraceIds = function (value) {
            if (value) {
                for (var i = 0; i < value.length; i++) {
                    var traceId = value[i];
                    this.m_responseTraceIds[traceId] = traceId;
                    var message = this.m_traceInfos[traceId];
                    if (!OfficeExtension.Utility.isNullOrUndefined(message)) {
                        this.m_responseTraceMessages.push(message);
                    }
                }
            }
        };
        ClientRequest.prototype.addAction = function (action) {
            if (action.isWriteOperation) {
                this.m_flags = this.m_flags | 1;
            }
            this.m_actions.push(action);
        };
        Object.defineProperty(ClientRequest.prototype, "hasActions", {
            get: function () {
                return this.m_actions.length > 0;
            },
            enumerable: true,
            configurable: true
        });
        ClientRequest.prototype.addTrace = function (actionId, message) {
            this.m_traceInfos[actionId] = message;
        };
        ClientRequest.prototype.addReferencedObjectPath = function (objectPath) {
            if (this.m_referencedObjectPaths[objectPath.objectPathInfo.Id]) {
                return;
            }
            if (!objectPath.isValid) {
                throw new OfficeExtension._Internal.RuntimeError({
                    code: OfficeExtension.ErrorCodes.invalidObjectPath,
                    message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.invalidObjectPath, OfficeExtension.Utility.getObjectPathExpression(objectPath)),
                    debugInfo: {
                        errorLocation: OfficeExtension.Utility.getObjectPathExpression(objectPath)
                    }
                });
            }
            while (objectPath) {
                if (objectPath.isWriteOperation) {
                    this.m_flags = this.m_flags | 1;
                }
                this.m_referencedObjectPaths[objectPath.objectPathInfo.Id] = objectPath;
                if (objectPath.objectPathInfo.ObjectPathType == 3) {
                    this.addReferencedObjectPaths(objectPath.argumentObjectPaths);
                }
                objectPath = objectPath.parentObjectPath;
            }
        };
        ClientRequest.prototype.addReferencedObjectPaths = function (objectPaths) {
            if (objectPaths) {
                for (var i = 0; i < objectPaths.length; i++) {
                    this.addReferencedObjectPath(objectPaths[i]);
                }
            }
        };
        ClientRequest.prototype.addActionResultHandler = function (action, resultHandler) {
            this.m_actionResultHandler[action.actionInfo.Id] = resultHandler;
        };
        ClientRequest.prototype.buildRequestMessageBody = function () {
            var objectPaths = {};
            for (var i in this.m_referencedObjectPaths) {
                objectPaths[i] = this.m_referencedObjectPaths[i].objectPathInfo;
            }
            var actions = [];
            for (var index = 0; index < this.m_actions.length; index++) {
                actions.push(this.m_actions[index].actionInfo);
            }
            var ret = {
                AutoKeepReference: this.m_context._autoCleanup,
                Actions: actions,
                ObjectPaths: objectPaths
            };
            return ret;
        };
        ClientRequest.prototype.processResponse = function (actionResults) {
            if (actionResults) {
                for (var i = 0; i < actionResults.length; i++) {
                    var actionResult = actionResults[i];
                    var handler = this.m_actionResultHandler[actionResult.ActionId];
                    if (handler) {
                        handler._handleResult(actionResult.Value);
                    }
                }
            }
        };
        ClientRequest.prototype.invalidatePendingInvalidObjectPaths = function () {
            for (var i in this.m_referencedObjectPaths) {
                if (this.m_referencedObjectPaths[i].isInvalidAfterRequest) {
                    this.m_referencedObjectPaths[i].isValid = false;
                }
            }
        };
        ClientRequest.prototype._addPendingEventHandlerAction = function (eventHandlers, action) {
            if (!this.m_pendingEventHandlerActions[eventHandlers._id]) {
                this.m_pendingEventHandlerActions[eventHandlers._id] = [];
                this.m_pendingProcessEventHandlers.push(eventHandlers);
            }
            this.m_pendingEventHandlerActions[eventHandlers._id].push(action);
        };
        Object.defineProperty(ClientRequest.prototype, "_pendingProcessEventHandlers", {
            get: function () {
                return this.m_pendingProcessEventHandlers;
            },
            enumerable: true,
            configurable: true
        });
        ClientRequest.prototype._getPendingEventHandlerActions = function (eventHandlers) {
            return this.m_pendingEventHandlerActions[eventHandlers._id];
        };
        ClientRequest.prototype._addPreSyncPromise = function (value) {
            this.m_preSyncPromises.push(value);
        };
        Object.defineProperty(ClientRequest.prototype, "_preSyncPromises", {
            get: function () {
                return this.m_preSyncPromises;
            },
            enumerable: true,
            configurable: true
        });
        return ClientRequest;
    }());
    OfficeExtension.ClientRequest = ClientRequest;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var SessionBase = (function () {
        function SessionBase() {
        }
        SessionBase.prototype._resolveRequestUrlAndHeaderInfo = function () {
            return OfficeExtension.Utility._createPromiseFromResult(null);
        };
        SessionBase.prototype._createRequestExecutorOrNull = function () {
            return null;
        };
        Object.defineProperty(SessionBase.prototype, "eventRegistration", {
            get: function () {
                return OfficeExtension._Internal.officeJsEventRegistration;
            },
            enumerable: true,
            configurable: true
        });
        return SessionBase;
    }());
    OfficeExtension.SessionBase = SessionBase;
    var ClientRequestContext = (function () {
        function ClientRequestContext(url) {
            this.m_customRequestHeaders = {};
            this._onRunFinishedNotifiers = [];
            this.m_nextId = 0;
            if (ClientRequestContext._overrideSession) {
                this.m_requestUrlAndHeaderInfoResolver = ClientRequestContext._overrideSession;
            }
            else {
                if (OfficeExtension.Utility.isNullOrUndefined(url) || typeof (url) === "string" && url.length === 0) {
                    url = ClientRequestContext.defaultRequestUrlAndHeaders;
                    if (!url) {
                        url = { url: OfficeExtension.Constants.localDocument, headers: {} };
                    }
                }
                if (typeof (url) === "string") {
                    this.m_requestUrlAndHeaderInfo = { url: url, headers: {} };
                }
                else if (ClientRequestContext.isRequestUrlAndHeaderInfoResolver(url)) {
                    this.m_requestUrlAndHeaderInfoResolver = url;
                }
                else if (ClientRequestContext.isRequestUrlAndHeaderInfo(url)) {
                    var requestInfo = url;
                    this.m_requestUrlAndHeaderInfo = { url: requestInfo.url, headers: {} };
                    OfficeExtension.Utility._copyHeaders(requestInfo.headers, this.m_requestUrlAndHeaderInfo.headers);
                }
                else {
                    throw OfficeExtension._Internal.RuntimeError._createInvalidArgError("url");
                }
            }
            if (this.m_requestUrlAndHeaderInfoResolver instanceof SessionBase) {
                this.m_session = this.m_requestUrlAndHeaderInfoResolver;
            }
            this._processingResult = false;
            this._customData = OfficeExtension.Constants.iterativeExecutor;
            this.sync = this.sync.bind(this);
        }
        Object.defineProperty(ClientRequestContext.prototype, "session", {
            get: function () {
                return this.m_session;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ClientRequestContext.prototype, "eventRegistration", {
            get: function () {
                if (this.m_session) {
                    return this.m_session.eventRegistration;
                }
                return OfficeExtension._Internal.officeJsEventRegistration;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ClientRequestContext.prototype, "_url", {
            get: function () {
                if (this.m_requestUrlAndHeaderInfo) {
                    return this.m_requestUrlAndHeaderInfo.url;
                }
                return null;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ClientRequestContext.prototype, "_pendingRequest", {
            get: function () {
                if (this.m_pendingRequest == null) {
                    this.m_pendingRequest = new OfficeExtension.ClientRequest(this);
                }
                return this.m_pendingRequest;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ClientRequestContext.prototype, "trackedObjects", {
            get: function () {
                if (!this.m_trackedObjects) {
                    this.m_trackedObjects = new OfficeExtension.TrackedObjects(this);
                }
                return this.m_trackedObjects;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ClientRequestContext.prototype, "requestHeaders", {
            get: function () {
                return this.m_customRequestHeaders;
            },
            enumerable: true,
            configurable: true
        });
        ClientRequestContext.prototype.load = function (clientObj, option) {
            OfficeExtension.Utility.validateContext(this, clientObj);
            var queryOption = ClientRequestContext.parseQueryOption(option);
            var action = OfficeExtension.ActionFactory.createQueryAction(this, clientObj, queryOption);
            this._pendingRequest.addActionResultHandler(action, clientObj);
        };
        ClientRequestContext.parseQueryOption = function (option) {
            var queryOption = {};
            if (typeof (option) == "string") {
                var select = option;
                queryOption.Select = OfficeExtension.Utility._parseSelectExpand(select);
            }
            else if (Array.isArray(option)) {
                queryOption.Select = option;
            }
            else if (typeof (option) == "object") {
                var loadOption = option;
                if (typeof (loadOption.select) == "string") {
                    queryOption.Select = OfficeExtension.Utility._parseSelectExpand(loadOption.select);
                }
                else if (Array.isArray(loadOption.select)) {
                    queryOption.Select = loadOption.select;
                }
                else if (!OfficeExtension.Utility.isNullOrUndefined(loadOption.select)) {
                    OfficeExtension.Utility.throwError(OfficeExtension.ResourceStrings.invalidArgument, "option.select");
                }
                if (typeof (loadOption.expand) == "string") {
                    queryOption.Expand = OfficeExtension.Utility._parseSelectExpand(loadOption.expand);
                }
                else if (Array.isArray(loadOption.expand)) {
                    queryOption.Expand = loadOption.expand;
                }
                else if (!OfficeExtension.Utility.isNullOrUndefined(loadOption.expand)) {
                    OfficeExtension.Utility.throwError(OfficeExtension.ResourceStrings.invalidArgument, "option.expand");
                }
                if (typeof (loadOption.top) == "number") {
                    queryOption.Top = loadOption.top;
                }
                else if (!OfficeExtension.Utility.isNullOrUndefined(loadOption.top)) {
                    OfficeExtension.Utility.throwError(OfficeExtension.ResourceStrings.invalidArgument, "option.top");
                }
                if (typeof (loadOption.skip) == "number") {
                    queryOption.Skip = loadOption.skip;
                }
                else if (!OfficeExtension.Utility.isNullOrUndefined(loadOption.skip)) {
                    OfficeExtension.Utility.throwError(OfficeExtension.ResourceStrings.invalidArgument, "option.skip");
                }
            }
            else if (!OfficeExtension.Utility.isNullOrUndefined(option)) {
                OfficeExtension.Utility.throwError(OfficeExtension.ResourceStrings.invalidArgument, "option");
            }
            return queryOption;
        };
        ClientRequestContext.prototype.loadRecursive = function (clientObj, options, maxDepth) {
            if (!OfficeExtension.Utility.isPlainJsonObject(options)) {
                throw OfficeExtension._Internal.RuntimeError._createInvalidArgError("options");
            }
            var quries = {};
            for (var key in options) {
                quries[key] = ClientRequestContext.parseQueryOption(options[key]);
            }
            var action = OfficeExtension.ActionFactory.createRecursiveQueryAction(this, clientObj, { Queries: quries, MaxDepth: maxDepth });
            this._pendingRequest.addActionResultHandler(action, clientObj);
        };
        ClientRequestContext.prototype.trace = function (message) {
            OfficeExtension.ActionFactory.createTraceAction(this, message, true);
        };
        ClientRequestContext.prototype._processOfficeJsErrorResponse = function (officeJsErrorCode, response) {
        };
        ClientRequestContext.prototype.syncPrivateMain = function () {
            var _this = this;
            return OfficeExtension.Utility._createPromiseFromResult(null)
                .then(function () {
                if (!_this.m_requestUrlAndHeaderInfo) {
                    return _this.m_requestUrlAndHeaderInfoResolver._resolveRequestUrlAndHeaderInfo()
                        .then(function (value) {
                        _this.m_requestUrlAndHeaderInfo = value;
                        if (!_this.m_requestUrlAndHeaderInfo) {
                            _this.m_requestUrlAndHeaderInfo = { url: OfficeExtension.Constants.localDocument, headers: {} };
                        }
                        if (OfficeExtension.Utility.isNullOrEmptyString(_this.m_requestUrlAndHeaderInfo.url)) {
                            _this.m_requestUrlAndHeaderInfo.url = OfficeExtension.Constants.localDocument;
                        }
                        if (!_this.m_requestUrlAndHeaderInfo.headers) {
                            _this.m_requestUrlAndHeaderInfo.headers = {};
                        }
                        if (typeof (_this.m_requestUrlAndHeaderInfoResolver._createRequestExecutorOrNull) === "function") {
                            var executor = _this.m_requestUrlAndHeaderInfoResolver._createRequestExecutorOrNull();
                            if (executor) {
                                _this._requestExecutor = executor;
                            }
                        }
                    });
                }
            })
                .then(function () {
                var req = _this._pendingRequest;
                _this.m_pendingRequest = null;
                return _this.processPreSyncPromises(req)
                    .then(function () { return _this.syncPrivate(req); });
            });
        };
        ClientRequestContext.prototype.syncPrivate = function (req) {
            var _this = this;
            if (!req.hasActions) {
                return this.processPendingEventHandlers(req);
            }
            var msgBody = req.buildRequestMessageBody();
            var requestFlags = req.flags;
            if (!this._requestExecutor) {
                if (OfficeExtension.Utility._isLocalDocumentUrl(this.m_requestUrlAndHeaderInfo.url)) {
                    this._requestExecutor = new OfficeExtension.OfficeJsRequestExecutor(this);
                }
                else {
                    this._requestExecutor = new OfficeExtension.HttpRequestExecutor();
                }
            }
            var requestExecutor = this._requestExecutor;
            var headers = {};
            OfficeExtension.Utility._copyHeaders(this.m_requestUrlAndHeaderInfo.headers, headers);
            OfficeExtension.Utility._copyHeaders(this.m_customRequestHeaders, headers);
            var requestExecutorRequestMessage = {
                Url: this.m_requestUrlAndHeaderInfo.url,
                Headers: headers,
                Body: msgBody
            };
            req.invalidatePendingInvalidObjectPaths();
            var errorFromResponse = null;
            var errorFromProcessEventHandlers = null;
            return requestExecutor.executeAsync(this._customData, requestFlags, requestExecutorRequestMessage)
                .then(function (response) {
                errorFromResponse = _this.processRequestExecutorResponseMessage(req, response);
                return _this.processPendingEventHandlers(req)
                    .catch(function (ex) {
                    OfficeExtension.Utility.log("Error in processPendingEventHandlers");
                    OfficeExtension.Utility.log(JSON.stringify(ex));
                    errorFromProcessEventHandlers = ex;
                });
            })
                .then(function () {
                if (errorFromResponse) {
                    OfficeExtension.Utility.log("Throw error from response: " + JSON.stringify(errorFromResponse));
                    throw errorFromResponse;
                }
                if (errorFromProcessEventHandlers) {
                    OfficeExtension.Utility.log("Throw error from ProcessEventHandler: " + JSON.stringify(errorFromProcessEventHandlers));
                    var transformedError = null;
                    if (errorFromProcessEventHandlers instanceof OfficeExtension._Internal.RuntimeError) {
                        transformedError = errorFromProcessEventHandlers;
                        transformedError.traceMessages = req._responseTraceMessages;
                    }
                    else {
                        var message = null;
                        if (typeof (errorFromProcessEventHandlers) === "string") {
                            message = errorFromProcessEventHandlers;
                        }
                        else {
                            message = errorFromProcessEventHandlers.message;
                        }
                        if (OfficeExtension.Utility.isNullOrEmptyString(message)) {
                            message = OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.cannotRegisterEvent);
                        }
                        transformedError = new OfficeExtension._Internal.RuntimeError({
                            code: OfficeExtension.ErrorCodes.cannotRegisterEvent,
                            message: message,
                            traceMessages: req._responseTraceMessages
                        });
                    }
                    throw transformedError;
                }
            });
        };
        ClientRequestContext.prototype.processRequestExecutorResponseMessage = function (req, response) {
            if (response.Body && response.Body.TraceIds) {
                req._setResponseTraceIds(response.Body.TraceIds);
            }
            var traceMessages = req._responseTraceMessages;
            if (response.Body) {
                var actionResults = null;
                if (response.Body.Results) {
                    actionResults = response.Body.Results;
                }
                else if (response.Body.ProcessedResults && response.Body.ProcessedResults.Results) {
                    actionResults = response.Body.ProcessedResults.Results;
                }
                if (actionResults) {
                    this._processingResult = true;
                    try {
                        req.processResponse(actionResults);
                    }
                    finally {
                        this._processingResult = false;
                    }
                }
            }
            if (!OfficeExtension.Utility.isNullOrEmptyString(response.ErrorCode)) {
                return new OfficeExtension._Internal.RuntimeError({
                    code: response.ErrorCode,
                    message: response.ErrorMessage,
                    traceMessages: traceMessages
                });
            }
            else if (response.Body && response.Body.Error) {
                return new OfficeExtension._Internal.RuntimeError({
                    code: response.Body.Error.Code,
                    message: response.Body.Error.Message,
                    traceMessages: traceMessages,
                    debugInfo: {
                        errorLocation: response.Body.Error.Location
                    }
                });
            }
            return null;
        };
        ClientRequestContext.prototype.processPendingEventHandlers = function (req) {
            var ret = OfficeExtension.Utility._createPromiseFromResult(null);
            for (var i = 0; i < req._pendingProcessEventHandlers.length; i++) {
                var eventHandlers = req._pendingProcessEventHandlers[i];
                ret = ret.then(this.createProcessOneEventHandlersFunc(eventHandlers, req));
            }
            return ret;
        };
        ClientRequestContext.prototype.createProcessOneEventHandlersFunc = function (eventHandlers, req) {
            return function () { return eventHandlers._processRegistration(req); };
        };
        ClientRequestContext.prototype.processPreSyncPromises = function (req) {
            var ret = OfficeExtension.Utility._createPromiseFromResult(null);
            for (var i = 0; i < req._preSyncPromises.length; i++) {
                var p = req._preSyncPromises[i];
                ret = ret.then(this.createProcessOneProSyncFunc(p));
            }
            return ret;
        };
        ClientRequestContext.prototype.createProcessOneProSyncFunc = function (p) {
            return function () { return p; };
        };
        ClientRequestContext.prototype.sync = function (passThroughValue) {
            return this.syncPrivateMain().then(function () { return passThroughValue; });
        };
        ClientRequestContext._run = function (ctxInitializer, batch, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure) {
            if (numCleanupAttempts === void 0) { numCleanupAttempts = 3; }
            if (retryDelay === void 0) { retryDelay = 5000; }
            return ClientRequestContext._runCommon("run", null, ctxInitializer, batch, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure);
        };
        ClientRequestContext.isRequestUrlAndHeaderInfo = function (value) {
            return (typeof (value) === "object" &&
                value !== null &&
                Object.getPrototypeOf(value) === Object.getPrototypeOf({}) &&
                !OfficeExtension.Utility.isNullOrUndefined(value.url));
        };
        ClientRequestContext.isRequestUrlAndHeaderInfoResolver = function (value) {
            return (typeof (value) === "object" &&
                value !== null &&
                typeof (value._resolveRequestUrlAndHeaderInfo) === "function");
        };
        ClientRequestContext._runBatch = function (functionName, receivedRunArgs, ctxInitializer, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure) {
            if (numCleanupAttempts === void 0) { numCleanupAttempts = 3; }
            if (retryDelay === void 0) { retryDelay = 5000; }
            var ctxRetriever;
            var batch;
            var requestInfo = null;
            var argOffset = 0;
            if (receivedRunArgs.length > 0 &&
                (typeof (receivedRunArgs[0]) === "string" ||
                    ClientRequestContext.isRequestUrlAndHeaderInfo(receivedRunArgs[0]) ||
                    ClientRequestContext.isRequestUrlAndHeaderInfoResolver(receivedRunArgs[0]))) {
                requestInfo = receivedRunArgs[0];
                argOffset = 1;
            }
            if (receivedRunArgs.length == argOffset + 1) {
                ctxRetriever = ctxInitializer;
                batch = receivedRunArgs[argOffset + 0];
            }
            else if (receivedRunArgs.length == argOffset + 2) {
                if (receivedRunArgs[argOffset + 0] instanceof OfficeExtension.ClientObject) {
                    ctxRetriever = function () { return receivedRunArgs[argOffset + 0].context; };
                }
                else if (Array.isArray(receivedRunArgs[argOffset + 0])) {
                    var array = receivedRunArgs[argOffset + 0];
                    if (array.length == 0) {
                        return ClientRequestContext.createErrorPromise(functionName);
                    }
                    for (var i = 0; i < array.length; i++) {
                        if (!(array[i] instanceof OfficeExtension.ClientObject)) {
                            return ClientRequestContext.createErrorPromise(functionName);
                        }
                        if (array[i].context != array[0].context) {
                            return ClientRequestContext.createErrorPromise(functionName, OfficeExtension.ResourceStrings.invalidRequestContext);
                        }
                    }
                    ctxRetriever = function () { return array[0].context; };
                }
                else {
                    return ClientRequestContext.createErrorPromise(functionName);
                }
                batch = receivedRunArgs[argOffset + 1];
            }
            else {
                return ClientRequestContext.createErrorPromise(functionName);
            }
            return ClientRequestContext._runCommon(functionName, requestInfo, ctxRetriever, batch, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure);
        };
        ClientRequestContext.createErrorPromise = function (functionName, code) {
            if (code === void 0) { code = OfficeExtension.ResourceStrings.invalidArgument; }
            return OfficeExtension._Internal.OfficePromise.reject(OfficeExtension.Utility.createRuntimeError(code, OfficeExtension.Utility._getResourceString(code), functionName));
        };
        ClientRequestContext._runCommon = function (functionName, requestInfo, ctxRetriever, batch, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure) {
            if (ClientRequestContext._overrideSession) {
                requestInfo = ClientRequestContext._overrideSession;
            }
            var starterPromise = new OfficeExtension._Internal.OfficePromise(function (resolve, reject) { resolve(); });
            var ctx;
            var succeeded = false;
            var resultOrError;
            return starterPromise
                .then(function () {
                ctx = ctxRetriever(requestInfo);
                if (ctx._autoCleanup) {
                    return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
                        ctx._onRunFinishedNotifiers.push(function () {
                            ctx._autoCleanup = true;
                            resolve();
                        });
                    });
                }
                else {
                    ctx._autoCleanup = true;
                }
            })
                .then(function () {
                if (typeof batch !== 'function') {
                    return ClientRequestContext.createErrorPromise(functionName);
                }
                var batchResult = batch(ctx);
                if (OfficeExtension.Utility.isNullOrUndefined(batchResult) || (typeof batchResult.then !== 'function')) {
                    OfficeExtension.Utility.throwError(OfficeExtension.ResourceStrings.runMustReturnPromise);
                }
                return batchResult;
            })
                .then(function (batchResult) {
                return ctx.sync(batchResult);
            })
                .then(function (result) {
                succeeded = true;
                resultOrError = result;
            })
                .catch(function (error) {
                resultOrError = error;
            })
                .then(function () {
                var itemsToRemove = ctx.trackedObjects._retrieveAndClearAutoCleanupList();
                ctx._autoCleanup = false;
                for (var key in itemsToRemove) {
                    itemsToRemove[key]._objectPath.isValid = false;
                }
                var cleanupCounter = 0;
                if (OfficeExtension.Utility._synchronousCleanup || ClientRequestContext.isRequestUrlAndHeaderInfoResolver(requestInfo)) {
                    return attemptCleanup();
                }
                else {
                    attemptCleanup();
                }
                function attemptCleanup() {
                    cleanupCounter++;
                    for (var key in itemsToRemove) {
                        ctx.trackedObjects.remove(itemsToRemove[key]);
                    }
                    return ctx.sync()
                        .then(function () {
                        if (onCleanupSuccess) {
                            onCleanupSuccess(cleanupCounter);
                        }
                    })
                        .catch(function () {
                        if (onCleanupFailure) {
                            onCleanupFailure(cleanupCounter);
                        }
                        if (cleanupCounter < numCleanupAttempts) {
                            setTimeout(function () {
                                attemptCleanup();
                            }, retryDelay);
                        }
                    });
                }
            })
                .then(function () {
                if (ctx._onRunFinishedNotifiers && ctx._onRunFinishedNotifiers.length > 0) {
                    var func = ctx._onRunFinishedNotifiers.shift();
                    func();
                }
                if (succeeded) {
                    return resultOrError;
                }
                else {
                    throw resultOrError;
                }
            });
        };
        ClientRequestContext.prototype._nextId = function () {
            return ++this.m_nextId;
        };
        return ClientRequestContext;
    }());
    OfficeExtension.ClientRequestContext = ClientRequestContext;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var ClientResult = (function () {
        function ClientResult(type) {
            this.m_type = type;
        }
        Object.defineProperty(ClientResult.prototype, "value", {
            get: function () {
                if (!this.m_isLoaded) {
                    throw new OfficeExtension._Internal.RuntimeError({
                        code: OfficeExtension.ErrorCodes.valueNotLoaded,
                        message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.valueNotLoaded),
                        debugInfo: {
                            errorLocation: "clientResult.value"
                        }
                    });
                }
                return this.m_value;
            },
            enumerable: true,
            configurable: true
        });
        ClientResult.prototype._handleResult = function (value) {
            this.m_isLoaded = true;
            if (typeof (value) === "object" && value && value._IsNull) {
                return;
            }
            if (this.m_type === 1) {
                this.m_value = OfficeExtension.Utility.adjustToDateTime(value);
            }
            else {
                this.m_value = value;
            }
        };
        return ClientResult;
    }());
    OfficeExtension.ClientResult = ClientResult;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var Constants = (function () {
        function Constants() {
        }
        Constants.flags = "flags";
        Constants.getItemAt = "GetItemAt";
        Constants.id = "Id";
        Constants.idPrivate = "_Id";
        Constants.index = "_Index";
        Constants.items = "_Items";
        Constants.iterativeExecutor = "IterativeExecutor";
        Constants.localDocument = "http://document.localhost/";
        Constants.localDocumentApiPrefix = "http://document.localhost/_api/";
        Constants.processQuery = "ProcessQuery";
        Constants.referenceId = "_ReferenceId";
        Constants.isTracked = "_IsTracked";
        Constants.sourceLibHeader = "SdkVersion";
        Constants.embeddingPageOrigin = "EmbeddingPageOrigin";
        Constants.embeddingPageSessionInfo = "EmbeddingPageSessionInfo";
        Constants.eventMessageCategory = 65536;
        return Constants;
    }());
    OfficeExtension.Constants = Constants;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var versionToken = 1;
    var internalConfiguration = {
        invokeRequestModifier: function (request) {
            request.DdaMethod.Version = versionToken;
            return request;
        },
        invokeResponseModifier: function (args) {
            versionToken = args.Version;
            if (args.Error) {
                args.error = {};
                args.error.Code = args.Error;
            }
            return args;
        }
    };
    var EmbeddedApiStatus;
    (function (EmbeddedApiStatus) {
        EmbeddedApiStatus[EmbeddedApiStatus["Success"] = 0] = "Success";
        EmbeddedApiStatus[EmbeddedApiStatus["Timeout"] = 1] = "Timeout";
        EmbeddedApiStatus[EmbeddedApiStatus["InternalError"] = 5001] = "InternalError";
    })(EmbeddedApiStatus || (EmbeddedApiStatus = {}));
    var CommunicationConstants;
    (function (CommunicationConstants) {
        CommunicationConstants.SendingId = "sId";
        CommunicationConstants.RespondingId = "rId";
        CommunicationConstants.CommandKey = "command";
        CommunicationConstants.SessionInfoKey = "sessionInfo";
        CommunicationConstants.ParamsKey = "params";
        CommunicationConstants.ApiReadyCommand = "apiready";
        CommunicationConstants.ExecuteMethodCommand = "executeMethod";
        CommunicationConstants.GetAppContextCommand = "getAppContext";
        CommunicationConstants.RegisterEventCommand = "registerEvent";
        CommunicationConstants.UnregisterEventCommand = "unregisterEvent";
        CommunicationConstants.FireEventCommand = "fireEvent";
    })(CommunicationConstants || (CommunicationConstants = {}));
    var EmbeddedSession = (function (_super) {
        __extends(EmbeddedSession, _super);
        function EmbeddedSession(url, options) {
            _super.call(this);
            this.m_chosenWindow = null;
            this.m_chosenOrigin = null;
            this.m_enabled = true;
            this.m_onMessageHandler = this._onMessage.bind(this);
            this.m_callbackList = {};
            this.m_id = 0;
            this.m_timeoutId = -1;
            this.m_appContext = null;
            this.m_url = url;
            this.m_options = options;
            if (!this.m_options) {
                this.m_options = { sessionKey: Math.random().toString() };
            }
            if (!this.m_options.sessionKey) {
                this.m_options.sessionKey = Math.random().toString();
            }
            if (!this.m_options.container) {
                this.m_options.container = document.body;
            }
            if (!this.m_options.timeoutInMilliseconds) {
                this.m_options.timeoutInMilliseconds = 60000;
            }
            if (!this.m_options.height) {
                this.m_options.height = "400px";
            }
            if (!this.m_options.width) {
                this.m_options.width = "100%";
            }
        }
        EmbeddedSession.prototype._getIFrameSrc = function () {
            var origin = window.location.protocol + "//" + window.location.host;
            var toAppend = OfficeExtension.Constants.embeddingPageOrigin + "=" + encodeURIComponent(origin) + "&" + OfficeExtension.Constants.embeddingPageSessionInfo + "=" + encodeURIComponent(this.m_options.sessionKey);
            var useHash = false;
            if (this.m_url.toLowerCase().indexOf("/_layouts/preauth.aspx") > 0) {
                useHash = true;
            }
            var a = document.createElement("a");
            a.href = this.m_url;
            if (useHash) {
                if (a.hash.length === 0 || a.hash === "#") {
                    a.hash = "#" + toAppend;
                }
                else {
                    a.hash = a.hash + "&" + toAppend;
                }
            }
            else {
                if (a.search.length === 0 || a.search === "?") {
                    a.search = "?" + toAppend;
                }
                else {
                    a.search = a.search + "&" + toAppend;
                }
            }
            var iframeSrc = a.href;
            return iframeSrc;
        };
        EmbeddedSession.prototype.init = function () {
            var _this = this;
            window.addEventListener("message", this.m_onMessageHandler);
            var iframeSrc = this._getIFrameSrc();
            return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
                var iframeElement = document.createElement("iframe");
                if (_this.m_options.id) {
                    iframeElement.id = _this.m_options.id;
                }
                iframeElement.style.height = _this.m_options.height;
                iframeElement.style.width = _this.m_options.width;
                iframeElement.src = iframeSrc;
                _this.m_options.container.appendChild(iframeElement);
                _this.m_timeoutId = setTimeout(function () {
                    _this.close();
                    var err = OfficeExtension.Utility.createRuntimeError(OfficeExtension.ErrorCodes.timeout, OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.timeout), "EmbeddedSession.init");
                    reject(err);
                }, _this.m_options.timeoutInMilliseconds);
                _this.m_promiseResolver = resolve;
            });
        };
        EmbeddedSession.prototype._invoke = function (method, callback, params) {
            if (!this.m_enabled) {
                callback(EmbeddedApiStatus.InternalError, null);
                return;
            }
            if (internalConfiguration.invokeRequestModifier) {
                params = internalConfiguration.invokeRequestModifier(params);
            }
            this._sendMessageWithCallback(this.m_id++, method, params, function (args) {
                if (internalConfiguration.invokeResponseModifier) {
                    args = internalConfiguration.invokeResponseModifier(args);
                }
                var errorCode = args["Error"];
                delete args["Error"];
                callback(errorCode || EmbeddedApiStatus.Success, args);
            });
        };
        EmbeddedSession.prototype.close = function () {
            window.removeEventListener("message", this.m_onMessageHandler);
            window.clearTimeout(this.m_timeoutId);
            this.m_enabled = false;
        };
        Object.defineProperty(EmbeddedSession.prototype, "eventRegistration", {
            get: function () {
                if (!this.m_sessionEventManager) {
                    this.m_sessionEventManager = new OfficeExtension.EventRegistration(this._registerEventImpl.bind(this), this._unregisterEventImpl.bind(this));
                }
                return this.m_sessionEventManager;
            },
            enumerable: true,
            configurable: true
        });
        EmbeddedSession.prototype._createRequestExecutorOrNull = function () {
            return new EmbeddedRequestExecutor(this);
        };
        EmbeddedSession.prototype._resolveRequestUrlAndHeaderInfo = function () {
            return OfficeExtension.Utility._createPromiseFromResult(null);
        };
        EmbeddedSession.prototype._registerEventImpl = function (eventId, targetId) {
            var _this = this;
            return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
                _this._sendMessageWithCallback(_this.m_id++, CommunicationConstants.RegisterEventCommand, { EventId: eventId, TargetId: targetId }, function () {
                    resolve(null);
                });
            });
        };
        EmbeddedSession.prototype._unregisterEventImpl = function (eventId, targetId) {
            var _this = this;
            return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
                _this._sendMessageWithCallback(_this.m_id++, CommunicationConstants.UnregisterEventCommand, { EventId: eventId, TargetId: targetId }, function () {
                    resolve();
                });
            });
        };
        EmbeddedSession.prototype._onMessage = function (event) {
            var _this = this;
            if (!this.m_enabled) {
                return;
            }
            if (this.m_chosenWindow
                && (this.m_chosenWindow !== event.source || this.m_chosenOrigin !== event.origin)) {
                return;
            }
            var eventData = event.data;
            if (eventData && eventData[CommunicationConstants.CommandKey] === CommunicationConstants.ApiReadyCommand) {
                if (!this.m_chosenWindow
                    && this._isValidDescendant(event.source)
                    && eventData[CommunicationConstants.SessionInfoKey] === this.m_options.sessionKey) {
                    this.m_chosenWindow = event.source;
                    this.m_chosenOrigin = event.origin;
                    this._sendMessageWithCallback(this.m_id++, CommunicationConstants.GetAppContextCommand, null, function (appContext) {
                        _this._setupContext(appContext);
                        window.clearTimeout(_this.m_timeoutId);
                        _this.m_promiseResolver();
                    });
                }
                return;
            }
            if (eventData && eventData[CommunicationConstants.CommandKey] === CommunicationConstants.FireEventCommand) {
                var msg = eventData[CommunicationConstants.ParamsKey];
                var eventId = msg["EventId"];
                var targetId = msg["TargetId"];
                var data = msg["Data"];
                if (this.m_sessionEventManager) {
                    var handlers = this.m_sessionEventManager.getHandlers(eventId, targetId);
                    for (var i = 0; i < handlers.length; i++) {
                        handlers[i](data);
                    }
                }
                return;
            }
            if (eventData && eventData.hasOwnProperty(CommunicationConstants.RespondingId)) {
                var rId = eventData[CommunicationConstants.RespondingId];
                var callback = this.m_callbackList[rId];
                if (typeof callback === "function") {
                    callback(eventData[CommunicationConstants.ParamsKey]);
                }
                delete this.m_callbackList[rId];
            }
        };
        EmbeddedSession.prototype._sendMessageWithCallback = function (id, command, data, callback) {
            this.m_callbackList[id] = callback;
            var message = {};
            message[CommunicationConstants.SendingId] = id;
            message[CommunicationConstants.CommandKey] = command;
            message[CommunicationConstants.ParamsKey] = data;
            this.m_chosenWindow.postMessage(JSON.stringify(message), this.m_chosenOrigin);
        };
        EmbeddedSession.prototype._isValidDescendant = function (wnd) {
            var container = this.m_options.container || document.body;
            function doesFrameWindow(containerWindow) {
                if (containerWindow === wnd) {
                    return true;
                }
                for (var i = 0, len = containerWindow.frames.length; i < len; i++) {
                    if (doesFrameWindow(containerWindow.frames[i])) {
                        return true;
                    }
                }
                return false;
            }
            var iframes = container.getElementsByTagName("iframe");
            for (var i = 0, len = iframes.length; i < len; i++) {
                if (doesFrameWindow(iframes[i].contentWindow)) {
                    return true;
                }
            }
            return false;
        };
        EmbeddedSession.prototype._setupContext = function (appContext) {
            if (!(this.m_appContext = appContext)) {
                return;
            }
        };
        return EmbeddedSession;
    }(OfficeExtension.SessionBase));
    OfficeExtension.EmbeddedSession = EmbeddedSession;
    var EmbeddedRequestExecutor = (function () {
        function EmbeddedRequestExecutor(session) {
            this.m_session = session;
        }
        EmbeddedRequestExecutor.prototype.executeAsync = function (customData, requestFlags, requestMessage) {
            var _this = this;
            var messageSafearray = OfficeExtension.RichApiMessageUtility.buildMessageArrayForIRequestExecutor(customData, requestFlags, requestMessage, EmbeddedRequestExecutor.SourceLibHeaderValue);
            return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
                _this.m_session._invoke(CommunicationConstants.ExecuteMethodCommand, function (status, result) {
                    OfficeExtension.Utility.log("Response:");
                    OfficeExtension.Utility.log(JSON.stringify(result));
                    var response;
                    if (status == EmbeddedApiStatus.Success) {
                        response = OfficeExtension.RichApiMessageUtility.buildResponseOnSuccess(OfficeExtension.RichApiMessageUtility.getResponseBodyFromSafeArray(result.Data), OfficeExtension.RichApiMessageUtility.getResponseHeadersFromSafeArray(result.Data));
                    }
                    else {
                        response = OfficeExtension.RichApiMessageUtility.buildResponseOnError(result.error.Code, result.error.Message);
                    }
                    resolve(response);
                }, EmbeddedRequestExecutor._transformMessageArrayIntoParams(messageSafearray));
            });
        };
        EmbeddedRequestExecutor._transformMessageArrayIntoParams = function (msgArray) {
            return {
                ArrayData: msgArray,
                DdaMethod: {
                    DispatchId: EmbeddedRequestExecutor.DispidExecuteRichApiRequestMethod
                }
            };
        };
        EmbeddedRequestExecutor.DispidExecuteRichApiRequestMethod = 93;
        EmbeddedRequestExecutor.SourceLibHeaderValue = "Embedded";
        return EmbeddedRequestExecutor;
    }());
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var _Internal;
    (function (_Internal) {
        var RuntimeError = (function (_super) {
            __extends(RuntimeError, _super);
            function RuntimeError(error) {
                _super.call(this, (typeof error === "string") ? error : error.message);
                this.name = "OfficeExtension.Error";
                if (typeof error === "string") {
                    this.message = error;
                }
                else {
                    this.code = error.code;
                    this.message = error.message;
                    this.traceMessages = error.traceMessages || [];
                    this.innerError = error.innerError || null;
                    this.debugInfo = this._createDebugInfo(error.debugInfo || {});
                }
            }
            RuntimeError.prototype.toString = function () {
                return this.code + ': ' + this.message;
            };
            RuntimeError.prototype._createDebugInfo = function (partialDebugInfo) {
                var debugInfo = {
                    code: this.code,
                    message: this.message,
                    toString: function () {
                        return JSON.stringify(this);
                    }
                };
                for (var key in partialDebugInfo) {
                    debugInfo[key] = partialDebugInfo[key];
                }
                if (this.innerError) {
                    if (this.innerError instanceof OfficeExtension.Error) {
                        debugInfo.innerError = this.innerError.debugInfo;
                    }
                    else {
                        debugInfo.innerError = this.innerError;
                    }
                }
                return debugInfo;
            };
            RuntimeError._createInvalidArgError = function (error) {
                return new _Internal.RuntimeError({
                    code: OfficeExtension.ErrorCodes.invalidArgument,
                    message: (OfficeExtension.Utility.isNullOrEmptyString(error.argumentName) ?
                        OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.invalidArgumentGeneric) :
                        OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.invalidArgument, error.argumentName)),
                    debugInfo: error.errorLocation ? { errorLocation: error.errorLocation } : {},
                    innerError: error.innerError
                });
            };
            return RuntimeError;
        }(Error));
        _Internal.RuntimeError = RuntimeError;
    })(_Internal = OfficeExtension._Internal || (OfficeExtension._Internal = {}));
    OfficeExtension.Error = _Internal.RuntimeError;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var ErrorCodes = (function () {
        function ErrorCodes() {
        }
        ErrorCodes.accessDenied = "AccessDenied";
        ErrorCodes.generalException = "GeneralException";
        ErrorCodes.activityLimitReached = "ActivityLimitReached";
        ErrorCodes.invalidObjectPath = "InvalidObjectPath";
        ErrorCodes.propertyNotLoaded = "PropertyNotLoaded";
        ErrorCodes.valueNotLoaded = "ValueNotLoaded";
        ErrorCodes.invalidRequestContext = "InvalidRequestContext";
        ErrorCodes.invalidArgument = "InvalidArgument";
        ErrorCodes.runMustReturnPromise = "RunMustReturnPromise";
        ErrorCodes.cannotRegisterEvent = "CannotRegisterEvent";
        ErrorCodes.apiNotFound = "ApiNotFound";
        ErrorCodes.connectionFailure = "ConnectionFailure";
        ErrorCodes.timeout = "Timeout";
        ErrorCodes.invalidOrTimedOutSession = "InvalidOrTimedOutSession";
        return ErrorCodes;
    }());
    OfficeExtension.ErrorCodes = ErrorCodes;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var EventHandlers = (function () {
        function EventHandlers(context, parentObject, name, eventInfo) {
            var _this = this;
            this.m_id = context._nextId();
            this.m_context = context;
            this.m_name = name;
            this.m_handlers = [];
            this.m_registered = false;
            this.m_eventInfo = eventInfo;
            this.m_callback = function (args) {
                _this.m_eventInfo.eventArgsTransformFunc(args)
                    .then(function (newArgs) { return _this.fireEvent(newArgs); });
            };
        }
        Object.defineProperty(EventHandlers.prototype, "_registered", {
            get: function () {
                return this.m_registered;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(EventHandlers.prototype, "_id", {
            get: function () {
                return this.m_id;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(EventHandlers.prototype, "_handlers", {
            get: function () {
                return this.m_handlers;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(EventHandlers.prototype, "_context", {
            get: function () {
                return this.m_context;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(EventHandlers.prototype, "_callback", {
            get: function () {
                return this.m_callback;
            },
            enumerable: true,
            configurable: true
        });
        EventHandlers.prototype.add = function (handler) {
            var action = OfficeExtension.ActionFactory.createTraceAction(this.m_context, null, false);
            this.m_context._pendingRequest._addPendingEventHandlerAction(this, { id: action.actionInfo.Id, handler: handler, operation: 0 });
            return new OfficeExtension.EventHandlerResult(this.m_context, this, handler);
        };
        EventHandlers.prototype.remove = function (handler) {
            var action = OfficeExtension.ActionFactory.createTraceAction(this.m_context, null, false);
            this.m_context._pendingRequest._addPendingEventHandlerAction(this, { id: action.actionInfo.Id, handler: handler, operation: 1 });
        };
        EventHandlers.prototype.removeAll = function () {
            var action = OfficeExtension.ActionFactory.createTraceAction(this.m_context, null, false);
            this.m_context._pendingRequest._addPendingEventHandlerAction(this, { id: action.actionInfo.Id, handler: null, operation: 2 });
        };
        EventHandlers.prototype._processRegistration = function (req) {
            var _this = this;
            var ret = OfficeExtension.Utility._createPromiseFromResult(null);
            var actions = req._getPendingEventHandlerActions(this);
            if (!actions) {
                return ret;
            }
            var handlersResult = [];
            for (var i = 0; i < this.m_handlers.length; i++) {
                handlersResult.push(this.m_handlers[i]);
            }
            var hasChange = false;
            for (var i = 0; i < actions.length; i++) {
                if (req._responseTraceIds[actions[i].id]) {
                    hasChange = true;
                    switch (actions[i].operation) {
                        case 0:
                            handlersResult.push(actions[i].handler);
                            break;
                        case 1:
                            for (var index = handlersResult.length - 1; index >= 0; index--) {
                                if (handlersResult[index] === actions[i].handler) {
                                    handlersResult.splice(index, 1);
                                    break;
                                }
                            }
                            break;
                        case 2:
                            handlersResult = [];
                            break;
                    }
                }
            }
            if (hasChange) {
                if (!this.m_registered && handlersResult.length > 0) {
                    ret = ret
                        .then(function () { return _this.m_eventInfo.registerFunc(_this.m_callback); })
                        .then(function () { return (_this.m_registered = true); });
                }
                else if (this.m_registered && handlersResult.length == 0) {
                    ret = ret
                        .then(function () { return _this.m_eventInfo.unregisterFunc(_this.m_callback); })
                        .catch(function (ex) {
                        OfficeExtension.Utility.log("Error when unregister event: " + JSON.stringify(ex));
                    })
                        .then(function () { return (_this.m_registered = false); });
                }
                ret = ret
                    .then(function () { return (_this.m_handlers = handlersResult); });
            }
            return ret;
        };
        EventHandlers.prototype.fireEvent = function (args) {
            var promises = [];
            for (var i = 0; i < this.m_handlers.length; i++) {
                var handler = this.m_handlers[i];
                var p = OfficeExtension.Utility._createPromiseFromResult(null)
                    .then(this.createFireOneEventHandlerFunc(handler, args))
                    .catch(function (ex) {
                    OfficeExtension.Utility.log("Error when invoke handler: " + JSON.stringify(ex));
                });
                promises.push(p);
            }
            OfficeExtension._Internal.OfficePromise.all(promises);
        };
        EventHandlers.prototype.createFireOneEventHandlerFunc = function (handler, args) {
            return function () { return handler(args); };
        };
        return EventHandlers;
    }());
    OfficeExtension.EventHandlers = EventHandlers;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var EventHandlerResult = (function () {
        function EventHandlerResult(context, handlers, handler) {
            this.m_context = context;
            this.m_allHandlers = handlers;
            this.m_handler = handler;
        }
        EventHandlerResult.prototype.remove = function () {
            if (this.m_allHandlers && this.m_handler) {
                this.m_allHandlers.remove(this.m_handler);
                this.m_allHandlers = null;
                this.m_handler = null;
            }
        };
        return EventHandlerResult;
    }());
    OfficeExtension.EventHandlerResult = EventHandlerResult;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var _Internal;
    (function (_Internal) {
        var OfficeJsEventRegistration = (function () {
            function OfficeJsEventRegistration() {
            }
            OfficeJsEventRegistration.prototype.register = function (eventId, targetId, handler) {
                switch (eventId) {
                    case 4:
                        return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.bindings.getByIdAsync(targetId, callback); })
                            .then(function (officeBinding) {
                            return OfficeExtension.Utility.promisify(function (callback) { return officeBinding.addHandlerAsync(Office.EventType.BindingDataChanged, handler, callback); });
                        });
                    case 3:
                        return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.bindings.getByIdAsync(targetId, callback); })
                            .then(function (officeBinding) {
                            return OfficeExtension.Utility.promisify(function (callback) { return officeBinding.addHandlerAsync(Office.EventType.BindingSelectionChanged, handler, callback); });
                        });
                    case 2:
                        return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, handler, callback); });
                    case 1:
                        return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.settings.addHandlerAsync(Office.EventType.SettingsChanged, handler, callback); });
                    case 5:
                        return OfficeExtension.Utility.promisify(function (callback) { return OSF.DDA.RichApi.richApiMessageManager.addHandlerAsync("richApiMessage", handler, callback); });
                    case 13:
                        return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.addHandlerAsync(Office.EventType.ObjectDeleted, handler, { id: targetId }, callback); });
                    case 14:
                        return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.addHandlerAsync(Office.EventType.ObjectSelectionChanged, handler, { id: targetId }, callback); });
                    case 15:
                        return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.addHandlerAsync(Office.EventType.ObjectDataChanged, handler, { id: targetId }, callback); });
                    case 16:
                        return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.addHandlerAsync(Office.EventType.ContentControlAdded, handler, { id: targetId }, callback); });
                    default:
                        throw _Internal.RuntimeError._createInvalidArgError("eventId");
                }
            };
            OfficeJsEventRegistration.prototype.unregister = function (eventId, targetId, handler) {
                switch (eventId) {
                    case 4:
                        return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.bindings.getByIdAsync(targetId, callback); })
                            .then(function (officeBinding) {
                            return OfficeExtension.Utility.promisify(function (callback) { return officeBinding.removeHandlerAsync(Office.EventType.BindingDataChanged, { handler: handler }, callback); });
                        });
                    case 3:
                        return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.bindings.getByIdAsync(targetId, callback); })
                            .then(function (officeBinding) {
                            return OfficeExtension.Utility.promisify(function (callback) { return officeBinding.removeHandlerAsync(Office.EventType.BindingSelectionChanged, { handler: handler }, callback); });
                        });
                    case 2:
                        return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, { handler: handler }, callback); });
                    case 1:
                        return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.settings.removeHandlerAsync(Office.EventType.SettingsChanged, { handler: handler }, callback); });
                    case 5:
                        return OfficeExtension.Utility.promisify(function (callback) { return OSF.DDA.RichApi.richApiMessageManager.removeHandlerAsync("richApiMessage", { handler: handler }, callback); });
                    case 13:
                        return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.removeHandlerAsync(Office.EventType.ObjectDeleted, { id: targetId, handler: handler }, callback); });
                    case 14:
                        return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.removeHandlerAsync(Office.EventType.ObjectSelectionChanged, { id: targetId, handler: handler }, callback); });
                    case 15:
                        return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.removeHandlerAsync(Office.EventType.ObjectDataChanged, { id: targetId, handler: handler }, callback); });
                    case 16:
                        return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.removeHandlerAsync(Office.EventType.ContentControlAdded, { id: targetId, handler: handler }, callback); });
                    default:
                        throw _Internal.RuntimeError._createInvalidArgError("eventId");
                }
            };
            return OfficeJsEventRegistration;
        }());
        _Internal.officeJsEventRegistration = new OfficeJsEventRegistration();
    })(_Internal = OfficeExtension._Internal || (OfficeExtension._Internal = {}));
    var EventRegistration = (function () {
        function EventRegistration(registerEventImpl, unregisterEventImpl) {
            this.m_handlersByEventByTarget = {};
            this.m_registerEventImpl = registerEventImpl;
            this.m_unregisterEventImpl = unregisterEventImpl;
        }
        EventRegistration.prototype.getHandlers = function (eventId, targetId) {
            if (OfficeExtension.Utility.isNullOrUndefined(targetId)) {
                targetId = "";
            }
            var handlersById = this.m_handlersByEventByTarget[eventId];
            if (!handlersById) {
                handlersById = {};
                this.m_handlersByEventByTarget[eventId] = handlersById;
            }
            var handlers = handlersById[targetId];
            if (!handlers) {
                handlers = [];
                handlersById[targetId] = handlers;
            }
            return handlers;
        };
        EventRegistration.prototype.register = function (eventId, targetId, handler) {
            if (!handler) {
                throw _Internal.RuntimeError._createInvalidArgError("handler");
            }
            var handlers = this.getHandlers(eventId, targetId);
            handlers.push(handler);
            if (handlers.length === 1) {
                return this.m_registerEventImpl(eventId, targetId);
            }
            return OfficeExtension.Utility._createPromiseFromResult(null);
        };
        EventRegistration.prototype.unregister = function (eventId, targetId, handler) {
            if (!handler) {
                throw _Internal.RuntimeError._createInvalidArgError("handler");
            }
            var handlers = this.getHandlers(eventId, targetId);
            for (var index = handlers.length - 1; index >= 0; index--) {
                if (handlers[index] === handler) {
                    handlers.splice(index, 1);
                    break;
                }
            }
            if (handlers.length === 0) {
                return this.m_unregisterEventImpl(eventId, targetId);
            }
            return OfficeExtension.Utility._createPromiseFromResult(null);
        };
        return EventRegistration;
    }());
    OfficeExtension.EventRegistration = EventRegistration;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var GenericEventRegistration = (function () {
        function GenericEventRegistration() {
            this.m_eventRegistration = new OfficeExtension.EventRegistration(this._registerEventImpl.bind(this), this._unregisterEventImpl.bind(this));
            this.m_richApiMessageHandler = this._handleRichApiMessage.bind(this);
        }
        GenericEventRegistration.prototype.ready = function () {
            var _this = this;
            if (!this.m_ready) {
                if (GenericEventRegistration._testReadyImpl) {
                    this.m_ready = GenericEventRegistration._testReadyImpl()
                        .then(function () {
                        _this.m_isReady = true;
                    });
                }
                else {
                    this.m_ready = OfficeExtension._Internal.officeJsEventRegistration.register(5, "", this.m_richApiMessageHandler)
                        .then(function () {
                        _this.m_isReady = true;
                    });
                }
            }
            return this.m_ready;
        };
        Object.defineProperty(GenericEventRegistration.prototype, "isReady", {
            get: function () {
                return this.m_isReady;
            },
            enumerable: true,
            configurable: true
        });
        GenericEventRegistration.prototype.register = function (eventId, targetId, handler) {
            var _this = this;
            return this.ready()
                .then(function () { return _this.m_eventRegistration.register(eventId, targetId, handler); });
        };
        GenericEventRegistration.prototype.unregister = function (eventId, targetId, handler) {
            var _this = this;
            return this.ready()
                .then(function () { return _this.m_eventRegistration.unregister(eventId, targetId, handler); });
        };
        GenericEventRegistration.prototype._registerEventImpl = function (eventId, targetId) {
            return OfficeExtension.Utility._createPromiseFromResult(null);
        };
        GenericEventRegistration.prototype._unregisterEventImpl = function (eventId, targetId) {
            return OfficeExtension.Utility._createPromiseFromResult(null);
        };
        GenericEventRegistration.prototype._handleRichApiMessage = function (msg) {
            if (msg && msg.entries) {
                for (var entryIndex = 0; entryIndex < msg.entries.length; entryIndex++) {
                    var entry = msg.entries[entryIndex];
                    if (entry.messageCategory == OfficeExtension.Constants.eventMessageCategory) {
                        if (OfficeExtension.Utility._logEnabled) {
                            OfficeExtension.Utility.log(JSON.stringify(entry));
                        }
                        var funcs = this.m_eventRegistration.getHandlers(entry.messageType, entry.targetId);
                        if (funcs.length > 0) {
                            var arg = JSON.parse(entry.message);
                            for (var i = 0; i < funcs.length; i++) {
                                funcs[i](arg);
                            }
                        }
                    }
                }
            }
        };
        GenericEventRegistration.getGenericEventRegistration = function () {
            if (!GenericEventRegistration.s_genericEventRegistration) {
                GenericEventRegistration.s_genericEventRegistration = new GenericEventRegistration();
            }
            return GenericEventRegistration.s_genericEventRegistration;
        };
        GenericEventRegistration.richApiMessageEventCategory = 65536;
        return GenericEventRegistration;
    }());
    function _testSetRichApiMessageReadyImpl(impl) {
        GenericEventRegistration._testReadyImpl = impl;
    }
    OfficeExtension._testSetRichApiMessageReadyImpl = _testSetRichApiMessageReadyImpl;
    function _testTriggerRichApiMessageEvent(msg) {
        GenericEventRegistration.getGenericEventRegistration()._handleRichApiMessage(msg);
    }
    OfficeExtension._testTriggerRichApiMessageEvent = _testTriggerRichApiMessageEvent;
    var GenericEventHandlers = (function (_super) {
        __extends(GenericEventHandlers, _super);
        function GenericEventHandlers(context, parentObject, name, eventInfo) {
            _super.call(this, context, parentObject, name, eventInfo);
            this.m_genericEventInfo = eventInfo;
        }
        GenericEventHandlers.prototype.add = function (handler) {
            var _this = this;
            if (this.m_genericEventInfo.registerFunc) {
                this.m_genericEventInfo.registerFunc();
            }
            if (!GenericEventRegistration.getGenericEventRegistration().isReady) {
                this._context._pendingRequest._addPreSyncPromise(GenericEventRegistration.getGenericEventRegistration().ready());
            }
            OfficeExtension.ActionFactory.createTraceMarkerForCallback(this._context, function () {
                _this._handlers.push(handler);
                if (_this._handlers.length == 1) {
                    GenericEventRegistration.getGenericEventRegistration().register(_this.m_genericEventInfo.eventType, _this.m_genericEventInfo.getTargetIdFunc(), _this._callback);
                }
            });
            return new OfficeExtension.EventHandlerResult(this._context, this, handler);
        };
        GenericEventHandlers.prototype.remove = function (handler) {
            var _this = this;
            if (this.m_genericEventInfo.unregisterFunc) {
                this.m_genericEventInfo.unregisterFunc();
            }
            OfficeExtension.ActionFactory.createTraceMarkerForCallback(this._context, function () {
                var handlers = _this._handlers;
                for (var index = handlers.length - 1; index >= 0; index--) {
                    if (handlers[index] === handler) {
                        handlers.splice(index, 1);
                        break;
                    }
                }
                if (handlers.length == 0) {
                    GenericEventRegistration.getGenericEventRegistration().unregister(_this.m_genericEventInfo.eventType, _this.m_genericEventInfo.getTargetIdFunc(), _this._callback);
                }
            });
        };
        GenericEventHandlers.prototype.removeAll = function () {
        };
        return GenericEventHandlers;
    }(OfficeExtension.EventHandlers));
    OfficeExtension.GenericEventHandlers = GenericEventHandlers;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var HttpRequestExecutor = (function () {
        function HttpRequestExecutor() {
        }
        HttpRequestExecutor.prototype.executeAsync = function (customData, requestFlags, requestMessage) {
            var requestMessageText = JSON.stringify(requestMessage.Body);
            var url = requestMessage.Url;
            if (url.charAt(url.length - 1) != "/") {
                url = url + "/";
            }
            url = url + OfficeExtension.Constants.processQuery;
            url = url + "?" + OfficeExtension.Constants.flags + "=" + requestFlags.toString();
            var requestInfo = {
                method: "POST",
                url: url,
                headers: {},
                body: requestMessageText
            };
            requestInfo.headers[OfficeExtension.Constants.sourceLibHeader] = HttpRequestExecutor.SourceLibHeaderValue;
            requestInfo.headers["CONTENT-TYPE"] = "application/json";
            if (requestMessage.Headers) {
                for (var key in requestMessage.Headers) {
                    requestInfo.headers[key] = requestMessage.Headers[key];
                }
            }
            return OfficeExtension.HttpUtility.sendRequest(requestInfo)
                .then(function (responseInfo) {
                var response;
                if (responseInfo.statusCode === 200) {
                    response = { ErrorCode: null, ErrorMessage: null, Headers: responseInfo.headers, Body: JSON.parse(responseInfo.body) };
                }
                else {
                    OfficeExtension.Utility.log("Error Response:" + responseInfo.body);
                    var error = OfficeExtension.Utility._parseErrorResponse(responseInfo);
                    response = {
                        ErrorCode: error.errorCode,
                        ErrorMessage: error.errorMessage,
                        Headers: responseInfo.headers,
                        Body: null
                    };
                }
                return response;
            });
        };
        HttpRequestExecutor.SourceLibHeaderValue = "officejs-rest";
        return HttpRequestExecutor;
    }());
    OfficeExtension.HttpRequestExecutor = HttpRequestExecutor;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var HttpUtility = (function () {
        function HttpUtility() {
        }
        HttpUtility.setCustomSendRequestFunc = function (func) {
            HttpUtility.s_customSendRequestFunc = func;
        };
        HttpUtility.xhrSendRequestFunc = function (request) {
            return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
                var xhr = new XMLHttpRequest();
                xhr.open(request.method, request.url);
                xhr.onload = function () {
                    var resp = {
                        statusCode: xhr.status,
                        headers: OfficeExtension.Utility._parseHttpResponseHeaders(xhr.getAllResponseHeaders()),
                        body: xhr.responseText
                    };
                    resolve(resp);
                };
                xhr.onerror = function () {
                    reject(new OfficeExtension._Internal.RuntimeError({
                        code: OfficeExtension.ErrorCodes.connectionFailure,
                        message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.connectionFailureWithStatus, xhr.statusText)
                    }));
                };
                if (request.headers) {
                    for (var key in request.headers) {
                        xhr.setRequestHeader(key, request.headers[key]);
                    }
                }
                xhr.send(request.body);
            });
        };
        HttpUtility.sendRequest = function (request) {
            HttpUtility.validateAndNormalizeRequest(request);
            var func = HttpUtility.s_customSendRequestFunc;
            if (!func) {
                func = HttpUtility.xhrSendRequestFunc;
            }
            return func(request);
        };
        HttpUtility.setCustomSendLocalDocumentRequestFunc = function (func) {
            HttpUtility.s_customSendLocalDocumentRequestFunc = func;
        };
        HttpUtility.sendLocalDocumentRequest = function (request) {
            HttpUtility.validateAndNormalizeRequest(request);
            var func;
            func = HttpUtility.s_customSendLocalDocumentRequestFunc || HttpUtility.officeJsSendLocalDocumentRequestFunc;
            return func(request);
        };
        HttpUtility.officeJsSendLocalDocumentRequestFunc = function (request) {
            request = OfficeExtension.Utility._validateLocalDocumentRequest(request);
            var requestSafeArray = OfficeExtension.Utility._buildRequestMessageSafeArray(request);
            return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
                OSF.DDA.RichApi.executeRichApiRequestAsync(requestSafeArray, function (asyncResult) {
                    var response;
                    if (asyncResult.status == "succeeded") {
                        response =
                            {
                                statusCode: OfficeExtension.RichApiMessageUtility.getResponseStatusCode(asyncResult),
                                headers: OfficeExtension.RichApiMessageUtility.getResponseHeaders(asyncResult),
                                body: OfficeExtension.RichApiMessageUtility.getResponseBody(asyncResult)
                            };
                    }
                    else {
                        response = OfficeExtension.RichApiMessageUtility.buildHttpResponseFromOfficeJsError(asyncResult.error.code, asyncResult.error.message);
                    }
                    OfficeExtension.Utility.log(JSON.stringify(response));
                    resolve(response);
                });
            });
        };
        HttpUtility.validateAndNormalizeRequest = function (request) {
            if (OfficeExtension.Utility.isNullOrUndefined(request)) {
                throw OfficeExtension._Internal.RuntimeError._createInvalidArgError({
                    argumentName: "request"
                });
            }
            if (OfficeExtension.Utility.isNullOrEmptyString(request.method)) {
                request.method = "GET";
            }
            request.method = request.method.toUpperCase();
        };
        HttpUtility.logRequest = function (request) {
            if (OfficeExtension.Utility._logEnabled) {
                OfficeExtension.Utility.log("---HTTP Request---");
                OfficeExtension.Utility.log(request.method + " " + request.url);
                if (request.headers) {
                    for (var key in request.headers) {
                        OfficeExtension.Utility.log(key + ": " + request.headers[key]);
                    }
                }
                if (HttpUtility._logBody) {
                    OfficeExtension.Utility.log(request.body);
                }
            }
        };
        HttpUtility.logResponse = function (response) {
            if (OfficeExtension.Utility._logEnabled) {
                OfficeExtension.Utility.log("---HTTP Response---");
                OfficeExtension.Utility.log("" + response.statusCode);
                if (response.headers) {
                    for (var key in response.headers) {
                        OfficeExtension.Utility.log(key + ": " + response.headers[key]);
                    }
                }
                if (HttpUtility._logBody) {
                    OfficeExtension.Utility.log(response.body);
                }
            }
        };
        HttpUtility._logBody = false;
        return HttpUtility;
    }());
    OfficeExtension.HttpUtility = HttpUtility;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var InstantiateActionResultHandler = (function () {
        function InstantiateActionResultHandler(clientObject) {
            this.m_clientObject = clientObject;
        }
        InstantiateActionResultHandler.prototype._handleResult = function (value) {
            this.m_clientObject._handleIdResult(value);
        };
        return InstantiateActionResultHandler;
    }());
    OfficeExtension.InstantiateActionResultHandler = InstantiateActionResultHandler;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    ;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var ObjectPath = (function () {
        function ObjectPath(objectPathInfo, parentObjectPath, isCollection, isInvalidAfterRequest) {
            this.m_objectPathInfo = objectPathInfo;
            this.m_parentObjectPath = parentObjectPath;
            this.m_isWriteOperation = false;
            this.m_isCollection = isCollection;
            this.m_isInvalidAfterRequest = isInvalidAfterRequest;
            this.m_isValid = true;
        }
        Object.defineProperty(ObjectPath.prototype, "objectPathInfo", {
            get: function () {
                return this.m_objectPathInfo;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ObjectPath.prototype, "isWriteOperation", {
            get: function () {
                return this.m_isWriteOperation;
            },
            set: function (value) {
                this.m_isWriteOperation = value;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ObjectPath.prototype, "isCollection", {
            get: function () {
                return this.m_isCollection;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ObjectPath.prototype, "isInvalidAfterRequest", {
            get: function () {
                return this.m_isInvalidAfterRequest;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ObjectPath.prototype, "parentObjectPath", {
            get: function () {
                return this.m_parentObjectPath;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ObjectPath.prototype, "argumentObjectPaths", {
            get: function () {
                return this.m_argumentObjectPaths;
            },
            set: function (value) {
                this.m_argumentObjectPaths = value;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ObjectPath.prototype, "isValid", {
            get: function () {
                return this.m_isValid;
            },
            set: function (value) {
                this.m_isValid = value;
                if (!value &&
                    this.m_objectPathInfo.ObjectPathType === 6 &&
                    this.m_savedObjectPathInfo) {
                    ObjectPath.copyObjectPathInfo(this.m_savedObjectPathInfo.pathInfo, this.m_objectPathInfo);
                    this.m_parentObjectPath = this.m_savedObjectPathInfo.parent;
                    this.m_isValid = true;
                    this.m_savedObjectPathInfo = null;
                }
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ObjectPath.prototype, "getByIdMethodName", {
            get: function () {
                return this.m_getByIdMethodName;
            },
            set: function (value) {
                this.m_getByIdMethodName = value;
            },
            enumerable: true,
            configurable: true
        });
        ObjectPath.prototype._updateAsNullObject = function () {
            this.m_isInvalidAfterRequest = false;
            this.m_isValid = true;
            this.m_objectPathInfo.ObjectPathType = 7;
            this.m_objectPathInfo.Name = "";
            this.m_objectPathInfo.ArgumentInfo = {};
            this.m_parentObjectPath = null;
            this.m_argumentObjectPaths = null;
        };
        ObjectPath.prototype.updateUsingObjectData = function (value) {
            var referenceId = value[OfficeExtension.Constants.referenceId];
            if (!OfficeExtension.Utility.isNullOrEmptyString(referenceId)) {
                if (!this.m_savedObjectPathInfo &&
                    !this.isInvalidAfterRequest &&
                    ObjectPath.isRestorableObjectPath(this.m_objectPathInfo.ObjectPathType)) {
                    var pathInfo = {};
                    ObjectPath.copyObjectPathInfo(this.m_objectPathInfo, pathInfo);
                    this.m_savedObjectPathInfo = {
                        pathInfo: pathInfo,
                        parent: this.m_parentObjectPath
                    };
                }
                this.m_isInvalidAfterRequest = false;
                this.m_isValid = true;
                this.m_objectPathInfo.ObjectPathType = 6;
                this.m_objectPathInfo.Name = referenceId;
                this.m_objectPathInfo.ArgumentInfo = {};
                this.m_parentObjectPath = null;
                this.m_argumentObjectPaths = null;
                return;
            }
            var parentIsCollection = this.parentObjectPath && this.parentObjectPath.isCollection;
            var getByIdMethodName = this.getByIdMethodName;
            if (parentIsCollection || !OfficeExtension.Utility.isNullOrEmptyString(getByIdMethodName)) {
                var id = value[OfficeExtension.Constants.id];
                if (OfficeExtension.Utility.isNullOrUndefined(id)) {
                    id = value[OfficeExtension.Constants.idPrivate];
                }
                if (!OfficeExtension.Utility.isNullOrUndefined(id)) {
                    this.m_isInvalidAfterRequest = false;
                    this.m_isValid = true;
                    if (!OfficeExtension.Utility.isNullOrEmptyString(getByIdMethodName)) {
                        this.m_objectPathInfo.ObjectPathType = 3;
                        this.m_objectPathInfo.Name = getByIdMethodName;
                        this.m_getByIdMethodName = null;
                    }
                    else {
                        this.m_objectPathInfo.ObjectPathType = 5;
                        this.m_objectPathInfo.Name = "";
                    }
                    this.isWriteOperation = false;
                    this.m_objectPathInfo.ArgumentInfo = {};
                    this.m_objectPathInfo.ArgumentInfo.Arguments = [id];
                    this.m_argumentObjectPaths = null;
                    return;
                }
            }
        };
        ObjectPath.isRestorableObjectPath = function (objectPathType) {
            return (objectPathType === 1 ||
                objectPathType === 5 ||
                objectPathType === 3 ||
                objectPathType === 4);
        };
        ObjectPath.copyObjectPathInfo = function (src, dest) {
            dest.Id = src.Id;
            dest.ArgumentInfo = src.ArgumentInfo;
            dest.Name = src.Name;
            dest.ObjectPathType = src.ObjectPathType;
            dest.ParentObjectPathId = src.ParentObjectPathId;
        };
        return ObjectPath;
    }());
    OfficeExtension.ObjectPath = ObjectPath;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var ObjectPathFactory = (function () {
        function ObjectPathFactory() {
        }
        ObjectPathFactory.createGlobalObjectObjectPath = function (context) {
            var objectPathInfo = { Id: context._nextId(), ObjectPathType: 1, Name: "" };
            return new OfficeExtension.ObjectPath(objectPathInfo, null, false, false);
        };
        ObjectPathFactory.createNewObjectObjectPath = function (context, typeName, isCollection) {
            var objectPathInfo = { Id: context._nextId(), ObjectPathType: 2, Name: typeName };
            return new OfficeExtension.ObjectPath(objectPathInfo, null, isCollection, false);
        };
        ObjectPathFactory.createPropertyObjectPath = function (context, parent, propertyName, isCollection, isInvalidAfterRequest) {
            var objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 4,
                Name: propertyName,
                ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
            };
            return new OfficeExtension.ObjectPath(objectPathInfo, parent._objectPath, isCollection, isInvalidAfterRequest);
        };
        ObjectPathFactory.createIndexerObjectPath = function (context, parent, args) {
            var objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 5,
                Name: "",
                ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            };
            objectPathInfo.ArgumentInfo.Arguments = args;
            return new OfficeExtension.ObjectPath(objectPathInfo, parent._objectPath, false, false);
        };
        ObjectPathFactory.createIndexerObjectPathUsingParentPath = function (context, parentObjectPath, args) {
            var objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 5,
                Name: "",
                ParentObjectPathId: parentObjectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            };
            objectPathInfo.ArgumentInfo.Arguments = args;
            return new OfficeExtension.ObjectPath(objectPathInfo, parentObjectPath, false, false);
        };
        ObjectPathFactory.createMethodObjectPath = function (context, parent, methodName, operationType, args, isCollection, isInvalidAfterRequest, getByIdMethodName) {
            var objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 3,
                Name: methodName,
                ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            };
            var argumentObjectPaths = OfficeExtension.Utility.setMethodArguments(context, objectPathInfo.ArgumentInfo, args);
            var ret = new OfficeExtension.ObjectPath(objectPathInfo, parent._objectPath, isCollection, isInvalidAfterRequest);
            ret.argumentObjectPaths = argumentObjectPaths;
            ret.isWriteOperation = (operationType != 1);
            ret.getByIdMethodName = getByIdMethodName;
            return ret;
        };
        ObjectPathFactory.createReferenceIdObjectPath = function (context, referenceId) {
            var objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 6,
                Name: referenceId,
                ArgumentInfo: {}
            };
            var ret = new OfficeExtension.ObjectPath(objectPathInfo, null, false, false);
            return ret;
        };
        ObjectPathFactory.createChildItemObjectPathUsingIndexerOrGetItemAt = function (hasIndexerMethod, context, parent, childItem, index) {
            var id = childItem[OfficeExtension.Constants.id];
            if (OfficeExtension.Utility.isNullOrUndefined(id)) {
                id = childItem[OfficeExtension.Constants.idPrivate];
            }
            if (hasIndexerMethod && !OfficeExtension.Utility.isNullOrUndefined(id)) {
                return ObjectPathFactory.createChildItemObjectPathUsingIndexer(context, parent, childItem);
            }
            else {
                return ObjectPathFactory.createChildItemObjectPathUsingGetItemAt(context, parent, childItem, index);
            }
        };
        ObjectPathFactory.createChildItemObjectPathUsingIndexer = function (context, parent, childItem) {
            var id = childItem[OfficeExtension.Constants.id];
            if (OfficeExtension.Utility.isNullOrUndefined(id)) {
                id = childItem[OfficeExtension.Constants.idPrivate];
            }
            var objectPathInfo = objectPathInfo =
                {
                    Id: context._nextId(),
                    ObjectPathType: 5,
                    Name: "",
                    ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
                    ArgumentInfo: {}
                };
            objectPathInfo.ArgumentInfo.Arguments = [id];
            return new OfficeExtension.ObjectPath(objectPathInfo, parent._objectPath, false, false);
        };
        ObjectPathFactory.createChildItemObjectPathUsingGetItemAt = function (context, parent, childItem, index) {
            var indexFromServer = childItem[OfficeExtension.Constants.index];
            if (indexFromServer) {
                index = indexFromServer;
            }
            var objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 3,
                Name: OfficeExtension.Constants.getItemAt,
                ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            };
            objectPathInfo.ArgumentInfo.Arguments = [index];
            return new OfficeExtension.ObjectPath(objectPathInfo, parent._objectPath, false, false);
        };
        return ObjectPathFactory;
    }());
    OfficeExtension.ObjectPathFactory = ObjectPathFactory;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var OfficeJsRequestExecutor = (function () {
        function OfficeJsRequestExecutor(context) {
            this.m_context = context;
        }
        OfficeJsRequestExecutor.prototype.executeAsync = function (customData, requestFlags, requestMessage) {
            var _this = this;
            var messageSafearray = OfficeExtension.RichApiMessageUtility.buildMessageArrayForIRequestExecutor(customData, requestFlags, requestMessage, OfficeJsRequestExecutor.SourceLibHeaderValue);
            return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
                OSF.DDA.RichApi.executeRichApiRequestAsync(messageSafearray, function (result) {
                    OfficeExtension.Utility.log("Response:");
                    OfficeExtension.Utility.log(JSON.stringify(result));
                    var response;
                    if (result.status == "succeeded") {
                        response = OfficeExtension.RichApiMessageUtility.buildResponseOnSuccess(OfficeExtension.RichApiMessageUtility.getResponseBody(result), OfficeExtension.RichApiMessageUtility.getResponseHeaders(result));
                    }
                    else {
                        response = OfficeExtension.RichApiMessageUtility.buildResponseOnError(result.error.code, result.error.message);
                        _this.m_context._processOfficeJsErrorResponse(result.error.code, response);
                    }
                    resolve(response);
                });
            });
        };
        OfficeJsRequestExecutor.SourceLibHeaderValue = "officejs";
        return OfficeJsRequestExecutor;
    }());
    OfficeExtension.OfficeJsRequestExecutor = OfficeJsRequestExecutor;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var _Internal;
    (function (_Internal) {
        var PromiseImpl;
        (function (PromiseImpl) {
            function Init() {
                return (function () {
                    "use strict";
                    function lib$es6$promise$utils$$objectOrFunction(x) {
                        return typeof x === 'function' || (typeof x === 'object' && x !== null);
                    }
                    function lib$es6$promise$utils$$isFunction(x) {
                        return typeof x === 'function';
                    }
                    function lib$es6$promise$utils$$isMaybeThenable(x) {
                        return typeof x === 'object' && x !== null;
                    }
                    var lib$es6$promise$utils$$_isArray;
                    if (!Array.isArray) {
                        lib$es6$promise$utils$$_isArray = function (x) {
                            return Object.prototype.toString.call(x) === '[object Array]';
                        };
                    }
                    else {
                        lib$es6$promise$utils$$_isArray = Array.isArray;
                    }
                    var lib$es6$promise$utils$$isArray = lib$es6$promise$utils$$_isArray;
                    var lib$es6$promise$asap$$len = 0;
                    var lib$es6$promise$asap$$toString = {}.toString;
                    var lib$es6$promise$asap$$vertxNext;
                    var lib$es6$promise$asap$$customSchedulerFn;
                    var lib$es6$promise$asap$$asap = function asap(callback, arg) {
                        lib$es6$promise$asap$$queue[lib$es6$promise$asap$$len] = callback;
                        lib$es6$promise$asap$$queue[lib$es6$promise$asap$$len + 1] = arg;
                        lib$es6$promise$asap$$len += 2;
                        if (lib$es6$promise$asap$$len === 2) {
                            if (lib$es6$promise$asap$$customSchedulerFn) {
                                lib$es6$promise$asap$$customSchedulerFn(lib$es6$promise$asap$$flush);
                            }
                            else {
                                lib$es6$promise$asap$$scheduleFlush();
                            }
                        }
                    };
                    function lib$es6$promise$asap$$setScheduler(scheduleFn) {
                        lib$es6$promise$asap$$customSchedulerFn = scheduleFn;
                    }
                    function lib$es6$promise$asap$$setAsap(asapFn) {
                        lib$es6$promise$asap$$asap = asapFn;
                    }
                    var lib$es6$promise$asap$$browserWindow = (typeof window !== 'undefined') ? window : undefined;
                    var lib$es6$promise$asap$$browserGlobal = lib$es6$promise$asap$$browserWindow || {};
                    var lib$es6$promise$asap$$BrowserMutationObserver = lib$es6$promise$asap$$browserGlobal.MutationObserver || lib$es6$promise$asap$$browserGlobal.WebKitMutationObserver;
                    var lib$es6$promise$asap$$isNode = typeof process !== 'undefined' && {}.toString.call(process) === '[object process]';
                    var lib$es6$promise$asap$$isWorker = typeof Uint8ClampedArray !== 'undefined' &&
                        typeof importScripts !== 'undefined' &&
                        typeof MessageChannel !== 'undefined';
                    function lib$es6$promise$asap$$useNextTick() {
                        var nextTick = process.nextTick;
                        var version = process.versions.node.match(/^(?:(\d+)\.)?(?:(\d+)\.)?(\*|\d+)$/);
                        if (Array.isArray(version) && version[1] === '0' && version[2] === '10') {
                            nextTick = setImmediate;
                        }
                        return function () {
                            nextTick(lib$es6$promise$asap$$flush);
                        };
                    }
                    function lib$es6$promise$asap$$useVertxTimer() {
                        return function () {
                            lib$es6$promise$asap$$vertxNext(lib$es6$promise$asap$$flush);
                        };
                    }
                    function lib$es6$promise$asap$$useMutationObserver() {
                        var iterations = 0;
                        var observer = new lib$es6$promise$asap$$BrowserMutationObserver(lib$es6$promise$asap$$flush);
                        var node = document.createTextNode('');
                        observer.observe(node, { characterData: true });
                        return function () {
                            node.data = (iterations = ++iterations % 2);
                        };
                    }
                    function lib$es6$promise$asap$$useMessageChannel() {
                        var channel = new MessageChannel();
                        channel.port1.onmessage = lib$es6$promise$asap$$flush;
                        return function () {
                            channel.port2.postMessage(0);
                        };
                    }
                    function lib$es6$promise$asap$$useSetTimeout() {
                        return function () {
                            setTimeout(lib$es6$promise$asap$$flush, 1);
                        };
                    }
                    var lib$es6$promise$asap$$queue = new Array(1000);
                    function lib$es6$promise$asap$$flush() {
                        for (var i = 0; i < lib$es6$promise$asap$$len; i += 2) {
                            var callback = lib$es6$promise$asap$$queue[i];
                            var arg = lib$es6$promise$asap$$queue[i + 1];
                            callback(arg);
                            lib$es6$promise$asap$$queue[i] = undefined;
                            lib$es6$promise$asap$$queue[i + 1] = undefined;
                        }
                        lib$es6$promise$asap$$len = 0;
                    }
                    var lib$es6$promise$asap$$scheduleFlush;
                    if (lib$es6$promise$asap$$isNode) {
                        lib$es6$promise$asap$$scheduleFlush = lib$es6$promise$asap$$useNextTick();
                    }
                    else if (lib$es6$promise$asap$$BrowserMutationObserver) {
                        lib$es6$promise$asap$$scheduleFlush = lib$es6$promise$asap$$useMutationObserver();
                    }
                    else if (lib$es6$promise$asap$$isWorker) {
                        lib$es6$promise$asap$$scheduleFlush = lib$es6$promise$asap$$useMessageChannel();
                    }
                    else {
                        lib$es6$promise$asap$$scheduleFlush = lib$es6$promise$asap$$useSetTimeout();
                    }
                    function lib$es6$promise$$internal$$noop() { }
                    var lib$es6$promise$$internal$$PENDING = void 0;
                    var lib$es6$promise$$internal$$FULFILLED = 1;
                    var lib$es6$promise$$internal$$REJECTED = 2;
                    var lib$es6$promise$$internal$$GET_THEN_ERROR = new lib$es6$promise$$internal$$ErrorObject();
                    function lib$es6$promise$$internal$$selfFullfillment() {
                        return new TypeError("You cannot resolve a promise with itself");
                    }
                    function lib$es6$promise$$internal$$cannotReturnOwn() {
                        return new TypeError('A promises callback cannot return that same promise.');
                    }
                    function lib$es6$promise$$internal$$getThen(promise) {
                        try {
                            return promise.then;
                        }
                        catch (error) {
                            lib$es6$promise$$internal$$GET_THEN_ERROR.error = error;
                            return lib$es6$promise$$internal$$GET_THEN_ERROR;
                        }
                    }
                    function lib$es6$promise$$internal$$tryThen(then, value, fulfillmentHandler, rejectionHandler) {
                        try {
                            then.call(value, fulfillmentHandler, rejectionHandler);
                        }
                        catch (e) {
                            return e;
                        }
                    }
                    function lib$es6$promise$$internal$$handleForeignThenable(promise, thenable, then) {
                        lib$es6$promise$asap$$asap(function (promise) {
                            var sealed = false;
                            var error = lib$es6$promise$$internal$$tryThen(then, thenable, function (value) {
                                if (sealed) {
                                    return;
                                }
                                sealed = true;
                                if (thenable !== value) {
                                    lib$es6$promise$$internal$$resolve(promise, value);
                                }
                                else {
                                    lib$es6$promise$$internal$$fulfill(promise, value);
                                }
                            }, function (reason) {
                                if (sealed) {
                                    return;
                                }
                                sealed = true;
                                lib$es6$promise$$internal$$reject(promise, reason);
                            }, 'Settle: ' + (promise._label || ' unknown promise'));
                            if (!sealed && error) {
                                sealed = true;
                                lib$es6$promise$$internal$$reject(promise, error);
                            }
                        }, promise);
                    }
                    function lib$es6$promise$$internal$$handleOwnThenable(promise, thenable) {
                        if (thenable._state === lib$es6$promise$$internal$$FULFILLED) {
                            lib$es6$promise$$internal$$fulfill(promise, thenable._result);
                        }
                        else if (thenable._state === lib$es6$promise$$internal$$REJECTED) {
                            lib$es6$promise$$internal$$reject(promise, thenable._result);
                        }
                        else {
                            lib$es6$promise$$internal$$subscribe(thenable, undefined, function (value) {
                                lib$es6$promise$$internal$$resolve(promise, value);
                            }, function (reason) {
                                lib$es6$promise$$internal$$reject(promise, reason);
                            });
                        }
                    }
                    function lib$es6$promise$$internal$$handleMaybeThenable(promise, maybeThenable) {
                        if (maybeThenable.constructor === promise.constructor) {
                            lib$es6$promise$$internal$$handleOwnThenable(promise, maybeThenable);
                        }
                        else {
                            var then = lib$es6$promise$$internal$$getThen(maybeThenable);
                            if (then === lib$es6$promise$$internal$$GET_THEN_ERROR) {
                                lib$es6$promise$$internal$$reject(promise, lib$es6$promise$$internal$$GET_THEN_ERROR.error);
                            }
                            else if (then === undefined) {
                                lib$es6$promise$$internal$$fulfill(promise, maybeThenable);
                            }
                            else if (lib$es6$promise$utils$$isFunction(then)) {
                                lib$es6$promise$$internal$$handleForeignThenable(promise, maybeThenable, then);
                            }
                            else {
                                lib$es6$promise$$internal$$fulfill(promise, maybeThenable);
                            }
                        }
                    }
                    function lib$es6$promise$$internal$$resolve(promise, value) {
                        if (promise === value) {
                            lib$es6$promise$$internal$$reject(promise, lib$es6$promise$$internal$$selfFullfillment());
                        }
                        else if (lib$es6$promise$utils$$objectOrFunction(value)) {
                            lib$es6$promise$$internal$$handleMaybeThenable(promise, value);
                        }
                        else {
                            lib$es6$promise$$internal$$fulfill(promise, value);
                        }
                    }
                    function lib$es6$promise$$internal$$publishRejection(promise) {
                        if (promise._onerror) {
                            promise._onerror(promise._result);
                        }
                        lib$es6$promise$$internal$$publish(promise);
                    }
                    function lib$es6$promise$$internal$$fulfill(promise, value) {
                        if (promise._state !== lib$es6$promise$$internal$$PENDING) {
                            return;
                        }
                        promise._result = value;
                        promise._state = lib$es6$promise$$internal$$FULFILLED;
                        if (promise._subscribers.length !== 0) {
                            lib$es6$promise$asap$$asap(lib$es6$promise$$internal$$publish, promise);
                        }
                    }
                    function lib$es6$promise$$internal$$reject(promise, reason) {
                        if (promise._state !== lib$es6$promise$$internal$$PENDING) {
                            return;
                        }
                        promise._state = lib$es6$promise$$internal$$REJECTED;
                        promise._result = reason;
                        lib$es6$promise$asap$$asap(lib$es6$promise$$internal$$publishRejection, promise);
                    }
                    function lib$es6$promise$$internal$$subscribe(parent, child, onFulfillment, onRejection) {
                        var subscribers = parent._subscribers;
                        var length = subscribers.length;
                        parent._onerror = null;
                        subscribers[length] = child;
                        subscribers[length + lib$es6$promise$$internal$$FULFILLED] = onFulfillment;
                        subscribers[length + lib$es6$promise$$internal$$REJECTED] = onRejection;
                        if (length === 0 && parent._state) {
                            lib$es6$promise$asap$$asap(lib$es6$promise$$internal$$publish, parent);
                        }
                    }
                    function lib$es6$promise$$internal$$publish(promise) {
                        var subscribers = promise._subscribers;
                        var settled = promise._state;
                        if (subscribers.length === 0) {
                            return;
                        }
                        var child, callback, detail = promise._result;
                        for (var i = 0; i < subscribers.length; i += 3) {
                            child = subscribers[i];
                            callback = subscribers[i + settled];
                            if (child) {
                                lib$es6$promise$$internal$$invokeCallback(settled, child, callback, detail);
                            }
                            else {
                                callback(detail);
                            }
                        }
                        promise._subscribers.length = 0;
                    }
                    function lib$es6$promise$$internal$$ErrorObject() {
                        this.error = null;
                    }
                    var lib$es6$promise$$internal$$TRY_CATCH_ERROR = new lib$es6$promise$$internal$$ErrorObject();
                    function lib$es6$promise$$internal$$tryCatch(callback, detail) {
                        try {
                            return callback(detail);
                        }
                        catch (e) {
                            lib$es6$promise$$internal$$TRY_CATCH_ERROR.error = e;
                            return lib$es6$promise$$internal$$TRY_CATCH_ERROR;
                        }
                    }
                    function lib$es6$promise$$internal$$invokeCallback(settled, promise, callback, detail) {
                        var hasCallback = lib$es6$promise$utils$$isFunction(callback), value, error, succeeded, failed;
                        if (hasCallback) {
                            value = lib$es6$promise$$internal$$tryCatch(callback, detail);
                            if (value === lib$es6$promise$$internal$$TRY_CATCH_ERROR) {
                                failed = true;
                                error = value.error;
                                value = null;
                            }
                            else {
                                succeeded = true;
                            }
                            if (promise === value) {
                                lib$es6$promise$$internal$$reject(promise, lib$es6$promise$$internal$$cannotReturnOwn());
                                return;
                            }
                        }
                        else {
                            value = detail;
                            succeeded = true;
                        }
                        if (promise._state !== lib$es6$promise$$internal$$PENDING) {
                        }
                        else if (hasCallback && succeeded) {
                            lib$es6$promise$$internal$$resolve(promise, value);
                        }
                        else if (failed) {
                            lib$es6$promise$$internal$$reject(promise, error);
                        }
                        else if (settled === lib$es6$promise$$internal$$FULFILLED) {
                            lib$es6$promise$$internal$$fulfill(promise, value);
                        }
                        else if (settled === lib$es6$promise$$internal$$REJECTED) {
                            lib$es6$promise$$internal$$reject(promise, value);
                        }
                    }
                    function lib$es6$promise$$internal$$initializePromise(promise, resolver) {
                        try {
                            resolver(function resolvePromise(value) {
                                lib$es6$promise$$internal$$resolve(promise, value);
                            }, function rejectPromise(reason) {
                                lib$es6$promise$$internal$$reject(promise, reason);
                            });
                        }
                        catch (e) {
                            lib$es6$promise$$internal$$reject(promise, e);
                        }
                    }
                    function lib$es6$promise$enumerator$$Enumerator(Constructor, input) {
                        var enumerator = this;
                        enumerator._instanceConstructor = Constructor;
                        enumerator.promise = new Constructor(lib$es6$promise$$internal$$noop);
                        if (enumerator._validateInput(input)) {
                            enumerator._input = input;
                            enumerator.length = input.length;
                            enumerator._remaining = input.length;
                            enumerator._init();
                            if (enumerator.length === 0) {
                                lib$es6$promise$$internal$$fulfill(enumerator.promise, enumerator._result);
                            }
                            else {
                                enumerator.length = enumerator.length || 0;
                                enumerator._enumerate();
                                if (enumerator._remaining === 0) {
                                    lib$es6$promise$$internal$$fulfill(enumerator.promise, enumerator._result);
                                }
                            }
                        }
                        else {
                            lib$es6$promise$$internal$$reject(enumerator.promise, enumerator._validationError());
                        }
                    }
                    lib$es6$promise$enumerator$$Enumerator.prototype._validateInput = function (input) {
                        return lib$es6$promise$utils$$isArray(input);
                    };
                    lib$es6$promise$enumerator$$Enumerator.prototype._validationError = function () {
                        return new _Internal.Error('Array Methods must be provided an Array');
                    };
                    lib$es6$promise$enumerator$$Enumerator.prototype._init = function () {
                        this._result = new Array(this.length);
                    };
                    var lib$es6$promise$enumerator$$default = lib$es6$promise$enumerator$$Enumerator;
                    lib$es6$promise$enumerator$$Enumerator.prototype._enumerate = function () {
                        var enumerator = this;
                        var length = enumerator.length;
                        var promise = enumerator.promise;
                        var input = enumerator._input;
                        for (var i = 0; promise._state === lib$es6$promise$$internal$$PENDING && i < length; i++) {
                            enumerator._eachEntry(input[i], i);
                        }
                    };
                    lib$es6$promise$enumerator$$Enumerator.prototype._eachEntry = function (entry, i) {
                        var enumerator = this;
                        var c = enumerator._instanceConstructor;
                        if (lib$es6$promise$utils$$isMaybeThenable(entry)) {
                            if (entry.constructor === c && entry._state !== lib$es6$promise$$internal$$PENDING) {
                                entry._onerror = null;
                                enumerator._settledAt(entry._state, i, entry._result);
                            }
                            else {
                                enumerator._willSettleAt(c.resolve(entry), i);
                            }
                        }
                        else {
                            enumerator._remaining--;
                            enumerator._result[i] = entry;
                        }
                    };
                    lib$es6$promise$enumerator$$Enumerator.prototype._settledAt = function (state, i, value) {
                        var enumerator = this;
                        var promise = enumerator.promise;
                        if (promise._state === lib$es6$promise$$internal$$PENDING) {
                            enumerator._remaining--;
                            if (state === lib$es6$promise$$internal$$REJECTED) {
                                lib$es6$promise$$internal$$reject(promise, value);
                            }
                            else {
                                enumerator._result[i] = value;
                            }
                        }
                        if (enumerator._remaining === 0) {
                            lib$es6$promise$$internal$$fulfill(promise, enumerator._result);
                        }
                    };
                    lib$es6$promise$enumerator$$Enumerator.prototype._willSettleAt = function (promise, i) {
                        var enumerator = this;
                        lib$es6$promise$$internal$$subscribe(promise, undefined, function (value) {
                            enumerator._settledAt(lib$es6$promise$$internal$$FULFILLED, i, value);
                        }, function (reason) {
                            enumerator._settledAt(lib$es6$promise$$internal$$REJECTED, i, reason);
                        });
                    };
                    function lib$es6$promise$promise$all$$all(entries) {
                        return new lib$es6$promise$enumerator$$default(this, entries).promise;
                    }
                    var lib$es6$promise$promise$all$$default = lib$es6$promise$promise$all$$all;
                    function lib$es6$promise$promise$race$$race(entries) {
                        var Constructor = this;
                        var promise = new Constructor(lib$es6$promise$$internal$$noop);
                        if (!lib$es6$promise$utils$$isArray(entries)) {
                            lib$es6$promise$$internal$$reject(promise, new TypeError('You must pass an array to race.'));
                            return promise;
                        }
                        var length = entries.length;
                        function onFulfillment(value) {
                            lib$es6$promise$$internal$$resolve(promise, value);
                        }
                        function onRejection(reason) {
                            lib$es6$promise$$internal$$reject(promise, reason);
                        }
                        for (var i = 0; promise._state === lib$es6$promise$$internal$$PENDING && i < length; i++) {
                            lib$es6$promise$$internal$$subscribe(Constructor.resolve(entries[i]), undefined, onFulfillment, onRejection);
                        }
                        return promise;
                    }
                    var lib$es6$promise$promise$race$$default = lib$es6$promise$promise$race$$race;
                    function lib$es6$promise$promise$resolve$$resolve(object) {
                        var Constructor = this;
                        if (object && typeof object === 'object' && object.constructor === Constructor) {
                            return object;
                        }
                        var promise = new Constructor(lib$es6$promise$$internal$$noop);
                        lib$es6$promise$$internal$$resolve(promise, object);
                        return promise;
                    }
                    var lib$es6$promise$promise$resolve$$default = lib$es6$promise$promise$resolve$$resolve;
                    function lib$es6$promise$promise$reject$$reject(reason) {
                        var Constructor = this;
                        var promise = new Constructor(lib$es6$promise$$internal$$noop);
                        lib$es6$promise$$internal$$reject(promise, reason);
                        return promise;
                    }
                    var lib$es6$promise$promise$reject$$default = lib$es6$promise$promise$reject$$reject;
                    var lib$es6$promise$promise$$counter = 0;
                    function lib$es6$promise$promise$$needsResolver() {
                        throw new TypeError('You must pass a resolver function as the first argument to the promise constructor');
                    }
                    function lib$es6$promise$promise$$needsNew() {
                        throw new TypeError("Failed to construct 'Promise': Please use the 'new' operator, this object constructor cannot be called as a function.");
                    }
                    var lib$es6$promise$promise$$default = lib$es6$promise$promise$$Promise;
                    function lib$es6$promise$promise$$Promise(resolver) {
                        this._id = lib$es6$promise$promise$$counter++;
                        this._state = undefined;
                        this._result = undefined;
                        this._subscribers = [];
                        if (lib$es6$promise$$internal$$noop !== resolver) {
                            if (!lib$es6$promise$utils$$isFunction(resolver)) {
                                lib$es6$promise$promise$$needsResolver();
                            }
                            if (!(this instanceof lib$es6$promise$promise$$Promise)) {
                                lib$es6$promise$promise$$needsNew();
                            }
                            lib$es6$promise$$internal$$initializePromise(this, resolver);
                        }
                    }
                    lib$es6$promise$promise$$Promise.all = lib$es6$promise$promise$all$$default;
                    lib$es6$promise$promise$$Promise.race = lib$es6$promise$promise$race$$default;
                    lib$es6$promise$promise$$Promise.resolve = lib$es6$promise$promise$resolve$$default;
                    lib$es6$promise$promise$$Promise.reject = lib$es6$promise$promise$reject$$default;
                    lib$es6$promise$promise$$Promise._setScheduler = lib$es6$promise$asap$$setScheduler;
                    lib$es6$promise$promise$$Promise._setAsap = lib$es6$promise$asap$$setAsap;
                    lib$es6$promise$promise$$Promise._asap = lib$es6$promise$asap$$asap;
                    lib$es6$promise$promise$$Promise.prototype = {
                        constructor: lib$es6$promise$promise$$Promise,
                        then: function (onFulfillment, onRejection) {
                            var parent = this;
                            var state = parent._state;
                            if (state === lib$es6$promise$$internal$$FULFILLED && !onFulfillment || state === lib$es6$promise$$internal$$REJECTED && !onRejection) {
                                return this;
                            }
                            var child = new this.constructor(lib$es6$promise$$internal$$noop);
                            var result = parent._result;
                            if (state) {
                                var callback = arguments[state - 1];
                                lib$es6$promise$asap$$asap(function () {
                                    lib$es6$promise$$internal$$invokeCallback(state, child, callback, result);
                                });
                            }
                            else {
                                lib$es6$promise$$internal$$subscribe(parent, child, onFulfillment, onRejection);
                            }
                            return child;
                        },
                        'catch': function (onRejection) {
                            return this.then(null, onRejection);
                        }
                    };
                    return lib$es6$promise$promise$$default;
                }).call(this);
            }
            PromiseImpl.Init = Init;
        })(PromiseImpl = _Internal.PromiseImpl || (_Internal.PromiseImpl = {}));
    })(_Internal = OfficeExtension._Internal || (OfficeExtension._Internal = {}));
    var _Internal;
    (function (_Internal) {
        function isEdgeLessThan14() {
            var userAgent = window.navigator.userAgent;
            var versionIdx = userAgent.indexOf("Edge/");
            if (versionIdx >= 0) {
                userAgent = userAgent.substring(versionIdx + 5, userAgent.length);
                if (userAgent < "14.14393")
                    return true;
                else
                    return false;
            }
            return false;
        }
        function determinePromise() {
            if (typeof (window) === "undefined" && typeof (Promise) === "function") {
                return Promise;
            }
            if (typeof (window) !== "undefined" && window.Promise) {
                if (isEdgeLessThan14()) {
                    return _Internal.PromiseImpl.Init();
                }
                else {
                    return window.Promise;
                }
            }
            else {
                return _Internal.PromiseImpl.Init();
            }
        }
        _Internal.OfficePromise = determinePromise();
    })(_Internal = OfficeExtension._Internal || (OfficeExtension._Internal = {}));
    var OfficePromise = _Internal.OfficePromise;
    OfficeExtension.Promise = OfficePromise;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var TrackedObjects = (function () {
        function TrackedObjects(context) {
            this._autoCleanupList = {};
            this.m_context = context;
        }
        TrackedObjects.prototype.add = function (param) {
            var _this = this;
            if (Array.isArray(param)) {
                param.forEach(function (item) { return _this._addCommon(item, true); });
            }
            else {
                this._addCommon(param, true);
            }
        };
        TrackedObjects.prototype._autoAdd = function (object) {
            this._addCommon(object, false);
            this._autoCleanupList[object._objectPath.objectPathInfo.Id] = object;
        };
        TrackedObjects.prototype._autoTrackIfNecessaryWhenHandleObjectResultValue = function (object, resultValue) {
            var shouldAutoTrack = (this.m_context._autoCleanup &&
                !object[OfficeExtension.Constants.isTracked] &&
                object !== this.m_context._rootObject &&
                resultValue &&
                !OfficeExtension.Utility.isNullOrEmptyString(resultValue[OfficeExtension.Constants.referenceId]));
            if (shouldAutoTrack) {
                this._autoCleanupList[object._objectPath.objectPathInfo.Id] = object;
                object[OfficeExtension.Constants.isTracked] = true;
            }
        };
        TrackedObjects.prototype._addCommon = function (object, isExplicitlyAdded) {
            if (object[OfficeExtension.Constants.isTracked]) {
                if (isExplicitlyAdded && this.m_context._autoCleanup) {
                    delete this._autoCleanupList[object._objectPath.objectPathInfo.Id];
                }
                return;
            }
            var referenceId = object[OfficeExtension.Constants.referenceId];
            if (OfficeExtension.Utility.isNullOrEmptyString(referenceId) && object._KeepReference) {
                object._KeepReference();
                OfficeExtension.ActionFactory.createInstantiateAction(this.m_context, object);
                if (isExplicitlyAdded && this.m_context._autoCleanup) {
                    delete this._autoCleanupList[object._objectPath.objectPathInfo.Id];
                }
                object[OfficeExtension.Constants.isTracked] = true;
            }
        };
        TrackedObjects.prototype.remove = function (param) {
            var _this = this;
            if (Array.isArray(param)) {
                param.forEach(function (item) { return _this._removeCommon(item); });
            }
            else {
                this._removeCommon(param);
            }
        };
        TrackedObjects.prototype._removeCommon = function (object) {
            var referenceId = object[OfficeExtension.Constants.referenceId];
            if (!OfficeExtension.Utility.isNullOrEmptyString(referenceId)) {
                var rootObject = this.m_context._rootObject;
                if (rootObject._RemoveReference) {
                    rootObject._RemoveReference(referenceId);
                }
                delete object[OfficeExtension.Constants.isTracked];
            }
        };
        TrackedObjects.prototype._retrieveAndClearAutoCleanupList = function () {
            var list = this._autoCleanupList;
            this._autoCleanupList = {};
            return list;
        };
        return TrackedObjects;
    }());
    OfficeExtension.TrackedObjects = TrackedObjects;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var ResourceStrings = (function () {
        function ResourceStrings() {
        }
        ResourceStrings.cannotRegisterEvent = "CannotRegisterEvent";
        ResourceStrings.connectionFailureWithStatus = "ConnectionFailureWithStatus";
        ResourceStrings.connectionFailureWithDetails = "ConnectionFailureWithDetails";
        ResourceStrings.invalidObjectPath = "InvalidObjectPath";
        ResourceStrings.invalidRequestContext = "InvalidRequestContext";
        ResourceStrings.invalidArgument = "InvalidArgument";
        ResourceStrings.invalidArgumentGeneric = "InvalidArgumentGeneric";
        ResourceStrings.propertyNotLoaded = "PropertyNotLoaded";
        ResourceStrings.runMustReturnPromise = "RunMustReturnPromise";
        ResourceStrings.timeout = "Timeout";
        ResourceStrings.propertyDoesNotExist = "PropertyDoesNotExist";
        ResourceStrings.attemptingToSetReadOnlyProperty = "AttemptingToSetReadOnlyProperty";
        ResourceStrings.moreInfoInnerError = "MoreInfoInnerError";
        ResourceStrings.cannotApplyPropertyThroughSetMethod = "CannotApplyPropertyThroughSetMethod";
        ResourceStrings.valueNotLoaded = "ValueNotLoaded";
        ResourceStrings.invalidOrTimedOutSessionMessage = "InvalidOrTimedOutSessionMessage";
        ResourceStrings.invalidOperationInCellEditMode = "InvalidOperationInCellEditMode";
        ResourceStrings.customFunctionDefintionMissing = "CustomFunctionDefintionMissing";
        ResourceStrings.customFunctionImplementationMissing = "CustomFunctionImplementationMissing";
        ResourceStrings.apiNotFoundDetails = "ApiNotFoundDetails";
        return ResourceStrings;
    }());
    OfficeExtension.ResourceStrings = ResourceStrings;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var ResourceStringValues = (function () {
        function ResourceStringValues() {
        }
        ResourceStringValues.CannotRegisterEvent = "The event handler cannot be registered.";
        ResourceStringValues.ConnectionFailureWithStatus = "The request failed with status code of {0}.";
        ResourceStringValues.ConnectionFailureWithDetails = "The request failed with status code of {0}, error code {1} and the following error message: {2}";
        ResourceStringValues.InvalidArgument = "The argument '{0}' doesn't work for this situation, is missing, or isn't in the right format.";
        ResourceStringValues.InvalidObjectPath = "The object path '{0}' isn't working for what you're trying to do. If you're using the object across multiple \"context.sync\" calls and outside the sequential execution of a \".run\" batch, please use the \"context.trackedObjects.add()\" and \"context.trackedObjects.remove()\" methods to manage the object's lifetime.";
        ResourceStringValues.InvalidRequestContext = "Cannot use the object across different request contexts.";
        ResourceStringValues.PropertyNotLoaded = "The property '{0}' is not available. Before reading the property's value, call the load method on the containing object and call \"context.sync()\" on the associated request context.";
        ResourceStringValues.RunMustReturnPromise = "The batch function passed to the \".run\" method didn't return a promise. The function must return a promise, so that any automatically-tracked objects can be released at the completion of the batch operation. Typically, you return a promise by returning the response from \"context.sync()\".";
        ResourceStringValues.Timeout = "The operation has timed out.";
        ResourceStringValues.ValueNotLoaded = "The value of the result object has not been loaded yet. Before reading the value property, call \"context.sync()\" on the associated request context.";
        ResourceStringValues.InvalidOrTimedOutSessionMessage = "Your Office Online session has expired or is invalid. To continue, refresh the page.";
        ResourceStringValues.InvalidOperationInCellEditMode = "Excel is in cell-editing mode. Please exit the edit mode by pressing ENTER or TAB or selecting another cell, and then try again.";
        ResourceStringValues.CustomFunctionDefintionMissing = "A property with this name that represents the function's definition must exist on Excel.CustomFunctions";
        ResourceStringValues.CustomFunctionImplementationMissing = "The property with this name on Excel.CustomFunctions that represents the function's definition must contain a 'call' property that implements the function.";
        ResourceStringValues.ApiNotFoundDetails = "The method or property {0} is part of the {1} requirement set, which is not available in your version of {2}.";
        return ResourceStringValues;
    }());
    OfficeExtension.ResourceStringValues = ResourceStringValues;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var RichApiMessageUtility = (function () {
        function RichApiMessageUtility() {
        }
        RichApiMessageUtility.buildMessageArrayForIRequestExecutor = function (customData, requestFlags, requestMessage, sourceLibHeaderValue) {
            var requestMessageText = JSON.stringify(requestMessage.Body);
            OfficeExtension.Utility.log("Request:");
            OfficeExtension.Utility.log(requestMessageText);
            var headers = {};
            headers[OfficeExtension.Constants.sourceLibHeader] = sourceLibHeaderValue;
            var messageSafearray = RichApiMessageUtility.buildRequestMessageSafeArray(customData, requestFlags, "POST", "ProcessQuery", headers, requestMessageText);
            return messageSafearray;
        };
        RichApiMessageUtility.buildResponseOnSuccess = function (responseBody, responseHeaders) {
            var response = { ErrorCode: '', ErrorMessage: '', Headers: null, Body: null };
            response.Body = JSON.parse(responseBody);
            response.Headers = responseHeaders;
            return response;
        };
        RichApiMessageUtility.buildResponseOnError = function (errorCode, message) {
            var response = { ErrorCode: '', ErrorMessage: '', Headers: null, Body: null };
            response.ErrorCode = OfficeExtension.ErrorCodes.generalException;
            response.ErrorMessage = message;
            if (errorCode == RichApiMessageUtility.OfficeJsErrorCode_ooeNoCapability) {
                response.ErrorCode = OfficeExtension.ErrorCodes.accessDenied;
            }
            else if (errorCode == RichApiMessageUtility.OfficeJsErrorCode_ooeActivityLimitReached) {
                response.ErrorCode = OfficeExtension.ErrorCodes.activityLimitReached;
            }
            else if (errorCode == RichApiMessageUtility.OfficeJsErrorCode_ooeInvalidOrTimedOutSession) {
                response.ErrorCode = OfficeExtension.ErrorCodes.invalidOrTimedOutSession;
                response.ErrorMessage = OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.invalidOrTimedOutSessionMessage);
            }
            return response;
        };
        RichApiMessageUtility.buildHttpResponseFromOfficeJsError = function (errorCode, message) {
            var statusCode = 500;
            var errorBody = {};
            errorBody["error"] = {};
            errorBody["error"]["code"] = OfficeExtension.ErrorCodes.generalException;
            errorBody["error"]["message"] = message;
            if (errorCode === RichApiMessageUtility.OfficeJsErrorCode_ooeNoCapability) {
                statusCode = 403;
                errorBody["error"]["code"] = OfficeExtension.ErrorCodes.accessDenied;
            }
            else if (errorCode === RichApiMessageUtility.OfficeJsErrorCode_ooeActivityLimitReached) {
                statusCode = 429;
                errorBody["error"]["code"] = OfficeExtension.ErrorCodes.activityLimitReached;
            }
            return { statusCode: statusCode, headers: {}, body: JSON.stringify(errorBody) };
        };
        RichApiMessageUtility.buildRequestMessageSafeArray = function (customData, requestFlags, method, path, headers, body) {
            var headerArray = [];
            if (headers) {
                for (var headerName in headers) {
                    headerArray.push(headerName);
                    headerArray.push(headers[headerName]);
                }
            }
            var appPermission = 0;
            var solutionId = "";
            var instanceId = "";
            var marketplaceType = "";
            return [
                customData,
                method,
                path,
                headerArray,
                body,
                appPermission,
                requestFlags,
                solutionId,
                instanceId,
                marketplaceType
            ];
        };
        RichApiMessageUtility.getResponseBody = function (result) {
            return RichApiMessageUtility.getResponseBodyFromSafeArray(result.value.data);
        };
        RichApiMessageUtility.getResponseHeaders = function (result) {
            return RichApiMessageUtility.getResponseHeadersFromSafeArray(result.value.data);
        };
        RichApiMessageUtility.getResponseBodyFromSafeArray = function (data) {
            var ret = data[2];
            if (typeof (ret) === "string") {
                return ret;
            }
            var arr = ret;
            return arr.join("");
        };
        RichApiMessageUtility.getResponseHeadersFromSafeArray = function (data) {
            var arrayHeader = data[1];
            if (!arrayHeader) {
                return null;
            }
            var headers = {};
            for (var i = 0; i < arrayHeader.length - 1; i += 2) {
                headers[arrayHeader[i]] = arrayHeader[i + 1];
            }
            return headers;
        };
        RichApiMessageUtility.getResponseStatusCode = function (result) {
            return RichApiMessageUtility.getResponseStatusCodeFromSafeArray(result.value.data);
        };
        RichApiMessageUtility.getResponseStatusCodeFromSafeArray = function (data) {
            return data[0];
        };
        RichApiMessageUtility.OfficeJsErrorCode_ooeInvalidOrTimedOutSession = 5012;
        RichApiMessageUtility.OfficeJsErrorCode_ooeActivityLimitReached = 5102;
        RichApiMessageUtility.OfficeJsErrorCode_ooeNoCapability = 7000;
        return RichApiMessageUtility;
    }());
    OfficeExtension.RichApiMessageUtility = RichApiMessageUtility;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var Utility = (function () {
        function Utility() {
        }
        Utility.checkArgumentNull = function (value, name) {
            if (Utility.isNullOrUndefined(value)) {
                throw OfficeExtension._Internal.RuntimeError._createInvalidArgError(name);
            }
        };
        Utility.isNullOrUndefined = function (value) {
            if (value === null) {
                return true;
            }
            if (typeof (value) === "undefined") {
                return true;
            }
            return false;
        };
        Utility.isUndefined = function (value) {
            if (typeof (value) === "undefined") {
                return true;
            }
            return false;
        };
        Utility.isNullOrEmptyString = function (value) {
            if (value === null) {
                return true;
            }
            if (typeof (value) === "undefined") {
                return true;
            }
            if (value.length == 0) {
                return true;
            }
            return false;
        };
        Utility.isPlainJsonObject = function (value) {
            if (Utility.isNullOrUndefined(value)) {
                return false;
            }
            if (typeof (value) !== "object") {
                return false;
            }
            return Object.getPrototypeOf(value) === Object.getPrototypeOf({});
        };
        Utility.trim = function (str) {
            return str.replace(new RegExp("^\\s+|\\s+$", "g"), "");
        };
        Utility.caseInsensitiveCompareString = function (str1, str2) {
            if (Utility.isNullOrUndefined(str1)) {
                return Utility.isNullOrUndefined(str2);
            }
            else {
                if (Utility.isNullOrUndefined(str2)) {
                    return false;
                }
                else {
                    return str1.toUpperCase() == str2.toUpperCase();
                }
            }
        };
        Utility.adjustToDateTime = function (value) {
            if (Utility.isNullOrUndefined(value)) {
                return null;
            }
            if (typeof (value) === "string") {
                return new Date(value);
            }
            if (Array.isArray(value)) {
                var arr = value;
                for (var i = 0; i < arr.length; i++) {
                    arr[i] = Utility.adjustToDateTime(arr[i]);
                }
                return arr;
            }
            throw OfficeExtension._Internal.RuntimeError._createInvalidArgError("date");
        };
        Utility.isReadonlyRestRequest = function (method) {
            return Utility.caseInsensitiveCompareString(method, "GET");
        };
        Utility.setMethodArguments = function (context, argumentInfo, args) {
            if (Utility.isNullOrUndefined(args)) {
                return null;
            }
            var referencedObjectPaths = new Array();
            var referencedObjectPathIds = new Array();
            var hasOne = Utility.collectObjectPathInfos(context, args, referencedObjectPaths, referencedObjectPathIds);
            argumentInfo.Arguments = args;
            if (hasOne) {
                argumentInfo.ReferencedObjectPathIds = referencedObjectPathIds;
                return referencedObjectPaths;
            }
            return null;
        };
        Utility.collectObjectPathInfos = function (context, args, referencedObjectPaths, referencedObjectPathIds) {
            var hasOne = false;
            for (var i = 0; i < args.length; i++) {
                if (args[i] instanceof OfficeExtension.ClientObject) {
                    var clientObject = args[i];
                    Utility.validateContext(context, clientObject);
                    args[i] = clientObject._objectPath.objectPathInfo.Id;
                    referencedObjectPathIds.push(clientObject._objectPath.objectPathInfo.Id);
                    referencedObjectPaths.push(clientObject._objectPath);
                    hasOne = true;
                }
                else if (Array.isArray(args[i])) {
                    var childArrayObjectPathIds = new Array();
                    var childArrayHasOne = Utility.collectObjectPathInfos(context, args[i], referencedObjectPaths, childArrayObjectPathIds);
                    if (childArrayHasOne) {
                        referencedObjectPathIds.push(childArrayObjectPathIds);
                        hasOne = true;
                    }
                    else {
                        referencedObjectPathIds.push(0);
                    }
                }
                else {
                    referencedObjectPathIds.push(0);
                }
            }
            return hasOne;
        };
        Utility.fixObjectPathIfNecessary = function (clientObject, value) {
            if (clientObject && clientObject._objectPath && value) {
                clientObject._objectPath.updateUsingObjectData(value);
            }
        };
        Utility.validateObjectPath = function (clientObject) {
            var objectPath = clientObject._objectPath;
            while (objectPath) {
                if (!objectPath.isValid) {
                    throw new OfficeExtension._Internal.RuntimeError({
                        code: OfficeExtension.ErrorCodes.invalidObjectPath,
                        message: Utility._getResourceString(OfficeExtension.ResourceStrings.invalidObjectPath, Utility.getObjectPathExpression(objectPath)),
                        debugInfo: {
                            errorLocation: Utility.getObjectPathExpression(objectPath)
                        }
                    });
                }
                objectPath = objectPath.parentObjectPath;
            }
        };
        Utility.validateReferencedObjectPaths = function (objectPaths) {
            if (objectPaths) {
                for (var i = 0; i < objectPaths.length; i++) {
                    var objectPath = objectPaths[i];
                    while (objectPath) {
                        if (!objectPath.isValid) {
                            throw new OfficeExtension._Internal.RuntimeError({
                                code: OfficeExtension.ErrorCodes.invalidObjectPath,
                                message: Utility._getResourceString(OfficeExtension.ResourceStrings.invalidObjectPath, Utility.getObjectPathExpression(objectPath))
                            });
                        }
                        objectPath = objectPath.parentObjectPath;
                    }
                }
            }
        };
        Utility.validateContext = function (context, obj) {
            if (obj && obj.context !== context) {
                throw new OfficeExtension._Internal.RuntimeError({
                    code: OfficeExtension.ErrorCodes.invalidRequestContext,
                    message: Utility._getResourceString(OfficeExtension.ResourceStrings.invalidRequestContext)
                });
            }
        };
        Utility.log = function (message) {
            if (Utility._logEnabled && typeof (console) !== "undefined" && console.log) {
                console.log(message);
            }
        };
        Utility.load = function (clientObj, option) {
            clientObj.context.load(clientObj, option);
        };
        Utility._parseSelectExpand = function (select) {
            var args = [];
            if (!Utility.isNullOrEmptyString(select)) {
                var propertyNames = select.split(",");
                for (var i = 0; i < propertyNames.length; i++) {
                    var propertyName = propertyNames[i];
                    propertyName = sanitizeForAnyItemsSlash(propertyName.trim());
                    if (propertyName.length > 0) {
                        args.push(propertyName);
                    }
                }
            }
            return args;
            function sanitizeForAnyItemsSlash(propertyName) {
                var propertyNameLower = propertyName.toLowerCase();
                if (propertyNameLower === "items" || propertyNameLower === "items/") {
                    return '*';
                }
                var itemsSlashLength = 6;
                if (propertyNameLower.substr(0, itemsSlashLength) === "items/") {
                    propertyName = propertyName.substr(itemsSlashLength);
                }
                return propertyName.replace(new RegExp("\/items\/", "gi"), "/");
            }
        };
        Utility.throwError = function (resourceId, arg, errorLocation) {
            throw new OfficeExtension._Internal.RuntimeError({
                code: resourceId,
                message: Utility._getResourceString(resourceId, arg),
                debugInfo: errorLocation ? { errorLocation: errorLocation } : undefined
            });
        };
        Utility.createRuntimeError = function (code, message, location) {
            return (new OfficeExtension._Internal.RuntimeError({
                code: code,
                message: message,
                debugInfo: { errorLocation: location }
            }));
        };
        Utility._getResourceString = function (resourceId, arg) {
            var ret;
            if (typeof (window) !== "undefined" && window.Strings && window.Strings.OfficeOM) {
                var stringName = "L_" + resourceId;
                var stringValue = window.Strings.OfficeOM[stringName];
                if (stringValue) {
                    ret = stringValue;
                }
            }
            if (!ret) {
                ret = OfficeExtension.ResourceStringValues[resourceId];
            }
            if (!ret) {
                ret = resourceId;
            }
            if (!Utility.isNullOrUndefined(arg)) {
                if (Array.isArray(arg)) {
                    var arrArg = arg;
                    ret = Utility._formatString(ret, arrArg);
                }
                else {
                    ret = ret.replace("{0}", arg);
                }
            }
            return ret;
        };
        Utility._formatString = function (format, arrArg) {
            return format.replace(/\{\d\}/g, function (v) {
                var position = parseInt(v.substr(1, v.length - 2));
                if (position < arrArg.length) {
                    return arrArg[position];
                }
                else {
                    throw OfficeExtension._Internal.RuntimeError._createInvalidArgError("format");
                }
            });
        };
        Utility.throwIfNotLoaded = function (propertyName, fieldValue, entityName, isNull) {
            if (!isNull && Utility.isUndefined(fieldValue) && propertyName.charCodeAt(0) != Utility.s_underscoreCharCode) {
                throw new OfficeExtension._Internal.RuntimeError({
                    code: OfficeExtension.ErrorCodes.propertyNotLoaded,
                    message: Utility._getResourceString(OfficeExtension.ResourceStrings.propertyNotLoaded, propertyName),
                    debugInfo: entityName ? { errorLocation: entityName + "." + propertyName } : undefined
                });
            }
        };
        Utility.throwIfApiNotSupported = function (apiFullName, apiSetName, apiSetVersion, hostName) {
            if (!Utility._doApiNotSupportedCheck) {
                return;
            }
            if (typeof (window) !== "undefined" && window.Office && window.Office.context) {
                if (!window.Office.context.requirements.isSetSupported(apiSetName, apiSetVersion)) {
                    var message = Utility._getResourceString(OfficeExtension.ResourceStrings.apiNotFoundDetails, [apiFullName, apiSetName + " " + apiSetVersion, hostName]);
                    throw new OfficeExtension._Internal.RuntimeError({
                        code: OfficeExtension.ErrorCodes.apiNotFound,
                        message: message,
                        debugInfo: { errorLocation: apiFullName }
                    });
                }
            }
        };
        Utility.getObjectPathExpression = function (objectPath) {
            var ret = "";
            while (objectPath) {
                switch (objectPath.objectPathInfo.ObjectPathType) {
                    case 1:
                        ret = ret;
                        break;
                    case 2:
                        ret = "new()" + (ret.length > 0 ? "." : "") + ret;
                        break;
                    case 3:
                        ret = Utility.normalizeName(objectPath.objectPathInfo.Name) + "()" + (ret.length > 0 ? "." : "") + ret;
                        break;
                    case 4:
                        ret = Utility.normalizeName(objectPath.objectPathInfo.Name) + (ret.length > 0 ? "." : "") + ret;
                        break;
                    case 5:
                        ret = "getItem()" + (ret.length > 0 ? "." : "") + ret;
                        break;
                    case 6:
                        ret = "_reference()" + (ret.length > 0 ? "." : "") + ret;
                        break;
                }
                objectPath = objectPath.parentObjectPath;
            }
            return ret;
        };
        Utility._createPromiseFromResult = function (value) {
            return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
                resolve(value);
            });
        };
        Utility._createTimeoutPromise = function (timeout) {
            return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
                setTimeout(function () {
                    resolve(null);
                }, timeout);
            });
        };
        Utility.promisify = function (action) {
            return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
                var callback = function (result) {
                    if (result.status == "failed") {
                        reject(result.error);
                    }
                    else {
                        resolve(result.value);
                    }
                };
                action(callback);
            });
        };
        Utility._addActionResultHandler = function (clientObj, action, resultHandler) {
            clientObj.context._pendingRequest.addActionResultHandler(action, resultHandler);
        };
        Utility._handleNavigationPropertyResults = function (clientObj, objectValue, propertyNames) {
            for (var i = 0; i < propertyNames.length - 1; i += 2) {
                if (!Utility.isUndefined(objectValue[propertyNames[i + 1]])) {
                    clientObj[propertyNames[i]]._handleResult(objectValue[propertyNames[i + 1]]);
                }
            }
        };
        Utility.normalizeName = function (name) {
            return name.substr(0, 1).toLowerCase() + name.substr(1);
        };
        Utility._isLocalDocumentUrl = function (url) {
            return Utility._getLocalDocumentUrlPrefixLength(url) > 0;
        };
        Utility._getLocalDocumentUrlPrefixLength = function (url) {
            var localDocumentPrefixes = ["http://document.localhost", "https://document.localhost", "//document.localhost"];
            var urlLower = url.toLowerCase().trim();
            for (var i = 0; i < localDocumentPrefixes.length; i++) {
                if (urlLower === localDocumentPrefixes[i]) {
                    return localDocumentPrefixes[i].length;
                }
                else if (urlLower.substr(0, localDocumentPrefixes[i].length + 1) === localDocumentPrefixes[i] + "/") {
                    return localDocumentPrefixes[i].length + 1;
                }
            }
            return 0;
        };
        Utility._validateLocalDocumentRequest = function (request) {
            var index = Utility._getLocalDocumentUrlPrefixLength(request.url);
            if (index <= 0) {
                throw OfficeExtension._Internal.RuntimeError._createInvalidArgError({
                    argumentName: "request"
                });
            }
            var path = request.url.substr(index);
            var pathLower = path.toLowerCase();
            if (pathLower === "_api") {
                path = "";
            }
            else if (pathLower.substr(0, "_api/".length) === "_api/") {
                path = path.substr("_api/".length);
            }
            return {
                method: request.method,
                url: path,
                headers: request.headers,
                body: request.body
            };
        };
        Utility._buildRequestMessageSafeArray = function (request) {
            var requestFlags = 0;
            if (!Utility.isReadonlyRestRequest(request.method)) {
                requestFlags = 1;
            }
            if (request.url.substr(0, OfficeExtension.Constants.processQuery.length).toLowerCase() === OfficeExtension.Constants.processQuery.toLowerCase()) {
                var index = request.url.indexOf("?");
                if (index > 0) {
                    var queryString = request.url.substr(index + 1);
                    var parts = queryString.split("&");
                    for (var i = 0; i < parts.length; i++) {
                        var keyvalue = parts[i].split("=");
                        if (keyvalue[0].toLowerCase() === OfficeExtension.Constants.flags) {
                            var flags = parseInt(keyvalue[1]);
                            requestFlags = flags;
                            requestFlags = requestFlags & 1;
                            break;
                        }
                    }
                }
            }
            return OfficeExtension.RichApiMessageUtility.buildRequestMessageSafeArray("", requestFlags, request.method, request.url, request.headers, request.body);
        };
        Utility._parseHttpResponseHeaders = function (allResponseHeaders) {
            var responseHeaders = {};
            if (!Utility.isNullOrEmptyString(allResponseHeaders)) {
                var regex = new RegExp("\r?\n");
                var entries = allResponseHeaders.split(regex);
                for (var i = 0; i < entries.length; i++) {
                    var entry = entries[i];
                    if (entry != null) {
                        var index = entry.indexOf(':');
                        if (index > 0) {
                            var key = entry.substr(0, index);
                            var value = entry.substr(index + 1);
                            key = Utility.trim(key);
                            value = Utility.trim(value);
                            responseHeaders[key.toUpperCase()] = value;
                        }
                    }
                }
            }
            return responseHeaders;
        };
        Utility._parseErrorResponse = function (responseInfo) {
            var errorObj = null;
            if (!Utility.isNullOrEmptyString(responseInfo.body)) {
                var errorResponseBody = Utility.trim(responseInfo.body);
                try {
                    errorObj = JSON.parse(errorResponseBody);
                }
                catch (e) {
                    Utility.log("Error when parse " + errorResponseBody);
                }
            }
            var errorMessage;
            var errorCode;
            if (!Utility.isNullOrUndefined(errorObj) && typeof (errorObj) === "object" && errorObj.error) {
                errorCode = errorObj.error.code;
                errorMessage = Utility._getResourceString(OfficeExtension.ResourceStrings.connectionFailureWithDetails, [responseInfo.statusCode.toString(), errorObj.error.code, errorObj.error.message]);
            }
            else {
                errorMessage = Utility._getResourceString(OfficeExtension.ResourceStrings.connectionFailureWithStatus, responseInfo.statusCode.toString());
            }
            if (Utility.isNullOrEmptyString(errorCode)) {
                errorCode = OfficeExtension.ErrorCodes.connectionFailure;
            }
            return { errorCode: errorCode, errorMessage: errorMessage };
        };
        Utility._copyHeaders = function (src, dest) {
            if (src && dest) {
                for (var key in src) {
                    dest[key] = src[key];
                }
            }
        };
        Utility._logEnabled = false;
        Utility._synchronousCleanup = false;
        Utility._doApiNotSupportedCheck = false;
        Utility.s_underscoreCharCode = "_".charCodeAt(0);
        return Utility;
    }());
    OfficeExtension.Utility = Utility;
})(OfficeExtension || (OfficeExtension = {}));
