var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var OfficeCore;
(function (OfficeCore) {
    var _hostName = "OfficeCore";
    var _defaultApiSetName = "ExperimentApi";
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
    var FlightingService = (function (_super) {
        __extends(FlightingService, _super);
        function FlightingService() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(FlightingService.prototype, "_className", {
            get: function () {
                return "FlightingService";
            },
            enumerable: true,
            configurable: true
        });
        FlightingService.prototype.getClientSessionId = function () {
            var action = _createMethodAction(this.context, this, "GetClientSessionId", 0, []);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        FlightingService.prototype.getDeferredFlights = function () {
            var action = _createMethodAction(this.context, this, "GetDeferredFlights", 0, []);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        FlightingService.prototype.getFeature = function (featureName, type, defaultValue, possibleValues) {
            return new OfficeCore.ABType(this.context, _createMethodObjectPath(this.context, this, "GetFeature", 0, [featureName, type, defaultValue, possibleValues], false, false, null));
        };
        FlightingService.prototype.getFeatureGate = function (featureName, scope) {
            return new OfficeCore.ABType(this.context, _createMethodObjectPath(this.context, this, "GetFeatureGate", 0, [featureName, scope], false, false, null));
        };
        FlightingService.prototype.resetOverride = function (featureName) {
            _createMethodAction(this.context, this, "ResetOverride", 0, [featureName]);
        };
        FlightingService.prototype.setOverride = function (featureName, type, value) {
            _createMethodAction(this.context, this, "SetOverride", 0, [featureName, type, value]);
        };
        FlightingService.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
        };
        FlightingService.newObject = function (context) {
            var ret = new OfficeCore.FlightingService(context, _createNewObjectObjectPath(context, "Microsoft.Experiment.FlightingService", false));
            return ret;
        };
        FlightingService.prototype.toJSON = function () {
            return {};
        };
        return FlightingService;
    }(OfficeExtension.ClientObject));
    OfficeCore.FlightingService = FlightingService;
    var ABType = (function (_super) {
        __extends(ABType, _super);
        function ABType() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ABType.prototype, "_className", {
            get: function () {
                return "ABType";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ABType.prototype, "value", {
            get: function () {
                _throwIfNotLoaded("value", this.m_value, "ABType", this._isNull);
                return this.m_value;
            },
            enumerable: true,
            configurable: true
        });
        ABType.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Value"])) {
                this.m_value = obj["Value"];
            }
        };
        ABType.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ABType.prototype.toJSON = function () {
            return {
                "value": this.m_value
            };
        };
        return ABType;
    }(OfficeExtension.ClientObject));
    OfficeCore.ABType = ABType;
    var FeatureType;
    (function (FeatureType) {
        FeatureType.boolean = "Boolean";
        FeatureType.integer = "Integer";
        FeatureType.string = "String";
    })(FeatureType = OfficeCore.FeatureType || (OfficeCore.FeatureType = {}));
    var ExperimentErrorCodes;
    (function (ExperimentErrorCodes) {
        ExperimentErrorCodes.generalException = "GeneralException";
    })(ExperimentErrorCodes = OfficeCore.ExperimentErrorCodes || (OfficeCore.ExperimentErrorCodes = {}));
})(OfficeCore || (OfficeCore = {}));
var OfficeCore;
(function (OfficeCore) {
    var RequestContext = (function (_super) {
        __extends(RequestContext, _super);
        function RequestContext(url) {
            _super.call(this, url);
        }
        Object.defineProperty(RequestContext.prototype, "flightingService", {
            get: function () {
                if (!this.m_flightingService) {
                    this.m_flightingService = OfficeCore.FlightingService.newObject(this);
                }
                return this.m_flightingService;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RequestContext.prototype, "flighting", {
            get: function () {
                return this.flightingService;
            },
            enumerable: true,
            configurable: true
        });
        return RequestContext;
    }(OfficeExtension.ClientRequestContext));
    OfficeCore.RequestContext = RequestContext;
})(OfficeCore || (OfficeCore = {}));
