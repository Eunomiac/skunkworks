/**
 * fixTypography.jsx
 *
 * Photoshop 26.4.1 ExtendScript utility for finding layers by flexible,
 * hard-coded criteria and running a sequential group of typography-oriented
 * actions against every matching layer.
 *
 * Configure LAYER_FILTER and ACTION_GROUP below, then run from Photoshop with:
 *   File > Scripts > Browse...
 */

/* --------------------------------------------------------------------------
 * CONFIGURATION: LAYER FILTER
 * --------------------------------------------------------------------------
 * The filter is intentionally broad and composable. Each section is optional:
 * - document: match general document properties.
 * - layer: match general layer qualities.
 * - text: match text-item properties. Any text matcher automatically excludes
 *   non-text layers unless text.required is false.
 * - custom: add arbitrary predicate functions for special project rules.
 *
 * Matching rules:
 * - enabled: turn the entire filter on/off.
 * - mode: "all" requires every configured matcher; "any" accepts one match.
 * - includeBackground: whether Photoshop background layers may be matched.
 * - includeHidden / includeLocked: whether those layers may be matched.
 * - includeGroups: whether layer sets themselves may be matched. Children of
 *   groups are still traversed either way.
 */
var LAYER_FILTER = {
    enabled: true,
    mode: "all", // "all" or "any"
    includeBackground: false,
    includeHidden: false,
    includeLocked: false,
    includeGroups: false,

    document: {
        name: null,              // string, RegExp, array, or matcher object
        colorMode: null,         // DocumentMode.RGB, DocumentMode.CMYK, etc.
        width: null,             // number or { min: 1000, max: 3000 }
        height: null,
        resolution: null
    },

    layer: {
        kind: LayerKind.TEXT,     // null for any kind
        name: null,              // example: /headline|title/i
        opacity: null,           // example: { min: 50, max: 100 }
        blendMode: null,         // example: BlendMode.NORMAL
        visible: null,           // true, false, or null
        bounds: {
            left: null,
            top: null,
            right: null,
            bottom: null,
            width: null,
            height: null
        },
        path: null               // full group path, e.g. /Desktop\/Hero/i
    },

    text: {
        required: true,
        contents: null,          // example: /\S/
        font: null,              // example: /Helvetica|Arial/i
        size: null,              // example: { min: 8, max: 72 }
        color: null,             // { r: 0, g: 0, b: 0 }, hex string, or fn
        justification: null,     // example: Justification.CENTER
        capitalization: null,    // best-effort Action Manager lookup
        antiAliasMethod: null,   // example: AntiAlias.SHARP
        direction: null          // Direction.HORIZONTAL or Direction.VERTICAL
    },

    custom: [
        // function (layer, context) { return layer.name.indexOf("TODO") !== -1; }
    ]
};

/* --------------------------------------------------------------------------
 * CONFIGURATION: ACTION GROUP
 * --------------------------------------------------------------------------
 * Actions run sequentially for each matching layer. Set enabled:false to skip
 * an action. Built-in action types are implemented below; use type:"custom"
 * with a run function for anything project-specific.
 */
var ACTION_GROUP = {
    name: "Typography cleanup",
    continueOnError: true,
    actions: [
        {
            type: "setTextFont",
            enabled: false,
            font: "ArialMT"
        },
        {
            type: "setTextSize",
            enabled: false,
            size: 24 // points
        },
        {
            type: "setTextColor",
            enabled: false,
            color: { r: 0, g: 0, b: 0 }
        },
        {
            type: "replaceText",
            enabled: false,
            find: /\s+/g,
            replace: " "
        },
        {
            type: "trimText",
            enabled: true
        },
        {
            type: "setJustification",
            enabled: false,
            justification: Justification.LEFT
        },
        {
            type: "runPhotoshopAction",
            enabled: false,
            action: "My Action",
            set: "My Action Set"
        },
        {
            type: "custom",
            enabled: false,
            run: function (layer, context) {
                // Example: layer.textItem.tracking = 0;
            }
        }
    ]
};

(function main() {
    if (app.documents.length === 0) {
        alert("No Photoshop document is open.");
        return;
    }

    var originalDialogs = app.displayDialogs;
    var originalActiveDocument = app.activeDocument;
    var originalActiveLayer = app.activeDocument.activeLayer;
    app.displayDialogs = DialogModes.NO;

    var report = {
        scanned: 0,
        matched: 0,
        actionRuns: 0,
        errors: []
    };

    try {
        var doc = app.activeDocument;
        var context = {
            document: doc,
            filter: LAYER_FILTER,
            actionGroup: ACTION_GROUP,
            report: report
        };

        visitLayers(doc.layers, "", context, function (layer, layerContext) {
            report.scanned++;
            if (matchesFilter(layer, layerContext)) {
                report.matched++;
                doc.activeLayer = layer;
                runActionGroup(layer, layerContext);
            }
        });

        alert(
            ACTION_GROUP.name + " complete.\n" +
            "Scanned layers: " + report.scanned + "\n" +
            "Matched layers: " + report.matched + "\n" +
            "Actions run: " + report.actionRuns + "\n" +
            "Errors: " + report.errors.length
        );
    } catch (err) {
        alert("fixTypography.jsx failed: " + describeError(err));
    } finally {
        try { originalActiveDocument.activeLayer = originalActiveLayer; } catch (restoreErr) {}
        app.displayDialogs = originalDialogs;
    }
}());

function visitLayers(layers, parentPath, context, visitor) {
    for (var i = layers.length - 1; i >= 0; i--) {
        var layer = layers[i];
        var path = parentPath ? parentPath + "/" + layer.name : layer.name;
        var layerContext = cloneContext(context);
        layerContext.path = path;
        layerContext.parentPath = parentPath;

        visitor(layer, layerContext);

        if (isLayerSet(layer)) {
            visitLayers(layer.layers, path, context, visitor);
        }
    }
}

function matchesFilter(layer, context) {
    var filter = context.filter;
    if (!filter || filter.enabled === false) {
        return true;
    }

    if (!filter.includeBackground && isBackgroundLayer(layer)) {
        return false;
    }
    if (!filter.includeHidden && layer.visible === false) {
        return false;
    }
    if (!filter.includeLocked && isLockedLayer(layer)) {
        return false;
    }
    if (!filter.includeGroups && isLayerSet(layer)) {
        return false;
    }

    var checks = [];
    addDocumentChecks(checks, filter.document, context.document);
    addLayerChecks(checks, filter.layer, layer, context);
    addTextChecks(checks, filter.text, layer, context);
    addCustomChecks(checks, filter.custom, layer, context);

    if (checks.length === 0) {
        return true;
    }

    var mode = String(filter.mode || "all").toLowerCase();
    for (var i = 0; i < checks.length; i++) {
        if (mode === "any" && checks[i]) {
            return true;
        }
        if (mode !== "any" && !checks[i]) {
            return false;
        }
    }
    return mode === "any" ? false : true;
}

function addDocumentChecks(checks, spec, doc) {
    if (!spec) { return; }
    pushMatcher(checks, spec.name, doc.name);
    pushMatcher(checks, spec.colorMode, doc.mode);
    pushMatcher(checks, spec.width, numericValue(doc.width));
    pushMatcher(checks, spec.height, numericValue(doc.height));
    pushMatcher(checks, spec.resolution, doc.resolution);
}

function addLayerChecks(checks, spec, layer, context) {
    if (!spec) { return; }
    pushMatcher(checks, spec.kind, layer.kind);
    pushMatcher(checks, spec.name, layer.name);
    pushMatcher(checks, spec.opacity, layer.opacity);
    pushMatcher(checks, spec.blendMode, layer.blendMode);
    pushMatcher(checks, spec.visible, layer.visible);
    pushMatcher(checks, spec.path, context.path);

    if (spec.bounds) {
        var b = getBounds(layer);
        pushMatcher(checks, spec.bounds.left, b.left);
        pushMatcher(checks, spec.bounds.top, b.top);
        pushMatcher(checks, spec.bounds.right, b.right);
        pushMatcher(checks, spec.bounds.bottom, b.bottom);
        pushMatcher(checks, spec.bounds.width, b.width);
        pushMatcher(checks, spec.bounds.height, b.height);
    }
}

function addTextChecks(checks, spec, layer, context) {
    if (!spec) { return; }
    var isText = isTextLayer(layer);
    if (spec.required !== false) {
        checks.push(isText);
    }
    if (!isText) { return; }

    var textItem = layer.textItem;
    pushMatcher(checks, spec.contents, textItem.contents);
    pushMatcher(checks, spec.font, textItem.font);
    pushMatcher(checks, spec.size, numericValue(textItem.size));
    if (spec.color !== null && typeof spec.color !== "undefined") {
        checks.push(matchColor(spec.color, textItem.color));
    }
    pushMatcher(checks, spec.justification, textItem.justification);
    pushMatcher(checks, spec.antiAliasMethod, textItem.antiAliasMethod);
    pushMatcher(checks, spec.direction, textItem.direction);

    if (spec.capitalization !== null && typeof spec.capitalization !== "undefined") {
        checks.push(matchValue(spec.capitalization, getTextCapitalization(context.document, layer)));
    }
}

function addCustomChecks(checks, predicates, layer, context) {
    if (!predicates || !predicates.length) { return; }
    for (var i = 0; i < predicates.length; i++) {
        if (typeof predicates[i] === "function") {
            checks.push(Boolean(predicates[i](layer, context)));
        }
    }
}

function runActionGroup(layer, context) {
    var group = context.actionGroup;
    if (!group || !group.actions) { return; }

    for (var i = 0; i < group.actions.length; i++) {
        var action = group.actions[i];
        if (!action || action.enabled === false) { continue; }

        try {
            runConfiguredAction(layer, context, action);
            context.report.actionRuns++;
        } catch (err) {
            context.report.errors.push({ layer: layer.name, action: action.type, error: describeError(err) });
            if (!group.continueOnError) {
                throw err;
            }
        }
    }
}

function runConfiguredAction(layer, context, action) {
    if (action.type === "setTextFont") {
        requireTextLayer(layer, action.type);
        layer.textItem.font = action.font;
    } else if (action.type === "setTextSize") {
        requireTextLayer(layer, action.type);
        layer.textItem.size = UnitValue(action.size, "pt");
    } else if (action.type === "setTextColor") {
        requireTextLayer(layer, action.type);
        layer.textItem.color = makeSolidColor(action.color);
    } else if (action.type === "replaceText") {
        requireTextLayer(layer, action.type);
        layer.textItem.contents = String(layer.textItem.contents).replace(action.find, action.replace);
    } else if (action.type === "trimText") {
        requireTextLayer(layer, action.type);
        layer.textItem.contents = trimString(layer.textItem.contents);
    } else if (action.type === "setJustification") {
        requireTextLayer(layer, action.type);
        layer.textItem.justification = action.justification;
    } else if (action.type === "runPhotoshopAction") {
        app.doAction(action.action, action.set);
    } else if (action.type === "custom" && typeof action.run === "function") {
        action.run(layer, context, action);
    } else {
        throw new Error("Unknown action type: " + action.type);
    }
}

function pushMatcher(checks, expected, actual) {
    if (expected === null || typeof expected === "undefined") { return; }
    checks.push(matchValue(expected, actual));
}

function matchValue(expected, actual) {
    if (typeof expected === "function") {
        return Boolean(expected(actual));
    }
    if (expected instanceof RegExp) {
        return expected.test(String(actual));
    }
    if (expected instanceof Array) {
        for (var i = 0; i < expected.length; i++) {
            if (matchValue(expected[i], actual)) { return true; }
        }
        return false;
    }
    if (typeof expected === "object") {
        if (isColorSpec(expected)) {
            return matchColor(expected, actual);
        }
        if (typeof expected.min !== "undefined" && numericValue(actual) < expected.min) {
            return false;
        }
        if (typeof expected.max !== "undefined" && numericValue(actual) > expected.max) {
            return false;
        }
        if (typeof expected.equals !== "undefined") {
            return matchValue(expected.equals, actual);
        }
        if (typeof expected.contains !== "undefined") {
            return String(actual).indexOf(String(expected.contains)) !== -1;
        }
        if (typeof expected.not !== "undefined") {
            return !matchValue(expected.not, actual);
        }
        return true;
    }
    return actual === expected;
}

function getBounds(layer) {
    var b = layer.bounds;
    var left = numericValue(b[0]);
    var top = numericValue(b[1]);
    var right = numericValue(b[2]);
    var bottom = numericValue(b[3]);
    return {
        left: left,
        top: top,
        right: right,
        bottom: bottom,
        width: right - left,
        height: bottom - top
    };
}

function isTextLayer(layer) {
    return !isLayerSet(layer) && layer.kind === LayerKind.TEXT;
}

function isLayerSet(layer) {
    return layer.typename === "LayerSet";
}

function isBackgroundLayer(layer) {
    try { return layer.isBackgroundLayer === true; } catch (err) { return false; }
}

function isLockedLayer(layer) {
    try { return layer.allLocked || layer.pixelsLocked || layer.positionLocked || layer.transparentPixelsLocked; }
    catch (err) { return false; }
}

function requireTextLayer(layer, actionType) {
    if (!isTextLayer(layer)) {
        throw new Error(actionType + " requires a text layer: " + layer.name);
    }
}

function makeSolidColor(spec) {
    if (spec && spec.typename === "SolidColor") { return spec; }
    var color = new SolidColor();
    if (typeof spec === "string") {
        spec = hexToRgb(spec);
    } else if (spec && typeof spec.hex !== "undefined") {
        spec = hexToRgb(spec.hex);
    }
    color.rgb.red = clampColor(spec.r);
    color.rgb.green = clampColor(spec.g);
    color.rgb.blue = clampColor(spec.b);
    return color;
}

function matchColor(expected, actual) {
    if (!actual || actual.typename !== "SolidColor") { return false; }
    var color = makeSolidColor(expected);
    return Math.round(actual.rgb.red) === Math.round(color.rgb.red) &&
        Math.round(actual.rgb.green) === Math.round(color.rgb.green) &&
        Math.round(actual.rgb.blue) === Math.round(color.rgb.blue);
}

function isColorSpec(value) {
    return value && (typeof value.r !== "undefined" || typeof value.hex !== "undefined" || value.typename === "SolidColor");
}

function hexToRgb(hex) {
    hex = String(hex).replace(/^#/, "");
    if (hex.length === 3) {
        hex = hex.charAt(0) + hex.charAt(0) + hex.charAt(1) + hex.charAt(1) + hex.charAt(2) + hex.charAt(2);
    }
    return {
        r: parseInt(hex.substr(0, 2), 16),
        g: parseInt(hex.substr(2, 2), 16),
        b: parseInt(hex.substr(4, 2), 16)
    };
}

function clampColor(value) {
    return Math.max(0, Math.min(255, Number(value)));
}

function numericValue(value) {
    if (value && value.as) {
        try { return value.as("px"); } catch (err) {}
        try { return value.as("pt"); } catch (errPt) {}
    }
    return Number(value);
}

function trimString(value) {
    return String(value).replace(/^\s+|\s+$/g, "");
}

function cloneContext(context) {
    var clone = {};
    for (var key in context) {
        if (context.hasOwnProperty(key)) {
            clone[key] = context[key];
        }
    }
    return clone;
}

function describeError(err) {
    if (!err) { return "Unknown error"; }
    return err.message ? err.message : String(err);
}

function getTextCapitalization(doc, layer) {
    // Photoshop exposes some character-style properties only through Action
    // Manager. This is best-effort so the rest of the filter remains portable.
    try {
        doc.activeLayer = layer;
        var ref = new ActionReference();
        ref.putProperty(charIDToTypeID("Prpr"), stringIDToTypeID("textKey"));
        ref.putEnumerated(charIDToTypeID("Lyr "), charIDToTypeID("Ordn"), charIDToTypeID("Trgt"));
        var desc = executeActionGet(ref).getObjectValue(stringIDToTypeID("textKey"));
        var styleRange = desc.getList(stringIDToTypeID("textStyleRange")).getObjectValue(0);
        var style = styleRange.getObjectValue(stringIDToTypeID("textStyle"));
        if (style.hasKey(stringIDToTypeID("fontCaps"))) {
            return typeIDToStringID(style.getEnumerationValue(stringIDToTypeID("fontCaps")));
        }
    } catch (err) {}
    return null;
}
