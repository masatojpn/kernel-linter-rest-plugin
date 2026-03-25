"use strict";
const STORAGE_KEY = "kernel-linter-rest-config";
const MARKER_NAME = "__kernel-linter-marker";
const UI_WIDTH = 360;
const UI_HEIGHT = 320;
figma.showUI(__html__, {
    width: UI_WIDTH,
    height: UI_HEIGHT,
    themeColors: true
});
void boot();
async function boot() {
    var config = await loadConfig();
    console.log("[code] boot loadConfig", {
        serverBaseUrl: config.serverBaseUrl,
        hasSessionToken: !!config.sessionToken,
        sessionTokenLength: config.sessionToken ? config.sessionToken.length : 0
    });
    postToUI({
        type: "init",
        config: config
    });
    postToUI({
        type: "connection-status",
        connected: !!config.sessionToken
    });
}
figma.ui.onmessage = async function (msg) {
    console.log("[code] ui.onmessage", msg);
    try {
        if (msg.type === "ui-ready") {
            var initConfig = await loadConfig();
            postToUI({
                type: "init",
                config: initConfig
            });
            postToUI({
                type: "connection-status",
                connected: !!initConfig.sessionToken
            });
            return;
        }
        if (msg.type === "save-session-token") {
            console.log("[code] save-session-token start", {
                serverBaseUrl: msg.serverBaseUrl,
                hasSessionToken: !!msg.sessionToken,
                sessionTokenLength: msg.sessionToken ? msg.sessionToken.length : 0
            });
            var currentConfigForSave = await loadConfig();
            console.log("[code] current config before save", {
                serverBaseUrl: currentConfigForSave.serverBaseUrl,
                hasSessionToken: !!currentConfigForSave.sessionToken
            });
            await saveConfig({
                serverBaseUrl: currentConfigForSave.serverBaseUrl,
                sessionToken: msg.sessionToken
            });
            var savedConfig = await loadConfig();
            console.log("[code] config after save", {
                serverBaseUrl: savedConfig.serverBaseUrl,
                hasSessionToken: !!savedConfig.sessionToken,
                sessionTokenLength: savedConfig.sessionToken ? savedConfig.sessionToken.length : 0
            });
            postToUI({
                type: "init",
                config: savedConfig
            });
            postToUI({
                type: "connection-status",
                connected: !!savedConfig.sessionToken
            });
            postToUI({
                type: "status",
                message: "接続済みです。token保存: " +
                    (savedConfig.sessionToken ? "OK" : "NG")
            });
            return;
        }
        if (msg.type === "run-lint") {
            var currentConfig = await loadConfig();
            var serverBaseUrl = normalizeBaseUrl(msg.serverBaseUrl);
            var sessionToken = currentConfig.sessionToken;
            console.log("[code] run-lint start", {
                inputServerBaseUrl: msg.serverBaseUrl,
                normalizedServerBaseUrl: serverBaseUrl,
                savedServerBaseUrl: currentConfig.serverBaseUrl,
                hasSessionToken: !!sessionToken,
                sessionTokenLength: sessionToken ? sessionToken.length : 0
            });
            if (!serverBaseUrl) {
                throw new Error("Server Base URL を入力してください。");
            }
            if (!sessionToken) {
                throw new Error("未接続です。Connect を実行してください。");
            }
            await saveConfig({
                serverBaseUrl: serverBaseUrl,
                sessionToken: sessionToken
            });
            postToUI({
                type: "status",
                message: "allowed components を取得しています..."
            });
            var sources = [
                {
                    fileKey: "eV4Tr84CWhKWPzpciqMUg4",
                    pageName: "__allowedComponentList"
                }
            ];
            var allowedKeys = await fetchAllowedKeys(serverBaseUrl, sessionToken, sources);
            postToUI({
                type: "status",
                message: "allowedKeys: " +
                    String(allowedKeys.size) +
                    " 件 / current page を走査しています..."
            });
            var result = await lintCurrentPage(allowedKeys);
            postToUI({
                type: "result",
                totalInstances: result.totalInstances,
                invalidCount: result.invalidCount,
                markerCount: result.markerCount
            });
            postToUI({
                type: "status",
                message: "完了しました。"
            });
            return;
        }
    }
    catch (error) {
        var message = error && error instanceof Error
            ? error.message
            : "不明なエラーが発生しました。";
        if (typeof message === "string" &&
            (message.indexOf("Invalid session") !== -1 ||
                message.indexOf("session not found") !== -1 ||
                message.indexOf("Unauthorized") !== -1)) {
            var currentConfigForClear = await loadConfig();
            await saveConfig({
                serverBaseUrl: currentConfigForClear.serverBaseUrl,
                sessionToken: ""
            });
            postToUI({
                type: "connection-status",
                connected: false
            });
            postToUI({
                type: "error",
                message: "接続が無効になりました。Connect をやり直してください。"
            });
            return;
        }
        postToUI({
            type: "error",
            message: message
        });
    }
};
function postToUI(message) {
    figma.ui.postMessage(message);
}
async function loadConfig() {
    var saved = await figma.clientStorage.getAsync(STORAGE_KEY);
    var serverBaseUrl = "https://kernel-linter-rest-server.vercel.app";
    var sessionToken = "";
    if (saved && typeof saved === "object") {
        if (typeof saved.serverBaseUrl === "string") {
            if (saved.serverBaseUrl.trim() !== "") {
                serverBaseUrl = saved.serverBaseUrl;
            }
        }
        if (typeof saved.sessionToken === "string") {
            sessionToken = saved.sessionToken;
        }
    }
    return {
        serverBaseUrl: serverBaseUrl,
        sessionToken: sessionToken
    };
}
async function saveConfig(config) {
    await figma.clientStorage.setAsync(STORAGE_KEY, config);
}
function normalizeBaseUrl(value) {
    return value.trim().replace(/\/+$/, "");
}
async function fetchAllowedKeys(serverBaseUrl, sessionToken, sources) {
    var url = serverBaseUrl + "/api/allowed-components";
    console.log("[code] fetchAllowedKeys request", {
        url: url,
        hasSessionToken: !!sessionToken,
        sessionTokenLength: sessionToken ? sessionToken.length : 0,
        sources: sources
    });
    var response;
    try {
        response = await fetch(url, {
            method: "POST",
            headers: {
                Authorization: "Bearer " + sessionToken,
                "Content-Type": "application/json"
            },
            body: JSON.stringify({
                sources: sources
            })
        });
    }
    catch (error) {
        throw new Error("allowed-components の取得に失敗しました。server が起動しているか、URL が正しいか確認してください。");
    }
    var json;
    try {
        json = await response.json();
        console.log("[code] fetchAllowedKeys response", {
            ok: response.ok,
            status: response.status,
            json: json
        });
    }
    catch (error) {
        console.log("[code] catch", {
            error: error,
            message: error && error instanceof Error
                ? error.message
                : "不明なエラーが発生しました。"
        });
        throw new Error("server のレスポンスが JSON ではありません。");
    }
    if (!response.ok || !json.ok) {
        throw new Error(json.error ||
            "allowed-components の取得に失敗しました。status=" +
                String(response.status));
    }
    var allowedKeys = Array.isArray(json.allowedKeys) ? json.allowedKeys : [];
    var normalizedKeys = [];
    var i;
    for (i = 0; i < allowedKeys.length; i += 1) {
        var key = String(allowedKeys[i]).trim();
        if (key) {
            normalizedKeys.push(key);
        }
    }
    return new Set(normalizedKeys);
}
async function lintCurrentPage(allowedKeys) {
    await figma.loadFontAsync({ family: "Inter", style: "Regular" });
    clearExistingMarkers(figma.currentPage);
    var instances = figma.currentPage.findAll(function (node) {
        return node.type === "INSTANCE";
    });
    var invalidCount = 0;
    var markerCount = 0;
    var i;
    for (i = 0; i < instances.length; i += 1) {
        var instance = instances[i];
        var key = await getInstanceMainComponentKey(instance);
        if (!key) {
            continue;
        }
        if (allowedKeys.has(key)) {
            continue;
        }
        invalidCount += 1;
        if (createMarkerForInstance(instance, key)) {
            markerCount += 1;
        }
    }
    return {
        totalInstances: instances.length,
        invalidCount: invalidCount,
        markerCount: markerCount
    };
}
function clearExistingMarkers(root) {
    const nodes = root.findAll(function (node) {
        return "name" in node && node.name === MARKER_NAME;
    });
    var i;
    for (i = 0; i < nodes.length; i += 1) {
        nodes[i].remove();
    }
}
function createMarkerForInstance(instance, key) {
    const parent = instance.parent;
    if (!parent || !("appendChild" in parent)) {
        return false;
    }
    const badge = figma.createFrame();
    badge.name = MARKER_NAME;
    badge.layoutMode = "HORIZONTAL";
    badge.primaryAxisSizingMode = "AUTO";
    badge.counterAxisSizingMode = "AUTO";
    badge.paddingLeft = 8;
    badge.paddingRight = 8;
    badge.paddingTop = 4;
    badge.paddingBottom = 4;
    badge.cornerRadius = 6;
    badge.fills = [{ type: "SOLID", color: hexToRgb("#E5484D") }];
    badge.strokes = [];
    badge.itemSpacing = 6;
    const text = figma.createText();
    text.name = MARKER_NAME + "-text";
    text.characters = "UNALLOWED";
    text.fontName = { family: "Inter", style: "Regular" };
    text.fontSize = 10;
    text.fills = [{ type: "SOLID", color: hexToRgb("#FFFFFF") }];
    badge.appendChild(text);
    const x = instance.x;
    const y = instance.y - 24;
    badge.x = x;
    badge.y = y;
    try {
        parent.appendChild(badge);
        badge.setPluginData("type", "kernel-linter-marker");
        badge.setPluginData("instanceKey", key);
        return true;
    }
    catch (error) {
        return false;
    }
}
function hexToRgb(hex) {
    const normalized = hex.replace("#", "");
    const bigint = parseInt(normalized, 16);
    return {
        r: ((bigint >> 16) & 255) / 255,
        g: ((bigint >> 8) & 255) / 255,
        b: (bigint & 255) / 255
    };
}
async function getInstanceMainComponentKey(instance) {
    try {
        var mainComponent = await instance.getMainComponentAsync();
        return mainComponent ? mainComponent.key : null;
    }
    catch (error) {
        return null;
    }
}
