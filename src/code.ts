const CLIENT_STORAGE_CONFIG_KEY = "kernel-rest-config";
const PROJECT_SOURCES_PLUGIN_DATA_KEY = "kernel-project-sources";


type IndexLibraryCandidate = {
  slug: string;
  label: string;
  fileKey: string;
};

function extractSlugAndField(variableName: string): {
  slug: string;
  field: string;
} | null {
  const parts = variableName.split("/");

  if (parts.length !== 3) {
    return null;
  }

  if (parts[0] !== "ds") {
    return null;
  }

  if (!parts[1] || !parts[2]) {
    return null;
  }

  return {
    slug: parts[1],
    field: parts[2]
  };
}

async function getImportedStringVariableValueByKey(
  variableKey: string
): Promise<string> {
  const importedVariable =
    await figma.variables.importVariableByKeyAsync(variableKey);

  if (!importedVariable) {
    return "";
  }

  if (importedVariable.resolvedType !== "STRING") {
    return "";
  }

  const collection = await figma.variables.getVariableCollectionByIdAsync(
    importedVariable.variableCollectionId
  );

  if (!collection) {
    return "";
  }

  const defaultModeId = collection.defaultModeId;
  const rawValue = importedVariable.valuesByMode[defaultModeId];

  if (typeof rawValue !== "string") {
    return "";
  }

  return rawValue.trim();
}

async function getKernelIndexCandidates(): Promise<IndexLibraryCandidate[]> {
  const collections =
    await figma.teamLibrary.getAvailableLibraryVariableCollectionsAsync();

  console.log("[kernel-index] available collections", collections);

  const targetCollections = collections.filter(function (collection) {
    return collection.name === "Kernel Index";
  });

  console.log("[kernel-index] target collections", targetCollections);

  const grouped = new Map<
    string,
    {
      label: string;
      fileKey: string;
    }
  >();

  for (const collection of targetCollections) {
    const variables =
      await figma.teamLibrary.getVariablesInLibraryCollectionAsync(
        collection.key
      );

    console.log("[kernel-index] variables", {
      collectionName: collection.name,
      libraryName: collection.libraryName,
      variables: variables
    });

    for (const variable of variables) {
      const parsed = extractSlugAndField(variable.name);

      if (!parsed) {
        continue;
      }

      if (variable.resolvedType !== "STRING") {
        continue;
      }

      const value = await getImportedStringVariableValueByKey(variable.key);

      if (!value) {
        continue;
      }

      const current = grouped.get(parsed.slug) || {
        label: "",
        fileKey: ""
      };

      if (parsed.field === "label") {
        current.label = value;
      }

      if (parsed.field === "fileKey") {
        current.fileKey = value;
      }

      grouped.set(parsed.slug, current);
    }
  }

  const result: IndexLibraryCandidate[] = [];

  for (const entry of grouped.entries()) {
    const slug = entry[0];
    const value = entry[1];

    if (!value.label || !value.fileKey) {
      continue;
    }

    result.push({
      slug: slug,
      label: value.label,
      fileKey: value.fileKey
    });
  }

  console.log("[kernel-index] candidates", result);

  return result;
}

type StoredConfig = {
  serverBaseUrl: string;
  sessionToken: string;
  globalSources: AllowedSource[];
  projectSources: AllowedSource[];
};

type AllowedSource = {
  fileKey: string;
  pageName: string;
};

type UIToCodeMessage =
  | { type: "ui-ready" }
  | {
      type: "save-session-token";
      serverBaseUrl: string;
      sessionToken: string;
    }
  | {
      type: "run-lint";
      serverBaseUrl: string;
    }
  | {
    type: "save-sources";
    globalSources: AllowedSource[];
    projectSources: AllowedSource[];
  };

type CodeToUIMessage =
  | {
      type: "init";
      config: StoredConfig;
      candidates: IndexLibraryCandidate[];
    }
  | {
      type: "connection-status";
      connected: boolean;
    }
  | {
      type: "status";
      message: string;
    }
  | {
      type: "result";
      totalInstances: number;
      invalidCount: number;
      markerCount: number;
    }
  | {
      type: "error";
      message: string;
    };

type LintResult = {
  totalInstances: number;
  invalidCount: number;
  markerCount: number;
};

const STORAGE_KEY = "kernel-linter-rest-config";
const MARKER_NAME = "__kernel-linter-marker";
const UI_WIDTH = 360;
const UI_HEIGHT = 320;

// (async function () {
//   try {
//     const candidates = await getKernelIndexCandidates();

//     console.log("[kernel-index] candidates", candidates);
//   } catch (error) {
//     console.log("[kernel-index] failed", error);
//   }
// })();

figma.showUI(__html__, {
  width: UI_WIDTH,
  height: UI_HEIGHT,
  themeColors: true
});

void boot();

async function boot(): Promise<void> {
  var config = await loadConfig();
  var candidates = await getKernelIndexCandidates();

  console.log("[code] boot loadConfig", {
    serverBaseUrl: config.serverBaseUrl,
    hasSessionToken: !!config.sessionToken,
    sessionTokenLength: config.sessionToken ? config.sessionToken.length : 0
  });

  postToUI({
    type: "init",
    config: config,
    candidates: candidates
  });

  postToUI({
    type: "connection-status",
    connected: !!config.sessionToken
  });
}

figma.ui.onmessage = async function (msg: UIToCodeMessage) {
  console.log("[code] ui.onmessage", msg);
  try {
    if (msg.type === "ui-ready") {
      var initConfig = await loadConfig();
      var initCandidates = await getKernelIndexCandidates();

      postToUI({
        type: "init",
        config: initConfig,
        candidates: initCandidates
      });

      postToUI({
        type: "connection-status",
        connected: !!initConfig.sessionToken
      });

      return;
    }

    if (msg.type === "save-sources") {
      var currentConfigForSources = await loadConfig();

      await saveConfig({
        serverBaseUrl: currentConfigForSources.serverBaseUrl,
        sessionToken: currentConfigForSources.sessionToken,
        globalSources: normalizeAllowedSources(msg.globalSources),
        projectSources: normalizeAllowedSources(msg.projectSources)
      });

      var savedConfigForSources = await loadConfig();
      var savedCandidatesForSources = await getKernelIndexCandidates();

      console.log("[code] sources saved", {
        globalSources: savedConfigForSources.globalSources,
        projectSources: savedConfigForSources.projectSources
      });

      postToUI({
        type: "init",
        config: savedConfigForSources,
        candidates: savedCandidatesForSources
      });

      postToUI({
        type: "status",
        message: "Sources を保存しました。"
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
        sessionToken: msg.sessionToken,
        globalSources: currentConfigForSave.globalSources,
        projectSources: currentConfigForSave.projectSources
      });

      var savedConfig = await loadConfig();
      var savedCandidates = await getKernelIndexCandidates();
      console.log("[code] config after save", {
        serverBaseUrl: savedConfig.serverBaseUrl,
        hasSessionToken: !!savedConfig.sessionToken,
        sessionTokenLength: savedConfig.sessionToken ? savedConfig.sessionToken.length : 0
      });

      postToUI({
        type: "init",
        config: savedConfig,
        candidates: savedCandidates
      });

      postToUI({
        type: "connection-status",
        connected: !!savedConfig.sessionToken
      });

      postToUI({
        type: "status",
        message:
          "接続済みです。token保存: " +
          (savedConfig.sessionToken ? "OK" : "NG")
      });

      return;
    }

    if (msg.type === "run-lint") {
      var currentConfig = await loadConfig();
      var serverBaseUrl = normalizeBaseUrl(msg.serverBaseUrl);
      var sessionToken = currentConfig.sessionToken;
      var sources = mergeSources(
        currentConfig.globalSources,
        currentConfig.projectSources
      );

      console.log("[code] run-lint start", {
        inputServerBaseUrl: msg.serverBaseUrl,
        normalizedServerBaseUrl: serverBaseUrl,
        savedServerBaseUrl: currentConfig.serverBaseUrl,
        hasSessionToken: !!sessionToken,
        sessionTokenLength: sessionToken ? sessionToken.length : 0
      });

      console.log("[code] run-lint sources", {
        globalSourceCount: currentConfig.globalSources.length,
        projectSourceCount: currentConfig.projectSources.length,
        mergedSourceCount: sources.length,
        sources: sources
      });

      if (!serverBaseUrl) {
        throw new Error("Server Base URL を入力してください。");
      }

      if (!sessionToken) {
        throw new Error("未接続です。Connect を実行してください。");
      }

      const currentConfigForUpdate = await loadConfig();

      await saveConfig({
        serverBaseUrl: serverBaseUrl,
        sessionToken: sessionToken,
        globalSources: currentConfigForUpdate.globalSources,
        projectSources: currentConfigForUpdate.projectSources
      });

      postToUI({
        type: "status",
        message: "allowed components を取得しています..."
      });

      // var sources: AllowedSource[] = [
      //   {
      //     fileKey: "eV4Tr84CWhKWPzpciqMUg4",
      //     pageName: "__allowedComponentList"
      //   }
      // ];

      if (sources.length === 0) {
        figma.ui.postMessage({
          type: "error",
          message: "参照先ライブラリが未設定です。"
        });
        return;
      }

      var allowedKeys = await fetchAllowedKeys(
        serverBaseUrl,
        sessionToken,
        sources
      );

      postToUI({
        type: "status",
        message:
          "allowedKeys: " +
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
  } catch (error) {
    var message =
      error && error instanceof Error
        ? error.message
        : "不明なエラーが発生しました。";

    if (
      typeof message === "string" &&
      (
        message.indexOf("Invalid session") !== -1 ||
        message.indexOf("session not found") !== -1 ||
        message.indexOf("Unauthorized") !== -1
      )
    ) {
      var currentConfigForClear = await loadConfig();

      await saveConfig({
        serverBaseUrl: currentConfigForClear.serverBaseUrl,
        sessionToken: "",
        globalSources: currentConfigForClear.globalSources,
        projectSources: currentConfigForClear.projectSources
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

function mergeSources(
  globalSources: AllowedSource[],
  projectSources: AllowedSource[]
): AllowedSource[] {
  const merged = new Map<string, AllowedSource>();

  for (const source of globalSources) {
    const key = source.fileKey + "::" + source.pageName;
    merged.set(key, source);
  }

  for (const source of projectSources) {
    const key = source.fileKey + "::" + source.pageName;
    merged.set(key, source);
  }

  return Array.from(merged.values());
}

function postToUI(message: CodeToUIMessage): void {
  figma.ui.postMessage(message);
}

function normalizeAllowedSources(input: unknown): AllowedSource[] {
  if (!Array.isArray(input)) {
    return [];
  }

  const result: AllowedSource[] = [];

  for (const item of input) {
    if (!item || typeof item !== "object") {
      continue;
    }

    const fileKey =
      "fileKey" in item && typeof item.fileKey === "string"
        ? item.fileKey.trim()
        : "";

    const pageName =
      "pageName" in item && typeof item.pageName === "string"
        ? item.pageName.trim()
        : "";

    if (!fileKey || !pageName) {
      continue;
    }

    result.push({
      fileKey: fileKey,
      pageName: pageName
    });
  }

  return result;
}

function loadProjectSources(): AllowedSource[] {
  const raw = figma.root.getPluginData(PROJECT_SOURCES_PLUGIN_DATA_KEY);

  if (!raw) {
    return [];
  }

  try {
    const parsed = JSON.parse(raw);
    return normalizeAllowedSources(parsed);
  } catch (error) {
    console.log("[plugin] failed to parse projectSources", error);
    return [];
  }
}

function saveProjectSources(sources: AllowedSource[]): void {
  figma.root.setPluginData(
    PROJECT_SOURCES_PLUGIN_DATA_KEY,
    JSON.stringify(sources)
  );
}

async function loadConfig(): Promise<StoredConfig> {
  const raw = await figma.clientStorage.getAsync(CLIENT_STORAGE_CONFIG_KEY);

  let parsed: Partial<StoredConfig> = {};

  if (raw && typeof raw === "object") {
    parsed = raw as Partial<StoredConfig>;
  }

  const serverBaseUrl =
    typeof parsed.serverBaseUrl === "string" ? parsed.serverBaseUrl : "";

  const sessionToken =
    typeof parsed.sessionToken === "string" ? parsed.sessionToken : "";

  const globalSources = normalizeAllowedSources(parsed.globalSources);
  const projectSources = loadProjectSources();

  return {
    serverBaseUrl: serverBaseUrl,
    sessionToken: sessionToken,
    globalSources: globalSources,
    projectSources: projectSources
  };
}

async function saveConfig(config: StoredConfig): Promise<void> {
  await figma.clientStorage.setAsync(CLIENT_STORAGE_CONFIG_KEY, {
    serverBaseUrl: config.serverBaseUrl,
    sessionToken: config.sessionToken,
    globalSources: config.globalSources
  });

  saveProjectSources(config.projectSources);
}

function normalizeBaseUrl(value: string): string {
  return value.trim().replace(/\/+$/, "");
}

async function fetchAllowedKeys(
  serverBaseUrl: string,
  sessionToken: string,
  sources: AllowedSource[]
): Promise<Set<string>> {
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
  } catch (error) {
    throw new Error(
      "allowed-components の取得に失敗しました。server が起動しているか、URL が正しいか確認してください。"
    );
  }

  var json;
  try {
    json = await response.json();
    console.log("[code] fetchAllowedKeys response", {
  ok: response.ok,
  status: response.status,
  json: json
});
  } catch (error) {
    console.log("[code] catch", {
  error: error,
  message:
    error && error instanceof Error
      ? error.message
      : "不明なエラーが発生しました。"
});
    throw new Error("server のレスポンスが JSON ではありません。");
  }

  if (!response.ok || !json.ok) {
    throw new Error(
      json.error ||
        "allowed-components の取得に失敗しました。status=" +
          String(response.status)
    );
  }

  var allowedKeys = Array.isArray(json.allowedKeys) ? json.allowedKeys : [];
  var normalizedKeys: string[] = [];
  var i: number;

  for (i = 0; i < allowedKeys.length; i += 1) {
    var key = String(allowedKeys[i]).trim();
    if (key) {
      normalizedKeys.push(key);
    }
  }

  return new Set(normalizedKeys);
}

async function lintCurrentPage(allowedKeys: Set<string>): Promise<LintResult> {
  await figma.loadFontAsync({ family: "Inter", style: "Regular" });

  clearExistingMarkers(figma.currentPage);

  var instances = figma.currentPage.findAll(function (node) {
    return node.type === "INSTANCE";
  }) as InstanceNode[];

  var invalidCount = 0;
  var markerCount = 0;
  var i: number;

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

function clearExistingMarkers(root: PageNode): void {
  const nodes = root.findAll(function (node) {
    return "name" in node && node.name === MARKER_NAME;
  });

  var i: number;
  for (i = 0; i < nodes.length; i += 1) {
    nodes[i].remove();
  }
}

function createMarkerForInstance(instance: InstanceNode, key: string): boolean {
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
  } catch (error) {
    return false;
  }
}

function hexToRgb(hex: string): RGB {
  const normalized = hex.replace("#", "");
  const bigint = parseInt(normalized, 16);

  return {
    r: ((bigint >> 16) & 255) / 255,
    g: ((bigint >> 8) & 255) / 255,
    b: (bigint & 255) / 255
  };
}

async function getInstanceMainComponentKey(instance: InstanceNode): Promise<string | null> {
  try {
    var mainComponent = await instance.getMainComponentAsync();
    return mainComponent ? mainComponent.key : null;
  } catch (error) {
    return null;
  }
}