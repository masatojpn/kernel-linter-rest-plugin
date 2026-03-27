const CLIENT_STORAGE_CONFIG_KEY = "kernel-rest-config";
const PROJECT_SOURCES_PLUGIN_DATA_KEY = "kernel-project-sources";
const TEST_PAGE_NAME = "__test";
const ALLOWED_PAGE_NAME = "__allowedComponentList";
const MARKER_ROOT_NAME = "__kernelLintMarkers";
const LABEL_FONT: FontName = { family: "Inter", style: "Regular" };

type IndexLibraryCandidate = {
  slug: string;
  label: string;
  fileKey: string;
};

type LocalAllowedRegistry = {
  allowedKeys: Set<string>;
  allowedSignatures: Set<string>;
};

type RemoteAllowedRegistry = {
  allowedKeys: Set<string>;
  allowedSignatures: Set<string>;
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
    type: "setup-pages";
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
      sessionToken: string;
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
      type: "setup-pages-result";
      testPageCreated: boolean;
      allowedPageCreated: boolean;
    }
  | {
      type: "error";
      message: string;
    };

type ViolationType =
  | "RAW_TEXT"
  | "NOT_REGISTERED"
  | "FILL_NEEDS_VARIABLE"
  | "STROKE_NEEDS_VARIABLE"
  | "WIDTH_NEEDS_VARIABLE"
  | "HEIGHT_NEEDS_VARIABLE"
  | "ITEM_SPACING_NEEDS_VARIABLE"
  | "COUNTER_AXIS_SPACING_NEEDS_VARIABLE"
  | "PADDING_TOP_NEEDS_VARIABLE"
  | "PADDING_RIGHT_NEEDS_VARIABLE"
  | "PADDING_BOTTOM_NEEDS_VARIABLE"
  | "PADDING_LEFT_NEEDS_VARIABLE"
  | "CORNER_RADIUS_NEEDS_VARIABLE";

type NumericField =
  | "width"
  | "height"
  | "itemSpacing"
  | "counterAxisSpacing"
  | "paddingTop"
  | "paddingRight"
  | "paddingBottom"
  | "paddingLeft"
  | "cornerRadius"
  | "topLeftRadius"
  | "topRightRadius"
  | "bottomRightRadius"
  | "bottomLeftRadius";

type Violation = {
  type: ViolationType;
  node: SceneNode;
  message: string;
};

type LintResult = {
  totalInstances: number;
  invalidCount: number;
  markerCount: number;
  violations: Violation[];
};

// const STORAGE_KEY = "kernel-linter-rest-config";
// const MARKER_NAME = "__kernel-linter-marker";
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

function findPageByName(name: string): PageNode | null {
  for (var i = 0; i < figma.root.children.length; i += 1) {
    var node = figma.root.children[i];

    if (node.type !== "PAGE") {
      continue;
    }

    if (node.name === name) {
      return node;
    }
  }

  return null;
}

function ensurePage(name: string): {
  page: PageNode;
  created: boolean;
} {
  var existingPage = findPageByName(name);

  if (existingPage) {
    return {
      page: existingPage,
      created: false
    };
  }

  var page = figma.createPage();
  page.name = name;

  return {
    page: page,
    created: true
  };
}

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
    candidates: candidates,
    sessionToken: config.sessionToken
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
        candidates: initCandidates,
        sessionToken: initConfig.sessionToken ? initConfig.sessionToken : ""
      });

      postToUI({
        type: "connection-status",
        connected: !!initConfig.sessionToken
      });

      return;
    }
    if (msg.type === "setup-pages") {
      var testPageResult = ensurePage(TEST_PAGE_NAME);
      var allowedPageResult = ensurePage(ALLOWED_PAGE_NAME);

      postToUI({
        type: "status",
        message: "Setup Pages 完了"
      });

      postToUI({
        type: "setup-pages-result",
        testPageCreated: testPageResult.created,
        allowedPageCreated: allowedPageResult.created
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
        candidates: savedCandidatesForSources,
        sessionToken: savedConfigForSources.sessionToken ? savedConfigForSources.sessionToken : ""
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
        candidates: savedCandidates,
        sessionToken: savedConfig.sessionToken ? savedConfig.sessionToken : ""
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

      var remoteAllowedRegistry =
        sources.length > 0
          ? await fetchAllowedRegistry(serverBaseUrl, sessionToken, sources)
          : {
              allowedKeys: new Set<string>(),
              allowedSignatures: new Set<string>()
            };

      var currentFileAllowedRegistry =
        await collectCurrentFileAllowedRegistry();

      var allowedKeys = mergeAllowedKeySets([
        remoteAllowedRegistry.allowedKeys,
        currentFileAllowedRegistry.allowedKeys
      ]);

      var allowedSignatures = mergeAllowedKeySets([
        remoteAllowedRegistry.allowedSignatures,
        currentFileAllowedRegistry.allowedSignatures
      ]);

      console.log("[code] allowedKeys detail", Array.from(allowedKeys));
      console.log(
        "[code] allowedSignatures detail",
        Array.from(allowedSignatures)
      );

      console.log("[code] allowed registry merged", {
        serverAllowedKeys: Array.from(remoteAllowedRegistry.allowedKeys),
        serverAllowedSignatures: Array.from(
          remoteAllowedRegistry.allowedSignatures
        ),
        currentFileAllowedKeys: Array.from(
          currentFileAllowedRegistry.allowedKeys
        ),
        currentFileAllowedSignatures: Array.from(
          currentFileAllowedRegistry.allowedSignatures
        ),
        allowedKeys: Array.from(allowedKeys),
        allowedSignatures: Array.from(allowedSignatures)
      });
      postToUI({
        type: "status",
        message:
          "allowedKeys: " +
          String(allowedKeys.size) +
          " 件 / current page を走査しています..."
      });

      var lintTargets: SceneNode[];
      var modeLabel: string;

      if (figma.currentPage.name === TEST_PAGE_NAME) {
        lintTargets = collectTestPageLintTargets();
        modeLabel = "Promote";
      } else {
        lintTargets = collectCurrentPageLintTargets();
        modeLabel = "Check";
      }

      postToUI({
        type: "status",
        message:
          "mode: " +
          modeLabel +
          "\n" +
          "allowedKeys: " +
          String(allowedKeys.size) +
          " 件 / lint 対象: " +
          String(lintTargets.length) +
          " 件"
      });

      var result = await lintSceneNodes(lintTargets, allowedKeys, allowedSignatures);

      if (figma.currentPage.name === TEST_PAGE_NAME) {
        if (result.invalidCount > 0) {
          postToUI({
            type: "status",
            message:
              "mode: Promote\n" +
              "違反があるため昇格しません。\n" +
              "invalid: " +
              String(result.invalidCount)
          });

          postToUI({
            type: "result",
            totalInstances: result.totalInstances,
            invalidCount: result.invalidCount,
            markerCount: result.markerCount
          });

          return;
        }

        var allowedPageResult = ensurePage(ALLOWED_PAGE_NAME);
        var promoteTargets =
          figma.currentPage.selection.length > 0
            ? collectPromoteTargetsFromSelection(figma.currentPage.selection)
            : collectPromoteTargetsFromPage(figma.currentPage);

        var promoteResult = await promoteNodesToAllowedPage(
          promoteTargets,
          allowedPageResult.page
        );

        postToUI({
          type: "status",
          message:
            "mode: Promote\n" +
            "昇格完了\n" +
            "moved: " +
            String(promoteResult.promotedCount) +
            "\n" +
            "replaced: " +
            String(promoteResult.replacedCount) +
            "\n" +
            "skipped: " +
            String(promoteResult.skippedCount)
        });

        postToUI({
          type: "result",
          totalInstances: result.totalInstances,
          invalidCount: result.invalidCount,
          markerCount: result.markerCount
        });

        return;
      }
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

async function fetchAllowedRegistry(
  serverBaseUrl: string,
  sessionToken: string,
  sources: AllowedSource[]
): Promise<RemoteAllowedRegistry> {
  var url = serverBaseUrl + "/api/allowed-components";

  var response = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: "Bearer " + sessionToken,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      sources: sources
    })
  });

  var json = await response.json();

  if (!response.ok || !json.ok) {
    throw new Error(
      json.error ||
        "allowed-components の取得に失敗しました。status=" +
          String(response.status)
    );
  }

  var allowedKeysRaw = Array.isArray(json.allowedKeys) ? json.allowedKeys : [];
  var allowedSignaturesRaw = Array.isArray(json.allowedSignatures)
    ? json.allowedSignatures
    : [];

  var allowedKeys = new Set<string>();
  var allowedSignatures = new Set<string>();
  var i: number;

  for (i = 0; i < allowedKeysRaw.length; i += 1) {
    var key = String(allowedKeysRaw[i]).trim();

    if (!key) {
      continue;
    }

    addAllowedKeyVariants(allowedKeys, key);
  }

  for (i = 0; i < allowedSignaturesRaw.length; i += 1) {
    var signature = String(allowedSignaturesRaw[i]).trim();

    if (!signature) {
      continue;
    }

    allowedSignatures.add(signature);
  }

  return {
    allowedKeys: allowedKeys,
    allowedSignatures: allowedSignatures
  };
}

function addAllowedKeyVariants(target: Set<string>, key: string): void {
  var normalized = key.trim();

  if (normalized === "") {
    return;
  }

  target.add(normalized);
  target.add("COMPONENT:" + normalized);
  target.add("COMPONENT_SET:" + normalized);
}

async function collectCurrentFileAllowedRegistry(): Promise<LocalAllowedRegistry> {
  var allowedKeys = new Set<string>();
  var allowedSignatures = new Set<string>();

  await figma.loadAllPagesAsync();

  for (var i = 0; i < figma.root.children.length; i += 1) {
    var page = figma.root.children[i];

    if (page.type !== "PAGE") {
      continue;
    }

    if (page.name !== ALLOWED_PAGE_NAME && page.name !== TEST_PAGE_NAME) {
      continue;
    }

    var nodes = page.findAll(function (node) {
      return (
        node.type === "INSTANCE" ||
        node.type === "COMPONENT" ||
        node.type === "COMPONENT_SET"
      );
    });

    for (var j = 0; j < nodes.length; j += 1) {
      var node = nodes[j];

      if (node.type === "COMPONENT_SET") {
        var componentSetKey = node.key ? node.key : node.id;
        allowedKeys.add(componentSetKey);
        allowedKeys.add("COMPONENT_SET:" + componentSetKey);
        continue;
      }

      if (node.type === "COMPONENT") {
        var componentCandidates = toAllowedKeyCandidatesFromComponent(node);
        var componentSignature = toLegacySignatureFromComponent(node);
        var k: number;

        for (k = 0; k < componentCandidates.length; k += 1) {
          allowedKeys.add(componentCandidates[k]);
        }

        if (componentSignature) {
          allowedSignatures.add(componentSignature);
        }

        continue;
      }

      if (node.type === "INSTANCE") {
        var instanceCandidates = await toAllowedKeyCandidatesFromInstance(node);
        var instanceSignature = await toLegacySignatureFromInstance(node);
        var m: number;

        for (m = 0; m < instanceCandidates.length; m += 1) {
          allowedKeys.add(instanceCandidates[m]);
        }

        if (instanceSignature) {
          allowedSignatures.add(instanceSignature);
        }
      }
    }
  }

  return {
    allowedKeys: allowedKeys,
    allowedSignatures: allowedSignatures
  };
}

function mergeAllowedKeySets(sets: Array<Set<string>>): Set<string> {
  var merged = new Set<string>();

  for (var i = 0; i < sets.length; i += 1) {
    var set = sets[i];

    set.forEach(function (value) {
      merged.add(value);
    });
  }

  return merged;
}

function collectCurrentPageLintTargets(): SceneNode[] {
  return figma.currentPage.findAll(function (node) {
    if (!node.visible) {
      return false;
    }

    if (isInsideMarkerLayer(node)) {
      return false;
    }

    return true;
  }) as SceneNode[];
}

function toAllowedKeyCandidatesFromComponent(component: ComponentNode): string[] {
  var result: string[] = [];
  var componentKey = component.key ? component.key : component.id;

  result.push(componentKey);
  result.push("COMPONENT:" + componentKey);

  var parent = component.parent;

  if (parent && parent.type === "COMPONENT_SET") {
    var componentSetKey = parent.key ? parent.key : parent.id;

    result.push(componentSetKey);
    result.push("COMPONENT_SET:" + componentSetKey);
  }

  return result;
}

async function toAllowedKeyCandidatesFromInstance(
  instance: InstanceNode
): Promise<string[]> {
  var mainComponent = await instance.getMainComponentAsync();

  if (!mainComponent) {
    return [];
  }

  return toAllowedKeyCandidatesFromComponent(mainComponent);
}

function hasAnyAllowedKey(
  allowedKeys: Set<string>,
  candidates: string[]
): boolean {
  var i: number;

  for (i = 0; i < candidates.length; i += 1) {
    if (allowedKeys.has(candidates[i])) {
      return true;
    }
  }

  return false;
}

function collectTestPageLintTargets(): SceneNode[] {
  var selection = figma.currentPage.selection;

  if (selection.length > 0) {
    var selectedTargets: SceneNode[] = [];
    var i: number;

    for (i = 0; i < selection.length; i += 1) {
      var node = selection[i];

      if (!node.visible) {
        continue;
      }

      if (isInsideMarkerLayer(node)) {
        continue;
      }

      selectedTargets.push(node);
    }

    return selectedTargets;
  }

  return figma.currentPage.findAll(function (node) {
    if (!node.visible) {
      return false;
    }

    if (isInsideMarkerLayer(node)) {
      return false;
    }

    return true;
  }) as SceneNode[];
}

async function lintSceneNodes(
  nodes: SceneNode[],
  allowedKeys: Set<string>,
  allowedSignatures: Set<string>
): Promise<LintResult> {
  await figma.loadFontAsync(LABEL_FONT);

  clearExistingMarkers(figma.currentPage);

  var totalInstances = 0;
  var violations: Violation[] = [];
  var i: number;

  for (i = 0; i < nodes.length; i += 1) {
    var node = nodes[i];

    if (!node.visible) {
      continue;
    }

    if (isInsideMarkerLayer(node)) {
      continue;
    }

    if (node.type === "TEXT") {
      var shouldFlagRawText =
        !hasAncestorType(node, "INSTANCE") &&
        node.characters.trim().length > 0;

      if (shouldFlagRawText) {
        violations.push({
          type: "RAW_TEXT",
          node: node,
          message: "RAW TEXT"
        });
      }
    }

    if (node.type === "INSTANCE") {
      totalInstances += 1;

      var keyCandidates = await toAllowedKeyCandidatesFromInstance(node);
      var keyMatched = hasAnyAllowedKey(allowedKeys, keyCandidates);

      var signature = await toLegacySignatureFromInstance(node);
      var signatureMatched = false;

      if (signature && allowedSignatures.has(signature)) {
        signatureMatched = true;
      }

      var isAllowed = keyMatched || signatureMatched;

      if (node.name === "Text") {
        console.log("[code] text instance check", {
          nodeId: node.id,
          nodeName: node.name,
          keyCandidates: keyCandidates,
          signature: signature,
          keyMatched: keyMatched,
          signatureMatched: signatureMatched,
          hasAllowed: isAllowed
        });
      }

      if (!isAllowed) {
        violations.push({
          type: "NOT_REGISTERED",
          node: node,
          message: "NOT REGISTERED"
        });
      }
    }

    var colorViolations = await checkColorVariableViolations(node);
    violations = violations.concat(colorViolations);

    var numericViolations = checkNumericVariableViolations(node);
    violations = violations.concat(numericViolations);
  }

  var unique = uniqueViolations(violations);
  var markerCount = await renderMarkers(figma.currentPage, unique);

  return {
    totalInstances: totalInstances,
    invalidCount: unique.length,
    markerCount: markerCount,
    violations: unique
  };
}

function checkCornerRadiusVariableViolation(node: SceneNode): Violation | null {
  if (!shouldCheckNumericForNode(node)) {
    return null;
  }

  const anyNode = node as any;

  const hasCornerRadius = typeof anyNode.cornerRadius === "number";
  const hasIndividualRadius =
    typeof anyNode.topLeftRadius === "number" ||
    typeof anyNode.topRightRadius === "number" ||
    typeof anyNode.bottomRightRadius === "number" ||
    typeof anyNode.bottomLeftRadius === "number";

  if (!hasCornerRadius && !hasIndividualRadius) {
    return null;
  }

  const cornerRadius =
    typeof anyNode.cornerRadius === "number"
      ? anyNode.cornerRadius
      : undefined;

  const corners = [
    {
      field: "topLeftRadius",
      value:
        typeof anyNode.topLeftRadius === "number"
          ? anyNode.topLeftRadius
          : undefined,
    },
    {
      field: "topRightRadius",
      value:
        typeof anyNode.topRightRadius === "number"
          ? anyNode.topRightRadius
          : undefined,
    },
    {
      field: "bottomRightRadius",
      value:
        typeof anyNode.bottomRightRadius === "number"
          ? anyNode.bottomRightRadius
          : undefined,
    },
    {
      field: "bottomLeftRadius",
      value:
        typeof anyNode.bottomLeftRadius === "number"
          ? anyNode.bottomLeftRadius
          : undefined,
    },
  ] as const;

  const allValues = [
    cornerRadius,
    corners[0].value,
    corners[1].value,
    corners[2].value,
    corners[3].value,
  ].filter(function (value): value is number {
    return typeof value === "number" && Number.isFinite(value);
  });

  if (allValues.length === 0) {
    return null;
  }

  const hasNonZeroRadius = allValues.some(function (value) {
    return value !== 0;
  });

  if (!hasNonZeroRadius) {
    return null;
  }

  if (hasBoundVariableForField(anyNode, "cornerRadius")) {
    return null;
  }

  for (let i = 0; i < corners.length; i += 1) {
    const corner = corners[i];

    if (typeof corner.value !== "number") {
      continue;
    }

    if (corner.value === 0) {
      continue;
    }

    if (!hasBoundVariableForField(anyNode, corner.field)) {
      return {
        type: "CORNER_RADIUS_NEEDS_VARIABLE",
        node,
        message: "CORNER RADIUS NEEDS VARIABLE",
      };
    }
  }

  return null;
}

function isInspectableSolidPaint(paint: Paint): paint is SolidPaint {
  if (paint.type !== "SOLID") {
    return false;
  }

  if (paint.visible === false) {
    return false;
  }

  return true;
}

function getBoundPaintAliases(
  node: any,
  field: "fills" | "strokes"
): any[] {
  if (!node || !node.boundVariables) {
    return [];
  }

  var boundVariables = node.boundVariables;
  var aliases = boundVariables[field];

  if (!Array.isArray(aliases)) {
    return [];
  }

  return aliases;
}

async function checkColorVariableViolations(
  node: SceneNode
): Promise<Violation[]> {
  const violations: Violation[] = [];
  const anyNode = node as any;

  if (
    node.type !== "TEXT" &&
    "fills" in anyNode &&
    anyNode.fills !== figma.mixed
  ) {
    const fills = anyNode.fills as readonly Paint[];
    const fillBound = getBoundPaintAliases(anyNode, "fills");

    for (let i = 0; i < fills.length; i += 1) {
      const paint = fills[i];

      if (!isInspectableSolidPaint(paint)) {
        continue;
      }

      const hasBoundVariable = !!fillBound[i];

      if (!hasBoundVariable) {
        violations.push({
          type: "FILL_NEEDS_VARIABLE",
          node,
          message: "FILL NEEDS VARIABLE",
        });
        break;
      }
    }
  }

  if ("strokes" in anyNode && anyNode.strokes !== figma.mixed) {
    const strokes = anyNode.strokes as readonly Paint[];
    const strokeBound = getBoundPaintAliases(anyNode, "strokes");

    for (let i = 0; i < strokes.length; i += 1) {
      const paint = strokes[i];

      if (!isInspectableSolidPaint(paint)) {
        continue;
      }

      const hasBoundVariable = !!strokeBound[i];

      if (!hasBoundVariable) {
        violations.push({
          type: "STROKE_NEEDS_VARIABLE",
          node,
          message: "STROKE NEEDS VARIABLE",
        });
        break;
      }
    }
  }

  if (node.type === "TEXT") {
    const textNodeAny = node as any;
    const nodeLevelTextFillBound = getBoundPaintAliases(textNodeAny, "fills");

    let hasNodeLevelTextFillVariable = false;

    for (let i = 0; i < nodeLevelTextFillBound.length; i += 1) {
      if (nodeLevelTextFillBound[i]) {
        hasNodeLevelTextFillVariable = true;
        break;
      }
    }

    if (!hasNodeLevelTextFillVariable) {
      const segments = node.getStyledTextSegments(["fills", "boundVariables"]);

      for (let i = 0; i < segments.length; i += 1) {
        const segment = segments[i];
        const segmentFills = segment.fills;

        if (!Array.isArray(segmentFills) || segmentFills.length === 0) {
          continue;
        }

        const maybeBoundVariables = (segment as any).boundVariables;
        const boundTextFills =
          maybeBoundVariables && Array.isArray(maybeBoundVariables.fills)
            ? maybeBoundVariables.fills
            : [];

        let hasUnboundSolidFill = false;

        for (let j = 0; j < segmentFills.length; j += 1) {
          const fill = segmentFills[j];

          if (!isInspectableSolidPaint(fill)) {
            continue;
          }

          const hasBoundVariable = !!boundTextFills[j];

          if (!hasBoundVariable) {
            hasUnboundSolidFill = true;
            break;
          }
        }

        if (hasUnboundSolidFill) {
          violations.push({
            type: "FILL_NEEDS_VARIABLE",
            node,
            message: "TEXT FILL NEEDS VARIABLE",
          });
          break;
        }
      }
    }
  }

  return violations;
}

function hasBoundVariableAlias(value: any): boolean {
  if (!value) {
    return false;
  }

  if (typeof value !== "object") {
    return false;
  }

  return true;
}

function shouldCheckAutoLayoutNumericField(
  node: SceneNode,
  field: NumericField
): boolean {
  if (!isAutoLayoutNode(node)) {
    return false;
  }

  return shouldCheckNumericField(node, field);
}

function isAutoLayoutNode(node: SceneNode): boolean {
  const anyNode = node as any;

  return (
    "layoutMode" in anyNode &&
    anyNode.layoutMode &&
    anyNode.layoutMode !== "NONE"
  );
}

function isZeroAllowedNumericField(field: NumericField): boolean {
  return (
    field === "itemSpacing" ||
    field === "counterAxisSpacing" ||
    field === "paddingTop" ||
    field === "paddingRight" ||
    field === "paddingBottom" ||
    field === "paddingLeft" ||
    field === "cornerRadius"
  );
}

function hasFiniteNumericField(node: any, field: NumericField): boolean {
  if (!(field in node)) {
    return false;
  }

  const value = node[field];

  return typeof value === "number" && Number.isFinite(value);
}

function hasBoundVariableForField(node: any, field: string): boolean {
  const value = node.boundVariables ? node.boundVariables[field] : null;

  if (!value) {
    return false;
  }

  if (Array.isArray(value)) {
    return value.some(Boolean);
  }

  return true;
}

function violationTypeFromNumericField(field: NumericField): ViolationType {
  if (field === "width") {
    return "WIDTH_NEEDS_VARIABLE";
  }

  if (field === "height") {
    return "HEIGHT_NEEDS_VARIABLE";
  }

  if (field === "itemSpacing") {
    return "ITEM_SPACING_NEEDS_VARIABLE";
  }

  if (field === "counterAxisSpacing") {
    return "COUNTER_AXIS_SPACING_NEEDS_VARIABLE";
  }

  if (field === "paddingTop") {
    return "PADDING_TOP_NEEDS_VARIABLE";
  }

  if (field === "paddingRight") {
    return "PADDING_RIGHT_NEEDS_VARIABLE";
  }

  if (field === "paddingBottom") {
    return "PADDING_BOTTOM_NEEDS_VARIABLE";
  }

  if (field === "paddingLeft") {
    return "PADDING_LEFT_NEEDS_VARIABLE";
  }

  return "CORNER_RADIUS_NEEDS_VARIABLE";
}

function shouldCheckFixedHeight(node: SceneNode): boolean {
  if (!hasFiniteNumericField(node, "height")) {
    return false;
  }

  const anyNode = node as any;
  const sizing = anyNode.layoutSizingVertical;

  if (sizing === "HUG" || sizing === "FILL") {
    return false;
  }

  if (sizing === "FIXED") {
    return true;
  }

  return true;
}

function shouldCheckFixedWidth(node: SceneNode): boolean {
  if (!hasFiniteNumericField(node, "width")) {
    return false;
  }

  const anyNode = node as any;
  const sizing = anyNode.layoutSizingHorizontal;

  if (sizing === "HUG" || sizing === "FILL") {
    return false;
  }

  if (sizing === "FIXED") {
    return true;
  }

  return true;
}

function shouldCheckNumericForNode(node: SceneNode): boolean {
  return (
    node.type === "FRAME" ||
    node.type === "GROUP" ||
    node.type === "COMPONENT" ||
    node.type === "COMPONENT_SET" ||
    node.type === "INSTANCE" ||
    node.type === "SECTION"
  );
}

function shouldCheckNumericField(
  node: SceneNode,
  field: NumericField
): boolean {
  if (!hasFiniteNumericField(node, field)) {
    return false;
  }

  const value = (node as any)[field];

  if (typeof value !== "number") {
    return false;
  }

  if (isZeroAllowedNumericField(field) && value === 0) {
    return false;
  }

  return true;
}

function checkNumericVariableViolations(node: SceneNode): Violation[] {
  if (!shouldCheckNumericForNode(node)) {
    return [];
  }

  const violations: Violation[] = [];

  const checks: Array<{
    field: NumericField;
    enabled: boolean;
    message: string;
  }> = [
    {
      field: "width",
      enabled: shouldCheckFixedWidth(node),
      message: "WIDTH NEEDS VARIABLE",
    },
    {
      field: "height",
      enabled: shouldCheckFixedHeight(node),
      message: "HEIGHT NEEDS VARIABLE",
    },
    {
      field: "itemSpacing",
      enabled: shouldCheckAutoLayoutNumericField(node, "itemSpacing"),
      message: "ITEM SPACING NEEDS VARIABLE",
    },
    {
      field: "counterAxisSpacing",
      enabled: shouldCheckAutoLayoutNumericField(node, "counterAxisSpacing"),
      message: "COUNTER AXIS SPACING NEEDS VARIABLE",
    },
    {
      field: "paddingTop",
      enabled: shouldCheckAutoLayoutNumericField(node, "paddingTop"),
      message: "PADDING TOP NEEDS VARIABLE",
    },
    {
      field: "paddingRight",
      enabled: shouldCheckAutoLayoutNumericField(node, "paddingRight"),
      message: "PADDING RIGHT NEEDS VARIABLE",
    },
    {
      field: "paddingBottom",
      enabled: shouldCheckAutoLayoutNumericField(node, "paddingBottom"),
      message: "PADDING BOTTOM NEEDS VARIABLE",
    },
    {
      field: "paddingLeft",
      enabled: shouldCheckAutoLayoutNumericField(node, "paddingLeft"),
      message: "PADDING LEFT NEEDS VARIABLE",
    },
  ];

  for (let i = 0; i < checks.length; i += 1) {
    const check = checks[i];

    if (!check.enabled) {
      continue;
    }

    if (hasBoundVariableForField(node as any, check.field)) {
      continue;
    }

    violations.push({
      type: violationTypeFromNumericField(check.field),
      node,
      message: check.message,
    });
  }

  const cornerRadiusViolation = checkCornerRadiusVariableViolation(node);

  if (cornerRadiusViolation) {
    violations.push(cornerRadiusViolation);
  }

  return violations;
}

function clearExistingMarkers(root: PageNode): void {
  var existing = root.children.find(function (node) {
    return node.name === MARKER_ROOT_NAME;
  });

  if (existing) {
    existing.remove();
  }
}

function uniqueViolations(violations: Violation[]): Violation[] {
  var map = new Map<string, Violation>();
  var i: number;

  for (i = 0; i < violations.length; i += 1) {
    var violation = violations[i];
    var key = violation.type + ":" + violation.node.id;

    if (!map.has(key)) {
      map.set(key, violation);
    }
  }

  return Array.from(map.values());
}

function hasAncestorType(node: BaseNode, type: NodeType): boolean {
  var current = node.parent;

  while (current) {
    if (current.type === type) {
      return true;
    }

    current = current.parent;
  }

  return false;
}

function isInsideMarkerLayer(node: BaseNode): boolean {
  var current: BaseNode | null = node;

  while (current) {
    if ("name" in current && current.name === MARKER_ROOT_NAME) {
      return true;
    }

    current = current.parent;
  }

  return false;
}

function getRenderableBounds(node: SceneNode): Rect | null {
  if ("absoluteRenderBounds" in node) {
    const renderBounds = node.absoluteRenderBounds;

    if (renderBounds) {
      return renderBounds;
    }
  }

  if ("absoluteBoundingBox" in node) {
    const boundingBox = node.absoluteBoundingBox;

    if (boundingBox) {
      return boundingBox;
    }
  }

  return null;
}

function getViolationColor(type: ViolationType): RGB {
  if (type === "RAW_TEXT") {
    return { r: 1, g: 0.231, b: 0.188 };
  }

  if (type === "NOT_REGISTERED") {
    return { r: 0.933, g: 0.267, b: 0.267 };
  }

  if (type === "FILL_NEEDS_VARIABLE" || type === "STROKE_NEEDS_VARIABLE") {
    return { r: 0.933, g: 0.565, b: 0.114 };
  }

  if (
    type === "WIDTH_NEEDS_VARIABLE" ||
    type === "HEIGHT_NEEDS_VARIABLE" ||
    type === "ITEM_SPACING_NEEDS_VARIABLE" ||
    type === "COUNTER_AXIS_SPACING_NEEDS_VARIABLE" ||
    type === "PADDING_TOP_NEEDS_VARIABLE" ||
    type === "PADDING_RIGHT_NEEDS_VARIABLE" ||
    type === "PADDING_BOTTOM_NEEDS_VARIABLE" ||
    type === "PADDING_LEFT_NEEDS_VARIABLE" ||
    type === "CORNER_RADIUS_NEEDS_VARIABLE"
  ) {
    return { r: 0.353, g: 0.608, b: 0.996 };
  }

  return { r: 0.933, g: 0.267, b: 0.267 };
}

async function renderMarkers(
  page: PageNode,
  violations: Violation[]
): Promise<number> {
  if (violations.length === 0) {
    return 0;
  }

  await figma.loadFontAsync(LABEL_FONT);

  var root = figma.createFrame();
  root.name = MARKER_ROOT_NAME;
  root.clipsContent = false;
  root.locked = true;
  root.fills = [];
  root.strokes = [];
  root.x = 0;
  root.y = 0;
  root.resizeWithoutConstraints(0, 0);

  page.appendChild(root);

  var markerCount = 0;
  var i: number;

  for (i = 0; i < violations.length; i += 1) {
    var marker = createMarkerForViolation(violations[i]);

    if (!marker) {
      continue;
    }

    root.appendChild(marker);
    markerCount += 1;
  }

  return markerCount;
}

function createMarkerForViolation(violation: Violation): FrameNode | null {
  var bounds = getRenderableBounds(violation.node);

  if (!bounds) {
    return null;
  }

  var color = getViolationColor(violation.type);

  var frame = figma.createFrame();
  frame.name = "marker:" + violation.type;
  frame.clipsContent = false;
  frame.fills = [];
  frame.strokes = [];
  frame.x = bounds.x;
  frame.y = bounds.y;
  frame.resizeWithoutConstraints(
    Math.max(bounds.width, 1),
    Math.max(bounds.height, 1)
  );

  var outline = figma.createRectangle();
  outline.name = "outline";
  outline.x = 0;
  outline.y = 0;
  outline.resizeWithoutConstraints(
    Math.max(bounds.width, 1),
    Math.max(bounds.height, 1)
  );
  outline.fills = [];
  outline.strokeWeight = 2;
  outline.strokes = [{ type: "SOLID", color: color }];
  outline.cornerRadius = 4;

  var badge = figma.createFrame();
  badge.name = "badge";
  badge.layoutMode = "HORIZONTAL";
  badge.primaryAxisSizingMode = "AUTO";
  badge.counterAxisSizingMode = "AUTO";
  badge.paddingLeft = 8;
  badge.paddingRight = 8;
  badge.paddingTop = 4;
  badge.paddingBottom = 4;
  badge.cornerRadius = 999;
  badge.fills = [{ type: "SOLID", color: color }];
  badge.strokes = [];
  badge.clipsContent = false;

  var text = figma.createText();
  text.name = "label";
  text.fontName = LABEL_FONT;
  text.fontSize = 10;
  text.characters = violation.message;
  text.fills = [{ type: "SOLID", color: { r: 1, g: 1, b: 1 } }];

  badge.appendChild(text);
  badge.x = 0;
  badge.y = -24;

  frame.appendChild(outline);
  frame.appendChild(badge);

  return frame;
}

function normalizeSignaturePart(value: string): string {
  return value.trim();
}

function buildVariantSignature(component: ComponentNode): string {
  var variantProperties = component.variantProperties;

  if (!variantProperties || typeof variantProperties !== "object") {
    return "";
  }

  var keys = Object.keys(variantProperties).sort();

  if (keys.length === 0) {
    return "";
  }

  var parts: string[] = [];
  var i: number;

  for (i = 0; i < keys.length; i += 1) {
    var key = keys[i];
    parts.push(key + "=" + String(variantProperties[key]));
  }

  return "::" + parts.join("|");
}

function toLegacySignatureFromComponent(component: ComponentNode): string {
  var parent = component.parent;

  if (parent && parent.type === "COMPONENT_SET") {
    return normalizeSignaturePart(parent.name);
  }

  return normalizeSignaturePart(component.name);
}

async function toLegacySignatureFromInstance(
  instance: InstanceNode
): Promise<string | null> {
  var mainComponent = await instance.getMainComponentAsync();

  if (!mainComponent) {
    return null;
  }

  return toLegacySignatureFromComponent(mainComponent);
}

function isPromotableNode(node: SceneNode): boolean {
  return (
    node.type === "INSTANCE" ||
    node.type === "COMPONENT" ||
    node.type === "COMPONENT_SET"
  );
}

function collectPromoteTargetsFromSelection(selection: readonly SceneNode[]): SceneNode[] {
  var result: SceneNode[] = [];
  var i: number;

  for (i = 0; i < selection.length; i += 1) {
    var node = selection[i];

    if (!isPromotableNode(node)) {
      continue;
    }

    result.push(node);
  }

  return result;
}

function collectPromoteTargetsFromPage(page: PageNode): SceneNode[] {
  return page.findAll(function (node) {
    return isPromotableNode(node);
  }) as SceneNode[];
}

async function getPromoteIdentity(node: SceneNode): Promise<string | null> {
  if (node.type === "COMPONENT_SET") {
    return node.key ? "COMPONENT_SET:" + node.key : "COMPONENT_SET:" + node.id;
  }

  if (node.type === "COMPONENT") {
    var parent = node.parent;

    if (parent && parent.type === "COMPONENT_SET") {
      return parent.key
        ? "COMPONENT_SET:" + parent.key
        : "COMPONENT_SET:" + parent.id;
    }

    return node.key ? "COMPONENT:" + node.key : "COMPONENT:" + node.id;
  }

  if (node.type === "INSTANCE") {
    var mainComponent = await node.getMainComponentAsync();

    if (!mainComponent) {
      return null;
    }

    var mainParent = mainComponent.parent;

    if (mainParent && mainParent.type === "COMPONENT_SET") {
      return mainParent.key
        ? "COMPONENT_SET:" + mainParent.key
        : "COMPONENT_SET:" + mainParent.id;
    }

    return mainComponent.key
      ? "COMPONENT:" + mainComponent.key
      : "COMPONENT:" + mainComponent.id;
  }

  return null;
}

async function collectExistingPromoteIdentities(page: PageNode): Promise<Set<string>> {
  var result = new Set<string>();
  var nodes = page.findAll(function (node) {
    return isPromotableNode(node);
  }) as SceneNode[];
  var i: number;

  for (i = 0; i < nodes.length; i += 1) {
    var identity = await getPromoteIdentity(nodes[i]);

    if (!identity) {
      continue;
    }

    result.add(identity);
  }

  return result;
}

async function promoteNodesToAllowedPage(
  nodes: SceneNode[],
  allowedPage: PageNode
): Promise<{
  promotedCount: number;
  replacedCount: number;
  skippedCount: number;
}> {
  var existingNodes = await collectExistingPromoteNodes(allowedPage);
  var promotedCount = 0;
  var replacedCount = 0;
  var skippedCount = 0;
  var i: number;

  for (i = 0; i < nodes.length; i += 1) {
    var node = nodes[i];
    var identity = await getPromoteIdentity(node);

    if (!identity) {
      skippedCount += 1;
      continue;
    }

    var existingNode = existingNodes.get(identity);

    if (existingNode) {
      existingNode.remove();
      replacedCount += 1;
    }

    allowedPage.appendChild(node);
    existingNodes.set(identity, node);
    promotedCount += 1;
  }

  return {
    promotedCount: promotedCount,
    replacedCount: replacedCount,
    skippedCount: skippedCount
  };
}
async function collectExistingPromoteNodes(
  page: PageNode
): Promise<Map<string, SceneNode>> {
  var result = new Map<string, SceneNode>();
  var nodes = page.findAll(function (node) {
    return isPromotableNode(node);
  }) as SceneNode[];
  var i: number;

  for (i = 0; i < nodes.length; i += 1) {
    var identity = await getPromoteIdentity(nodes[i]);

    if (!identity) {
      continue;
    }

    result.set(identity, nodes[i]);
  }

  return result;
}