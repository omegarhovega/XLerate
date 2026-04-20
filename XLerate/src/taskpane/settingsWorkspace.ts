import { DEFAULT_AUTO_COLOR_PALETTE, type AutoColorPalette } from "../core/autoColor";
import { DEFAULT_CELL_FORMATS, type CellFormatDefinition } from "../core/cellFormatCycle";
import { DEFAULT_DATE_FORMATS, type DateFormatDefinition } from "../core/dateFormatCycle";
import {
  buildDefaultFormatSettings,
  cloneResolvedFormatSettings,
  getFormatSettingsValidationError,
  type ResolvedFormatSettings,
} from "../core/formatSettings";
import { DEFAULT_NUMBER_FORMATS, type NumberFormatDefinition } from "../core/numberFormatCycle";
import { DEFAULT_TEXT_STYLES, type TextStyleDefinition } from "../core/textStyleCycle";
import { MAX_TRACE_MAX_DEPTH, MAX_TRACE_SAFETY_LIMIT } from "../core/traceUtils";

type SettingsTabKey =
  | "numberFormats"
  | "dateFormats"
  | "cellFormats"
  | "textStyles"
  | "autoColorPalette"
  | "trace";

type ListTabKey = Exclude<SettingsTabKey, "autoColorPalette" | "trace">;

type SettingsView = "index" | SettingsTabKey;

type SettingsWorkspaceOptions = {
  initialSettings: ResolvedFormatSettings;
  loadSavedSettings: () => Promise<ResolvedFormatSettings>;
  exportSettings: (settings: ResolvedFormatSettings) => Promise<void>;
  saveSettings: (settings: ResolvedFormatSettings) => Promise<void>;
  onStatus: (message: string) => void;
};

type SettingsWorkspaceState = {
  draft: ResolvedFormatSettings;
  savedSnapshot: ResolvedFormatSettings;
  currentView: SettingsView;
  selectedIndex: Record<ListTabKey, number | null>;
};

type TabMeta = {
  title: string;
  description: string;
  group: "Formatting" | "Behavior";
};

const LIST_TABS: ListTabKey[] = ["numberFormats", "dateFormats", "cellFormats", "textStyles"];

const TAB_META: Record<SettingsTabKey, TabMeta> = {
  numberFormats: {
    title: "Number Format Cycle",
    description: "Ordered presets for Cycle Number on the XLerate ribbon.",
    group: "Formatting",
  },
  dateFormats: {
    title: "Date Format Cycle",
    description: "Ordered presets for Cycle Date on the XLerate ribbon.",
    group: "Formatting",
  },
  cellFormats: {
    title: "Cell Format Cycle",
    description: "Fill, font, and border combinations for Cycle Cell.",
    group: "Formatting",
  },
  textStyles: {
    title: "Text Style Cycle",
    description: "Session-cycled presets for headings, sums, and normal rows.",
    group: "Formatting",
  },
  autoColorPalette: {
    title: "Auto-color Palette",
    description: "Semantic font colors used by Auto-color on the ribbon.",
    group: "Behavior",
  },
  trace: {
    title: "Trace Settings",
    description: "Default depth and safety limits for the trace dialog.",
    group: "Behavior",
  },
};

const AUTO_COLOR_FIELDS: Array<{
  key: keyof AutoColorPalette;
  label: string;
  description: string;
}> = [
  {
    key: "input",
    label: "Inputs",
    description: "Typed numeric values that are not formulas.",
  },
  {
    key: "formula",
    label: "Formula",
    description: "Formulas with no external references.",
  },
  {
    key: "worksheetLink",
    label: "Same-sheet link",
    description: "Formulas referencing cells on the active sheet.",
  },
  {
    key: "workbookLink",
    label: "Cross-sheet link",
    description: "Formulas referencing other sheets in the workbook.",
  },
  {
    key: "external",
    label: "External link",
    description: "Formulas referencing another workbook or external source.",
  },
  {
    key: "hyperlink",
    label: "Hyperlink",
    description: "Cells with a hyperlink attached.",
  },
  {
    key: "partialInput",
    label: "Partial input",
    description: "Constants mixed with small references, such as =100+A1.",
  },
];

const COMMON_EXCEL_FONTS = [
  "Aptos",
  "Aptos Display",
  "Aptos Narrow",
  "Arial",
  "Arial Narrow",
  "Bahnschrift",
  "Calibri",
  "Cambria",
  "Candara",
  "Century Gothic",
  "Consolas",
  "Constantia",
  "Corbel",
  "Courier New",
  "Franklin Gothic Book",
  "Garamond",
  "Georgia",
  "Gill Sans MT",
  "Helvetica",
  "Lucida Sans Unicode",
  "Segoe UI",
  "Tahoma",
  "Times New Roman",
  "Trebuchet MS",
  "Verdana",
];

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function isListTab(tab: SettingsTabKey): tab is ListTabKey {
  return LIST_TABS.includes(tab as ListTabKey);
}

function isDetailView(view: SettingsView): view is SettingsTabKey {
  return view !== "index";
}

function serializeSettings(settings: ResolvedFormatSettings): string {
  return JSON.stringify(settings);
}

function safeColor(value: string, fallback = "#F97316"): string {
  const trimmed = value.trim();
  if (/^#?[0-9a-fA-F]{6}$/.test(trimmed)) {
    return trimmed.startsWith("#") ? trimmed : `#${trimmed}`;
  }
  return fallback;
}

function clampIndex(index: number, length: number): number {
  if (length <= 0) {
    return 0;
  }
  if (!Number.isInteger(index) || index < 0) {
    return 0;
  }
  if (index >= length) {
    return length - 1;
  }
  return index;
}

function normalizeSelection(state: SettingsWorkspaceState): void {
  state.selectedIndex.numberFormats =
    state.selectedIndex.numberFormats === null
      ? null
      : clampIndex(state.selectedIndex.numberFormats, state.draft.numberFormats.length);
  state.selectedIndex.dateFormats =
    state.selectedIndex.dateFormats === null
      ? null
      : clampIndex(state.selectedIndex.dateFormats, state.draft.dateFormats.length);
  state.selectedIndex.cellFormats =
    state.selectedIndex.cellFormats === null
      ? null
      : clampIndex(state.selectedIndex.cellFormats, state.draft.cellFormats.length);
  state.selectedIndex.textStyles =
    state.selectedIndex.textStyles === null
      ? null
      : clampIndex(state.selectedIndex.textStyles, state.draft.textStyles.length);
}

function getListItems(state: SettingsWorkspaceState, tab: ListTabKey) {
  switch (tab) {
    case "numberFormats":
      return state.draft.numberFormats;
    case "dateFormats":
      return state.draft.dateFormats;
    case "cellFormats":
      return state.draft.cellFormats;
    case "textStyles":
      return state.draft.textStyles;
  }
}

function createNumberFormatPreset(count: number): NumberFormatDefinition {
  return {
    ...DEFAULT_NUMBER_FORMATS[0],
    name: `Number Format ${count + 1}`,
  };
}

function createDateFormatPreset(count: number): DateFormatDefinition {
  return {
    ...DEFAULT_DATE_FORMATS[0],
    name: `Date Format ${count + 1}`,
  };
}

function createCellFormatPreset(count: number): CellFormatDefinition {
  return {
    ...DEFAULT_CELL_FORMATS[0],
    name: `Cell Format ${count + 1}`,
  };
}

function createTextStylePreset(count: number): TextStyleDefinition {
  return {
    ...DEFAULT_TEXT_STYLES[0],
    name: `Text Style ${count + 1}`,
  };
}

function addPreset(state: SettingsWorkspaceState, tab: ListTabKey): void {
  switch (tab) {
    case "numberFormats":
      state.draft.numberFormats.push(createNumberFormatPreset(state.draft.numberFormats.length));
      state.selectedIndex.numberFormats = state.draft.numberFormats.length - 1;
      return;
    case "dateFormats":
      state.draft.dateFormats.push(createDateFormatPreset(state.draft.dateFormats.length));
      state.selectedIndex.dateFormats = state.draft.dateFormats.length - 1;
      return;
    case "cellFormats":
      state.draft.cellFormats.push(createCellFormatPreset(state.draft.cellFormats.length));
      state.selectedIndex.cellFormats = state.draft.cellFormats.length - 1;
      return;
    case "textStyles":
      state.draft.textStyles.push(createTextStylePreset(state.draft.textStyles.length));
      state.selectedIndex.textStyles = state.draft.textStyles.length - 1;
      return;
  }
}

function movePreset(
  state: SettingsWorkspaceState,
  tab: ListTabKey,
  index: number,
  direction: number
): void {
  const items = getListItems(state, tab);
  const targetIndex = index + direction;
  if (targetIndex < 0 || targetIndex >= items.length) {
    return;
  }

  const [moved] = items.splice(index, 1);
  items.splice(targetIndex, 0, moved);
  state.selectedIndex[tab] = targetIndex;
}

function deletePreset(state: SettingsWorkspaceState, tab: ListTabKey, index: number): void {
  const items = getListItems(state, tab);
  if (items.length <= 1) {
    return;
  }
  items.splice(index, 1);
  state.selectedIndex[tab] = clampIndex(index - 1, items.length);
}

function updateSelectedListField(
  state: SettingsWorkspaceState,
  field: string,
  value: string | number | boolean
): void {
  if (!isDetailView(state.currentView) || !isListTab(state.currentView)) {
    return;
  }

  const index = state.selectedIndex[state.currentView];
  if (index === null) {
    return;
  }
  const items = getListItems(state, state.currentView);
  const target = items[index] as Record<string, unknown> | undefined;
  if (!target) {
    return;
  }

  target[field] = value;
}

function getSectionSummary(state: SettingsWorkspaceState, tab: SettingsTabKey): string {
  switch (tab) {
    case "numberFormats":
      return `${state.draft.numberFormats.length} presets`;
    case "dateFormats":
      return `${state.draft.dateFormats.length} presets`;
    case "cellFormats":
      return `${state.draft.cellFormats.length} presets`;
    case "textStyles":
      return `${state.draft.textStyles.length} presets`;
    case "autoColorPalette":
      return `${AUTO_COLOR_FIELDS.length} colors`;
    case "trace":
      return `Depth ${state.draft.trace.maxDepth} / Limit ${state.draft.trace.safetyLimit}`;
  }
}

function renderNumberDatePreview(item: NumberFormatDefinition | DateFormatDefinition): string {
  return `<span class="accordion-item-subtitle accordion-code">${escapeHtml(item.formatCode)}</span>`;
}

function renderCellFormatPreview(item: CellFormatDefinition): string {
  const background =
    item.fillPattern === "Solid" ? safeColor(item.fillColor, "#FFFFFF") : "#FFFFFF";
  const border =
    item.borderStyle === "Continuous"
      ? `1px solid ${safeColor(item.borderColor, "#D6D3D1")}`
      : "1px dashed #D6D3D1";
  const textDecoration = [
    item.fontUnderline ? " underline" : "",
    item.fontStrikethrough ? " line-through" : "",
  ].join("");

  return `
    <div class="mini-preview" style="background:${background}; border:${border}; color:${safeColor(item.fontColor, "#1C1917")}; font-weight:${item.fontBold ? 700 : 500}; font-style:${item.fontItalic ? "italic" : "normal"}; text-decoration:${textDecoration.trim() || "none"};">
      Aa
    </div>
  `;
}

function renderTextStylePreview(item: TextStyleDefinition): string {
  const border =
    item.borderStyle === "None" ? "none" : `1px solid ${safeColor(item.fontColor, "#1C1917")}`;
  const style = [
    `background:${item.fillPattern === "None" ? "transparent" : safeColor(item.backColor, "#FFFFFF")};`,
    `color:${safeColor(item.fontColor, "#1C1917")};`,
    `font-size:${Number.isFinite(item.fontSize) ? item.fontSize : 11}px;`,
    item.bold ? "font-weight:700;" : "",
    item.italic ? "font-style:italic;" : "",
    item.underline ? "text-decoration:underline;" : "",
    `border-top:${item.borderTop ? border : "none"};`,
    `border-bottom:${item.borderBottom ? border : "none"};`,
    `border-left:${item.borderLeft ? border : "none"};`,
    `border-right:${item.borderRight ? border : "none"};`,
  ].join(" ");

  return `<div class="mini-preview text-style-preview" style="${style}">Sample</div>`;
}

function renderField(label: string, control: string, helpText?: string, wide = false): string {
  return `
    <label class="field${wide ? " field-wide" : ""}">
      <span class="field-label">${escapeHtml(label)}</span>
      ${control}
      ${helpText ? `<span class="field-help">${escapeHtml(helpText)}</span>` : ""}
    </label>
  `;
}

function parseEditorValue(
  target: HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement
): string | number | boolean {
  if (target instanceof HTMLInputElement && target.type === "checkbox") {
    return target.checked;
  }

  if (target instanceof HTMLInputElement && target.type === "number") {
    return target.value === "" ? NaN : Number(target.value);
  }

  return target.value;
}

function syncDraftFromDom(root: HTMLDivElement, state: SettingsWorkspaceState): void {
  const controls = root.querySelectorAll<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>(
    "[data-action='update-list-field'], [data-action='update-palette'], [data-action='update-trace']"
  );

  controls.forEach((target) => {
    const action = target.dataset.action;
    if (!action) {
      return;
    }

    if (action === "update-list-field") {
      updateSelectedListField(state, String(target.dataset.field ?? ""), parseEditorValue(target));
      return;
    }

    if (action === "update-palette") {
      const key = target.dataset.key as keyof AutoColorPalette | undefined;
      if (key) {
        state.draft.autoColorPalette[key] = String(parseEditorValue(target));
      }
      return;
    }

    if (action === "update-trace") {
      const key = target.dataset.key as keyof ResolvedFormatSettings["trace"] | undefined;
      if (key) {
        state.draft.trace[key] = Number(parseEditorValue(target));
      }
    }
  });
}

function renderFontOptions(currentFont: string): string {
  const options = COMMON_EXCEL_FONTS.includes(currentFont)
    ? COMMON_EXCEL_FONTS
    : [currentFont, ...COMMON_EXCEL_FONTS];

  return options
    .map(
      (fontName) =>
        `<option value="${escapeHtml(fontName)}" ${fontName === currentFont ? "selected" : ""}>${escapeHtml(fontName)}</option>`
    )
    .join("");
}

function renderNumberDateEditorContent(
  item: NumberFormatDefinition | DateFormatDefinition,
  typeLabel: string
): string {
  return `
    <div class="field-grid">
      ${renderField(
        "Preset name",
        `<input data-action="update-list-field" data-field="name" type="text" value="${escapeHtml(item.name)}" />`
      )}
      ${renderField(
        "Excel format code",
        `<input data-action="update-list-field" data-field="formatCode" type="text" value="${escapeHtml(item.formatCode)}" />`,
        `This code is applied when ${typeLabel.toLowerCase()} cycles reach this preset.`,
        true
      )}
    </div>
    <div class="preview-panel">
      <span class="preview-label">Format code</span>
      <code>${escapeHtml(item.formatCode)}</code>
    </div>
  `;
}

function renderCellFormatEditorContent(item: CellFormatDefinition): string {
  return `
    <div class="field-grid">
      ${renderField(
        "Preset name",
        `<input data-action="update-list-field" data-field="name" type="text" value="${escapeHtml(item.name)}" />`
      )}
      ${renderField(
        "Fill pattern",
        `<select data-action="update-list-field" data-field="fillPattern">
          <option value="Solid" ${item.fillPattern === "Solid" ? "selected" : ""}>Solid</option>
          <option value="None" ${item.fillPattern === "None" ? "selected" : ""}>None</option>
        </select>`
      )}
      ${renderField(
        "Fill color",
        `<input data-action="update-list-field" data-field="fillColor" type="color" value="${safeColor(item.fillColor, "#FFFFFF")}" />`
      )}
      ${renderField(
        "Font color",
        `<input data-action="update-list-field" data-field="fontColor" type="color" value="${safeColor(item.fontColor, "#1C1917")}" />`
      )}
      ${renderField(
        "Border style",
        `<select data-action="update-list-field" data-field="borderStyle">
          <option value="None" ${item.borderStyle === "None" ? "selected" : ""}>None</option>
          <option value="Continuous" ${item.borderStyle === "Continuous" ? "selected" : ""}>Continuous</option>
        </select>`
      )}
      ${renderField(
        "Border color",
        `<input data-action="update-list-field" data-field="borderColor" type="color" value="${safeColor(item.borderColor, "#D6D3D1")}" />`
      )}
    </div>
    <div class="toggle-grid">
      <label class="toggle"><input data-action="update-list-field" data-field="fontBold" type="checkbox" ${item.fontBold ? "checked" : ""} /> Bold</label>
      <label class="toggle"><input data-action="update-list-field" data-field="fontItalic" type="checkbox" ${item.fontItalic ? "checked" : ""} /> Italic</label>
      <label class="toggle"><input data-action="update-list-field" data-field="fontUnderline" type="checkbox" ${item.fontUnderline ? "checked" : ""} /> Underline</label>
      <label class="toggle"><input data-action="update-list-field" data-field="fontStrikethrough" type="checkbox" ${item.fontStrikethrough ? "checked" : ""} /> Strike</label>
    </div>
    <div class="preview-panel">
      <span class="preview-label">Preview</span>
      ${renderCellFormatPreview(item)}
    </div>
  `;
}

function renderTextStyleEditorContent(item: TextStyleDefinition): string {
  return `
    <div class="field-grid">
      ${renderField(
        "Preset name",
        `<input data-action="update-list-field" data-field="name" type="text" value="${escapeHtml(item.name)}" />`
      )}
      ${renderField(
        "Font name",
        `<select data-action="update-list-field" data-field="fontName">
          ${renderFontOptions(item.fontName)}
        </select>`
      )}
      ${renderField(
        "Font size",
        `<input data-action="update-list-field" data-field="fontSize" type="number" min="1" step="1" value="${item.fontSize}" />`
      )}
      ${renderField(
        "Font color",
        `<input data-action="update-list-field" data-field="fontColor" type="color" value="${safeColor(item.fontColor, "#1C1917")}" />`
      )}
      ${renderField(
        "Background fill",
        `<select data-action="update-list-field" data-field="fillPattern">
          <option value="Solid" ${item.fillPattern === "Solid" ? "selected" : ""}>Solid</option>
          <option value="None" ${item.fillPattern === "None" ? "selected" : ""}>No fill</option>
        </select>`
      )}
      ${renderField(
        "Background color",
        `<input data-action="update-list-field" data-field="backColor" type="color" value="${safeColor(item.backColor, "#FFFFFF")}" ${item.fillPattern === "None" ? "disabled" : ""} />`,
        item.fillPattern === "None" ? "Choose Solid to enable a fill color." : undefined
      )}
      ${renderField(
        "Border style",
        `<select data-action="update-list-field" data-field="borderStyle">
          <option value="None" ${item.borderStyle === "None" ? "selected" : ""}>None</option>
          <option value="Continuous" ${item.borderStyle === "Continuous" ? "selected" : ""}>Continuous</option>
          <option value="Double" ${item.borderStyle === "Double" ? "selected" : ""}>Double</option>
          <option value="Dash" ${item.borderStyle === "Dash" ? "selected" : ""}>Dash</option>
          <option value="Dot" ${item.borderStyle === "Dot" ? "selected" : ""}>Dot</option>
        </select>`
      )}
    </div>
    <div class="toggle-grid">
      <label class="toggle"><input data-action="update-list-field" data-field="bold" type="checkbox" ${item.bold ? "checked" : ""} /> Bold</label>
      <label class="toggle"><input data-action="update-list-field" data-field="italic" type="checkbox" ${item.italic ? "checked" : ""} /> Italic</label>
      <label class="toggle"><input data-action="update-list-field" data-field="underline" type="checkbox" ${item.underline ? "checked" : ""} /> Underline</label>
      <label class="toggle"><input data-action="update-list-field" data-field="borderTop" type="checkbox" ${item.borderTop ? "checked" : ""} /> Top border</label>
      <label class="toggle"><input data-action="update-list-field" data-field="borderBottom" type="checkbox" ${item.borderBottom ? "checked" : ""} /> Bottom border</label>
      <label class="toggle"><input data-action="update-list-field" data-field="borderLeft" type="checkbox" ${item.borderLeft ? "checked" : ""} /> Left border</label>
      <label class="toggle"><input data-action="update-list-field" data-field="borderRight" type="checkbox" ${item.borderRight ? "checked" : ""} /> Right border</label>
    </div>
    <div class="preview-panel">
      <span class="preview-label">Preview</span>
      ${renderTextStylePreview(item)}
    </div>
  `;
}

function renderPaletteEditor(palette: AutoColorPalette): string {
  return `
    <div class="content-card">
      <div class="content-card-header">
        <div>
          <p class="section-kicker">Auto-color palette</p>
          <h3>Semantic workbook colors</h3>
        </div>
      </div>
      <div class="palette-grid">
        ${AUTO_COLOR_FIELDS.map(
          (field) => `
            <label class="palette-row">
              <span class="palette-copy">
                <span class="palette-label">${escapeHtml(field.label)}</span>
                <span class="palette-help">${escapeHtml(field.description)}</span>
              </span>
              <span class="palette-inputs">
                <input data-action="update-palette" data-key="${field.key}" type="color" value="${safeColor(palette[field.key], DEFAULT_AUTO_COLOR_PALETTE[field.key])}" />
                <span class="palette-value">${escapeHtml(safeColor(palette[field.key], DEFAULT_AUTO_COLOR_PALETTE[field.key]).toUpperCase())}</span>
              </span>
            </label>
          `
        ).join("")}
      </div>
    </div>
  `;
}

function renderTraceEditor(settings: ResolvedFormatSettings["trace"]): string {
  return `
    <div class="content-card">
      <div class="content-card-header">
        <div>
          <p class="section-kicker">Trace dialog defaults</p>
          <h3>Expansion controls</h3>
        </div>
      </div>
      <div class="field-grid">
        ${renderField(
          "Maximum depth",
          `<input data-action="update-trace" data-key="maxDepth" type="number" min="1" max="${MAX_TRACE_MAX_DEPTH}" step="1" value="${settings.maxDepth}" />`,
          `XLerate clamps this between 1 and ${MAX_TRACE_MAX_DEPTH}.`
        )}
        ${renderField(
          "Safety limit",
          `<input data-action="update-trace" data-key="safetyLimit" type="number" min="1" max="${MAX_TRACE_SAFETY_LIMIT}" step="1" value="${settings.safetyLimit}" />`,
          `Hard stop for rendered rows. XLerate clamps this between 1 and ${MAX_TRACE_SAFETY_LIMIT}.`
        )}
      </div>
      <div class="preview-panel trace-stats-panel">
        <span class="preview-label">Current defaults</span>
        <div class="trace-stat-grid">
          <div><strong>${settings.maxDepth}</strong><span>levels</span></div>
          <div><strong>${settings.safetyLimit}</strong><span>rows</span></div>
        </div>
      </div>
    </div>
  `;
}

function renderDetailContent(state: SettingsWorkspaceState, tab: SettingsTabKey): string {
  switch (tab) {
    case "numberFormats":
      return renderExpandablePresetList(state, tab);
    case "dateFormats":
      return renderExpandablePresetList(state, tab);
    case "cellFormats":
      return renderExpandablePresetList(state, tab);
    case "textStyles":
      return renderExpandablePresetList(state, tab);
    case "autoColorPalette":
      return `<div class="detail-grid detail-grid-single">${renderPaletteEditor(state.draft.autoColorPalette)}</div>`;
    case "trace":
      return `<div class="detail-grid detail-grid-single">${renderTraceEditor(state.draft.trace)}</div>`;
  }
}

function renderExpandablePresetEditor(
  state: SettingsWorkspaceState,
  tab: ListTabKey,
  index: number
): string {
  switch (tab) {
    case "numberFormats":
      return renderNumberDateEditorContent(state.draft.numberFormats[index], "Number format");
    case "dateFormats":
      return renderNumberDateEditorContent(state.draft.dateFormats[index], "Date format");
    case "cellFormats":
      return renderCellFormatEditorContent(state.draft.cellFormats[index]);
    case "textStyles":
      return renderTextStyleEditorContent(state.draft.textStyles[index]);
  }
}

function renderExpandablePresetList(state: SettingsWorkspaceState, tab: ListTabKey): string {
  const items = getListItems(state, tab);
  const openIndex = state.selectedIndex[tab];

  return `
    <div class="detail-grid detail-grid-single">
      <section class="accordion-panel">
        <div class="accordion-panel-header">
          <div>
            <p class="section-kicker">Presets</p>
            <h3>${escapeHtml(TAB_META[tab].title)}</h3>
          </div>
          <button type="button" class="secondary-btn compact-btn" data-action="add-item">Add</button>
        </div>
        <div class="accordion-list">
          ${items
            .map((item, index) => {
              const isOpen = openIndex === index;
              const title = "name" in item ? escapeHtml(String(item.name)) : `Preset ${index + 1}`;
              const subtitle =
                tab === "numberFormats" || tab === "dateFormats"
                  ? renderNumberDatePreview(item as NumberFormatDefinition | DateFormatDefinition)
                  : tab === "cellFormats"
                    ? renderCellFormatPreview(item as CellFormatDefinition)
                    : renderTextStylePreview(item as TextStyleDefinition);

              return `
                <div class="accordion-item${isOpen ? " accordion-item-open" : ""}">
                  <div class="accordion-item-row">
                    <button type="button" class="accordion-item-button" data-action="select-item" data-index="${index}">
                      <span class="accordion-item-order">${index + 1}</span>
                      <span class="accordion-item-copy">
                        <span class="accordion-item-title">${title}</span>
                        ${subtitle}
                      </span>
                      <span class="accordion-item-chevron">${isOpen ? "▾" : "▸"}</span>
                    </button>
                  </div>
                  <div class="accordion-item-actions">
                    <button type="button" class="mini-btn" data-action="move-item" data-index="${index}" data-direction="-1" ${index === 0 ? "disabled" : ""}>Up</button>
                    <button type="button" class="mini-btn" data-action="move-item" data-index="${index}" data-direction="1" ${index === items.length - 1 ? "disabled" : ""}>Down</button>
                    <button type="button" class="mini-btn danger-btn" data-action="delete-item" data-index="${index}" ${items.length <= 1 ? "disabled" : ""}>Delete</button>
                  </div>
                  ${
                    isOpen
                      ? `
                        <div class="accordion-item-body">
                          ${renderExpandablePresetEditor(state, tab, index)}
                        </div>
                      `
                      : ""
                  }
                </div>
              `;
            })
            .join("")}
        </div>
      </section>
    </div>
  `;
}

function renderIndex(state: SettingsWorkspaceState): string {
  const groups: Array<"Formatting" | "Behavior"> = ["Formatting", "Behavior"];

  return `
    <div class="settings-shell">
      <section class="surface-card">
        <div class="settings-home-copy">
          <p class="section-kicker">Workbook</p>
          <h1>Settings</h1>
          <p class="section-copy">
            Choose one settings area to edit. Each page focuses on a single part of the workbook configuration.
          </p>
        </div>
      </section>
      <section class="surface-card settings-index-card">
        ${groups
          .map((group) => {
            const tabs = (Object.keys(TAB_META) as SettingsTabKey[]).filter(
              (tab) => TAB_META[tab].group === group
            );

            return `
              <div class="index-group">
                <p class="index-group-label">${escapeHtml(group)}</p>
                <div class="index-link-list">
                  ${tabs
                    .map(
                      (tab) => `
                        <button type="button" class="index-link" data-action="open-section" data-tab="${tab}">
                          <span class="index-link-copy">
                            <span class="index-link-title">${escapeHtml(TAB_META[tab].title)}</span>
                          </span>
                          <span class="index-link-meta">
                            <span class="index-link-summary">${escapeHtml(getSectionSummary(state, tab))}</span>
                            <span class="index-link-arrow">&rsaquo;</span>
                          </span>
                        </button>
                      `
                    )
                    .join("")}
                </div>
              </div>
            `;
          })
          .join("")}
      </section>
    </div>
  `;
}

function renderDetail(state: SettingsWorkspaceState, tab: SettingsTabKey): string {
  const dirty = serializeSettings(state.draft) !== serializeSettings(state.savedSnapshot);
  const validationError = getFormatSettingsValidationError(state.draft);

  return `
    <div class="settings-shell">
      <section class="surface-card detail-surface">
        <div class="detail-topbar">
          <button type="button" class="back-button" data-action="go-home" aria-label="Back to settings">
            <span aria-hidden="true">&larr;</span>
          </button>
          <div class="detail-topbar-copy">
            <p class="section-kicker">Settings</p>
            <h1>${escapeHtml(TAB_META[tab].title)}</h1>
            <p class="section-copy">${escapeHtml(TAB_META[tab].description)}</p>
          </div>
        </div>
        <div class="detail-toolbar">
          <button type="button" class="secondary-btn" data-action="load-saved">Load Saved Settings</button>
          <button type="button" class="secondary-btn" data-action="export-settings">Export Settings</button>
          <button type="button" class="secondary-btn" data-action="restore-defaults">Restore Defaults</button>
          <button type="button" class="primary-btn" data-action="save-settings" ${validationError ? "disabled" : ""}>Save Settings</button>
        </div>
        <div class="detail-meta">
          <span class="state-pill">${dirty ? "Unsaved changes" : "Saved draft"}</span>
          <span class="state-pill">${escapeHtml(getSectionSummary(state, tab))}</span>
          ${validationError ? `<span class="state-pill warning-pill">${escapeHtml(validationError)}</span>` : ""}
        </div>
        ${renderDetailContent(state, tab)}
      </section>
    </div>
  `;
}

function renderWorkspace(state: SettingsWorkspaceState): string {
  normalizeSelection(state);
  return state.currentView === "index" ? renderIndex(state) : renderDetail(state, state.currentView);
}

export function initSettingsWorkspace(options: SettingsWorkspaceOptions): void {
  const root = document.getElementById("settings-workspace");
  if (!(root instanceof HTMLDivElement)) {
    return;
  }

  const state: SettingsWorkspaceState = {
    draft: cloneResolvedFormatSettings(options.initialSettings),
    savedSnapshot: cloneResolvedFormatSettings(options.initialSettings),
    currentView: "index",
    selectedIndex: {
      numberFormats: null,
      dateFormats: null,
      cellFormats: null,
      textStyles: null,
    },
  };

  const render = (): void => {
    root.innerHTML = renderWorkspace(state);
  };

  const runAsync = async (work: () => Promise<void>): Promise<void> => {
    try {
      await work();
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      options.onStatus(`Error: ${message}`);
    }
  };

  root.addEventListener("click", (event) => {
    const target = event.target instanceof Element ? event.target.closest("[data-action]") : null;
    if (!(target instanceof HTMLElement)) {
      return;
    }

    const action = target.dataset.action;
    if (!action) {
      return;
    }

    syncDraftFromDom(root, state);

    if (action === "open-section") {
      const nextTab = target.dataset.tab as SettingsTabKey | undefined;
      if (nextTab) {
        state.currentView = nextTab;
        render();
      }
      return;
    }

    if (action === "go-home") {
      state.currentView = "index";
      render();
      return;
    }

    if (!isDetailView(state.currentView)) {
      return;
    }

    if (action === "select-item" && isListTab(state.currentView)) {
      const index = Number(target.dataset.index ?? "0");
      const normalizedIndex = clampIndex(index, getListItems(state, state.currentView).length);
      state.selectedIndex[state.currentView] =
        state.selectedIndex[state.currentView] === normalizedIndex ? null : normalizedIndex;
      render();
      return;
    }

    if (action === "add-item" && isListTab(state.currentView)) {
      addPreset(state, state.currentView);
      options.onStatus(`${TAB_META[state.currentView].title.slice(0, -1)} preset added.`);
      render();
      return;
    }

    if (action === "move-item" && isListTab(state.currentView)) {
      const index = Number(target.dataset.index ?? "0");
      const direction = Number(target.dataset.direction ?? "0");
      movePreset(state, state.currentView, index, direction);
      render();
      return;
    }

    if (action === "delete-item" && isListTab(state.currentView)) {
      const index = Number(target.dataset.index ?? "0");
      deletePreset(state, state.currentView, index);
      render();
      return;
    }

    if (action === "load-saved") {
      void runAsync(async () => {
        const imported = await options.loadSavedSettings();
        state.draft = cloneResolvedFormatSettings(imported);
        normalizeSelection(state);
        options.onStatus("Settings file loaded. Click Save Settings to apply it to this workbook.");
        render();
      });
      return;
    }

    if (action === "export-settings") {
      void runAsync(async () => {
        await options.exportSettings(cloneResolvedFormatSettings(state.draft));
        options.onStatus("Settings exported.");
      });
      return;
    }

    if (action === "restore-defaults") {
      state.draft = buildDefaultFormatSettings();
      normalizeSelection(state);
      options.onStatus("Built-in defaults loaded into the editor. Save Settings to apply them.");
      render();
      return;
    }

    if (action === "save-settings") {
      void runAsync(async () => {
        const validationError = getFormatSettingsValidationError(state.draft);
        if (validationError) {
          options.onStatus(validationError);
          render();
          return;
        }

        const toSave = cloneResolvedFormatSettings(state.draft);
        await options.saveSettings(toSave);
        state.savedSnapshot = cloneResolvedFormatSettings(toSave);
        options.onStatus("Workbook settings saved. Text style cycle reset.");
        render();
      });
    }
  });

  root.addEventListener("change", (event) => {
    const target = event.target;
    if (
      !(
        target instanceof HTMLInputElement ||
        target instanceof HTMLSelectElement ||
        target instanceof HTMLTextAreaElement
      )
    ) {
      return;
    }

    const action = target.dataset.action;
    if (!action) {
      return;
    }

    if (action === "update-list-field") {
      updateSelectedListField(state, String(target.dataset.field ?? ""), parseEditorValue(target));
      render();
      return;
    }

    if (action === "update-palette") {
      const key = target.dataset.key as keyof AutoColorPalette | undefined;
      if (key) {
        state.draft.autoColorPalette[key] = target.value;
        render();
      }
      return;
    }

    if (action === "update-trace") {
      const key = target.dataset.key as keyof ResolvedFormatSettings["trace"] | undefined;
      if (key) {
        state.draft.trace[key] = target.value === "" ? NaN : Number(target.value);
        render();
      }
    }
  });

  render();
}
