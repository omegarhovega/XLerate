import { CellFormatMutation, CellMutation, ExcelPort } from "../adapters/excelPort";
import {
  computeNextTextStyle,
  DEFAULT_TEXT_STYLES,
  mapBorderWeight,
  TextStyleDefinition,
} from "../core/textStyleCycle";

function styleToMutation(style: TextStyleDefinition): CellFormatMutation {
  const weight = mapBorderWeight(style.borderStyle);
  const edge =
    style.borderStyle === "None"
      ? undefined
      : { style: style.borderStyle, color: style.fontColor, weight };

  return {
    font: {
      name: style.fontName,
      size: style.fontSize,
      color: style.fontColor,
      bold: style.bold,
      italic: style.italic,
      underline: style.underline,
    },
    fill: { pattern: "Solid", color: style.backColor },
    borders: {
      clearAll: true,
      top: edge && style.borderTop ? edge : undefined,
      bottom: edge && style.borderBottom ? edge : undefined,
      left: edge && style.borderLeft ? edge : undefined,
      right: edge && style.borderRight ? edge : undefined,
    },
  };
}

/**
 * Implements spec §3.10. Advances the cycle index and applies the new
 * text style to the selection. Returns the new index so the caller can
 * persist it for the next invocation (session-scoped per spec §4.2).
 */
export async function runCycleTextStyle(
  port: ExcelPort,
  currentIndex: number,
  configuredStyles: TextStyleDefinition[] = DEFAULT_TEXT_STYLES
): Promise<{ index: number; style: TextStyleDefinition }> {
  const snaps = await port.getSelectionFormatting();
  if (snaps.length === 0) {
    return {
      index: currentIndex,
      style: configuredStyles[Math.max(currentIndex, 0)],
    };
  }

  const { index, style } = computeNextTextStyle(currentIndex, configuredStyles);
  const mutation = styleToMutation(style);

  const mutations: CellMutation[] = snaps.map((s) => ({
    address: s.address,
    kind: "formatBundle",
    format: mutation,
  }));
  await port.applyMutations(mutations);

  return { index, style };
}
