import { describe, expect, it } from "vitest";
import {
  computeNextNumberFormat,
  DEFAULT_NUMBER_FORMATS,
  hasMixedNumberFormats
} from "../src/core/numberFormatCycle";

describe("cycle number format baseline parity", () => {
  it("detects mixed format selections", () => {
    const first = DEFAULT_NUMBER_FORMATS[0].formatCode;
    const second = DEFAULT_NUMBER_FORMATS[1].formatCode;

    expect(hasMixedNumberFormats([first])).toBe(false);
    expect(hasMixedNumberFormats([first, first, first])).toBe(false);
    expect(hasMixedNumberFormats([first, second])).toBe(true);
  });

  it("cycles to next configured format", () => {
    const current = DEFAULT_NUMBER_FORMATS[0].formatCode;
    const next = DEFAULT_NUMBER_FORMATS[1].formatCode;

    expect(computeNextNumberFormat(current, false)).toBe(next);
  });

  it("wraps to first format at the end", () => {
    const current = DEFAULT_NUMBER_FORMATS[DEFAULT_NUMBER_FORMATS.length - 1].formatCode;
    const next = DEFAULT_NUMBER_FORMATS[0].formatCode;

    expect(computeNextNumberFormat(current, false)).toBe(next);
  });

  it("uses first format when selection has mixed formats", () => {
    const current = DEFAULT_NUMBER_FORMATS[1].formatCode;
    const first = DEFAULT_NUMBER_FORMATS[0].formatCode;

    expect(computeNextNumberFormat(current, true)).toBe(first);
  });

  it("uses first format when current format is unknown", () => {
    const first = DEFAULT_NUMBER_FORMATS[0].formatCode;

    expect(computeNextNumberFormat("General", false)).toBe(first);
  });

  it("uses provided custom configured formats", () => {
    const customFormats = [
      { name: "Custom-1", formatCode: "0.0" },
      { name: "Custom-2", formatCode: "0.00" }
    ];

    expect(computeNextNumberFormat("0.0", false, customFormats)).toBe("0.00");
    expect(computeNextNumberFormat("0.00", false, customFormats)).toBe("0.0");
    expect(computeNextNumberFormat("General", false, customFormats)).toBe("0.0");
  });
});
