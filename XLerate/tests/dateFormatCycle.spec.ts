import { describe, expect, it } from "vitest";
import { computeNextDateFormat, DEFAULT_DATE_FORMATS, hasMixedDateFormats } from "../src/core/dateFormatCycle";

describe("cycle date format baseline parity", () => {
  it("detects mixed format selections", () => {
    const first = DEFAULT_DATE_FORMATS[0].formatCode;
    const second = DEFAULT_DATE_FORMATS[1].formatCode;

    expect(hasMixedDateFormats([first])).toBe(false);
    expect(hasMixedDateFormats([first, first, first])).toBe(false);
    expect(hasMixedDateFormats(["dd-mmm-yy", "d-mmm-yy"])).toBe(false);
    expect(hasMixedDateFormats([first, second])).toBe(true);
  });

  it("cycles to next configured format", () => {
    const current = DEFAULT_DATE_FORMATS[0].formatCode;
    const next = DEFAULT_DATE_FORMATS[1].formatCode;

    expect(computeNextDateFormat(current, false)).toBe(next);
  });

  it("wraps to first format at the end", () => {
    const current = DEFAULT_DATE_FORMATS[DEFAULT_DATE_FORMATS.length - 1].formatCode;
    const next = DEFAULT_DATE_FORMATS[0].formatCode;

    expect(computeNextDateFormat(current, false)).toBe(next);
  });

  it("uses first format when selection has mixed formats", () => {
    const current = DEFAULT_DATE_FORMATS[1].formatCode;
    const first = DEFAULT_DATE_FORMATS[0].formatCode;

    expect(computeNextDateFormat(current, true)).toBe(first);
  });

  it("uses first format when current format is unknown", () => {
    const first = DEFAULT_DATE_FORMATS[0].formatCode;

    expect(computeNextDateFormat("General", false)).toBe(first);
  });

  it("uses provided custom configured formats", () => {
    const customFormats = [
      { name: "Custom-1", formatCode: "yyyy-mm" },
      { name: "Custom-2", formatCode: "dd/mm/yyyy" }
    ];

    expect(computeNextDateFormat("yyyy-mm", false, customFormats)).toBe("dd/mm/yyyy");
    expect(computeNextDateFormat("dd/mm/yyyy", false, customFormats)).toBe("yyyy-mm");
    expect(computeNextDateFormat("General", false, customFormats)).toBe("yyyy-mm");
  });

  it("treats Excel-normalized day tokens as the same preset", () => {
    const customFormats = [
      { name: "Year Only", formatCode: "yyyy" },
      { name: "Month Year", formatCode: "mmm-yyyy" },
      { name: "Full Date", formatCode: "dd-mmm-yy" },
      { name: "Verbose", formatCode: "dddd-mm-yyyy" }
    ];

    expect(computeNextDateFormat("d-mmm-yy", false, customFormats)).toBe("dddd-mm-yyyy");
  });
});
