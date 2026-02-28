import { describe, expect, it } from "vitest";
import {
  computeNextTextStyle,
  computeNextTextStyleIndex,
  DEFAULT_TEXT_STYLES,
  mapBorderWeight
} from "../src/core/textStyleCycle";

describe("cycle text style baseline parity", () => {
  it("increments and wraps style index", () => {
    expect(computeNextTextStyleIndex(-1)).toBe(0);
    expect(computeNextTextStyleIndex(0)).toBe(1);
    expect(computeNextTextStyleIndex(1)).toBe(2);
    expect(computeNextTextStyleIndex(2)).toBe(3);
    expect(computeNextTextStyleIndex(3)).toBe(0);
  });

  it("returns the matching next style payload", () => {
    const first = computeNextTextStyle(-1);
    expect(first.index).toBe(0);
    expect(first.style.name).toBe("Heading");

    const second = computeNextTextStyle(first.index);
    expect(second.index).toBe(1);
    expect(second.style.name).toBe("Subheading");
  });

  it("uses default style count", () => {
    expect(DEFAULT_TEXT_STYLES).toHaveLength(4);
    expect(DEFAULT_TEXT_STYLES[3].name).toBe("Normal");
  });

  it("maps border weight from style", () => {
    expect(mapBorderWeight("Continuous")).toBe("Medium");
    expect(mapBorderWeight("Double")).toBe("Thick");
    expect(mapBorderWeight("Dash")).toBe("Thin");
    expect(mapBorderWeight("Dot")).toBe("Thin");
    expect(mapBorderWeight("None")).toBe("Thin");
  });

  it("uses provided custom configured styles", () => {
    const customStyles = [
      {
        name: "Body",
        fontName: "Calibri",
        fontSize: 11,
        bold: false,
        italic: false,
        underline: false,
        fontColor: "#000000",
        backColor: "#FFFFFF",
        borderStyle: "None" as const,
        borderTop: false,
        borderBottom: false,
        borderLeft: false,
        borderRight: false
      },
      {
        name: "Emphasis",
        fontName: "Calibri",
        fontSize: 12,
        bold: true,
        italic: false,
        underline: false,
        fontColor: "#1F4E78",
        backColor: "#EAF2F8",
        borderStyle: "Continuous" as const,
        borderTop: false,
        borderBottom: true,
        borderLeft: false,
        borderRight: false
      }
    ];

    expect(computeNextTextStyle(-1, customStyles).style.name).toBe("Body");
    expect(computeNextTextStyle(0, customStyles).style.name).toBe("Emphasis");
    expect(computeNextTextStyle(1, customStyles).style.name).toBe("Body");
  });
});
