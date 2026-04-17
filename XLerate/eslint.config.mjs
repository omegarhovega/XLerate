import officeAddins from "eslint-plugin-office-addins";
import tsParser from "@typescript-eslint/parser";

export default [
  ...officeAddins.configs.recommended,
  {
    plugins: {
      "office-addins": officeAddins,
    },
    languageOptions: {
      parser: tsParser,
    },
  },
  // Live adapter: the one intentional Office.js file. Declare the Excel global
  // and disable office-addins load/sync rules that produce false positives for
  // correctly-written Excel.run() code (we do call load + sync in the right
  // order; the plugin just can't track navigation properties like range.worksheet).
  {
    files: ["src/adapters/excelPortLive.ts"],
    languageOptions: {
      globals: {
        Excel: "readonly",
      },
    },
    rules: {
      "no-undef": "off",
      "office-addins/call-sync-before-read": "off",
      "office-addins/load-object-before-read": "off",
    },
  },
  // Boundary rule: domain code (core/, services/) must not import office-js
  // or reference the Excel / Office globals. Use the ExcelPort interface instead.
  {
    files: ["src/core/**/*.ts", "src/services/**/*.ts"],
    rules: {
      "no-restricted-imports": [
        "error",
        {
          paths: [
            {
              name: "office-js",
              message:
                "Domain code (core/, services/) must not import office-js. Use the ExcelPort interface from src/adapters/excelPort.ts instead.",
            },
          ],
        },
      ],
      "no-restricted-globals": [
        "error",
        {
          name: "Excel",
          message: "Domain code must not reference the Excel global. Use ExcelPort.",
        },
        {
          name: "Office",
          message: "Domain code must not reference the Office global. Use ExcelPort.",
        },
      ],
    },
  },
];
