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
