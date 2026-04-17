module.exports = {
  forbidden: [
    {
      name: "no-office-js-outside-adapters",
      severity: "error",
      comment:
        "office-js may only be imported from src/adapters/. Domain code depends on the ExcelPort interface.",
      from: { pathNot: "^src/adapters/" },
      to: { path: "office-js" }
    },
    {
      name: "no-taskpane-in-core-or-services",
      severity: "error",
      comment: "core/ and services/ must not import from taskpane/. The UI layer is the outer dependency.",
      from: { path: "^src/(core|services)/" },
      to: { path: "^src/taskpane/" }
    },
    {
      name: "no-adapters-in-core",
      severity: "error",
      comment: "core/ must be standalone pure domain logic and must not know about adapters.",
      from: { path: "^src/core/" },
      to: { path: "^src/adapters/" }
    },
    {
      name: "no-circular",
      severity: "error",
      comment: "Circular dependencies are forbidden anywhere in src/.",
      from: {},
      to: { circular: true }
    }
  ],
  options: {
    doNotFollow: { path: "node_modules" },
    tsConfig: { fileName: "tsconfig.json" },
    enhancedResolveOptions: { exportsFields: ["exports"], conditionNames: ["import", "require", "node"] }
  }
};
