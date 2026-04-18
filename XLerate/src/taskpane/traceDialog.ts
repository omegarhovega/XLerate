/* global Office */

// Phase B skeleton: confirms end-to-end dialog lifecycle. Logic (trace
// computation, keyboard nav, messageParent protocol) lands in subsequent
// commits of the plan. For now this file just proves that the webpack
// entry emits a valid HTML page Office.js can load as a dialog.

function setDialogStatus(message: string): void {
  const el = document.getElementById("trace-dialog-status");
  if (el) el.textContent = message;
}

function parseDialogParams(): { direction: string; address: string } {
  const params = new URLSearchParams(window.location.search);
  return {
    direction: params.get("direction") ?? "precedents",
    address: params.get("address") ?? "",
  };
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) {
    setDialogStatus("Trace dialog requires Excel.");
    return;
  }

  const { direction, address } = parseDialogParams();
  const title = document.getElementById("trace-dialog-title");
  if (title) {
    title.textContent = `Trace ${direction === "dependents" ? "dependents" : "precedents"}`;
  }
  setDialogStatus(
    address
      ? `Ready to trace ${direction} from ${address}. (Trace logic lands in the next commit.)`
      : `Ready to trace ${direction}. (No start address supplied.)`
  );
});
