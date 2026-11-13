/**
 * Offizin studio keyboard shortcuts and the state of the Undo and Redo buttons.
 * Plain logic with no browser dependencies, so the tests run it in Node.
 *
 *   Ctrl+Z or Command+Z            Offizin undo
 *   Ctrl+Shift+Z or Command+Shift+Z, Ctrl+Y   Offizin redo
 *   + (or =) and -                 zoom in and out
 *   0                              fit the page to the pasteboard
 *   G                              show or hide guides
 *   P                              print preview
 *
 * While you type in a text field, the field keeps these keys (so Ctrl+Z undoes your typing).
 * The single keys (plus, minus, 0, G, and P) work only while the pasteboard or the page
 * navigator has focus, so they never surprise anyone working in the panel (WCAG 2.1.4).
 * Ctrl and Command shortcuts work anywhere outside a text field.
 */

export const SHORTCUTS = [
  { keys: "Ctrl+Z", action: "undo", text: "Undo the last Offizin change" },
  { keys: "Ctrl+Shift+Z", action: "redo", text: "Redo it" },
  { keys: "+", action: "zoom-in", text: "Zoom in" },
  { keys: "-", action: "zoom-out", text: "Zoom out" },
  { keys: "0", action: "zoom-fit", text: "Fit the page" },
  { keys: "G", action: "guides", text: "Show or hide guides" },
  { keys: "P", action: "print-preview", text: "Print preview" },
];

const TEXT_INPUTS = new Set(["text", "search", "email", "url", "tel", "number", "password", "date", "month", "time", "week", "datetime-local"]);

/** True when key presses belong to the element: text fields, text areas, selects, editable text. */
export function isEditableTarget(t) {
  if (!t) return false;
  if (t.isContentEditable) return true;
  const tag = String(t.tagName || "").toUpperCase();
  if (tag === "TEXTAREA" || tag === "SELECT") return true;
  if (tag === "INPUT") return TEXT_INPUTS.has(String(t.type || "text").toLowerCase());
  return false;
}

/** Where the single-key shortcuts are active: the pasteboard and the page navigator. */
export const SHORTCUT_SCOPE = "#st-board, #st-nav-list";

/** True when the element is inside the pasteboard or the page navigator (or is one of them). */
export function inShortcutScope(t) {
  return !!(t && typeof t.closest === "function" && t.closest(SHORTCUT_SCOPE));
}

/**
 * What Escape does in the studio: "menu" closes the install menu, "sheet" closes the panel
 * sheet on a phone, and null leaves the key alone. A press the panel already handled
 * (defaultPrevented, for its confirm or its history list) is left alone, so one Escape
 * closes one thing.
 */
export function escapeAction(e, { menuOpen = false, sheetOpen = false, phone = false } = {}) {
  if (!e || e.key !== "Escape" || e.defaultPrevented) return null;
  if (menuOpen) return "menu";
  if (sheetOpen && phone) return "sheet";
  return null;
}

/**
 * The studio action for a key event, or null. `e` needs key, ctrlKey, metaKey, shiftKey,
 * altKey, and target (all optional). Single keys count only inside SHORTCUT_SCOPE.
 */
export function shortcutFor(e) {
  if (!e || e.defaultPrevented || e.isComposing) return null;
  const key = String(e.key || "");
  const lower = key.toLowerCase();
  const mod = !!(e.ctrlKey || e.metaKey);
  if (isEditableTarget(e.target)) return null;
  if (mod && !e.altKey) {
    if (lower === "z") return e.shiftKey ? "redo" : "undo";
    if (lower === "y" && e.ctrlKey && !e.metaKey && !e.shiftKey) return "redo";
    return null;
  }
  if (mod || e.altKey) return null;
  if (!inShortcutScope(e.target)) return null;
  if (key === "+" || key === "=") return "zoom-in";
  if (key === "-" || key === "_") return "zoom-out";
  if (key === "0") return "zoom-fit";
  if (lower === "g" && !e.shiftKey) return "guides";
  if (lower === "p" && !e.shiftKey) return "print-preview";
  return null;
}

/**
 * The Undo and Redo buttons' state from the panel API. The panel offers undo(), redo(), and
 * onHistoryChange(fn); without all three, the buttons stay hidden. `snap` is the history
 * snapshot the panel reported ({ canUndo, canRedo, undoLabel, redoLabel, busy }).
 */
export function historyView(panel, snap) {
  const supported = !!panel && typeof panel.undo === "function" && typeof panel.redo === "function" && typeof panel.onHistoryChange === "function";
  if (!supported) return { supported: false, busy: false, canUndo: false, canRedo: false, undoTitle: "", redoTitle: "" };
  let s = snap;
  if (!s && panel.history) {
    try {
      const h = typeof panel.history === "function" ? panel.history() : panel.history;
      s = h && typeof h.snapshot === "function" ? h.snapshot() : h;
    } catch (e) {
      s = null;
    }
  }
  s = s || {};
  const busy = !!s.busy;
  const canUndo = !busy && !!s.canUndo;
  const canRedo = !busy && !!s.canRedo;
  return {
    supported: true,
    busy,
    canUndo,
    canRedo,
    undoTitle: canUndo && s.undoLabel ? `Undo: ${s.undoLabel}` : "Nothing to undo",
    redoTitle: canRedo && s.redoLabel ? `Redo: ${s.redoLabel}` : "Nothing to redo",
  };
}
