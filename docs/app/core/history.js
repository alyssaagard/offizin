/**
 * Offizin history: undo and redo for the changes Offizin makes to a document.
 *
 * Word and Google Docs keep their own undo lists. Offizin keeps a separate one that records only
 * what Offizin did: creating a publication, inserting a page part, applying a design, and
 * changing the page setup or columns. Undoing an entry asks the host to reverse exactly that
 * change, so the editor's own undo list and the user's own edits are left alone.
 *
 * The history knows nothing about any editor. It is created with a "performer", usually the
 * host itself, that has these methods (each may return a promise):
 *
 *   undo(op)       reverses op. Returns a message string or { message, op? }.
 *   redo(op)       applies op again. Returns a message string or { message, op? }. A returned op
 *                  replaces the stored one, for example when the editor gave a part a new anchor.
 *   discard(ops)   optional. Called with the ops that leave the history for good (a redo branch
 *                  is dropped, the oldest entries are trimmed, or the history is cleared), so the
 *                  host can remove anchors or snapshots it kept for them.
 *
 * Ops come from the host: { kind, label, data }. Kinds are "create", "insert", "restyle",
 * "pageSetup", and "columns". The history keeps at most `limit` entries (50 by default).
 *
 * Positions: entries[0 .. position - 1] are applied and entries[position ..] are undone. Moving
 * to position 0 undoes everything; moving to entries.length redoes everything.
 *
 * This module has no browser dependencies, so it runs in Node for tests.
 */

export const HISTORY_LIMIT = 50;

const KIND_LABELS = {
  create: "Create publication",
  insert: "Insert page part",
  restyle: "Apply design",
  pageSetup: "Change page setup",
  columns: "Change columns",
};

/** A plain label for an op that arrived without one. */
export function defaultLabel(kind) {
  return KIND_LABELS[kind] || "Offizin change";
}

/** True for an op the history can store: an object with a kind. */
export function isOp(op) {
  return !!op && typeof op === "object" && !Array.isArray(op) && typeof op.kind === "string" && op.kind.trim() !== "";
}

/**
 * Turns any host result into { message, op }. Hosts return a message string (contract version 1)
 * or { message, op } (contract version 2). Anything else counts as "no message, no op".
 */
export function normalizeResult(result, fallback = "") {
  if (typeof result === "string") return { message: result.trim() || fallback, op: null };
  if (result && typeof result === "object" && !Array.isArray(result)) {
    const message = typeof result.message === "string" && result.message.trim() ? result.message.trim() : fallback;
    return { message, op: isOp(result.op) ? result.op : null };
  }
  return { message: fallback, op: null };
}

/** The readable text of an error from a host, including Apps Script failures. */
export function errorMessage(e) {
  if (!e) return "Something went wrong.";
  if (typeof e === "string") return e;
  return e.message || String(e) || "Something went wrong.";
}

/**
 * Creates a history.
 * @param performer  { undo(op), redo(op), discard?(ops) }, usually the host.
 * @param options    { limit, now } where now() returns the time in milliseconds (for tests).
 */
export function createHistory(performer, options = {}) {
  const limit = Math.max(1, Math.floor(Number(options.limit) || HISTORY_LIMIT));
  const clock = typeof options.now === "function" ? options.now : () => Date.now();
  let entries = [];
  let position = 0;
  let busy = false;
  let counter = 0;
  const listeners = new Set();

  const view = (entry) => {
    const index = entries.indexOf(entry);
    return { id: entry.id, label: entry.label, kind: entry.kind, at: entry.at, done: index >= 0 && index < position, failed: entry.failed };
  };

  function snapshot() {
    const last = position > 0 ? entries[position - 1] : null;
    const next = position < entries.length ? entries[position] : null;
    return {
      entries: entries.map(view),
      position,
      size: entries.length,
      limit,
      busy,
      canUndo: !!last && !busy,
      canRedo: !!next && !busy,
      undoLabel: last ? last.label : "",
      redoLabel: next ? next.label : "",
      undoKind: last ? last.kind : "",
      redoKind: next ? next.kind : "",
    };
  }

  function emit() {
    const snap = snapshot();
    for (const fn of [...listeners]) {
      try {
        fn(snap);
      } catch (e) {
        // A listener that fails must not break the history for everyone else.
      }
    }
  }

  /** Hands ops that leave the history to the host. Cleanup is best effort and never throws. */
  function discard(list) {
    const ops = list.map((e) => e.op).filter(Boolean);
    if (!ops.length || !performer || typeof performer.discard !== "function") return;
    try {
      const pending = performer.discard(ops);
      if (pending && typeof pending.then === "function") pending.then(null, () => {});
    } catch (e) {
      // The entries are gone from the history either way.
    }
  }

  function guard() {
    if (busy) throw new Error("Offizin is still working on the last change. Try again in a moment.");
  }

  /** Records an op that the host just performed. Drops any undone entries (the redo branch). */
  function push(op, label) {
    guard();
    if (!isOp(op)) throw new Error("Offizin can only record changes reported by the editor.");
    const dropped = entries.slice(position);
    entries = entries.slice(0, position);
    const entry = {
      id: "h" + ++counter,
      op,
      label: String(label || op.label || "").trim() || defaultLabel(op.kind),
      kind: op.kind,
      at: clock(),
      failed: "",
    };
    entries.push(entry);
    let trimmed = [];
    if (entries.length > limit) {
      trimmed = entries.slice(0, entries.length - limit);
      entries = entries.slice(entries.length - limit);
    }
    position = entries.length;
    discard(dropped.concat(trimmed));
    emit();
    return view(entry);
  }

  /** One step back (direction -1) or forward (+1). The caller holds the busy flag. */
  async function perform(direction) {
    const index = direction < 0 ? position - 1 : position;
    const entry = entries[index];
    if (!entry) return { ok: false, message: direction < 0 ? "There is nothing to undo." : "There is nothing to redo.", entry: null };
    const method = direction < 0 ? "undo" : "redo";
    if (!performer || typeof performer[method] !== "function") {
      throw new Error(`This editor cannot ${method} Offizin changes.`);
    }
    try {
      const result = await performer[method](entry.op);
      const { message, op } = normalizeResult(result, "");
      // A refreshed op replaces the old one. The host manages the old op's anchors itself.
      if (op) entry.op = op;
      entry.failed = "";
      position = direction < 0 ? index : index + 1;
      emit();
      return { ok: true, message: message || `${direction < 0 ? "Undone" : "Redone"}: ${entry.label}.`, entry: view(entry) };
    } catch (e) {
      entry.failed = errorMessage(e);
      emit();
      throw e;
    }
  }

  async function exclusive(fn) {
    guard();
    busy = true;
    emit();
    try {
      return await fn();
    } finally {
      busy = false;
      emit();
    }
  }

  /** Undoes the most recent applied entry. */
  function undo() {
    return exclusive(() => perform(-1));
  }

  /** Redoes the most recent undone entry. */
  function redo() {
    return exclusive(() => perform(1));
  }

  /**
   * Moves to a position by undoing or redoing one entry at a time. Stops at the first failure
   * (and rethrows it); the entries already moved stay moved.
   */
  function jump(target) {
    const goal = Math.floor(Number(target));
    if (!Number.isFinite(goal)) return Promise.reject(new Error("That is not a place in the history."));
    return exclusive(async () => {
      const to = Math.max(0, Math.min(entries.length, goal));
      const back = to < position;
      const steps = [];
      while (position > to) steps.push(await perform(-1));
      while (position < to) steps.push(await perform(1));
      const last = steps[steps.length - 1];
      return {
        ok: true,
        steps: steps.length,
        messages: steps.map((s) => s.message),
        message: !steps.length ? "Nothing changed." : steps.length === 1 ? last.message : `${steps.length} changes ${back ? "undone" : "redone"}.`,
      };
    });
  }

  /** The entries a jump to `target` would undo (newest first) and redo (oldest first). */
  function plan(target) {
    const to = Math.max(0, Math.min(entries.length, Math.floor(Number(target)) || 0));
    return {
      undo: entries.slice(to, position).reverse().map(view),
      redo: entries.slice(position, to).map(view),
    };
  }

  /** The position just after an entry, which is where a jump "to" that entry goes. */
  function positionOf(id) {
    const i = entries.findIndex((e) => e.id === id);
    return i < 0 ? -1 : i + 1;
  }

  /** Removes one entry, for example one the editor can no longer undo. */
  function forget(id) {
    guard();
    const i = entries.findIndex((e) => e.id === id);
    if (i < 0) return false;
    const [entry] = entries.splice(i, 1);
    if (i < position) position--;
    discard([entry]);
    emit();
    return true;
  }

  /** Empties the history and lets the host clean up. */
  function clear() {
    guard();
    const all = entries;
    entries = [];
    position = 0;
    discard(all);
    emit();
  }

  /** Calls fn(snapshot) after every change. Returns a function that stops the calls. */
  function onChange(fn) {
    if (typeof fn !== "function") return () => {};
    listeners.add(fn);
    return () => listeners.delete(fn);
  }

  return {
    push,
    undo,
    redo,
    jump,
    plan,
    positionOf,
    forget,
    clear,
    onChange,
    snapshot,
    get entries() {
      return entries.map(view);
    },
    get position() {
      return position;
    },
    get size() {
      return entries.length;
    },
    get limit() {
      return limit;
    },
    get busy() {
      return busy;
    },
    get canUndo() {
      return position > 0 && !busy;
    },
    get canRedo() {
      return position < entries.length && !busy;
    },
  };
}
