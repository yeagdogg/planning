"""The undo stack behind the State Summary's live levers (W2c).

John's condition on fan-out edits was "as long as there's an easy way to
undo it" — so every lever edit pushes its OLD values here before the
mutation, and one button walks back the stack. Entries locate rows by
CONTENT match (never list position): the Inputs page replaces row lists
wholesale on its own edits, and a positional locator would silently
restore into the wrong row. A row that can no longer be matched exactly
once is SKIPPED LOUDLY — the pop reports it by name rather than guessing.

The stack clears on ctx.data_rev (scenario load / master paste replace
the world; locators would dangle). Scenario files remain the durable
history — this is the what-if eraser, not the audit trail.
"""
from __future__ import annotations

from dataclasses import dataclass, field


@dataclass(frozen=True)
class UndoEntry:
    """One field restoration: the row is found by matching every key in
    ``match`` (content AFTER the edit, so the current state locates it),
    then ``fields`` are restored to their old values."""
    lob: str
    table: str                 # 'lr_rows' | 'rate_rows' | 'mod_rows'
    match: tuple               # ((key, value-as-str), ...) — identity
    fields: tuple              # ((field, old value), ...)


def entry(lob: str, table: str, row: dict, fields: dict,
          after: dict | None = None) -> UndoEntry:
    """Build an entry from the row AS IT WILL LOOK after the edit.

    ``fields`` maps field -> OLD value; ``after`` overrides the match
    identity for fields the edit itself changes (matching must use the
    post-edit content or pop can never find the row)."""
    ident = dict(row)
    ident.update(after or {})
    match = tuple(sorted((k, repr(v)) for k, v in ident.items()))
    return UndoEntry(lob=lob, table=table, match=match,
                     fields=tuple(sorted(fields.items(),
                                         key=lambda kv: kv[0])))


class UndoStack:
    """LIFO, capped; one push per user gesture (a fan-out is ONE entry
    list)."""

    def __init__(self, cap: int = 20):
        self.cap = cap
        self._stack: list = []          # (description, [UndoEntry, ...])

    def push(self, description: str, entries: list) -> None:
        if not entries:
            return
        self._stack.append((description, list(entries)))
        del self._stack[:-self.cap]

    def peek(self) -> str | None:
        return self._stack[-1][0] if self._stack else None

    def clear(self) -> None:
        self._stack.clear()

    def __len__(self) -> int:
        return len(self._stack)

    def pop_apply(self, session) -> tuple:
        """Undo the top gesture. Returns (restored count, [skip messages]).
        Rows are matched by full content, exactly-once: zero or several
        matches skip that entry LOUDLY (the row was edited or replaced
        since — restoring by guess would corrupt)."""
        if not self._stack:
            return 0, ["nothing to undo"]
        _desc, entries = self._stack.pop()
        restored, skipped = 0, []
        for e in entries:
            scn = session.page.book.get(e.lob)
            if scn is None:
                skipped.append(f"{e.lob}: line no longer loaded")
                continue
            want = dict(e.match)
            rows = [r for r in getattr(scn, e.table)
                    if all(repr(r.get(k)) == v for k, v in want.items())]
            if len(rows) != 1:
                skipped.append(
                    f"{e.lob} {e.table}: row changed since the edit "
                    f"({len(rows)} matches) — left as-is")
                continue
            for f, old in e.fields:
                rows[0][f] = old
            restored += 1
        session.bus.bump()
        return restored, skipped
