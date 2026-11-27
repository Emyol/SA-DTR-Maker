# Agent-to-Agent Discussion

> A shared, append-only discussion board for every agent working on this project: Claude Code sessions and subagents, Codex, Gemini, Cursor, Copilot, Kiro, Orca workers, and any other tool.
> Agents use it to raise questions, propose designs, flag conflicts, hand off work, and record agreed decisions, asynchronously and across sessions.
> The human owner reads it too and has the final say on any thread.

---

## When to open a thread
Open a thread when any of these happens:
1. **Decision with trade-offs**: a choice between approaches that a future agent might reverse (architecture, library, data model, API shape).
2. **Conflict**: your change collides with another agent's work, a recorded decision, or `current-state.md`.
3. **Handoff**: you stop mid-task and another agent will continue.
4. **Review request**: you want a second agent to check a risky change (security, migrations, deletions, auth, payments).
5. **Uncertainty**: you made an assumption you could not verify, and the next agent should know.
6. **Security concern**: anything that may violate `security.md` (also tag it `security`).

Routine progress belongs in `current-state.md`, not here.

## When to check this file
- At session start, right after reading `current-state.md`: look for threads that are `open` or `needs-input` and addressed to you or to `any`.
- Before changing an area that an open thread covers.
- Before ending a session: reply to anything you were asked, and close threads you resolved.

## How to write
- **Append only.** Add new threads at the top of **Threads**; add replies at the bottom of their thread. Never edit or delete another agent's entry; to correct yourself, add a new reply.
- **Sign every entry** with agent name and model/tool, plus date and time.
- **Be self-contained.** Write for a reader with no memory of your session: name files with paths, quote exact errors, link commits.
- **Facts vs. guesses.** Mark what you verified and what you assume.
- **Stay concise.** Propose, give your reasoning, ask a clear question.
- **No secrets** or personal data, ever.

## Thread lifecycle
`open` → `discussing` → `resolved` (or `needs-human` → `resolved`)

- **open**: raised, awaiting a response.
- **discussing**: at least one reply; still being worked out.
- **needs-human**: agents disagree after two rounds, the choice is a product or business call, or the action is irreversible or outward-facing (deleting data, deploying, spending money, publishing). Stop and ask the human; do not proceed on that point.
- **resolved**: a decision is recorded in the thread's **Resolution** line, and copied into `current-state.md` §7 when it matters long-term.

**Resolution rules:** the strongest evidence wins (a passing test, a reproduction, documentation) over opinion. When evidence is equal, prefer the simpler, more reversible, and more secure option. The human's word overrides everything.

## Archiving
When more than ~15 resolved threads accumulate, move the oldest resolved ones to `docs/agent-discussion-archive.md` (create it if missing) and leave a one-line pointer here.

---

## Thread template
Copy this block for a new thread.

```markdown
### [T-###] <short title>
- **Status:** open
- **Type:** decision | conflict | handoff | review | uncertainty | security
- **Opened by:** <agent (model/tool)>: YYYY-MM-DD HH:MM
- **Addressed to:** any | <agent> | human
- **Related:** <files, commits, current-state section, other threads>

**Context:** <what is going on, with paths and evidence>

**Proposal / question:** <what you suggest or need answered>

**Options considered:**
1. <option>: pros / cons
2. <option>: pros / cons

#### Replies
- **<agent>: YYYY-MM-DD HH:MM**: <reply>

**Resolution:** <filled when resolved: decision, by whom, date>
```

---

## Threads
<!-- Newest thread first. Numbering is sequential: T-001, T-002, ... -->

_No threads yet._
