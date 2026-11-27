# Current State

> The project's working memory. Read this first at the start of every session; update it at the end of every meaningful unit of work and before the session ends.
> Keep it **true and short**: rewrite stale sections instead of appending to them. History goes in the Session Log, not in the sections above it.

**Last updated:** 2026-11-27 (UTC+8): by Claude Code
**Branch:** `main` · **Last commit:** `668957b`: changed facilitator name (template/defaults update below is uncommitted)
**Overall status:** 🟢 on track

---

## 1. Project goal
DTR Helper: turns a spreadsheet (or pasted table) of student-assistant time logs into a filled, printable Word DTR form (`EMPTY_DTR.docx`), via a Flask web app or CLI. Done for the current milestone = forms carry the correct 1st term 2026-2027 signatories.

## 2. Current milestone
- **Milestone:** 1st term 2026-2027 template refresh
- **Definition of done:** template + defaults updated; live site (https://sa-dtr-maker.vercel.app/) redeployed with them
- **Progress:** template + defaults updated and verified locally; not yet committed or redeployed

## 3. What works (verified)
- Upload `SAMPLE DTR.xlsx` → `/generate` returns a docx with 1st / 2026-2027 / Ms. Maria Fatima Dela Cruz headers and the new signatories in all 18 forms: verified 2026-11-27 via Flask test client + Word PDF export.
- Live site matched commit `668957b` byte-for-byte (homepage, CSS, generated docx): verified 2026-11-27.

## 4. In progress
- Template/defaults change: `EMPTY_DTR.docx`, `dtr/defaults.py`, `README.md`: next action: commit, push, and have the Vercel project owner redeploy.

## 5. Next up (prioritised)
1. Redeploy https://sa-dtr-maker.vercel.app/ after the change is pushed.

## 6. Blockers and open questions
- The Vercel project behind sa-dtr-maker.vercel.app is not on the owner's "Emyel's projects" team; likely hosted by collaborator Romeo-04. Redeploys depend on that account.
- User typed "Enr. Kennedy P. Limbo" for the Senior Director; applied as "Engr." (matches his prior title in the template). Confirm.

## 7. Key decisions
| Date | Decision | Why | Alternatives rejected |
|------|----------|-----|-----------------------|
| 2026-11-27 | Edit `EMPTY_DTR.docx` by swapping `<w:t>` run text in `word/document.xml` only | Keeps all formatting/layout | Re-saving in Word: risks layout drift |

## 8. Known issues and tech debt
- SA Facilitator line still reads "Ms. Ma. Fatima Dela Cruz" (not requested to change): severity: low.

## 9. Environment and how to run
- Run web app: `python run_web.py`
- CLI: `python dtr_generator.py` (see README)
- Deploy: Vercel Python runtime via `api/index.py` + `vercel.json`
- Gotchas: Term/School Year/Supervisor header lines are overwritten at generation time from `dtr/defaults.py`; signature-block names are static text in `EMPTY_DTR.docx`. Change both places together.
- Required env vars: none.

## 10. Architecture snapshot
- `dtr/engine.py`: parses xlsx/pasted input, groups into Mon–Sat weeks, fills the docx.
- `dtr/defaults.py`: identity defaults and constants.
- `webapp/`: Flask blueprint (`/`, `/generate`), templates, CSS.
- `api/index.py`: Vercel entrypoint wrapping the same Flask app.

---

## Session log

### 2026-11-27: Claude Code
- Did: verified live Vercel site matches repo; updated template + defaults to 1st term 2026-2027 signatories; created project docs.
- Verified: Flask test-client generation + Word PDF render of template.
- Left open: commit/push/redeploy; confirm "Engr." for Kennedy P. Limbo.
- Next session should: check whether the live site has been redeployed.

---

## Update rules (for agents)
- **When:** after finishing a feature or fix, after any decision in §7, when a blocker appears or clears, and before ending a session or handing off.
- **How:** edit sections 1–10 so they describe *now*, then add one Session Log entry. Remove anything no longer true.
- **Honesty:** list something under "What works" only if it was run or tested. Partial or unverified work goes under "In progress", with that stated.
- **No secrets:** record env var *names*, never values, tokens, keys, or personal data.
