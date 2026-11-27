# Security

> The security baseline for this project. Every agent and contributor applies it when writing, reviewing, or deploying code.
> **Part A** is project-specific: fill it in when the project is created and keep it current. **Part B** is the universal baseline, applied in full unless Part A records a justified exception.
> This file never contains real secrets, keys, tokens, internal URLs with credentials, or personal data.

---

# Part A: Project security profile

## A1. Assets and data classification
| Data / asset | Classification | Where stored | Who can access |
|--------------|----------------|--------------|----------------|
| Uploaded time logs + student names | Confidential (PII) | Request memory only; not persisted | The uploading user |
| `EMPTY_DTR.docx` (staff names, signatory titles) | Internal | Repo root | Repo collaborators |

Classes: **Public** · **Internal** · **Confidential** (PII, business data) · **Restricted** (credentials, payment data, health data, secrets).

## A2. Trust boundaries and attack surface
- Entry points: `GET /`, `POST /generate` (multipart .xlsx upload or pasted text, form fields capped at 200 chars)
- Third parties: Vercel (hosting) only
- Privileged paths: none (no auth, no admin, no secrets)

## A3. Threat model (STRIDE quick pass)
| Threat | Where it applies here | Mitigation | Status |
|--------|----------------------|------------|--------|
| **S**poofing | No accounts | n/a | n/a |
| **T**ampering | Crafted xlsx | openpyxl parse, row/label caps in `dtr/defaults.py` | done |
| **R**epudiation | n/a | n/a | n/a |
| **I**nformation disclosure | Uploaded PII | Processed in memory, docx returned inline, not stored | done |
| **D**enial of service | Huge/pathological workbooks | `MAX_DATA_ROWS`, `MAX_LABEL_SEARCH_ROWS` caps | done |
| **E**levation of privilege | n/a | n/a | n/a |

## A4. Stack-specific controls
- Flask on Vercel Python runtime; filename sanitised in `_safe_download_name`; extension allowlist `.xlsx`.

## A5. Accepted exceptions
| Rule (Part B §) | Exception | Reason | Compensating control | Review by |
|-----------------|-----------|--------|----------------------|-----------|

## A6. Security contacts and incident notes
- Owner: repo owner (Emyol)
- Where to report: GitHub issues on Emyol/SA-DTR-Maker
- Past incidents: none recorded

# Part B: Universal baseline

## B1. Secrets and credentials
- Keep secrets in environment variables or a secret manager. Commit a `.env.example` with **names only**.
- `.gitignore` covers `.env*` (except `.env.example`), key files (`*.pem`, `*.key`, `*.p12`), credential JSON, and local DB dumps, from the very first commit.
- Keep server-only secrets out of client bundles. Only intentionally public values use client prefixes (`NEXT_PUBLIC_`, `VITE_`, `EXPO_PUBLIC_`). Service-role, admin, and private API keys stay server-side.
- Before every commit, scan the diff for secrets (`gitleaks`, `git-secrets`, or `trufflehog`); add a pre-commit hook where the project allows.
- A leaked secret gets **rotated immediately**: deleting it from history does not un-leak it.
- Give each environment (dev / preview / prod) its own credentials, scoped to least privilege.

## B2. Authentication
- Use a proven auth provider or library; write no custom crypto or session schemes.
- Hash passwords with **Argon2id** (or bcrypt cost ≥ 12 / scrypt). Never store or log plaintext.
- Offer MFA for privileged accounts; support passkeys/WebAuthn where practical.
- Sessions: `HttpOnly`, `Secure`, `SameSite=Lax` or `Strict` cookies; rotate the session ID on login and privilege change; enforce idle and absolute timeouts; server-side logout invalidation.
- Tokens: short-lived access tokens, rotating refresh tokens, verify signature, `exp`, `aud`, and `iss`; reject `alg: none`.
- Rate-limit and lock out progressively on login, signup, password reset, and OTP endpoints. Use generic error messages ("invalid credentials") to prevent account enumeration.
- Reset tokens and magic links: single-use, short expiry, high entropy, stored hashed.

## B3. Authorization
- **Deny by default.** Every endpoint, server action, RPC, and query checks authorization **server-side**; client checks are UX only.
- Check object ownership on every access to prevent IDOR (`WHERE owner_id = current_user`).
- Database-level enforcement where available (e.g. Postgres/Supabase **Row Level Security enabled on every table**, with explicit policies; no table left open to `anon`).
- Least privilege for users, service accounts, API keys, CI tokens, and cloud IAM roles.
- Admin functions live behind separate role checks and are audit-logged.

## B4. Input handling and injection
- Validate all input at trust boundaries with a schema (Zod, Pydantic, Joi, etc.): type, length, format, range, allow-lists over deny-lists.
- **SQL:** parameterised queries or a safe ORM only; no string-built SQL.
- **XSS:** rely on framework auto-escaping; avoid `dangerouslySetInnerHTML` / `v-html` / `innerHTML`; sanitise with DOMPurify when HTML is unavoidable.
- **Command injection:** avoid shelling out; when necessary, pass argument arrays, never interpolated strings.
- **Path traversal:** resolve and verify paths stay inside an allowed base directory.
- **SSRF:** allow-list outbound hosts for any user-influenced URL; block private/link-local/metadata IP ranges.
- **Deserialisation:** never unpickle / `eval` / `yaml.load` untrusted data; use safe loaders.
- **Mass assignment:** map request fields explicitly onto models.
- **Open redirects:** redirect only to allow-listed or relative paths.

## B5. Web and API protections
- HTTPS everywhere; HSTS (`max-age=63072000; includeSubDomains; preload`) in production.
- Security headers: `Content-Security-Policy` (nonce- or hash-based, no `unsafe-inline` scripts), `X-Content-Type-Options: nosniff`, `Referrer-Policy: strict-origin-when-cross-origin`, `Permissions-Policy`, `frame-ancestors 'none'` (or a specific allow-list).
- CORS: explicit origin allow-list; never `*` together with credentials.
- CSRF protection for cookie-authenticated state-changing requests (SameSite plus tokens or origin checks).
- Rate limiting and request size limits on all public endpoints; pagination caps on list endpoints.
- Return generic errors to clients; keep stack traces and internals in server logs.
- Webhooks: verify signatures, check timestamps to block replays, process idempotently.

## B6. File uploads
- Validate type by content (magic bytes), not just extension or `Content-Type`; enforce size limits.
- Store outside the web root or in object storage with private buckets and signed, expiring URLs.
- Rename files to random IDs; strip metadata from images where privacy matters; scan for malware where risk warrants.

## B7. Cryptography
- TLS 1.2+ in transit; AES-256-GCM (or the platform KMS) at rest for Restricted data.
- Use vetted libraries and the platform CSPRNG (`crypto.randomBytes`, `secrets`); never `Math.random()` for security.
- Keys live in a KMS/secret manager with a rotation plan.

## B8. Dependencies and supply chain
- Commit lockfiles and install with frozen lockfiles in CI (`npm ci`, `pnpm i --frozen-lockfile`, `pip install --require-hashes`).
- Before adding a package, check: maintenance activity, download count, typosquat-like names, install scripts, and licence.
- Run `npm audit` / `pip-audit` / `osv-scanner` regularly and in CI; enable Dependabot or Renovate.
- Pin GitHub Actions to commit SHAs; give workflow tokens minimal `permissions:`.
- Pin container base images by digest; run as a non-root user; use minimal images.

## B9. Data protection and privacy
- Collect the minimum data needed; define retention periods and delete on schedule.
- Mask PII in logs, analytics, and error reports; never log passwords, tokens, full card numbers, or session IDs.
- Back up databases with encryption, and test restores.
- Honour applicable regulations (e.g. GDPR, Philippine Data Privacy Act, CCPA): consent, access, and deletion requests.

## B10. Logging, monitoring, and incident response
- Log security events: logins (success and failure), permission denials, admin actions, password and MFA changes, and key usage, with timestamp, actor, IP, and request ID.
- Protect logs from tampering and from containing secrets.
- Alert on anomalies (auth-failure spikes, unusual data exports).
- Incident steps: **contain → rotate credentials → assess scope → fix root cause → notify where required → record in A6.**

## B11. AI / LLM features (when present)
- Treat all model input that includes user content, retrieved documents, web pages, or tool output as **untrusted** (prompt injection).
- Keep secrets and other users' data out of prompts; enforce authorization before retrieval, not after.
- Give model-invoked tools least privilege; require human confirmation for destructive or outward-facing actions.
- Validate and encode model output before rendering it as HTML, running it as code or SQL, or passing it to a shell.
- Cap tokens, cost, and request rate per user.

## B12. Mobile and desktop clients (when present)
- Store tokens in Keychain / Keystore / OS credential vault, not plain storage.
- Ship no secrets in the app binary: anything shipped is public.
- Use certificate pinning for high-risk apps; disable debug flags and verbose logging in release builds.

## B13. Infrastructure and deployment
- Separate dev / preview / production environments and data; preview deployments get no production secrets.
- Protect the main branch: required reviews and passing checks before merge.
- Close unused ports; require auth on admin and database consoles; disable default credentials.
- Use infrastructure as code where possible, reviewed like application code.

---

## Pre-commit / pre-deploy checklist
Run through this before every commit that touches auth, data access, input handling, dependencies, config, or deployment, and before every production deploy:

- [ ] No secrets, keys, or personal data in the diff or the client bundle
- [ ] New inputs validated with a schema at the trust boundary
- [ ] Every new endpoint/action/query has a server-side authorization check (and RLS where applicable)
- [ ] No string-built SQL, shell commands, HTML, or file paths from user input
- [ ] New dependencies vetted; lockfile updated; audit clean or exceptions recorded in A5
- [ ] Errors returned to clients are generic; sensitive data stays out of logs
- [ ] Security headers, CORS, and rate limits are intact
- [ ] Part A updated if the attack surface, data, or stack changed
- [ ] Security-relevant changes noted in `current-state.md`

## Review cadence
- Re-check Part A whenever a feature adds a new entry point, data class, or third party.
- Full pass over Part B before each production release, and at least monthly on active projects.
