# Lesson Grader launcher (GitHub Pages)

**Status (25 September 2026):** implemented and covered by repository tests. **Not verified against a live
Apps Script deployment** — no real deployment, Google account or Gemini key is reachable from this repository.

A single static page, `site/index.html`, hosted on GitHub Pages, that opens the Lesson Grader Apps Script web app
and lets the teacher change **which** deployment it opens from the page's own **Settings** — stored in the browser,
never in code. Nothing in GitHub has to be edited when a deployment address changes.

```
  Browser
  ├── https://<user>.github.io/lesson-grader-gas/      ← site/index.html (static, no secrets)
  │     └── Settings: deployment address → localStorage
  │
  └── <iframe or new tab>
        └── https://script.google.com/macros/s/…/exec   ← your Apps Script deployment
              ├── WebApp.html + google.script.run       (unchanged grading UI)
              ├── WebAccess.requireTeacher()            (LESSON_GRADER_ALLOWED_EMAILS)
              └── doGet?lgapi=status                    (read-only handshake used by the launcher)
```

## What the launcher does

| Does | Does not |
|---|---|
| Remembers one or more deployment addresses and labels in this browser | Store a Gemini key, password, allowlist or student record |
| Embeds your real deployment (or opens it in a new tab) | Re-implement grading, Classroom, Drive or Sheets access |
| Asks the deployment one read-only question: alive? allowlisted? key configured? | Request, receive or display student work |
| Exports/imports its own settings as JSON, or as a shareable link | Change anything in Apps Script |
| Explains why the app is not loading (sign-in, allowlist, cookies) | Replace `LESSON_GRADER_ALLOWED_EMAILS` enforcement |

The grading UI still runs **inside the Apps Script deployment**, so `google.script.run`, the OAuth consent screen,
the teacher allowlist and the Script-Property Gemini key all keep working exactly as they do today. The launcher is
a front door, not a second copy of the app.

## Set it up once

1. **Publish the page.** Repository → *Settings → Pages → Build and deployment → Source*: **GitHub Actions**.
   `.github/workflows/pages.yml` publishes the `site/` folder after `npm run check` and `npm test` pass.
   The page will be at `https://<user>.github.io/lesson-grader-gas/`.
   *No Actions?* Copy `site/index.html` to the repository root or `docs/` and use *Source: Deploy from a branch*
   instead — it is a single self-contained file with no build step.
2. **Update the Apps Script project** with the current `WebAccess.gs` and `WebAppIntegration.gs`, then
   **Deploy → New deployment → Web app** (or a new *version* of the existing one). Copy the **Web app URL** ending
   in `/exec`. Keep the existing rules: *Execute as* **user accessing the web app**, *Who has access* as narrow as
   the school allows — never "Anyone" with "Execute as me".
3. **Open the page**, paste that URL into *Apps Script web app URL*, optionally give it a label, and press
   **Save & connect**. The address is written to this browser's `localStorage` and the app loads underneath.

To change the address later: **Settings → Deployment address → paste → Update connection**. No GitHub edit, no
redeploy of the page. A testing deployment and the live one can be saved side by side and switched with **Use**.

## The connection check

`Settings → Test connection` (and, by default, page load) asks:

```
GET https://script.google.com/macros/s/…/exec?lgapi=status
```

The page first tries a normal cross-origin `fetch`; if the browser blocks it (Apps Script sends no
`Access-Control-Allow-Origin` header), it falls back to JSONP, which Apps Script supports without CORS. Either way:

- the request is **read-only** — `doGet` only accepts `lgapi=status`, and refuses any other value;
- the answer contains `authorised`, `accountDetected`, a **masked** account (`da***@school.edu.au`),
  `keyConfigured` and `allowlistConfigured` — never the key, the allowlist or a full address;
- JSONP is only ever sent to `script.google.com` / `script.googleusercontent.com`, so a mistyped or hostile address
  cannot make this page execute code from somewhere else.

| Status pill | Meaning | What to do |
|---|---|---|
| **Connected** | Deployment answered and this account is allowlisted | Use it |
| **No Google account seen** | Google cookies were not visible to the check (common in Safari/strict Chrome) | Sign in to Google, or use **Open in tab** |
| **Not allowlisted** | The deployment answered, the account is not in `LESSON_GRADER_ALLOWED_EMAILS` | Add the exact address in *Project settings → Script properties* |
| **Not reachable from here** | No answer within 12 s | Try **Open in tab** before assuming the deployment is down |
| **Test skipped** | The saved address is not a `script.google.com` host | Check the address |

## Known limits

- **Embedding is off by default.** HtmlService blocks framing by other sites unless the owner sets the Script
  Property `LESSON_GRADER_ALLOW_EMBED=true` (which also allows clickjacking-style framing by any site). Without it,
  use *Open in tab*.

- **Third-party cookies.** An embedded Apps Script app needs Google cookies inside a frame. Safari blocks them;
  Chrome does so under strict tracking protection. Symptom: a blank frame or a sign-in loop. Fix:
  *Settings → Always open in a new tab*. This is a browser rule, not a deployment fault.
- **The check can be wrong when the app is fine.** The probe is a cross-site request; the embedded app is a
  document. Treat the pill as a hint and confirm with **Open in tab**.
- **`/dev` addresses** only work for the script owner while the editor project exists, and change often. The page
  accepts one but flags it.
- **No sandbox attribute on the frame.** Apps Script needs its own nested sandbox frame and sign-in pop-ups;
  adding `sandbox` to the iframe breaks them. The page's Content-Security-Policy restricts frames and scripts to
  Google's own hosts instead.
- **One browser, one set of settings.** Use *Export JSON* / *Copy share link* to move settings to another device.
  A share link contains the deployment address only; check any address someone else sends you before saving it.
- Publishing a **new version** of the same deployment keeps the `/exec` URL, so the launcher needs no change.
  Creating a **new deployment** gives a new URL — paste it in Settings.

## Verify in this repository

```bash
npm ci
npm run check   # parses the launcher's scripts and refuses a hardcoded deployment id
npm test        # tests/launcher.test.cjs: settings persistence, URL rules, probe, doGet handshake
npm run preview # http://localhost:4173/launcher — the page; http://localhost:4173/ — mock app UI
```

`tests/launcher.test.cjs` drives the page in jsdom (save → reload → switch → delete, JSONP and fetch paths,
untrusted hosts, corrupt storage) and runs `doGet` in a mocked Apps Script sandbox to prove the probe answers
without leaking the key, the allowlist or a full email. It cannot prove anything about a live Google deployment.
