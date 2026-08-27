# UCM Service Lines Dashboard — automation setup

This repo turns the weekly "All Projects" export into the live burn-down
dashboard automatically. Three pieces make that happen, and each one is a
one-time setup:

1. A `Dashboard Bucket` column added to the Excel export (a human decision,
   maintained weekly).
2. A Zap that drops the emailed export into this repo.
3. A GitHub Action that turns that file into `data.json`, which
   `index.html` reads on every page load.

Once all three are wired up, updating the live dashboard is just: the
export arrives by email → everything else happens on its own within a
minute or two.

## 1. Add the "Dashboard Bucket" column

This is the piece that actually decides what shows up where — the script
does not guess from the project type text (it's already been wrong that
way a few times: a "Media: Reporting" project turned out to be real
analytics work, a "Media: Labor" project turned out to be a no-charge
value-add). A person needs to make that call once per project, in the
spreadsheet itself.

Add a column titled exactly `Dashboard Bucket` to the export template (any
position, after the existing columns is simplest) and fill in one of these
values for every row that belongs to the `2026 UCM Service Lines
Campaigns` campaign:

| Value | What it means |
|---|---|
| `Active` | Counts against the $1,865,000 budget; shown under Active Projects |
| `Completed` | Counts against the budget; shown under Completed Projects (stays visible after closing) |
| `Added Value` | Shown in the Added Value section; **excluded** from the burn-down math even if it has a real dollar budget |
| `Exclude` | Left off the dashboard entirely — paid media/media labor, sales-pipeline ("Oppty-") rows, $0 admin rows, anything not relevant |

Rows in that campaign with a blank or unrecognized value are automatically
left off the dashboard **and** surfaced as a yellow warning banner at the
top of the page, so a newly added project never silently vanishes — it
shows up as "needs classification" until someone sets its bucket.

A worked example is in `templates/All_Projects_with_Dashboard_Bucket_example.xlsx`
— it has the column filled in exactly as this dashboard is configured
today (as of this write-up):

- Active: 26-UCMC-052 (Analytics & Reporting), 26-UCMC-058 (Account Oversight & Advisory)
- Added Value: 26-UCMC-034 (Media Transition & Campaign Builds), 26-UCMC-061 (Market Assessment Tool Trial), 26-UCMC-066 (Digestive Diseases & Transplant Media Plan)
- Exclude: every paid-media/media-labor line item, every "Oppty-" pipeline row, and the $0 admin rows (OOP Expenses Only, NB Client Admin)

Whoever owns the export (you, an AM, or finance) should add this column
to the export template so every future email already includes it.

## 2. Zapier: Workamajig email → this repo

The Tuesday 12am ET email from Workamajig contains a **link** to the
report, not a file attached directly — so before wiring the Zap below,
confirm the link is actually fetchable by a robot:

**Check this first.** Open the link from a real copy of the email in a
private/incognito browser window (i.e. logged out of Workamajig). If it
downloads the file straight away, it's a pre-authorized link and the Zap
below will work as-is. If it redirects you to a Workamajig login page
instead, Zapier can't get past that on its own (it can't hold a login
session), and you have two better options: (a) check the report's
schedule settings in Workamajig for a "send as attachment" option instead
of "send as link" — many report schedulers have this toggle, and it turns
this back into the simple attachment case; or (b) use Workamajig's Reports
API (`GET /reports?reportKey=...`, documented at
support.workamajig.com/hc/en-us/articles/360022768792) if your plan
includes it — that skips email entirely and the Action can pull data
directly on a schedule. Ask Workamajig support/your admin if you're not
sure which applies.

Assuming the link works logged out, build one Zap:

1. **Trigger — Gmail: "New Email Matching Search."** Search string should
   pin down the Workamajig sender address and the report's subject line
   exactly, e.g. `from:reports@workamajig.com subject:"All Projects"` —
   pull the real values from an actual email you've received, so nothing
   else in your inbox can fire this Zap. Since the report always lands at
   the same time (Tuesday 12am ET), no separate schedule step is needed —
   the email's arrival is the trigger, and Gmail polling on Zapier
   typically picks it up within a few minutes.

2. **Action — Code by Zapier (Run JavaScript).** This single step pulls
   the report URL out of the email, downloads it, and base64-encodes it
   (GitHub's file API requires base64 text, not raw binary) — doing all
   three in one step avoids Zapier's native webhook step mangling binary
   file responses. Map the trigger's **Body Plain** (or **Body HTML**, if
   Plain doesn't contain the link) field to an input named `emailBody`,
   then use:

   ```js
   // Adjust this regex once you see a real email — it needs to match
   // your Workamajig instance's actual link domain/path.
   const urlMatch = inputData.emailBody.match(/https:\/\/[^\s"<>]*workamajig[^\s"<>]*/i);
   if (!urlMatch) {
     throw new Error('No Workamajig report link found in the email body');
   }
   const reportUrl = urlMatch[0];

   const response = await fetch(reportUrl);
   if (!response.ok) {
     throw new Error(`Report download failed: ${response.status} ${response.statusText}`);
   }
   const arrayBuffer = await response.arrayBuffer();
   const base64Content = Buffer.from(arrayBuffer).toString('base64');

   output = { base64Content, reportUrl };
   ```

3. **Action — GitHub: "Create/Update File."** Connect your GitHub
   account, point it at this repo, set the path to
   `data/All_Projects.xlsx`, branch to `main`, and map the file content
   field to step 2's `base64Content` output. Zapier's GitHub action
   handles create-vs-update automatically.

That's it — every Tuesday, this repo's `data/All_Projects.xlsx` gets
overwritten with the new export, which is exactly the trigger the GitHub
Action below is waiting for.

**Before turning this on for real:** run the Zap once with "Test trigger"
using an actual received email, and check step 2's output is a long
base64 string (not an error, and not something short like a login page's
HTML re-encoded) before letting it write to GitHub automatically.

## 3. GitHub Actions: rebuild data.json automatically

Already included in this repo at
`.github/workflows/update-dashboard.yml`. It fires the moment
`data/All_Projects.xlsx` changes on `main`, runs
`scripts/parse_projects.py`, and commits the refreshed `data.json` back to
the repo. No setup needed beyond pushing this repo's contents to GitHub —
Actions is enabled by default.

One thing to check once: **Settings → Actions → General → Workflow
permissions** should be set to "Read and write permissions," or the
workflow's commit-back step will fail with a permissions error.

## 4. Turn on GitHub Pages

Settings → Pages → Source: **Deploy from a branch** → Branch: `main` →
Folder: `/ (root)` → Save. GitHub gives you a URL like
`https://<your-username>.github.io/<repo-name>/` — that's the live
dashboard. Pages redeploys within about a minute of any push, including
the automatic `data.json` commits from step 3.

## Changing contract terms (total budget / scope end date)

These live as two constants near the top of `scripts/parse_projects.py`
(`TOTAL_BUDGET`, `SCOPE_END`) rather than in the spreadsheet, since
they're contract terms that change rarely, not weekly project data. Edit
them there and push — the next Action run (or a manual "Run workflow" from
the Actions tab) picks up the new values.

## Verifying it end to end

1. Push this repo to GitHub, enable Pages (step 4).
2. Confirm the Action ran once already (Actions tab) and `data.json`
   exists at the repo root.
3. Open the Pages URL — you should see the same dashboard shown in this
   conversation.
4. Edit a `Dashboard Bucket` value in `data/All_Projects.xlsx`, commit
   that change (or send yourself the test email through the real Zap),
   and watch the Actions tab run and the live page update.
