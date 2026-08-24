# T2W Pipeline Asset Register

Searchable register for the Toowoomba to Warwick pipeline — 1,188 assets across six
source registers, keyed on chainage in metres.

The Excel workbook in this repository is the master. Everything the site serves is
generated from it at deploy time, so the published data can never drift from the
workbook it came from.

---

## How updating works

```
edit the workbook  →  git push  →  Cloudflare builds data.json  →  site is live
```

There is no `data.json` in this repository, and there shouldn't be. It is built on
every deploy by `scripts/build-data.mjs`. That means:

- no build step for you to remember, and no stale committed copy
- no workflow committing generated files back, which would retrigger itself
- the workbook and the site cannot disagree

To publish a change: open `T2W Pipeline Master Asset Register.xlsx`, edit it, commit,
push. Cloudflare rebuilds within about a minute.

---

## One-time Cloudflare setup

1. In the Cloudflare dashboard: **Compute → Workers & Pages → Create application → Pages
   → Connect to Git**.
2. Pick this repository.
3. Set the build settings exactly:

   | Setting | Value |
   |---|---|
   | Framework preset | None |
   | Build command | `npm ci && npm run build` |
   | Build output directory | `public` |

4. **Save and Deploy.**

Every push to the default branch redeploys from then on. Pull requests get their own
preview URL, which is a good way to check a data change before it goes live.

> Choose **Connect to Git**, not *Upload assets*. A project created as direct-upload
> cannot be switched to Git later — you would have to start again under a new name.

---

## Access

The site is public once deployed. Anyone with the URL can read the whole register,
including chainages, depths and the utilities the pipeline crosses.

If that isn't what you want, put **Cloudflare Access** in front of it: *Zero Trust →
Access → Applications → Add an application → Self-hosted*, point it at the Pages
hostname, and add a policy (email domain, or a named list). Free for up to 50 users.
It takes about five minutes and requires no change to this repository.

Worth checking separately: a substantial part of the foreign service data derives from
BYDA enquiries, and those usually carry conditions on redistribution.

---

## What's in here

```
T2W Pipeline Master Asset Register.xlsx   the master. Edit this.
public/
  index.html                              the register page
  _headers                                Cloudflare headers (no-cache on data.json)
  data.json                               generated at build time — not committed
scripts/
  build-data.mjs                          workbook -> data.json
package.json
.github/workflows/check-data.yml          runs the build on every push and PR
```

`_headers` is the Cloudflare header file. If you see a `staticwebapp.config.json`
anywhere, it's an Azure Static Web Apps file and Cloudflare ignores it — delete it.

---

## What the build script actually does

It is not a spreadsheet-to-JSON dump. It reproduces the derivations the register
depends on:

- **Alignment sheets.** Each sheet covers 700 m of chainage — a figure taken from the
  source workbook's own assumption block, not invented. 176 assets had no drawing
  number in any source register, so theirs is derived from chainage and marked as
  derived so an inferred sheet is never mistaken for a stated one. 20 records (the PPH
  and RHF thrust blocks) can't be placed on any sheet, because no PHW or PWG series
  covers them.
- **Data quality exceptions.** The 67 exceptions from the workbook are carried across
  and linked to the 157 specific records they affect, so a flag appears against the
  asset it concerns.
- **Sheet index.** Includes sheets that hold no records at all — currently
  `D-DWG-PWG-009` and `D-DWG-PWG-034` — so a gap is visible rather than absent.
- **Trimming.** 26 columns are empty in every record (the survey block, permits, depth
  of cover). They're dropped from the payload and reappear automatically once they
  carry data.

Run it yourself with `npm run build`.

---

## If a build fails

The script fails loudly rather than shipping bad data. The usual causes:

| Message | Cause |
|---|---|
| `No .xlsx file found` | workbook not committed, or renamed |
| `has no "Master Register" sheet` | sheet renamed in Excel |
| `missing a required column` | a column was renamed or deleted — check the header row |

Cloudflare keeps the previous successful deployment live when a build fails, so a bad
push takes the site stale rather than down.

---

## Chainage

Metres throughout, in every register. The `KP` column label in the original sources is
misleading — observed values run 0 to 27,012, which is metres, not kilometres. No
conversion is applied anywhere.
