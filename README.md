# NewNetOne DA Eventlog Exporter

A small Windows desktop tool that logs in to the NewNetOne **DirectEvent** portal, searches the Eventlog for a date/time range, follows every result page automatically, and exports everything to an Excel (`.xlsx`) file — with errors grouped onto their own sheets.

## Download

Grab the latest installer from the [Releases page](https://github.com/suryawanshi-dipak/Eventlog-scraper/releases/latest):

**[⬇ Download NewNetOneEventlog.exe](https://github.com/suryawanshi-dipak/Eventlog-scraper/releases/latest)**

> ⚠️ The installer is unsigned, so Windows/Defender may show an **"Unknown Publisher"** warning on first run. That's expected — click **More info → Run anyway** to continue.

## Usage

1. Run the installer and launch **NewNetOne Eventlog** from the Start menu or desktop shortcut.
2. In the login window, enter:
   - **Username / Password** — your DirectEvent portal credentials
   - **Start Date / End Date** — format `d-M-yyyy` (e.g. `5-3-2026`), defaults to today
   - **Start Time / End Time** — format `HH:mm`, defaults to `00:00`–`23:59`
3. Click **Start Scraping**. A tiger animation tracks progress while it walks through every result page.
4. When finished, the app reports how many rows were imported and where the file was saved.

### Output

Everything is written to:

```
%USERPROFILE%\Desktop\Eventlog\
├── Eventlog.xlsx      the exported data
└── app_log.txt        run log (see Troubleshooting)
```

`Eventlog.xlsx` contains:
- **Eventlog** — every row, in the order returned by the site
- **Unique Errors** — a summary count per error type, plus all one-off errors
- One sheet per **repeated error type** (grouped by the first few words of the error text)

## Troubleshooting

If a run fails or produces no output, check `%USERPROFILE%\Desktop\Eventlog\app_log.txt`. It records every step of the run: login result, each HTTP request with its status code, per-page row counts, and the full stack trace of any error. Any unexpected failure also shows a popup pointing at this file.

Common issues:
- **Login Failed dialog** — double-check username/password.
- **"Failed to launch JVM" when starting the app** — the installed runtime is corrupted or was blocked by antivirus during install; reinstall the latest release, or see the CI workflow notes below.
- **Nothing happens after clicking Start Scraping** — check `app_log.txt` for the first `WARNING`/`FATAL ERROR` line.

## Configuration

A few things are currently hardcoded in `NewNetOneEventlog.java` rather than exposed in the login window — edit and rebuild if you need to change them:

| Setting | Constant | Default |
|---|---|---|
| Server | `SERVER` | `"live"` |
| Service filter | `SERVICE` | all services |
| Event types shown | `SHOW_INFORMATION` / `SHOW_WARNING` / `SHOW_ERROR` / `SHOW_PENDING` | Info off, others on |
| Output filename | `OUTPUT_FILENAME` | `Eventlog.xlsx` |
| Page limit | `MAX_PAGES` | unlimited (follows "Next >>" until it runs out) |

## Building from source

Requires JDK 17+.

```bat
:: compile + run directly
run.bat
```

or manually:

```bat
javac -encoding UTF-8 -cp "lib/*" NewNetOneEventlog.java
java  -Dfile.encoding=UTF-8 -cp "lib/*;." NewNetOneEventlog
```

Dependencies (already included in `lib/`): Apache POI (`poi`, `poi-ooxml`, `poi-ooxml-lite`), `xmlbeans`, `commons-collections4`, `commons-compress`, `commons-io`, `commons-codec`, `commons-lang3`, `commons-math3`, `curvesapi`, `SparseBitSet`, `log4j-api`.

### Building the Windows installer

The `.github/workflows/main.yml` GitHub Actions workflow builds the installer `.exe` via `jpackage` + WiX and publishes it as a release. Trigger it by pushing a tag:

```bash
git tag v1.0.8
git push origin v1.0.8
```

or run it manually from the **Actions** tab (`workflow_dispatch`). The workflow builds the app-image first, then repairs the bundled JVM launcher from the JDK on the runner if antivirus stripped it during packaging, before producing the final installer — see the workflow file for details.
