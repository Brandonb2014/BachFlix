# TV Show Resolution and Format smart scan

Under **Update Google Sheet > DB Sheet**, enter **61** to update TV show media
values. Enter **61f** to force FFProbe to inspect every current episode again
(for exceptional content replacements with identical size and timestamp).
Option **59** remains the existing Profile Requests workflow.

The scanner reads DB headers from row 2 and show records from row 3. It validates
that **AE is Resolution** and **AG is Format** before scanning and again before
each write. Missing, duplicate or moved required headers stop the operation.
Only changed individual AE/AG cells are sent through the existing authenticated
Google Sheets batch writer, with RAW values. Formula cells are preserved.

Resolution is `480p`, `576p`, `720p`, or `1080p` only when every current episode has the
same supported resolution. When episodes have more than one resolution, Resolution
is `Mixed` (like Format). When every episode has the same unsupported resolution
(for example all 2160p), Resolution is cleared. NReco FFProbe reads video stream dimensions without decoding
the video. The existing `FindResolution()` and TV scanner share one classifier:
canonical widths retain the existing treatment of cropped HD video, and other
dimensions use the established height ranges. For example, `720x544` and `960x540`
normalize to `480p`, while `1918x1078` normalizes to `720p` according to the existing
height thresholds. This is the application's existing classification convention,
not rounding to whichever supported number is nearest. Sources wider than 1920
or at least 1440 pixels high remain unsupported in the TV scanner, including
cropped 4K. PAL SD at 576 pixels high supports
widths up to 1024 and up to two pixels of vertical cropping, including `718x576`
and `718x574`. Previously cached unsupported dimensions are reclassified using
the shared rules without another probe; the existing cache and checkpoints remain usable.
Mixed resolutions write `Mixed`. Format is explicitly `MP4`, `MKV`, or
`Mixed`, determined from all episode file extensions. Other supported video file
extensions contribute `Mixed` and are logged.

**Found Locations** (comma, semicolon or newline separated) is the primary source
of show directories. When populated, every listed path is scanned directly;
Hard Drive Letter and unrelated network shares are not checked. Commas inside
show folder names are preserved. An unreadable listed path is a true error and
preserves that show's existing sheet values.

When **Found Locations is blank or whitespace**, the show is immediately skipped
in yellow, preserving its existing values. No drive/share discovery or episode
probing is attempted, even if Hard Drive Letter is populated. All listed
locations contribute episodes, including
Specials and Season 0. Overlapping paths and mapped-drive/UNC aliases are
deduplicated. Recognized extras directories, sample/trailer files, sidecars and
non-video files are excluded; numbered episodes with titles such as "The
Interview" remain included. Directory links/junctions and file links are
reported for review rather than risking a duplicate or incomplete inventory.

An inaccessible listed location, season folder, or episode, a
failed probe, or a show with no episode files preserves that show's existing
sheet cells and counts as a true error. No
TVDB episode counts or existing sheet media values substitute for a disk scan.

The persistent, versioned cache is:

```text
%LOCALAPPDATA%\BachFlixNfo\TvEpisodeMediaCache.json
```

Each episode stores its canonical full path, byte size, UTC modification time,
resolution and format, grouped by TVDB ID (or normalized name/year fallback).
Normal scans enumerate all files and compare each file's metadata. Only new or
changed files are probed. Renames invalidate the old path; successful scans prune
deleted files and recalculate the show. The initial run always probes episodes.

Atomic per-show JSON checkpoints live beside the main cache in
`TvEpisodeMediaCache.json.checkpoints`. They save every 25 probe attempts or
10 seconds between attempts, and at the end of each changed show. This avoids
repeatedly rewriting the entire library cache. On exit, checkpoints are compacted
into an atomically replaced main JSON file. If the application is closed abruptly,
the next normal scan recovers completed checkpoints. At most the current
checkpoint interval must be re-probed. Successful individual episodes from an
otherwise failed show remain reusable; failed episodes are never cached as
successful. A lock prevents concurrent writers.

Press **Ctrl+C** to stop safely. Completed metadata survives cancellation and
Google Sheets failures. Unsent batches can be recalculated from the cache on the
next run. The console reports per-show progress, update status and final counters,
with cyan progress/headings, magenta changes, green completed updates, yellow
skips, warnings and pending batches, red errors and gray unchanged results.
The final counts separate **TV Shows skipped**, **TV Shows with errors**, and
scan-level errors. Normal skips do not set a failure exit code; true errors and
cancellation do. Saved logs
remain plain text.
BachFlixLog saves the same output, including problematic paths, under
`Desktop\BachFlixNfo Logs\TV Media Smart Scan`.
When errors occur, an **Errors to review** recap appears after the totals, with
show titles, DB rows, failed paths and reasons. An additional
`BachFlixNfo_TvMediaSmartScan_Errors_<timestamp>.txt` file in the same folder
contains just those errors; its full path is printed as **Errors-only log**.

Run the offline automated tests with Visual Studio C# build tools installed:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Run-TvEpisodeMediaTests.ps1
```

The tests use fake filesystem metadata, fake probes and a fake batch writer.
Only temporary test-owned cache files are written; no real media, network shares,
Google credentials or live sheet data are accessed.

The verified Debug build is in `BachFlixNfo\bin\Debug`. Launch
`BachFlixNfo.exe` with your normal working directory/credentials setup, then
select 61 (or 61f for a force rescan).
