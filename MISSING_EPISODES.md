# Tally Missing Episodes

The feature compares DB show rows with owned Emby episodes and TheTVDB v4's
complete official-order inventory. It writes only verified results: a compact list of the missing episodes
(e.g. `S01E04-08, 10 & 13, S02E03 & 04`; runs of three or more become a range),
or `None` when nothing aired is missing.
The output column is always resolved from the **Missing Episodes** header in
**DB row 2**; column P and the collection size are not hardcoded.

## Modes and commands

Run from the repository root containing `credentials.json`:

```powershell
# Read-only previews: these commands never write Google Sheet values.
.\BachFlixNfo\bin\Debug\BachFlixNfo.exe --missing-episodes-preview 3
.\BachFlixNfo\bin\Debug\BachFlixNfo.exe --missing-episodes-preview --row 742

# LIVE UPDATE: writes verified changed counts for the selected shows.
.\BachFlixNfo\bin\Debug\BachFlixNfo.exe --missing-episodes-update 3
.\BachFlixNfo\bin\Debug\BachFlixNfo.exe --missing-episodes-update --row 742
```

Omit the limit/row arguments to process all show rows. A limit selects the first
N nonempty show rows; `--row` selects one physical DB row. They cannot be combined.
Menu **28m** remains preview-only. **28mu** is explicitly labeled **LIVE UPDATE**.
No command refreshes Emby, renames media, or edits Emby metadata.

Use the existing Google authorization setup. Set `EMBY_SERVER_URL` (or its
`EMBY_SERVER` alias), `EMBY_API_KEY`, and `TVDB_API_KEY` in the ignored
`BachFlixNfo/BachFlixNfo.local.env` or the process environment. Never put keys
in URLs. The existing TVDB project-key login is retained.

DB needs `Missing Episodes`, `TVDB ID`, and `Series Name` or `Clean Name` headers.
Optional identity/review fields are `IMDb ID`/`IMDB ID`, `TMDB ID`, `Year`, and
`Continuing?`. Relevant duplicate headers and duplicate normalized TVDB IDs stop
processing. Rows are mapped by physical position, including intervening blanks.

## Verification rules

- TVDB provider IDs are the primary show identity. Emby internal item IDs are
  used only to retrieve that series' episodes. Existing fallback matching remains
  available only when stronger IDs are absent; a failed ID never falls back to
  a guessed title. Duplicate Emby matches fail.
- Conflicting available IMDb/TMDB IDs require review. Title differences alone
  and year differences alone are warnings; a differing title together with a
  year difference greater than one requires review. Accents, punctuation and a
  trailing title year are normalized for this comparison.
- Emby requests `Path,MediaSources,ProviderIds`. This was verified against the
  current server, which omits `LocationType` from episode responses. Ownership
  requires a positive-size, playable `File`/`Default` media source with an
  absolute path and explicit `IsRemote=false`. Virtual items, remote sources,
  STRM references, placeholders and explicit empty source lists are excluded.
  Missing/uncertain ownership evidence blocks the show's result.
- Regular owned episodes need valid season/start/end numbering. Inclusive
  combined ranges expand to individual pairs; duplicates deduplicate. Any
  unusable owned regular episode blocks the count, including a partially valid
  inventory. Specials are excluded before validating their numbering.
- Both APIs must complete pagination. Emby verifies totals, unique item IDs and
  episode parent identity. TVDB verifies series/episode IDs, next-page progression,
  available totals across pages, and rejects empty inventories and repeated IDs.
- TVDB regular episodes need positive numbering and valid `yyyy-MM-dd` air dates.
  Unknown dates, malformed numbering and conflicting duplicate dates require
  manual attention. Specials and dates after the run's captured local date do
  not enter the aired tally. A nonempty, completely dated future-only inventory
  can legitimately verify zero aired episodes; an empty response cannot.
- Owned numbers absent from the entire official inventory block the count as a
  numbering/order conflict. Known future owned episodes are reported separately.
- Missing episodes are the deduplicated aired season/episode pairs minus owned
  pairs. Only a successful result has a proposed value, including `None`.

No client can detect omissions that upstream services present as a complete,
valid inventory. Incorrect but internally consistent source metadata still
requires manual correction. No media filesystem probes are performed.

## Writes, recovery and reports

Updates use bounded batches of at most **50 single-cell RAW text values**.
Immediately before every batch the program rereads DB, resolves headers and
checks each destination's physical row, TVDB ID, output column, current value,
and formula status. Changed identities/headers, concurrent edits, and formulas
are preserved and reported. Unchanged values issue no writes; old numeric counts are replaced by lists. Failed
show checks are never included in a batch.

Do not sort, insert/delete rows, move headers or edit the output column while an
update is committing. Sheets has no conditional compare-and-set operation for
this batch endpoint: a small race remains between preflight and commit.

API reads use at most three transient-failure attempts. Emby/TVDB honor bounded
`Retry-After` delays; delays over 30 seconds stop rather than retry early. TVDB
retains one 401 reauthentication attempt. Google SDK implicit write retries are
disabled. A failed/uncertain Sheet write is read back; confirmed desired values
are recorded and unconfirmed rows are reported. No blind write retry occurs.
The run stops after such a batch; a fresh run skips already matching values.

Authentication/shared service failures stop conservatively after retries;
individual inventory/numbering/match failures allow other shows to continue.
Ctrl+C preserves uncommitted rows and generates a partial report after the
current API/show operation returns. Completed batches remain applied. Abrupt
process termination can prevent report creation; rerunning is still safe.
Only compact normalized inventories are cached during a run.

Progress and final output include total/selected/examined shows, successful
checks, shows with/without missing episodes, total missing episodes, failed
checks, updated and unchanged Sheet rows, Sheet conflicts and unchecked shows.
Timestamped `missing-episodes-preview-*.json` or `missing-episodes-update-*.json`
reports contain every checked show, actual missing identifiers, warnings,
failures, write outcomes and manual-attention flags. Reports contain library
names/provider IDs but no credentials or media paths. Failures/partial runs and
report-save failures return exit code 1.

## Automated and live validation

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Run-MissingEpisodeTests.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Run-MissingEpisodeSheetTests.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Run-TvdbTests.ps1
```

The Sheets adapter tests use the real Google request builder with an offline
HTTP handler and no credential loader. They verify integer zero/RAW payloads,
preview write rejection, mapped columns/formulas, header failures and write
acknowledgements. Workflow tests cover bounded batches, failures, row changes,
concurrent edits, partial commits, readback, cancellation and reruns.

Current verification: the Debug application builds; **65 comparison/workflow,
8 Sheets adapter, and 24 TVDB tests pass (97 total)**. Existing unrelated build
warnings remain. CLI argument validation and both menu labels were checked.

A bounded read-only comparison on 2026-10-17 used Pokemon's verified TVDB ID
76703 and Emby item 1516305: 1,255 raw Emby records, 1,254 regular owned pairs,
1,457 raw TVDB records, 1,373 aired regular pairs, and a proposed missing count
of 119. No Google Sheets call was made by that check. These are API-derived
results, awaiting the user's real-world episode review. Google authorization,
live Sheet reads and real writes still require production validation.

For a small live-write test: build, preview three shows, review their report,
run `--missing-episodes-update 3`, compare only the corresponding output cells,
then rerun the same command and confirm zero further writes. Use `--row 742`
for an isolated Pokemon test after reviewing its preview. Never begin with an
unbounded update. Production Sheet writes have not been performed by Codex.

API references: [Emby item fields](https://dev.emby.media/doc/restapi/Item-Information.html),
[Emby Items API](https://dev.emby.media/reference/RestAPI/ItemsService/getItems.html),
[TVDB v4](https://thetvdb.github.io/v4-api/).
