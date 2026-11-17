Tl;dr:  
“I wrote this console app to read from my Google Sheet to grab .nfo file information to save it with the movie. It then grew to converting the movie from the HandBrake CLI, and then grabbing the plot summary from themoviedb.org and adding it to the Google Sheet.  
The next step I hope to implement is signing into my Plex account to download the movie posters and background images.”  
  
The long version:  
I still love buying DVDs. I use a Google Sheet to keep track of all the ones I've purchased, but I hate actually using the DVD.  
So, I rip and convert the DVD then use Plex to stream the movie.  
I've used Plex for a few years now and after having to replace a couple failed hard drives, I've grown tired of using the plex.tv interface to change the sort order of my movies. Or spend hours picking the perfect movie poster and background just to lose all that and have to do it all over again.  
So, I spent many hours scouring the web to find a solution. I still want to use Plex, I've had my monthly pass for a few years now and I share my library with a number of family and close friends.  
Unfortunately, Plex doesn't natively support .nfo files, small xml files that help the player know more about the movie, but then I stumbled on a couple wonderful projects here:  
https://github.com/gboudreau/XBMCnfoMoviesImporter.bundle  
https://github.com/gboudreau/XBMCnfoTVImporter.bundle  
These are plugins that you can add to your Plex plugins folder, then with a couple small tweaks to your Plex settings, they will find the .nfo files with your movies and import that information.  
Haza! Now I use my Google Sheet to create the .nfo file with the basic information that I hate having to redo every time a hard drive fails.  
Here is a sample .nfo file I use: (I know this isn't the only way of using these, and I could add a lot more info, but this works for my needs)  
  
    <?xml version='1.0' encoding='UTF-8' standalone='yes'?>
    <movie xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'>
      <id>tt0147800</id>
      <title>10 Things I Hate About You</title>
      <sorttitle>10 Things I Hate About You (1999)</sorttitle>
      <plot>On the first day at his new school, Cameron instantly falls for Bianca, the gorgeous girl of his dreams. The only problem is that Bianca is forbidden to date until her ill-tempered, completely un-dateable older sister Kat goes out, too. In an attempt to solve his problem, Cameron singles out the only guy who could possibly be a match for Kat: a mysterious bad boy with a nasty reputation of his own.
      -- Did you experience any issues while playing this movie? Notice it's missing subtitles? Or simply have a movie request? Then please text me, Brandon, about it at (***)***-***** --</plot>
      <certification>USA:PG-13</certification>
    </movie>
  
First, I have the IMDB ID so Plex can easily identify the movie.  
Next, is just the movie title.  
Then, I have the sort title (That's one of the main reasons for this project)  
Finally, I add the plot so that way I can concatenate a small note from myself hoping to encourage my friends and family to let me know if anything is wrong with the movie (I removed my number from this demo)  
And the certification is just to make sure the right content rating gets applied to the movie. i.e. ‘The Internship’ is rated PG-13, but I have both the theatrical PG-13 version and Unrated version.  
  
This method worked great for a little while. I would simply copy in the text from the Google Sheet and create the .nfo file as I convert and add the movie to my library. Not too bad, but if I ever made ANY update to the movie info in the Google Sheet, I would have to edit the current .nfo file for that movie and copy in the new text. A little annoying but didn't happen often enough to bother me. However, when I decided to make a universal change to ALL of the .nfo files, it proved to be too tedious to update each .nfo file manually.  
So then came the reason of creating this console app. I wrote it for the sole purpose of reading from my Google Sheet to grab the text from the .nfo file column and then create the .nfo file in with the movie. This runs through in a matter of seconds, saving me countless hours of doing it by hand.  
  
Then it grows from there. I now have it convert my movies for me too. Once I rip the disc, I input which title # it uses to convert the movie and then it sends it to HandBrake to convert it.
And then I got an API account with themoviedb.org and now it gathers all the IMDB ID’s from the Google Sheet to send them to themoviedb API to get the movie plot and then plug it back into my Google Sheet.
Future plans. At this time, the next feature I hope to add is that it will go into my Plex account and download the movie posters and backgrounds in with the movie folder.
Then who knows from there.

## TheTVDB v4

TV media updates are available as **61 - Update TV Show Resolution and Format
(Smart Scan)** under DB Sheet, with **61f** for a force rescan. See
[cache behavior, safeguards, usage and tests](TV_MEDIA_SMART_SCAN.md).

Missing-episode tracking is available as **28m — Tally Missing Episodes (preview
only)** under DB Sheet. **28mu** is the explicit **LIVE UPDATE** option. The
`--missing-episodes-preview` command remains read-only; the separate
`--missing-episodes-update` command writes verified changed counts in bounded
batches, preserving failed/conflicted rows. Both support a show limit or
`--row <DB-row>` for a small test.
See [setup, verification rules, tests and live-write steps](MISSING_EPISODES.md).

Option **55m** scans the recording PCs' output shares directly:
`\\RAPH\RaphOutput`, `\\LEO\LeoOutput`, `\\MIKEY\MikeyOutput`, and
`\\DONNY\DonnyOutput`. The Downloads `checked` and QUADPLEX `Other\Output`
sources are also scanned as before. Use `<output>\<provider>\<title>` (or loose
movie files under the provider); the old `<output>\Ready\<provider>` layout is
still supported. Movie renaming, sidecars, destination merging, and subsequent
sheet/library processing follow the existing workflow.

Movie matching tries exact titles first, then treats standalone `and` / `&`
and a leading `The ` as equivalent on either title. These variations retain the
year before falling back to the existing title-only matching. Multiple equally
good candidates require a choice; the selected sheet Clean Title remains the
final filename. Run `tests/Run-MovieTitleMatcherTests.ps1` for offline checks.

Provider aliases live in `BachFlixNfo/Features/OutputMediaFolders.cs`.
`Amazon` maps to `Prime` for working-area paths and subsequent sheet updates;
matching ignores case. Other provider names pass through unchanged. Add further
aliases there when StreamFab and the sheet use different names.

Put unfinished items inside a folder named **Skip**. Both discovery and cleanup
ignore `Skip` (case-insensitive), `.sync`, and directory links/junctions. If a
title contains one of these folders, the entire title stays at the source because
title moves transfer the whole tree. Empty source folders are removed after the
move pass, even when there is nothing to move; configured roots and excluded
folders are preserved. Nonempty folders are never removed by cleanup.

Folder transfers remove each source file after copying (or finding an existing
destination of the same size), verifying the destination size and confirming the
source size and modification time have not changed. New files arriving during
the transfer remain for the next run. Changed or failed files are retained and
reported while other files continue; cleanup only deletes empty directories.

Offline checks for 55m folder discovery, aliases, protected folders, cleanup,
and destination merging:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Run-OutputMediaFoldersTests.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Run-MediaMoveProgressTests.ps1
```

TV metadata provided by [TheTVDB](https://thetvdb.com).

Set `TVDB_API_KEY` in your ignored `BachFlixNfo/BachFlixNfo.local.env`
(or as a process environment variable). See `BachFlixNfo.local.env.example`.
Use a project API key; no subscriber PIN is needed. The application loads it
through `LocalEnvironment` and obtains a token only when TVDB is used.
Tokens stay in memory; no Google Sheets token cell is read or written.

The client uses `https://api4.thetvdb.com/v4` with:
- `POST /login` (only `apikey` in the JSON body)
- `GET /search?type=series&query=...` (requires an exact slug match)
- `GET /series/{id}/extended`
- `GET /series/{id}/episodes/official?page=0&season=...&episodeNumber=...`

Official episode ordering preserves aired SxxExx numbering, including season 0
specials; it does not switch to DVD/absolute ordering. Season counts use distinct
positive official season numbers, excluding specials and alternative orders.
The compatibility response retains `seriesName`, `firstAired`, `season`, `rating`,
`status`, `episodeName`, and `overview` for existing sheet workflows. Content
ratings select USA ratings; absent text becomes an empty string. Missing first-air
dates leave the sheet Year unchanged. Existing IDs bypass slug search; a changed
or unresolved slug requires an explicit ID rather than guessing a result.

A 401 permits one fresh login and one retry. Persistent authentication failures,
missing records, network failures, and API errors are reported without response
bodies, keys, or tokens. Request timeouts are 30 seconds and redirects are disabled.

Run offline regression tests with Visual Studio C# build tools installed:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Run-TvdbTests.ps1
```

The tests use a fake key and transport and never load local credentials or modify
Sheets. Manual validation should check a known slug, an ID-only row, a regular
episode and a special, plus the resulting metadata in your normal sheet workflow.
API reference: https://thetvdb.github.io/v4-api/
