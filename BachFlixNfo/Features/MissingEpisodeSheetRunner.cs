using BachFlixNfo;
using BachFlixNfo.Features;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Net.Http;
using System.Threading;
using Newtonsoft.Json;
using TvdbApiCall;

namespace SheetsQuickstart
{
    // Preview instances reject writes even if called accidentally.
    internal sealed class GoogleMissingEpisodeSheet : IMissingEpisodeSheet
    {
        private readonly SheetsService service;
        private readonly string spreadsheet;
        private readonly bool update;
        internal GoogleMissingEpisodeSheet(SheetsService service, string spreadsheet, bool update)
        {
            this.service = service; this.spreadsheet = spreadsheet; this.update = update;
            // The SDK defaults to retrying 503 responses. Own read retries here and
            // prohibit implicit retries of writes whose commit status is uncertain.
            service.HttpClient.MessageHandler.NumTries = 1;
            service.HttpClient.Timeout = TimeSpan.FromSeconds(30);
        }

        private T ReadWithRetry<T>(Func<T> read)
        {
            for (int attempt = 0; ; attempt++)
            {
                try { return read(); }
                catch (Google.GoogleApiException ex)
                { if (attempt == 2 || !ApiRetry.Transient((int)ex.HttpStatusCode)) throw; }
                catch (HttpRequestException) { if (attempt == 2) throw; }
                catch (System.Threading.Tasks.TaskCanceledException) { if (attempt == 2) throw; }
                ApiRetry.Pause(attempt, null, null);
            }
        }
        public IList<MissingEpisodeShow> Read()
        {
            var headers = ReadWithRetry(() => service.Spreadsheets.Values.Get(spreadsheet, "DB!A2:2").Execute().Values);
            MissingEpisodePreview.ReadShows(headers, null);
            var request = service.Spreadsheets.Values.Get(spreadsheet, "DB!A3:" + SheetHeaderMapping.ColumnLetter(headers[0].Count - 1));
            request.ValueRenderOption = SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum.UNFORMATTEDVALUE;
            var rows = ReadWithRetry(() => request.Execute().Values);
            var shows = MissingEpisodePreview.ReadShows(headers, rows);
            int column = headers[0].ToList().FindIndex(h => Convert.ToString(h) == "Missing Episodes");
            var formulas = service.Spreadsheets.Values.Get(spreadsheet, "DB!" + SheetHeaderMapping.ColumnLetter(column) + "3:" + SheetHeaderMapping.ColumnLetter(column));
            formulas.ValueRenderOption = SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum.FORMULA;
            var output = ReadWithRetry(() => formulas.Execute().Values);
            foreach (var show in shows)
                show.OutputIsFormula = output != null && show.Row - 3 < output.Count && output[show.Row - 3].Count > 0 &&
                    Convert.ToString(output[show.Row - 3][0]).StartsWith("=", StringComparison.Ordinal);
            // Header changes during a read invalidate the entire snapshot.
            var after = ReadWithRetry(() => service.Spreadsheets.Values.Get(spreadsheet, "DB!A2:2").Execute().Values);
            if (JsonConvert.SerializeObject(headers) != JsonConvert.SerializeObject(after))
                throw new MissingEpisodeSheetException("DB headers changed during read; rerun from a fresh snapshot.");
            return shows;
        }
        public void Write(IList<MissingEpisodeCellChange> changes)
        {
            if (!update) throw new InvalidOperationException("Preview mode cannot write Sheet values.");
            if (changes.Count == 0) return;
            if (changes.Count > MissingEpisodeWorkflow.BatchSize || changes.Select(c => c.Range).Distinct().Count() != changes.Count ||
                changes.Any(c => string.IsNullOrWhiteSpace(c.Value) || c.Result.Show.Row < 3 || c.Result.Show.OutputColumn < 0))
                throw new MissingEpisodeSheetException("Refusing an invalid missing-episode batch.");
            var response = service.Spreadsheets.Values.BatchUpdate(new BatchUpdateValuesRequest {
                ValueInputOption = "RAW", IncludeValuesInResponse = false,
                Data = changes.Select(c => new ValueRange { Range = c.Range, MajorDimension = "ROWS",
                    Values = new List<IList<object>> { new List<object> { c.Value } } }).ToList()
            }, spreadsheet).Execute();
            if (response == null || response.TotalUpdatedCells != changes.Count)
                throw new MissingEpisodeSheetException("Sheet did not acknowledge every requested cell; readback required.");
        }
    }

    partial class Program
    {
        private static void TallyMissingEpisodes(int limit = int.MaxValue, bool update = false, int? row = null)
        {
            Environment.ExitCode = 0;
            var run = new MissingEpisodeRun { UpdateMode = update };
            DateTime today = DateTime.Today;
            Type("Tally Missing Episodes - " + (update ? "LIVE UPDATE: verified counts will be written to Google Sheets." : "PREVIEW ONLY: no Google Sheet writes."), 0, 0, 1, "Yellow");
            using (var cancellation = new CancellationTokenSource())
            {
                ConsoleCancelEventHandler cancel = (sender, e) => { e.Cancel = true; cancellation.Cancel(); };
                Console.CancelKeyPress += cancel;
                try
                {
                    using (var sheets = CreateSheetsService())
                    {
                        var sheet = new GoogleMissingEpisodeSheet(sheets, SPREADSHEET_ID, update);
                        var all = sheet.Read();
                        run.TotalShows = all.Count;
                        var shows = (row.HasValue ? all.Where(s => s.Row == row.Value) : all.Take(limit)).ToList();
                        run.SelectedShows = shows.Count;
                        run.SelectedRows = shows;
                        if (shows.Count == 0) throw new MissingEpisodeSheetException("No show rows selected; nothing to process.");
                        Type("DB shows: " + all.Count + " | Selected: " + shows.Count + " | As of: " + today.ToString("yyyy-MM-dd"), 0, 0, 1);
                        using (var emby = new EmbyClient(GetEmbyServerUrl(), GetEmbyApiKey()))
                        {
                            cancellation.Token.ThrowIfCancellationRequested();
                            var inventory = emby.GetSeries();
                            // An empty library would otherwise mark every row "Not in Emby".
                            if (inventory.Count == 0) throw new EmbyException("Emby returned no TV series; check the server before writing \"Not in Emby\".");
                            Type("Loaded " + inventory.Count + " Emby series.", 0, 0, 1, "Blue");
                            var tvdb = new TvdbClient(() => LocalEnvironment.GetRequired("TVDB_API_KEY"));
                            var preview = new MissingEpisodePreview(inventory, emby.GetEpisodes, tvdb.GetOfficialEpisodes);
                            MissingEpisodeWorkflow.Execute(run, shows, show => {
                                Type("Checking " + (run.Results.Count + 1) + "/" + shows.Count + ": " + show.Title + " (DB row " + show.Row + ")", 0, 0, 1, "Blue");
                                return preview.Check(show, today);
                            }, update ? sheet : null, cancellation.Token, (result, index) => {
                                if (result.NotInEmby)
                                {
                                    Type("Not in Emby: no Emby series has this ID | " + result.Show.OutputCell + " | " + result.SheetOutcome, 0, 0, 1, "Yellow");
                                    if (result.MatchCandidates.Count > 0) Type("Review: " + string.Join("; ", result.MatchCandidates), 0, 0, 1, "Yellow");
                                }
                                else if (!result.ProposedValue.HasValue)
                                {
                                    Type("FAILED [" + result.Status + "]: " + result.Message + " Existing value preserved.", 0, 0, 1, "Red");
                                    if (result.MatchCandidates.Count > 0) Type("Review: " + string.Join("; ", result.MatchCandidates), 0, 0, 1, "Yellow");
                                }
                                else
                                {
                                    Type("Owned: " + result.OwnedEpisodeCount + " | Aired: " + result.AiredEpisodeCount + " | Missing: " + result.ProposedValue + " | " + result.Show.OutputCell + " | " + result.SheetOutcome, 0, 0, 1);
                                    Type("Missing episodes: " + result.ProposedText, 0, 0, 1);
                                }
                                foreach (string warning in result.Warnings) Type("Review: " + warning, 0, 0, 1, "Yellow");
                                Type("Progress: " + index + "/" + shows.Count + " | Verified: " + run.SuccessfullyChecked + " | Updated so far: " + run.Updated, 0, 0, 1);
                            });
                        }
                    }
                }
                catch (OperationCanceledException) { run.Aborted = true; run.Error = "Cancelled; uncommitted rows preserved."; }
                catch (MissingEpisodeSheetException ex) { run.Aborted = true; run.Error = ex.Message; }
                catch (EmbyException ex) { run.Aborted = true; run.Error = ex.Message; }
                catch (TvdbException ex) { run.Aborted = true; run.Error = ex.Message; }
                catch (Exception ex)
                {
                    run.Aborted = true;
                    var cause = ex.GetBaseException();
                    var google = cause as Google.GoogleApiException;
                    run.Error = "Service/configuration failure (" + (google == null ? cause.GetType().Name : "HTTP " + (int)google.HttpStatusCode) + "). Check authorization and configuration; uncommitted values preserved.";
                }
                finally { Console.CancelKeyPress -= cancel; }
            }
            Type(run.Aborted ? "RUN ABORTED - partial results." : run.HasFailures ? "RUN FINISHED WITH FAILURES." : "RUN COMPLETE.", 0, 0, 1, run.HasFailures ? "Yellow" : "Green");
            if (run.Error != null) Type(run.Error, 0, 0, 1, "Red");
            Type("Total shows: " + run.TotalShows + " | Examined: " + run.Results.Count + "/" + run.SelectedShows + " | Successfully checked: " + run.SuccessfullyChecked, 0, 0, 1);
            Type("With missing episodes: " + run.WithMissing + " | Zero missing: " + run.ZeroMissing + " | Total missing episodes: " + run.TotalMissing + " | Not in Emby: " + run.NotInEmby + " | Failed: " + run.Failed, 0, 0, 1);
            Type("Updated Sheet rows: " + run.Updated + " | Unchanged Sheet rows: " + run.Unchanged + " | Sheet conflicts/failures: " + run.SheetFailures + " | Not checked: " + (run.SelectedShows - run.Results.Count), 0, 0, 1);
            if (!update) Type("Preview changes proposed: " + run.ProposedChanges + " | Google Sheet rows updated: 0", 0, 0, 1);
            try
            {
                string path = Path.GetFullPath("missing-episodes-" + (update ? "update-" : "preview-") + DateTime.Now.ToString("yyyyMMdd-HHmmss-fff") + ".json");
                File.WriteAllText(path, JsonConvert.SerializeObject(run.Report(today), Formatting.Indented));
                Type("Report (all checked shows, missing numbers and manual attention): " + path, 0, 0, 1, "Yellow");
            }
            catch (Exception) { DisplayMessage("error", "Could not save the report; console results remain available."); Environment.ExitCode = 1; }
            if (run.HasFailures) Environment.ExitCode = 1;
        }
    }
}
