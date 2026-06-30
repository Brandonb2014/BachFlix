using System;
using System.Collections.Generic;

namespace BachFlixNfo.Features
{
    public interface IWhisperAlignmentService
    {
        ExternalToolResolution ResolveWhisperX(
            string explicitPath,
            Action<string, string, int> log,
            IList<string> commandLog);

        WhisperAlignmentResult AlignWords(WhisperAlignmentRequest request);
    }

    public sealed class WhisperAlignmentService : IWhisperAlignmentService
    {
        public ExternalToolResolution ResolveWhisperX(
            string explicitPath,
            Action<string, string, int> log,
            IList<string> commandLog)
        {
            return ExternalToolResolver.Resolve(explicitPath, "whisperx", log, commandLog);
        }

        public WhisperAlignmentResult AlignWords(WhisperAlignmentRequest request)
        {
            throw new NotSupportedException("WhisperX alignment is intentionally not implemented in this phase.");
        }
    }

    public sealed class WhisperAlignmentRequest
    {
        public string MkvPath { get; set; }
        public string SubtitlePath { get; set; }
        public int AudioTrackIndex { get; set; }
        public IReadOnlyList<ProfanityOccurrence> OccurrencesToAlign { get; set; }

        public WhisperAlignmentRequest()
        {
            MkvPath = "";
            SubtitlePath = "";
            OccurrencesToAlign = new List<ProfanityOccurrence>();
        }
    }

    public sealed class WhisperAlignmentResult
    {
        public List<AlignedProfanityWord> AlignedWords { get; private set; }

        public WhisperAlignmentResult()
        {
            AlignedWords = new List<AlignedProfanityWord>();
        }
    }

    public sealed class AlignedProfanityWord
    {
        public ProfanityOccurrence SourceOccurrence { get; set; }
        public string Word { get; set; }
        public TimeSpan Start { get; set; }
        public TimeSpan End { get; set; }

        public AlignedProfanityWord()
        {
            Word = "";
        }
    }
}
