using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace BachFlixNfo.Features
{
    public sealed class AudioMuteFilterBuilder
    {
        public string BuildMuteFilter(IEnumerable<AudioMuteSegment> segments)
        {
            if (segments == null)
                return string.Empty;

            List<AudioMuteSegment> usableSegments = segments
                .Where(s => s != null && s.End > s.Start)
                .OrderBy(s => s.Start)
                .ThenBy(s => s.End)
                .ToList();

            if (usableSegments.Count == 0)
                return string.Empty;

            return string.Join(",", usableSegments.Select(BuildVolumeFilter));
        }

        public List<AudioMuteSegment> CreateSegments(WhisperAlignmentResult alignmentResult)
        {
            var segments = new List<AudioMuteSegment>();

            if (alignmentResult == null || alignmentResult.AlignedWords == null)
                return segments;

            foreach (AlignedProfanityWord word in alignmentResult.AlignedWords)
            {
                if (word == null || word.End <= word.Start)
                    continue;

                segments.Add(new AudioMuteSegment
                {
                    Start = word.Start,
                    End = word.End,
                    SourceOccurrence = word.SourceOccurrence
                });
            }

            return segments;
        }

        private static string BuildVolumeFilter(AudioMuteSegment segment)
        {
            return string.Format(
                CultureInfo.InvariantCulture,
                "volume=enable='between(t\\,{0}\\,{1})':volume=0",
                FormatSeconds(segment.Start),
                FormatSeconds(segment.End));
        }

        private static string FormatSeconds(TimeSpan value)
        {
            return value.TotalSeconds.ToString("0.000", CultureInfo.InvariantCulture);
        }
    }

    public sealed class AudioMuteSegment
    {
        public TimeSpan Start { get; set; }
        public TimeSpan End { get; set; }
        public ProfanityOccurrence SourceOccurrence { get; set; }
    }
}
