using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace BachFlixNfo.Features
{
    public sealed class AudioMuteFilterBuilder
    {
        public static readonly TimeSpan DefaultPaddingBefore = TimeSpan.FromMilliseconds(120);
        public static readonly TimeSpan DefaultPaddingAfter = TimeSpan.FromMilliseconds(160);
        private static readonly TimeSpan MinimumWordMuteDuration = TimeSpan.FromMilliseconds(350);

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

        public List<AudioMuteSegment> CreateSegments(IEnumerable<TranscriptProfanityHit> approvedHits)
        {
            var segments = new List<AudioMuteSegment>();

            if (approvedHits == null)
                return segments;

            foreach (TranscriptProfanityHit hit in approvedHits)
            {
                if (hit == null || !hit.Approved || hit.End <= hit.Start)
                    continue;

                TimeSpan before = hit.PaddingBefore < TimeSpan.Zero ? TimeSpan.Zero : hit.PaddingBefore;
                TimeSpan after = hit.PaddingAfter < TimeSpan.Zero ? TimeSpan.Zero : hit.PaddingAfter;
                TimeSpan start = hit.Start - before;
                if (start < TimeSpan.Zero)
                    start = TimeSpan.Zero;

                TimeSpan end = hit.End + after;
                ExpandToMinimumDuration(ref start, ref end, MinimumWordMuteDuration);

                if (end <= start)
                    continue;

                segments.Add(new AudioMuteSegment
                {
                    Start = start,
                    End = end,
                    SourceHit = hit
                });
            }

            return segments;
        }

        private static void ExpandToMinimumDuration(ref TimeSpan start, ref TimeSpan end, TimeSpan minimumDuration)
        {
            TimeSpan duration = end - start;
            if (duration >= minimumDuration)
                return;

            TimeSpan midpoint = TimeSpan.FromTicks((start.Ticks + end.Ticks) / 2);
            TimeSpan half = TimeSpan.FromTicks(minimumDuration.Ticks / 2);
            start = midpoint - half;
            if (start < TimeSpan.Zero)
                start = TimeSpan.Zero;

            end = start + minimumDuration;
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
        public TranscriptProfanityHit SourceHit { get; set; }
    }
}
