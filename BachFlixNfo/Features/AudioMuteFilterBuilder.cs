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
                TimeSpan minimumDuration = hit.MinMuteDuration > TimeSpan.Zero ? hit.MinMuteDuration : MinimumWordMuteDuration;
                ExpandToMinimumDuration(ref start, ref end, minimumDuration);
                if (hit.MaxMuteDuration > TimeSpan.Zero)
                    ClampToMaximumDuration(ref start, ref end, hit.MaxMuteDuration, hit.Start, hit.End);

                if (end <= start)
                    continue;

                var segment = new AudioMuteSegment
                {
                    Start = start,
                    End = end,
                    SourceHit = hit
                };
                segment.SourceHits.Add(hit);
                segments.Add(segment);
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

        private static void ClampToMaximumDuration(ref TimeSpan start, ref TimeSpan end, TimeSpan maximumDuration, TimeSpan anchorStart, TimeSpan anchorEnd)
        {
            if (maximumDuration <= TimeSpan.Zero || end - start <= maximumDuration)
                return;

            TimeSpan midpoint = anchorEnd > anchorStart
                ? TimeSpan.FromTicks((anchorStart.Ticks + anchorEnd.Ticks) / 2)
                : TimeSpan.FromTicks((start.Ticks + end.Ticks) / 2);
            TimeSpan half = TimeSpan.FromTicks(maximumDuration.Ticks / 2);
            start = midpoint - half;
            if (start < TimeSpan.Zero)
                start = TimeSpan.Zero;

            end = start + maximumDuration;
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

    public sealed class CensorIntervalMerger
    {
        public List<AudioMuteSegment> Merge(IEnumerable<AudioMuteSegment> segments, TimeSpan mergeGap)
        {
            TimeSpan gap = mergeGap < TimeSpan.Zero ? TimeSpan.Zero : mergeGap;
            var ordered = (segments ?? Enumerable.Empty<AudioMuteSegment>())
                .Where(s => s != null && s.End > s.Start)
                .OrderBy(s => s.Start)
                .ThenBy(s => s.End)
                .ToList();

            var merged = new List<AudioMuteSegment>();
            foreach (AudioMuteSegment segment in ordered)
            {
                AudioMuteSegment last = merged.LastOrDefault();
                if (last != null && segment.Start <= last.End + gap)
                {
                    if (segment.End > last.End)
                        last.End = segment.End;

                    AddSourceHits(last, segment);
                    continue;
                }

                merged.Add(CloneSegment(segment));
            }

            return merged;
        }

        private static AudioMuteSegment CloneSegment(AudioMuteSegment segment)
        {
            var clone = new AudioMuteSegment
            {
                Start = segment.Start,
                End = segment.End,
                SourceHit = segment.SourceHit
            };
            AddSourceHits(clone, segment);
            return clone;
        }

        private static void AddSourceHits(AudioMuteSegment target, AudioMuteSegment source)
        {
            if (target == null || source == null)
                return;

            if (source.SourceHits.Count > 0)
            {
                foreach (TranscriptProfanityHit hit in source.SourceHits)
                {
                    if (hit != null && !target.SourceHits.Contains(hit))
                        target.SourceHits.Add(hit);
                }
            }
            else if (source.SourceHit != null && !target.SourceHits.Contains(source.SourceHit))
            {
                target.SourceHits.Add(source.SourceHit);
            }

            if (target.SourceHit == null && target.SourceHits.Count > 0)
                target.SourceHit = target.SourceHits[0];
        }
    }

    public sealed class AudioMuteSegment
    {
        public TimeSpan Start { get; set; }
        public TimeSpan End { get; set; }
        public TranscriptProfanityHit SourceHit { get; set; }
        public List<TranscriptProfanityHit> SourceHits { get; private set; }

        public AudioMuteSegment()
        {
            SourceHits = new List<TranscriptProfanityHit>();
        }
    }
}
