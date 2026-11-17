namespace BachFlixNfo.Features
{
    /// <summary>The existing FindResolution rules, shared without console output.</summary>
    internal static class VideoResolutionClassifier
    {
        internal static string Classify(int width, int height)
        {
            if (width == 1280) return "720p";
            if (width == 1920) return "1080p";
            if (height == 576) return "576p";
            if (height < 720) return "480p";
            if (height < 1080) return "720p";
            if (height < 1440) return "1080p";
            if (height < 2160) return "1440p";
            if (height < 4320) return "2160p";
            if (height >= 4320) return "4320p";
            return "N/A";
        }
    }
}
