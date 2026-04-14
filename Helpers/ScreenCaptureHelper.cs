using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Windows;

namespace AiInterviewAssistant;

/// <summary>
/// Captures the virtual screen (all monitors) to PNG bytes, downscaling if wider/taller than <paramref name="maxDimension"/>.
/// </summary>
internal static class ScreenCaptureHelper
{
    public static byte[] CaptureVirtualScreenToPngBytes(int maxDimension = 1920)
    {
        var x = (int)SystemParameters.VirtualScreenLeft;
        var y = (int)SystemParameters.VirtualScreenTop;
        var w = (int)SystemParameters.VirtualScreenWidth;
        var h = (int)SystemParameters.VirtualScreenHeight;

        if (w <= 0 || h <= 0)
            throw new InvalidOperationException("Invalid virtual screen size.");

        using var full = new Bitmap(w, h, PixelFormat.Format32bppArgb);
        using (var g = Graphics.FromImage(full))
        {
            g.CopyFromScreen(x, y, 0, 0, new System.Drawing.Size(w, h), CopyPixelOperation.SourceCopy);
        }

        using var toEncode = ResizeIfNeeded(full, maxDimension);
        using var ms = new MemoryStream();
        toEncode.Save(ms, ImageFormat.Png);
        return ms.ToArray();
    }

    private static Bitmap ResizeIfNeeded(Bitmap src, int maxDimension)
    {
        if (src.Width <= maxDimension && src.Height <= maxDimension)
            return new Bitmap(src);

        var scale = Math.Min((double)maxDimension / src.Width, (double)maxDimension / src.Height);
        var nw = Math.Max(1, (int)(src.Width * scale));
        var nh = Math.Max(1, (int)(src.Height * scale));
        var dst = new Bitmap(nw, nh, PixelFormat.Format32bppArgb);
        using (var g = Graphics.FromImage(dst))
        {
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.SmoothingMode = SmoothingMode.HighQuality;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;
            g.DrawImage(src, 0, 0, nw, nh);
        }

        return dst;
    }
}
