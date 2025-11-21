using System;
using System.Drawing;

class Program
{
    static int Main(string[] args)
    {
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: checkpngalpha <path-to-png>");
            return 2;
        }

        var path = args[0];
        try
        {
            using var bmp = new Bitmap(path);
            bool hasAlpha = false;
            for (int y = 0; y < bmp.Height; y++)
            {
                for (int x = 0; x < bmp.Width; x++)
                {
                    var c = bmp.GetPixel(x, y);
                    if (c.A < 255)
                    {
                        hasAlpha = true;
                        break;
                    }
                }
                if (hasAlpha) break;
            }
            Console.WriteLine($"HasAlpha: {hasAlpha}");
            return hasAlpha ? 0 : 1;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
            return 3;
        }
    }
}
