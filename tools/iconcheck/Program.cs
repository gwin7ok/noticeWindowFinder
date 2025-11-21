using System;
using System.Drawing;
using System.IO;

namespace IconCheck
{
    class Program
    {
        static int Main(string[] args)
        {
            if (args.Length < 1)
            {
                Console.Error.WriteLine("Usage: iconcheck <ico-path>");
                return 2;
            }
            var ico = args[0];
            if (!File.Exists(ico))
            {
                Console.Error.WriteLine($"File not found: {ico}");
                return 3;
            }

            try
            {
                bool hasAlpha = CheckIcoForAlpha(ico);
                Console.WriteLine($"ICO alpha present: {hasAlpha}");
                return hasAlpha ? 0 : 1;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("Error: " + ex.Message);
                return 100;
            }
        }

        private static bool CheckIcoForAlpha(string icoPath)
        {
            using (var fs = new FileStream(icoPath, FileMode.Open, FileAccess.Read))
            using (var br = new BinaryReader(fs))
            {
                ushort reserved = br.ReadUInt16();
                ushort type = br.ReadUInt16();
                ushort count = br.ReadUInt16();
                for (int i = 0; i < count; i++)
                {
                    byte width = br.ReadByte();
                    byte height = br.ReadByte();
                    byte colors = br.ReadByte();
                    byte reservedEntry = br.ReadByte();
                    ushort planes = br.ReadUInt16();
                    ushort bitCount = br.ReadUInt16();
                    uint bytesInRes = br.ReadUInt32();
                    uint imageOffset = br.ReadUInt32();

                    long pos = fs.Position;
                    fs.Seek(imageOffset, SeekOrigin.Begin);
                    byte[] header = br.ReadBytes((int)Math.Min(8, bytesInRes));
                    // PNG signature
                    if (header.Length >= 8 && header[0] == 0x89 && header[1] == 0x50 && header[2] == 0x4E && header[3] == 0x47 && header[4] == 0x0D && header[5] == 0x0A && header[6] == 0x1A && header[7] == 0x0A)
                    {
                        fs.Seek(imageOffset, SeekOrigin.Begin);
                        byte[] pngData = br.ReadBytes((int)bytesInRes);
                        using (var ms = new MemoryStream(pngData))
                        using (var img = Image.FromStream(ms))
                        using (var bmp = new Bitmap(img))
                        {
                            for (int y = 0; y < bmp.Height; y++)
                            for (int x = 0; x < bmp.Width; x++)
                            {
                                if (bmp.GetPixel(x, y).A != 255) return true;
                            }
                        }
                    }
                    fs.Seek(pos, SeekOrigin.Begin);
                }
            }
            return false;
        }
    }
}