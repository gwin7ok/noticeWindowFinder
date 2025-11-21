using System;
using System.IO;
using System.Drawing;

namespace IconExtract
{
    class Program
    {
        static int Main(string[] args)
        {
            if (args.Length < 1)
            {
                Console.Error.WriteLine("Usage: iconextract <ico-path> [out-dir]");
                return 2;
            }
            var ico = args[0];
            var outDir = args.Length >= 2 ? args[1] : Path.Combine(Directory.GetCurrentDirectory(), "out");
            if (!File.Exists(ico))
            {
                Console.Error.WriteLine($"File not found: {ico}");
                return 3;
            }
            Directory.CreateDirectory(outDir);

            try
            {
                using (var fs = new FileStream(ico, FileMode.Open, FileAccess.Read))
                using (var br = new BinaryReader(fs))
                {
                    ushort reserved = br.ReadUInt16();
                    ushort type = br.ReadUInt16();
                    ushort count = br.ReadUInt16();
                    Console.WriteLine($"Found {count} entries in ICO");
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

                        int logicalWidth = width == 0 ? 256 : width;
                        int logicalHeight = height == 0 ? 256 : height;

                        Console.WriteLine($"Entry {i}: {logicalWidth}x{logicalHeight} bytes={bytesInRes} offset={imageOffset}");

                        long pos = fs.Position;
                        fs.Seek(imageOffset, SeekOrigin.Begin);
                        byte[] header = br.ReadBytes((int)Math.Min(8, bytesInRes));
                        fs.Seek(imageOffset, SeekOrigin.Begin);

                        // PNG signature
                        if (header.Length >= 8 && header[0] == 0x89 && header[1] == 0x50 && header[2] == 0x4E && header[3] == 0x47 && header[4] == 0x0D && header[5] == 0x0A && header[6] == 0x1A && header[7] == 0x0A)
                        {
                            byte[] pngData = br.ReadBytes((int)bytesInRes);
                            var outPath = Path.Combine(outDir, Path.GetFileNameWithoutExtension(ico) + $"_{logicalWidth}x{logicalHeight}.png");
                            File.WriteAllBytes(outPath, pngData);
                            Console.WriteLine($"Wrote PNG frame: {outPath}");
                        }
                        else
                        {
                            // Not PNG: try to read as BMP/DIB and convert
                            try
                            {
                                byte[] dib = br.ReadBytes((int)bytesInRes);
                                using (var ms = new MemoryStream(dib))
                                {
                                    // Some ICOs store BMP/DIB starting with BITMAPINFOHEADER; Image.FromStream may accept it
                                    try
                                    {
                                        using (var img = Image.FromStream(ms))
                                        {
                                            var bmp = new Bitmap(img);
                                            var outPath = Path.Combine(outDir, Path.GetFileNameWithoutExtension(ico) + $"_{logicalWidth}x{logicalHeight}.png");
                                            bmp.Save(outPath, System.Drawing.Imaging.ImageFormat.Png);
                                            Console.WriteLine($"Converted BMP/DIB frame to PNG: {outPath}");
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"Skipping non-PNG frame {i} (could not decode as image): {ex.Message}");
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Failed to read non-PNG frame {i}: {ex.Message}");
                            }
                        }

                        fs.Seek(pos, SeekOrigin.Begin);
                    }
                }

                Console.WriteLine("Done.");
                return 0;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("Error: " + ex.Message);
                return 100;
            }
        }
    }
}