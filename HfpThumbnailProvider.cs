using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text.RegularExpressions;

namespace HfpThumbnailHandler
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [Guid("A2ECD1CB-B5D5-4136-82CA-9E4D994DABB3")]
    public class HfpThumbnailProvider : IThumbnailProvider, IInitializeWithStream, IInitializeWithFile
    {
        private IStream _stream;
        private string _filePath;
        private static readonly string LogPath = Path.Combine(Path.GetTempPath(), "hfp_thumbnail_debug.log");

        static HfpThumbnailProvider()
        {
            try
            {
                string path = typeof(HfpThumbnailProvider).Assembly.Location;
                File.AppendAllText(LogPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] DLL loaded from: {path}\n");
            }
            catch { }
        }

        private void LogMessage(string message)
        {
            try
            {
                using (var writer = new StreamWriter(LogPath, append: true))
                {
                    writer.AutoFlush = true;
                    writer.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}");
                }
            }
            catch { /* Ignore logging errors */ }
        }
        public void Initialize(string pszFilePath, uint grfMode)
        {
            _filePath = pszFilePath;
            LogMessage($"InitializeWithFile called for: {_filePath}");
            try
            {
                Stream fileStream = File.OpenRead(_filePath);
                _stream = new StreamWrapper(fileStream); // Wraps Stream in IStream for compatibility
            }
            catch (Exception ex)
            {
                LogMessage($"Error initializing from file: {ex.Message}");
            }
        }


        public void Initialize(IStream stream, uint grfMode)
        {
            LogMessage($"Initialize called with grfMode={grfMode}");
            _stream = stream;
        }

        public void GetThumbnail(uint cx, out IntPtr hBitmap, out WTS_ALPHATYPE bitmapType)
        {
            LogMessage($"GetThumbnail called with cx={cx}");

            hBitmap = IntPtr.Zero;
            bitmapType = WTS_ALPHATYPE.WTSAT_UNKNOWN;

            try
            {
                // Convert IStream to .NET Stream
                Stream netStream = ConvertToStream(_stream);
                LogMessage($"Stream converted, length: {netStream.Length}");

                // Read JSON content
                string jsonContent;
                using (var reader = new StreamReader(netStream))
                {
                    jsonContent = reader.ReadToEnd();
                }
                LogMessage($"JSON content read, length: {jsonContent.Length}");

                // Extract Base64 image data
                string base64Image = ExtractBase64Image(jsonContent);
                if (base64Image == null || base64Image.Length == 0)
                {
                    LogMessage("No base64 image found");
                    return;
                }
                LogMessage($"Base64 preview: {base64Image.Substring(0, Math.Min(60, base64Image.Length))}");
                LogMessage($"Base64 tail: {base64Image.Substring(Math.Max(0, base64Image.Length - 20))}");

                LogMessage($"Base64 image extracted, length: {base64Image.Length}");

                // Decode Base64 string
                byte[] imageBytes;
                try
                {
                    LogMessage("Sanitizing base64...");

                    // JSON-escaped backslashes (e.g., \\n or \\/) get through regex
                    base64Image = base64Image
                        .Replace("\\n", "")
                        .Replace("\\r", "")
                        .Replace("\\t", "")
                        .Replace("\\/", "/");  // undo JSON escaping

                    // Then remove *real* whitespace just in case
                    base64Image = Regex.Replace(base64Image, @"\s+", "");

                    // Base64 must be a multiple of 4 chars
                    int mod = base64Image.Length % 4;
                    if (mod != 0)
                    {
                        base64Image = base64Image.PadRight(base64Image.Length + (4 - mod), '=');
                        LogMessage($"Base64 string padded to length: {base64Image.Length}");
                    }

                    LogMessage($"Final sanitized length: {base64Image.Length}");

                    imageBytes = Convert.FromBase64String(base64Image);
                    LogMessage($"Base64 decoded successfully, byte length: {imageBytes.Length}");
                }
                catch (Exception ex)
                {
                    LogMessage($"Base64 decode error: {ex.Message}");
                    LogMessage($"Base64 string length: {base64Image.Length}");
                    LogMessage($"Base64 string preview: {base64Image.Substring(0, Math.Min(50, base64Image.Length))}...");
                    return;
                }


                // Log first 32 bytes for format analysis
                string hex = BitConverter.ToString(imageBytes.Take(Math.Min(32, imageBytes.Length)).ToArray());
                LogMessage($"First bytes: {hex}");

                // Try to create thumbnail
                Bitmap thumbnailBitmap = CreateThumbnail(imageBytes, (int)cx);
                if (thumbnailBitmap != null)
                {
                    hBitmap = thumbnailBitmap.GetHbitmap();
                    bitmapType = WTS_ALPHATYPE.WTSAT_ARGB;
                    LogMessage("Thumbnail created successfully");
                }
                else
                {
                    LogMessage("Failed to create thumbnail");
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Error in GetThumbnail: {ex.Message}");
                LogMessage($"Stack trace: {ex.StackTrace}");
            }
        }

        private Bitmap CreateThumbnail(byte[] imageBytes, int size)
        {
            try
            {
                // First try standard image formats
                using (var ms = new MemoryStream(imageBytes))
                {
                    try
                    {
                        using (var originalImage = Image.FromStream(ms))
                        {
                            LogMessage($"Standard image format detected: {originalImage.Width}x{originalImage.Height}");
                            return ResizeImage(originalImage, size);
                        }
                    }
                    catch (ArgumentException)
                    {
                        LogMessage("Standard image format failed, trying QImage format");
                        return DecodeQImageFormat(imageBytes, size);
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Error creating thumbnail: {ex.Message}");
                return CreateFallbackThumbnail(size);
            }
        }

        private Bitmap DecodeQImageFormat(byte[] qimageBytes, int size)
        {
            try
            {
                // For now, create a simple colored rectangle as a placeholder
                // This confirms the COM integration works while we figure out QImage parsing
                LogMessage("Creating placeholder thumbnail for QImage format");

                var bitmap = new Bitmap(size, size);
                using (var graphics = Graphics.FromImage(bitmap))
                {
                    // Create a distinctive red-to-yellow gradient background
                    var brush = new System.Drawing.Drawing2D.LinearGradientBrush(
                        new Rectangle(0, 0, size, size),
                        Color.Red,
                        Color.Yellow,
                        45f);
                    graphics.FillRectangle(brush, 0, 0, size, size);

                    // Add some text to identify this as a QImage placeholder
                    using (var font = new Font("Arial", Math.Max(8, size / 8)))
                    using (var textBrush = new SolidBrush(Color.Black))
                    {
                        var text = "QImg";
                        var textSize = graphics.MeasureString(text, font);
                        var x = (size - textSize.Width) / 2;
                        var y = (size - textSize.Height) / 2;
                        graphics.DrawString(text, font, textBrush, x, y);
                    }
                }
                return bitmap;
            }
            catch (Exception ex)
            {
                LogMessage($"Error in DecodeQImageFormat: {ex.Message}");
                return CreateFallbackThumbnail(size);
            }
        }

        private Bitmap CreateFallbackThumbnail(int size)
        {
            var bitmap = new Bitmap(size, size);
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.FillRectangle(Brushes.Gray, 0, 0, size, size);
                graphics.DrawString("?", new Font("Arial", size / 4), Brushes.White, size / 3, size / 3);
            }
            return bitmap;
        }

        private Bitmap ResizeImage(Image originalImage, int size)
        {
            // Calculate the aspect ratio
            float aspectRatio = (float)originalImage.Width / originalImage.Height;

            int newWidth, newHeight;
            if (aspectRatio > 1) // Width is greater than height
            {
                newWidth = size;
                newHeight = (int)(size / aspectRatio);
            }
            else // Height is greater than or equal to width
            {
                newWidth = (int)(size * aspectRatio);
                newHeight = size;
            }

            // Create a square bitmap with the requested size
            var bitmap = new Bitmap(size, size);
            using (var graphics = Graphics.FromImage(bitmap))
            {
                // Fill with a transparent or white background
                graphics.Clear(Color.Transparent); // or Color.White if you prefer

                // Calculate centering offsets
                int x = (size - newWidth) / 2;
                int y = (size - newHeight) / 2;

                graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

                // Draw the image centered in the square
                graphics.DrawImage(originalImage, x, y, newWidth, newHeight);
            }
            return bitmap;
        }

        private Stream ConvertToStream(IStream comStream)
        {
            const int bufferSize = 4096;
            var memoryStream = new MemoryStream();
            var buffer = new byte[bufferSize];
            int bytesRead;
            IntPtr bytesReadPtr = Marshal.AllocHGlobal(sizeof(int));

            try
            {
                do
                {
                    comStream.Read(buffer, bufferSize, bytesReadPtr);
                    bytesRead = Marshal.ReadInt32(bytesReadPtr);
                    if (bytesRead > 0)
                        memoryStream.Write(buffer, 0, bytesRead);
                } while (bytesRead > 0);

                memoryStream.Position = 0;
                return memoryStream;
            }
            finally
            {
                Marshal.FreeHGlobal(bytesReadPtr);
            }
        }

        private string ExtractBase64Image(string jsonContent)
        {
            try
            {
                // Simple regex-based JSON parsing to avoid dependencies
                var thumbnailMatch = Regex.Match(jsonContent, @"""thumbnail""\s*:\s*""((?:\\.|[^""])*)""", RegexOptions.IgnoreCase);

                if (thumbnailMatch.Success)
                {
                    return thumbnailMatch.Groups[1].Value;
                }

                LogMessage("No thumbnail or imageData property found in JSON");
                return null;
            }
            catch (Exception ex)
            {
                LogMessage($"Error parsing JSON: {ex.Message}");
                return null;
            }
        }
    }
}