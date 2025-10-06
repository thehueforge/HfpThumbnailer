using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text.RegularExpressions;
using Microsoft.Win32;

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

        // Precompiled Regex patterns for OneDrive path detection (performance optimization)
        private static readonly Regex[] OneDrivePathPatterns = {
            new Regex(@"\\onedrive\\", RegexOptions.Compiled | RegexOptions.IgnoreCase),
            new Regex(@"\\onedrive - ", RegexOptions.Compiled | RegexOptions.IgnoreCase),
            new Regex(@"\\one drive\\", RegexOptions.Compiled | RegexOptions.IgnoreCase),
            new Regex(@"users\\[^\\]+\\onedrive", RegexOptions.Compiled | RegexOptions.IgnoreCase),
            new Regex(@"documents\\onedrive", RegexOptions.Compiled | RegexOptions.IgnoreCase)
        };

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
            
            // Major limitation was that OneDrive "on-demand" files don't show thumbnails
            // This detection helps identify OneDrive files to handle them differently
            bool isOneDrive = IsOneDriveFile(_filePath);
            LogMessage($"OneDrive detection: {isOneDrive}");
            
            if (isOneDrive)
            {
                // Check if the file is actually downloaded locally vs just a placeholder
                // OneDrive "on-demand" files appear in Explorer but aren't fully downloaded
                bool isLocallyAvailable = IsOneDriveFileAvailableLocally(_filePath);
                LogMessage($"OneDrive file locally available: {isLocallyAvailable}");
                
                if (!isLocallyAvailable)
                {
                    LogMessage("OneDrive file is on-demand only - will attempt to download");
                }
            }
            
            try
            {
                // Use smart file access that can handle OneDrive on-demand files
                // Standard File.OpenRead() fails on OneDrive placeholders
                Stream fileStream = CreateFileStream(_filePath);
                _stream = new StreamWrapper(fileStream); // Wraps Stream in IStream for compatibility
                LogMessage("File stream created successfully");
            }
            catch (Exception ex)
            {
                LogMessage($"Error initializing from file: {ex.Message}");
                if (isOneDrive)
                {
                    // Provide user-friendly error message specific to OneDrive issues
                    LogMessage("OneDrive-specific error: This may be due to on-demand sync. " +
                             "Try setting the file to 'Always keep on this device' in OneDrive settings.");
                }
                throw; // Re-throw to let Windows handle the error appropriately
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

            // Additional OneDrive context during thumbnail generation
            // This helps with debugging and provides better error context
            if (!string.IsNullOrEmpty(_filePath))
            {
                bool isOneDrive = IsOneDriveFile(_filePath);
                if (isOneDrive)
                {
                    LogMessage($"Processing OneDrive file for thumbnail: {_filePath}");
                    
                    // Double-check file availability at thumbnail time
                    // File status can change between Initialize() and GetThumbnail() calls
                    bool isAvailable = IsOneDriveFileAvailableLocally(_filePath);
                    LogMessage($"OneDrive file availability status: {(isAvailable ? "Available" : "On-demand")}");
                }
            }

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
                
                // Provide visual feedback for OneDrive-specific failures
                // Instead of generic error, show user this is a OneDrive file with distinctive icon
                if (!string.IsNullOrEmpty(_filePath) && IsOneDriveFile(_filePath))
                {
                    LogMessage("Creating OneDrive-specific fallback thumbnail");
                    return CreateOneDriveFallbackThumbnail(size);
                }
                
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

        private Bitmap CreateOneDriveFallbackThumbnail(int size)
        {
            var bitmap = new Bitmap(size, size);
            using (var graphics = Graphics.FromImage(bitmap))
            {
                // Create a distinctive OneDrive-themed background
                var brush = new System.Drawing.Drawing2D.LinearGradientBrush(
                    new Rectangle(0, 0, size, size),
                    Color.FromArgb(0, 120, 215), // OneDrive blue
                    Color.FromArgb(40, 160, 255), // Lighter blue
                    45f);
                graphics.FillRectangle(brush, 0, 0, size, size);

                // Add OneDrive cloud icon representation
                using (var whiteBrush = new SolidBrush(Color.White))
                {
                    // Simple cloud shape
                    int cloudSize = size / 3;
                    int x = (size - cloudSize) / 2;
                    int y = (size - cloudSize) / 2 - size / 8;
                    
                    graphics.FillEllipse(whiteBrush, x, y, cloudSize, cloudSize / 2);
                    graphics.FillEllipse(whiteBrush, x - cloudSize / 4, y + cloudSize / 6, cloudSize / 2, cloudSize / 3);
                    graphics.FillEllipse(whiteBrush, x + cloudSize / 2, y + cloudSize / 6, cloudSize / 2, cloudSize / 3);

                    // Add text below
                    using (var font = new Font("Arial", Math.Max(6, size / 12), FontStyle.Bold))
                    {
                        var text = "OneDrive";
                        var textSize = graphics.MeasureString(text, font);
                        var textX = (size - textSize.Width) / 2;
                        var textY = y + cloudSize / 2 + 5;
                        graphics.DrawString(text, font, whiteBrush, textX, textY);
                    }
                }
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

        #region OneDrive Detection
        
        // Core limitation was that OneDrive files don't show thumbnails
        // Unlike Google Drive (which downloads files locally), OneDrive uses "on-demand" sync
        // where files appear in Explorer but aren't actually downloaded until accessed
        // This section implements comprehensive OneDrive detection and handling
        //
        // ESSENTIAL FEATURES
        // - IsOneDriveFile(): Basic OneDrive path detection
        // - IsOneDriveFileAvailableLocally(): Check if file is downloaded
        // - CreateFileStream(): Smart file access that works with OneDrive
        //
        // OPTIONAL ENHANCEMENTS
        // - GetOneDrivePaths(): Comprehensive registry scanning for all OneDrive typesW
        // - TryForceOneDriveDownload(): Attempt automatic download of on-demand files
        // - CreateOneDriveFallbackThumbnail(): Visual feedback for OneDrive-specific errors
        // - Extensive logging and error handling for debugging

        /// <summary>
        /// Detects if a file path is located within OneDrive folders
        /// REASON: Need to identify OneDrive files to handle them with special logic
        /// </summary>
        private bool IsOneDriveFile(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
                return false;

            try
            {
                string normalizedPath = Path.GetFullPath(filePath).ToLowerInvariant();
                
                // Check against known OneDrive paths
                var oneDrivePaths = GetOneDrivePaths();
                
                foreach (var oneDrivePath in oneDrivePaths)
                {
                    if (normalizedPath.StartsWith(oneDrivePath.ToLowerInvariant()))
                    {
                        LogMessage($"File detected in OneDrive path: {oneDrivePath}");
                        return true;
                    }
                }

                // Additional heuristic checks
                if (IsOneDrivePathHeuristic(normalizedPath))
                {
                    LogMessage($"File detected as OneDrive via heuristics");
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                LogMessage($"Error detecting OneDrive path: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Gets OneDrive folder paths from registry and environment
        /// REASON: OneDrive installs in different locations (Personal, Business, Multiple accounts)
        /// Must check all possible locations to reliably detect OneDrive files
        /// </summary>
        private string[] GetOneDrivePaths()
        {
            var paths = new List<string>();

            try
            {
                // Method 1: Registry - Personal OneDrive
                var personalPath = GetOneDrivePathFromRegistry("PersonalFolder");
                if (!string.IsNullOrEmpty(personalPath))
                    paths.Add(personalPath);

                // Method 2: Registry - Business OneDrive
                var businessPath = GetOneDrivePathFromRegistry("BusinessFolder");
                if (!string.IsNullOrEmpty(businessPath))
                    paths.Add(businessPath);

                // Method 3: Registry - OneDrive Commercial
                var commercialPaths = GetOneDriveCommercialPaths();
                paths.AddRange(commercialPaths);

                // Method 4: Environment variables
                var envPath = Environment.GetEnvironmentVariable("OneDriveCommercial");
                if (!string.IsNullOrEmpty(envPath))
                    paths.Add(envPath);

                envPath = Environment.GetEnvironmentVariable("OneDriveConsumer");
                if (!string.IsNullOrEmpty(envPath))
                    paths.Add(envPath);

                // Method 5: Common default locations
                var userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                paths.Add(Path.Combine(userProfile, "OneDrive"));
                paths.Add(Path.Combine(userProfile, "OneDrive - Personal"));

            }
            catch (Exception ex)
            {
                LogMessage($"Error getting OneDrive paths: {ex.Message}");
            }

            // Remove duplicates and invalid paths
            return paths.Where(p => !string.IsNullOrEmpty(p) && Directory.Exists(p))
                       .Distinct()
                       .ToArray();
        }

        /// <summary>
        /// Gets OneDrive path from registry
        /// </summary>
        private string GetOneDrivePathFromRegistry(string folderType)
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\OneDrive\Accounts\Personal"))
                {
                    if (key != null)
                    {
                        var path = key.GetValue("UserFolder") as string ?? key.GetValue(folderType) as string;
                        if (!string.IsNullOrEmpty(path))
                            return path;
                    }
                }

                using (var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\OneDrive"))
                {
                    if (key != null)
                    {
                        var path = key.GetValue("UserFolder") as string;
                        if (!string.IsNullOrEmpty(path))
                            return path;
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Error reading OneDrive registry for {folderType}: {ex.Message}");
            }

            return null;
        }

        /// <summary>
        /// Gets OneDrive for Business paths from registry
        /// </summary>
        private string[] GetOneDriveCommercialPaths()
        {
            var paths = new List<string>();

            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\OneDrive\Accounts\Business1"))
                {
                    if (key != null)
                    {
                        var path = key.GetValue("UserFolder") as string;
                        if (!string.IsNullOrEmpty(path))
                            paths.Add(path);
                    }
                }

                // Check for multiple business accounts
                for (int i = 1; i <= 5; i++)
                {
                    try
                    {
                        using (var key = Registry.CurrentUser.OpenSubKey($@"Software\Microsoft\OneDrive\Accounts\Business{i}"))
                        {
                            if (key != null)
                            {
                                var path = key.GetValue("UserFolder") as string;
                                if (!string.IsNullOrEmpty(path) && !paths.Contains(path))
                                    paths.Add(path);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogMessage($"Error reading OneDrive Business{i} registry: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Error reading OneDrive commercial paths: {ex.Message}");
            }

            return paths.ToArray();
        }

        /// <summary>
        /// Uses heuristics to detect OneDrive paths
        /// PERFORMANCE: Uses precompiled Regex objects for better performance
        /// </summary>
        private bool IsOneDrivePathHeuristic(string normalizedPath)
        {
            // Use precompiled Regex patterns for better performance
            return OneDrivePathPatterns.Any(regex => regex.IsMatch(normalizedPath));
        }

        /// <summary>
        /// Checks if a OneDrive file is available locally (not on-demand only)
        /// REASON: Critical check - OneDrive files can appear in Explorer but be "on-demand" placeholders
        /// FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS flag indicates file needs to be downloaded
        /// </summary>
        private bool IsOneDriveFileAvailableLocally(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                    return false;

                var attributes = File.GetAttributes(filePath);
                
                // FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS (0x400000) indicates on-demand file
                const FileAttributes RECALL_ON_DATA_ACCESS = (FileAttributes)0x400000;
                
                if ((attributes & RECALL_ON_DATA_ACCESS) != 0)
                {
                    LogMessage($"OneDrive file is on-demand only: {filePath}");
                    return false;
                }

                // Additional check: try to get file size without triggering download
                var fileInfo = new FileInfo(filePath);
                if (fileInfo.Length == 0 && (attributes & FileAttributes.SparseFile) != 0)
                {
                    LogMessage($"OneDrive file appears to be sparse/placeholder: {filePath}");
                    return false;
                }

                LogMessage($"OneDrive file is available locally: {filePath}");
                return true;
            }
            catch (Exception ex)
            {
                LogMessage($"Error checking OneDrive file availability: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Attempts to force download of a OneDrive on-demand file
        /// REASON: Optional enhancement - try to automatically download on-demand files
        /// When user accesses file, OneDrive should download it automatically
        /// </summary>
        private bool TryForceOneDriveDownload(string filePath)
        {
            try
            {
                LogMessage($"Attempting to force download OneDrive file: {filePath}");

                // Method 1: Try to read the file content to trigger download
                using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    // Try to read first few bytes to trigger download
                    byte[] buffer = new byte[1024];
                    int bytesRead = fileStream.Read(buffer, 0, buffer.Length);
                    
                    if (bytesRead > 0)
                    {
                        LogMessage($"Successfully triggered download by reading {bytesRead} bytes");
                        
                        // Wait a bit for download to complete
                        System.Threading.Thread.Sleep(100);
                        
                        // Check if file is now available
                        return IsOneDriveFileAvailableLocally(filePath);
                    }
                }

                return false;
            }
            catch (IOException ioEx)
            {
                LogMessage($"IO Exception during OneDrive download attempt: {ioEx.Message}");
                
                // Method 2: Try using File.ReadAllBytes which might trigger download differently
                try
                {
                    LogMessage("Trying alternative download method...");
                    var bytes = File.ReadAllBytes(filePath);
                    LogMessage($"Alternative method read {bytes.Length} bytes");
                    return bytes.Length > 0;
                }
                catch (Exception altEx)
                {
                    LogMessage($"Alternative download method failed: {altEx.Message}");
                    return false;
                }
            }
            catch (UnauthorizedAccessException authEx)
            {
                LogMessage($"Access denied during OneDrive download: {authEx.Message}");
                return false;
            }
            catch (Exception ex)
            {
                LogMessage($"Unexpected error during OneDrive download: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Handles OneDrive file initialization with download fallback
        /// </summary>
        private Stream GetOneDriveFileStream(string filePath)
        {
            try
            {
                // First, check if file is locally available
                if (IsOneDriveFileAvailableLocally(filePath))
                {
                    LogMessage("OneDrive file is already locally available");
                    return File.OpenRead(filePath);
                }

                // File is on-demand, try to force download
                LogMessage("OneDrive file is on-demand, attempting to download...");
                
                if (TryForceOneDriveDownload(filePath))
                {
                    LogMessage("OneDrive file download successful, opening stream");
                    return File.OpenRead(filePath);
                }
                else
                {
                    LogMessage("OneDrive file download failed");
                    
                    // Last resort: try to open anyway and see what happens
                    LogMessage("Attempting to open OneDrive file anyway as last resort...");
                    try
                    {
                        return File.OpenRead(filePath);
                    }
                    catch (Exception lastEx)
                    {
                        LogMessage($"Last resort failed: {lastEx.Message}");
                        throw new IOException($"OneDrive file '{filePath}' is on-demand and could not be downloaded. " +
                                            "Please ensure the file is set to 'Always keep on this device' or is synced locally.", lastEx);
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Error getting OneDrive file stream: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Enhanced file stream creation with OneDrive support
        /// REASON: Essential - replaces standard File.OpenRead() with OneDrive-aware version
        /// This is where the main fix happens - smart file access that works with OneDrive
        /// </summary>
        private Stream CreateFileStream(string filePath)
        {
            try
            {
                // Check if this is a OneDrive file
                if (IsOneDriveFile(filePath))
                {
                    LogMessage("Using OneDrive-aware file access");
                    return GetOneDriveFileStream(filePath);
                }
                else
                {
                    LogMessage("Using standard file access");
                    return File.OpenRead(filePath);
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Error creating file stream: {ex.Message}");
                throw;
            }
        }

        #endregion

        private Stream ConvertToStream(IStream comStream)
        {
            // Enhanced stream conversion with OneDrive context
            // When OneDrive files fail to load, this provides better error messages
            const int bufferSize = 4096;
            var memoryStream = new MemoryStream();
            var buffer = new byte[bufferSize];
            int bytesRead;
            IntPtr bytesReadPtr = Marshal.AllocHGlobal(sizeof(int));

            try
            {
                int totalBytesRead = 0;
                do
                {
                    comStream.Read(buffer, bufferSize, bytesReadPtr);
                    bytesRead = Marshal.ReadInt32(bytesReadPtr);
                    if (bytesRead > 0)
                    {
                        memoryStream.Write(buffer, 0, bytesRead);
                        totalBytesRead += bytesRead;
                    }
                } while (bytesRead > 0);

                LogMessage($"ConvertToStream: Read {totalBytesRead} total bytes from IStream");
                
                if (totalBytesRead == 0)
                {
                    LogMessage("Warning: No data read from IStream - this may indicate OneDrive sync issues");
                }

                memoryStream.Position = 0;
                return memoryStream;
            }
            catch (Exception ex)
            {
                LogMessage($"Error converting IStream to Stream: {ex.Message}");
                if (!string.IsNullOrEmpty(_filePath) && IsOneDriveFile(_filePath))
                {
                    LogMessage("Stream conversion failed for OneDrive file - this may be due to sync issues");
                }
                throw;
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