# HFP Thumbnail Handler

Windows Shell Extension for displaying thumbnails of HueForge Project files (.hfp) in Windows Explorer.

## Features

- ✅ Shows thumbnail previews of .hfp files in Windows Explorer
- ✅ **Full OneDrive support** - works seamlessly with OneDrive folders and on-demand files
- ✅ **Universal cloud storage** - supports OneDrive Personal, Business, and Google Drive
- ✅ High-performance with precompiled Regex patterns and optimized file detection
- ✅ Robust error handling with detailed logging
- ✅ Automatic thumbnail cache management and COM cleanup

## Quick Setup

**Prerequisites:**
- .NET SDK: https://dotnet.microsoft.com/download
- .NET Framework 4.8 Dev Pack: https://dotnet.microsoft.com/en-us/download/dotnet-framework/thank-you/net481-developer-pack-offline-installer

**Build & Install:**
1. Right-click `build.bat` → "Run as administrator"
2. Choose option 2 (Build and attach)
3. Done! Thumbnails work in Explorer

## Build Options

Run `build.bat` as Administrator and choose:

- **Option 1**: Build only
  - Creates optimized DLL with OneDrive support
  - No registration or file association

- **Option 2**: Build and attach (recommended)
  - Builds the DLL
  - Registers COM component
  - Sets up .hfp file association
  - Registers Windows shell extension
  - Clears thumbnail cache
  - Restarts Explorer

- **Option 3**: Remove
  - Unregisters COM component
  - Removes all registry entries
  - Cleans up completely

## Key Features

### Cloud Storage Support
- **OneDrive Integration**: Full support for OneDrive Personal and Business accounts
- **On-Demand Files**: Automatically handles files that aren't downloaded locally
- **Smart Path Detection**: Uses registry scanning and intelligent path matching
- **Google Drive Compatible**: Works with Google Drive folders as before

### Performance & Reliability
- **Optimized Detection**: Precompiled Regex patterns for fast path recognition
- **Robust File Handling**: Advanced file availability checking and error recovery
- **Automatic Cleanup**: Built-in COM process management prevents build conflicts
- **Enhanced Logging**: Detailed debug information in `%TEMP%\hfp_thumbnail_debug.log`

### Technical Features
- Base64 image extraction from HFP JSON files
- Intelligent fallback thumbnail generation
- Windows Explorer integration (all icon sizes)
- Multi-layered cloud storage detection system

## Project Structure

- `build.bat` - All-in-one build and setup script
- `HfpThumbnailProvider.cs` - Main thumbnail provider with OneDrive optimizations
- `Interfaces.cs` - COM interface definitions
- `StreamWrapper.cs` - IStream wrapper for .NET compatibility

## License

MIT License