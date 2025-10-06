# HFP Thumbnail Handler

Windows Shell Extension for displaying thumbnails of HueForge Project files (.hfp) with OneDrive support.

## Features

- ✅ Shows thumbnail previews of .hfp files in Windows Explorer
- ✅ **OneDrive support** - works with on-demand synced files
- ✅ Optimized performance with precompiled Regex patterns
- ✅ Comprehensive error handling and logging
- ✅ Automatic thumbnail cache management

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

### OneDrive Support
- **Smart Detection**: Automatically detects OneDrive folders
- **On-Demand Files**: Handles files that aren't downloaded locally
- **Multiple Accounts**: Supports Personal and Business OneDrive
- **Performance Optimized**: Precompiled Regex patterns for path detection

### Technical Features
- Base64 image extraction from HFP JSON files
- Fallback thumbnail generation for corrupted files
- Comprehensive error logging (`%TEMP%\hfp_thumbnail_debug.log`)
- Windows Explorer integration (Large Icons, Extra Large Icons)

## Project Structure

- `build.bat` - All-in-one build and setup script
- `HfpThumbnailProvider.cs` - Main thumbnail provider with OneDrive optimizations
- `Interfaces.cs` - COM interface definitions
- `StreamWrapper.cs` - IStream wrapper for .NET compatibility

## License

MIT License