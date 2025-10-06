# Changelog

All notable changes to this project will be documented in this file.

## [1.1.0] - 2024-10-06

### Added
- **Complete OneDrive Support**: Full compatibility with OneDrive Personal and Business accounts
- **Advanced Cloud Detection**: Multi-layered system using registry scanning, environment variables, and intelligent path matching
- **On-Demand File Handling**: Automatic detection and handling of files that aren't downloaded locally
- **Performance Optimizations**: Precompiled Regex patterns for faster path recognition and reduced CPU usage
- **Enhanced Build System**: Unified `build.bat` script with automatic COM cleanup and Explorer restart
- **Robust Error Handling**: Comprehensive logging and fallback mechanisms for cloud storage files

### Improved
- **File Detection Logic**: Enhanced algorithm that works across all major cloud storage providers
- **Thumbnail Generation**: Better handling of corrupted or inaccessible files with intelligent fallbacks
- **Development Workflow**: Streamlined build process with automatic cleanup of COM processes
- **Code Architecture**: Modular design with separated concerns for better maintainability

### Technical Enhancements
- Static readonly Regex arrays with compiled patterns for optimal performance
- Registry-based OneDrive path discovery with multiple fallback methods
- FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS checking for on-demand file detection
- Automatic COM Surrogate process cleanup during builds
- Enhanced debug logging with detailed cloud storage information

### Fixed
- **OneDrive Compatibility**: Resolved issue where thumbnails wouldn't appear in OneDrive folders
- **Build Reliability**: Eliminated file locking issues during development builds
- **Performance Issues**: Optimized Regex compilation and path matching algorithms
- **Error Recovery**: Improved handling of network-dependent and virtualized files

---

## [1.0.0] - Initial Release

### Features
- Basic HFP file thumbnail generation
- Windows Shell Extension integration
- Google Drive folder support
- Base64 image extraction from JSON
- Simple build and registration system