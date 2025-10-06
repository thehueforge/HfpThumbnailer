using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace HfpThumbnailHandler
{
    // Define the IThumbnailProvider interface
    [ComImport]
    [Guid("e357fccd-a995-4576-b01f-234630154e96")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IThumbnailProvider
    {
        void GetThumbnail(uint cx, out IntPtr hBitmap, out WTS_ALPHATYPE bitmapType);
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("b7d14566-0509-4cce-a71f-0a554233bd9b")]
    public interface IInitializeWithFile
    {
        void Initialize(
            [MarshalAs(UnmanagedType.LPWStr)] string pszFilePath,
            uint grfMode);
    }

    // Define the IInitializeWithStream interface
    [ComImport]
    [Guid("b824b49d-22ac-4161-ac8a-9916e8fa3f7f")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IInitializeWithStream
    {
        void Initialize(IStream stream, uint grfMode);
    }

    // Define the WTS_ALPHATYPE enum
    public enum WTS_ALPHATYPE
    {
        WTSAT_UNKNOWN = 0,
        WTSAT_RGB = 1,
        WTSAT_ARGB = 2,
    }
}

