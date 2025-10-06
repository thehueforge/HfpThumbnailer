using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

public class StreamWrapper : IStream
{
    private readonly Stream _baseStream;

    public StreamWrapper(Stream stream)
    {
        _baseStream = stream;
    }

    public void Read(byte[] pv, int cb, IntPtr pcbRead)
    {
        int bytesRead = _baseStream.Read(pv, 0, cb);
        if (pcbRead != IntPtr.Zero)
            Marshal.WriteInt32(pcbRead, bytesRead);
    }

    public void Write(byte[] pv, int cb, IntPtr pcbWritten)
    {
        _baseStream.Write(pv, 0, cb);
        if (pcbWritten != IntPtr.Zero)
            Marshal.WriteInt32(pcbWritten, cb);
    }

    public void Seek(long dlibMove, int dwOrigin, IntPtr plibNewPosition)
    {
        long pos = _baseStream.Seek(dlibMove, (SeekOrigin)dwOrigin);
        if (plibNewPosition != IntPtr.Zero)
            Marshal.WriteInt64(plibNewPosition, pos);
    }

    public void SetSize(long libNewSize)
    {
        _baseStream.SetLength(libNewSize);
    }

    public void CopyTo(IStream pstm, long cb, IntPtr pcbRead, IntPtr pcbWritten)
    {
        const int bufferSize = 4096;
        byte[] buffer = new byte[bufferSize];
        long bytesRead = 0;
        long bytesWritten = 0;

        while (cb > 0)
        {
            int read = _baseStream.Read(buffer, 0, (int)Math.Min(bufferSize, cb));
            if (read == 0) break;

            bytesRead += read;
            cb -= read;

            pstm.Write(buffer, read, IntPtr.Zero);
            bytesWritten += read;
        }

        if (pcbRead != IntPtr.Zero)
            Marshal.WriteInt64(pcbRead, bytesRead);
        if (pcbWritten != IntPtr.Zero)
            Marshal.WriteInt64(pcbWritten, bytesWritten);
    }

    public void Commit(int grfCommitFlags) => _baseStream.Flush();
    public void Revert() => throw new NotSupportedException();
    public void LockRegion(long libOffset, long cb, int dwLockType) => throw new NotSupportedException();
    public void UnlockRegion(long libOffset, long cb, int dwLockType) => throw new NotSupportedException();
    public void Stat(out System.Runtime.InteropServices.ComTypes.STATSTG pstatstg, int grfStatFlag)
    {
        pstatstg = new System.Runtime.InteropServices.ComTypes.STATSTG
        {
            cbSize = _baseStream.Length,
            type = 2 // STGTY_STREAM
        };
    }


    public void Clone(out IStream ppstm) => throw new NotSupportedException();

}
