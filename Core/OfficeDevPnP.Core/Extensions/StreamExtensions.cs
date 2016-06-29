using System.IO;

namespace OfficeDevPnP.Core.Extensions
{
    public static class StreamExtensions
    {
        public static MemoryStream ToMemoryStream(this Stream source)
        {
            var stream = source as MemoryStream;
            if (stream != null) return stream;
            var target = new MemoryStream();
            const int bufSize = 65535;
            var buf = new byte[bufSize];
            int bytesRead;
            while ((bytesRead = source.Read(buf, 0, bufSize)) > 0)
                target.Write(buf, 0, bytesRead);
            target.Position = 0;
            return target;
        }
    }
}