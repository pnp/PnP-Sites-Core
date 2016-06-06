using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Extensions
{
    public static class StreamExtensions
    {
        public static MemoryStream ToMemoryStream(this Stream source)
        {
            var stream = source as MemoryStream;
            if (stream != null) return stream;
            MemoryStream target = new MemoryStream();
            const int bufSize = 65535;
            byte[] buf = new byte[bufSize];
            int bytesRead = -1;
            while ((bytesRead = source.Read(buf, 0, bufSize)) > 0)
                target.Write(buf, 0, bytesRead);
            return target;
        }
    }
}
