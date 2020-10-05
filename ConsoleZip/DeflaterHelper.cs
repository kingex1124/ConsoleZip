using ICSharpCode.SharpZipLib.Zip.Compression;
using ICSharpCode.SharpZipLib.Zip.Compression.Streams;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleZip
{
    public class DeflaterHelper
    {
        public static byte[] Compress(byte[] pBytes)
        {
            MemoryStream mMemory = new MemoryStream();
            Deflater mDeflater = new Deflater(Deflater.BEST_COMPRESSION);
            using (DeflaterOutputStream mStream = new DeflaterOutputStream(mMemory, mDeflater, 131072))
            {
                mStream.Write(pBytes, 0, pBytes.Length);
            }

            return mMemory.ToArray();
        }

        public static byte[] DeCompress(byte[] pBytes)
        {
            MemoryStream mMemory = new MemoryStream();
            using (InflaterInputStream mStream = new InflaterInputStream(new MemoryStream(pBytes)))
            {
                Int32 mSize;
                byte[] mWriteData = new byte[4096];
                while (true)
                {
                    mSize = mStream.Read(mWriteData, 0, mWriteData.Length);
                    if (mSize > 0)
                        mMemory.Write(mWriteData, 0, mSize);
                    else
                        break;
                }
            }
            return mMemory.ToArray();
        }
    }
}
