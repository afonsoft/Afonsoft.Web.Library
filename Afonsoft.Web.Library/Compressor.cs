using System.IO;
using System.IO.Compression;

namespace Afonsoft.Web.Library
{
    /// <summary>
    /// Compactador de bytes
    /// </summary>
    public class Compressor
    {
        /// <summary>
        /// Compactar um array de byte
        /// </summary>
        public static byte[] Compress(byte[] data)
        {
            try
            {
                using (MemoryStream output = new MemoryStream())
                {
                    using (GZipStream gzip = new GZipStream(output, CompressionMode.Compress))
                    {
                        gzip.Write(data, 0, data.Length);
                        gzip.Close();
                        return output.ToArray();
                    }
                }
            }
            catch { return data; }
        }
        /// <summary>
        /// Descompactar um array de byte
        /// </summary>
        public static byte[] Decompress(byte[] data)
        {
            try
            {
                using (MemoryStream input = new MemoryStream())
                {
                    input.Write(data, 0, data.Length);
                    input.Position = 0;
                    using (GZipStream gzip = new GZipStream(input, CompressionMode.Decompress))
                    {
                        using (MemoryStream output = new MemoryStream())
                        {
                            byte[] buff = new byte[1024];
                            int read = -1;
                            read = gzip.Read(buff, 0, buff.Length);
                            while (read > 0)
                            {
                                output.Write(buff, 0, read);
                                read = gzip.Read(buff, 0, buff.Length);
                            }
                            gzip.Close();
                            return output.ToArray();
                        }
                    }
                }
            }
            catch { return data; }
        }

    }
}
