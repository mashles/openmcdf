using System;

namespace OpenMcdf.Test
{
    public static class Helpers
    {
        public static byte[] GetBuffer(int count)
        {
            var r = new Random();
            var b = new byte[count];
            r.NextBytes(b);
            return b;
        }
        public static void FillBufferWithRandomData(byte[] buffer)
        {
            var r = new Random();
            r.NextBytes(buffer);
        }
        public static void FillBuffer(byte[] buffer,byte c)
        {
            for (var i = 0; i < buffer.Length; i++)
            {
                buffer[i] = c;
            }
        }

        public static byte[] GetBuffer(int count, byte c)
        {
            var b = new byte[count];
            FillBuffer(b,c);
            return b;
        }

        public static bool CompareBuffer(byte[] b, byte[] p)
        {
            var res = CompareBuffer(b, p, b.Length);
            return res && (b.Length == p.Length);
        }

        public static bool CompareBuffer(byte[] b, byte[] p, int count)
        {
            if (b == null && p == null)
                throw new Exception("Null buffers");

            if (b == null && p != null)
                return false;

            if (b != null && p == null)
                return false;


            for (var i = 0; i < count; i++)
            {
                if (b[i] != p[i])
                    return false;
            }

            return true;
        }
    }
}
