using System;

namespace OpenMcdf.PerfTest
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

        public static byte[] GetBuffer(int count, byte c)
        {
            var b = new byte[count];
            for (var i = 0; i < b.Length; i++)
            {
                b[i] = c;
            }

            return b;
        }

        public static bool CompareBuffer(byte[] b, byte[] p)
        {
            if (b == null && p == null)
                throw new Exception("Null buffers");

            if (b == null && p != null) 
                return false;

            if (b != null && p == null) 
                return false;

            if (b.Length != p.Length)
                return false;

            for (var i = 0; i < b.Length; i++)
            {
                if (b[i] != p[i])
                    return false;
            }

            return true;
        }
    }
}
