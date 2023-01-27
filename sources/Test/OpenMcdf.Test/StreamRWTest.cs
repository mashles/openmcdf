using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace OpenMcdf.Test
{
    [TestClass]
    public class StreamRwTest
    {
        [TestMethod]
        public void ReadInt64_MaxSizeRead()
        {
            var input = long.MaxValue;
            var bytes = BitConverter.GetBytes(input);
            long actual = 0;
            using (var memStream = new MemoryStream(bytes))
            {
                var reader = new StreamRw(memStream);
                actual = reader.ReadInt64();
            }
            Assert.AreEqual(input, actual);
        }

        [TestMethod]
        public void ReadInt64_SmallNumber()
        {
            long input = 1234;
            var bytes = BitConverter.GetBytes(input);
            long actual = 0;
            using (var memStream = new MemoryStream(bytes))
            {
                var reader = new StreamRw(memStream);
                actual = reader.ReadInt64();
            }
            Assert.AreEqual(input, actual);
        }

        [TestMethod]
        public void ReadInt64_Int32MaxPlusTen()
        {
            var input = (long)int.MaxValue + 10;
            var bytes = BitConverter.GetBytes(input);
            long actual = 0;
            using (var memStream = new MemoryStream(bytes))
            {
                var reader = new StreamRw(memStream);
                actual = reader.ReadInt64();
            }
            Assert.AreEqual(input, actual);
        }
    }
}
