using System;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace OpenMcdf.Extensions.Test
{
    /// <summary>
    /// Summary description for UnitTest1
    /// </summary>
    [TestClass]
    public class CfsStreamExtensionsTest
    {
        private TestContext _testContextInstance;

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get => _testContextInstance;
            set => _testContextInstance = value;
        }

        #region Additional test attributes
        //
        // You can use the following additional attributes as you write your tests:
        //
        // Use ClassInitialize to run code before running the first test in the class
        // [ClassInitialize()]
        // public static void MyClassInitialize(TestContext testContext) { }
        //
        // Use ClassCleanup to run code after all tests in a class have run
        // [ClassCleanup()]
        // public static void MyClassCleanup() { }
        //
        // Use TestInitialize to run code before running each test 
        // [TestInitialize()]
        // public void MyTestInitialize() { }
        //
        // Use TestCleanup to run code after each test has run
        // [TestCleanup()]
        // public void MyTestCleanup() { }
        //
        #endregion

        [TestMethod]
        public void Test_AS_IOSTREAM_READ()
        {
            var cf = new CompoundFile("MultipleStorage.cfs");

            var s = cf.RootStorage.GetStorage("MyStorage").GetStream("MyStream").AsIoStream();
            var br = new BinaryReader(s);
            var result = br.ReadBytes(32);
            Assert.IsTrue(Helpers.CompareBuffer(Helpers.GetBuffer(32, 1), result));
        }

        [TestMethod]
        public void Test_AS_IOSTREAM_WRITE()
        {
            const string cmp = "Hello World of BinaryWriter !";

            var cf = new CompoundFile();
            var s = cf.RootStorage.AddStream("ANewStream").AsIoStream();
            var bw = new BinaryWriter(s);
            bw.Write(cmp);
            cf.Save("$ACFFile.cfs");
            cf.Close();

            cf = new CompoundFile("$ACFFile.cfs");
            var br = new BinaryReader(cf.RootStorage.GetStream("ANewStream").AsIoStream());
            var st = br.ReadString();
            Assert.IsTrue(st == cmp);
            cf.Close();
        }

        [TestMethod]
        public void Test_AS_IOSTREAM_MULTISECTOR_WRITE()
        {
            var data = new byte[670];
            for (var i = 0; i < data.Length; i++)
            {
                data[i] = (byte)(i % 255);
            }
            
            using (var cf = new CompoundFile())
            {
                using (var s = cf.RootStorage.AddStream("ANewStream").AsIoStream())
                {
                    using (var bw = new BinaryWriter(s))
                    {
                        bw.Write(data);
                        cf.Save("$ACFFile2.cfs");
                        cf.Close();
                    }
                }
            }

            // Works
            using (var cf = new CompoundFile("$ACFFile2.cfs"))
            {
                using (var br = new BinaryReader(cf.RootStorage.GetStream("ANewStream").AsIoStream()))
                {
                    var readData = new byte[data.Length];
                    var readCount = br.Read(readData, 0, readData.Length);
                    Assert.IsTrue(readCount == readData.Length);
                    Assert.IsTrue(data.SequenceEqual(readData));
                    cf.Close();
                }
            }

            // Won't work until #88 is fixed.
            using (var cf = new CompoundFile("$ACFFile2.cfs"))
            {
                using (var readStream = cf.RootStorage.GetStream("ANewStream").AsIoStream())
                {
                    byte[] readData;
                    using (var ms = new MemoryStream())
                    {
                        readStream.CopyTo(ms);
                        readData = ms.ToArray();
                    }

                    Assert.IsTrue(data.SequenceEqual(readData));
                    cf.Close();
                }
            }
        }
    }
}
