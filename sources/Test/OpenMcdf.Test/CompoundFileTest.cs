using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace OpenMcdf.Test
{
    /// <summary>
    /// Summary description for CompoundFileTest
    /// </summary>
    [TestClass]
    public class CompoundFileTest
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
        public void Test_COMPRESS_SPACE()
        {
            var filename = "MultipleStorage3.cfs"; // 22Kb

            var srcFile = new FileInfo(filename);

            File.Copy(filename, "MultipleStorage_Deleted_Compress.cfs", true);

            var cf = new CompoundFile("MultipleStorage_Deleted_Compress.cfs", CfsUpdateMode.Update, CfsConfiguration.SectorRecycle | CfsConfiguration.EraseFreeSectors);

            var st = cf.RootStorage.GetStorage("MyStorage");
            st = st.GetStorage("AnotherStorage");

            Assert.IsNotNull(st);
            st.Delete("Another2Stream");
            cf.Commit();
            cf.Close();

            CompoundFile.ShrinkCompoundFile("MultipleStorage_Deleted_Compress.cfs"); // -> 7Kb

            var dstFile = new FileInfo("MultipleStorage_Deleted_Compress.cfs");

            Assert.IsTrue(srcFile.Length > dstFile.Length);

        }

        [TestMethod]
        public void Test_ENTRY_NAME_LENGTH()
        {
            //Thanks to Mark Bosold for bug fix and unit

            var cf = new CompoundFile();

            // Cannot be equal.
            var maxCharactersStreamName = "1234567890123456789A12345678901"; // 31 chars
            var maxCharactersStorageName = "1234567890123456789012345678901"; // 31 chars

            // Try Storage entry name with max characters.
            Assert.IsNotNull(cf.RootStorage.AddStorage(maxCharactersStorageName));
            var strg = cf.RootStorage.GetStorage(maxCharactersStorageName);
            Assert.IsNotNull(strg);
            Assert.IsTrue(strg.Name == maxCharactersStorageName);


            // Try Stream entry name with max characters.
            Assert.IsNotNull(cf.RootStorage.AddStream(maxCharactersStreamName));
            var strm = cf.RootStorage.GetStream(maxCharactersStreamName);
            Assert.IsNotNull(strm);
            Assert.IsTrue(strm.Name == maxCharactersStreamName);

            var tooManyCharactersEntryName = "12345678901234567890123456789012"; // 32 chars

            try
            {
                // Try Storage entry name with too many characters.
                cf.RootStorage.AddStorage(tooManyCharactersEntryName);
                Assert.Fail();
            }
            catch (Exception ex)
            {
                Assert.IsTrue(ex is CfException);
            }

            try
            {
                // Try Stream entry name with too many characters.
                cf.RootStorage.AddStream(tooManyCharactersEntryName);
                Assert.Fail();
            }
            catch (Exception ex)
            {
                Assert.IsTrue(ex is CfException);
            }

            cf.Save("EntryNameLength");
            cf.Close();
        }

        [TestMethod]
        public void Test_DELETE_WITHOUT_COMPRESSION()
        {
            var filename = "MultipleStorage3.cfs";

            var srcFile = new FileInfo(filename);

            var cf = new CompoundFile(filename);

            var st = cf.RootStorage.GetStorage("MyStorage");
            st = st.GetStorage("AnotherStorage");

            Assert.IsNotNull(st);

            st.Delete("Another2Stream"); //17Kb

            //cf.CompressFreeSpace();
            cf.Save("MultipleStorage_Deleted_Compress.cfs");

            cf.Close();
            var dstFile = new FileInfo("MultipleStorage_Deleted_Compress.cfs");

            Assert.IsFalse(srcFile.Length > dstFile.Length);

        }

        [TestMethod]
        public void Test_WRITE_AND_READ_CFS_VERSION_4()
        {
            var filename = "WRITE_AND_READ_CFS_V4.cfs";

            var cf = new CompoundFile(CfsVersion.Ver4, CfsConfiguration.EraseFreeSectors | CfsConfiguration.SectorRecycle);

            var st = cf.RootStorage.AddStorage("MyStorage");
            var sm = st.AddStream("MyStream");
            var b = new byte[220];
            sm.SetData(b);

            cf.Save(filename);
            cf.Close();

            var cf2 = new CompoundFile(filename);
            var st2 = cf2.RootStorage.GetStorage("MyStorage");
            var sm2 = st2.GetStream("MyStream");

            Assert.IsNotNull(sm2);
            Assert.IsTrue(sm2.Size == 220);

            cf2.Close();
        }

        [TestMethod]
        public void Test_WRITE_READ_CFS_VERSION_4_STREAM()
        {
            var filename = "WRITE_COMMIT_READ_CFS_V4.cfs";

            var cf = new CompoundFile(CfsVersion.Ver4, CfsConfiguration.SectorRecycle | CfsConfiguration.EraseFreeSectors);

            var st = cf.RootStorage.AddStorage("MyStorage");
            var sm = st.AddStream("MyStream");
            var b = Helpers.GetBuffer(227);
            sm.SetData(b);

            cf.Save(filename);
            cf.Close();

            var cf2 = new CompoundFile(filename);
            var st2 = cf2.RootStorage.GetStorage("MyStorage");
            var sm2 = st2.GetStream("MyStream");

            Assert.IsNotNull(sm2);
            Assert.IsTrue(sm2.Size == b.Length);

            cf2.Close();
        }

        [TestMethod]
        public void Test_OPEN_FROM_STREAM()
        {
            const string filename = "reportREAD.xls";

            using (var fs = new FileStream(filename, FileMode.Open))
            {
                using (var cf = new CompoundFile(fs))
                {
                    var foundStream = cf.RootStorage.GetStream("Workbook");
                    var temp = foundStream.GetData();
                    Assert.IsNotNull(temp);
                    cf.Close();
                }
            }


        }

        [TestMethod]
        public void Test_MULTIPLE_SAVE()
        {
            var file = new CompoundFile();

            file.Save("test.mdf");

            var meta = file.
                RootStorage.
                AddStream("meta");

            meta.Append(BitConverter.GetBytes(DateTime.Now.ToBinary()));
            meta.Append(BitConverter.GetBytes(DateTime.Now.ToBinary()));

            file.Save("test.mdf");
        }

        [TestMethod]
        public void Test_OPEN_COMPOUND_BUG_FIX_133()
        {
            var f = new CompoundFile("testbad.ole");
            var cfs = f.RootStorage.GetStream("\x01Ole10Native");
            var data = cfs.GetData();
            Assert.IsTrue(data.Length == 18140);
        }

        [TestMethod]
        public void Test_COMPARE_DIR_ENTRY_NAME_BUG_FIX_ID_3487353()
        {
            var f = new CompoundFile("report_name_fix.xls", CfsUpdateMode.Update, CfsConfiguration.SectorRecycle | CfsConfiguration.EraseFreeSectors);
            var cfs = f.RootStorage.AddStream("Poorbook");
            cfs.Append(Helpers.GetBuffer(20));
            f.Commit();
            f.Close();

            f = new CompoundFile("report_name_fix.xls", CfsUpdateMode.Update, CfsConfiguration.SectorRecycle | CfsConfiguration.EraseFreeSectors);
            cfs = f.RootStorage.GetStream("Workbook");
            Assert.IsTrue(cfs.Name == "Workbook");
            f.RootStorage.Delete("PoorBook");
            f.Commit();
            f.Close();

        }

        [TestMethod]
        public void Test_GET_COMPOUND_VERSION()
        {
            var f = new CompoundFile("report_name_fix.xls");
            var ver = f.Version;

            Assert.IsTrue(ver == CfsVersion.Ver3);

            f.Close();
        }

        [TestMethod]
        public void Test_FUNCTIONAL_BEHAVIOUR()
        {
            //System.Diagnostics.Trace.Listeners.Add(new ConsoleTraceListener());

            const int nFactor = 1;

            var bA = Helpers.GetBuffer(20 * 1024 * nFactor, 0x0A);
            var bB = Helpers.GetBuffer(5 * 1024, 0x0B);
            var bC = Helpers.GetBuffer(5 * 1024, 0x0C);
            var bD = Helpers.GetBuffer(5 * 1024, 0x0D);
            var bE = Helpers.GetBuffer(8 * 1024 * nFactor + 1, 0x1A);
            var bF = Helpers.GetBuffer(16 * 1024 * nFactor, 0x1B);
            var bG = Helpers.GetBuffer(14 * 1024 * nFactor, 0x1C);
            var bH = Helpers.GetBuffer(12 * 1024 * nFactor, 0x1D);
            var bE2 = Helpers.GetBuffer(8 * 1024 * nFactor, 0x2A);
            var bMini = Helpers.GetBuffer(1027, 0xEE);

            var sw = new Stopwatch();
            sw.Start();

            //############

            // Phase 1
            var cf = new CompoundFile(CfsVersion.Ver3, CfsConfiguration.SectorRecycle);
            cf.RootStorage.AddStream("A").SetData(bA);
            cf.Save("OneStream.cfs");
            cf.Close();

            // Test Phase 1
            var cfTest = new CompoundFile("OneStream.cfs");
            var testSt = cfTest.RootStorage.GetStream("A");

            Assert.IsNotNull(testSt);
            Assert.IsTrue(testSt.Size == bA.Length);
            Assert.IsTrue(Helpers.CompareBuffer(bA, testSt.GetData()));

            cfTest.Close();

            //###########

            //Phase 2
            cf = new CompoundFile("OneStream.cfs", CfsUpdateMode.ReadOnly, CfsConfiguration.SectorRecycle);

            cf.RootStorage.AddStream("B").SetData(bB);
            cf.RootStorage.AddStream("C").SetData(bC);
            cf.RootStorage.AddStream("D").SetData(bD);
            cf.RootStorage.AddStream("E").SetData(bE);
            cf.RootStorage.AddStream("F").SetData(bF);
            cf.RootStorage.AddStream("G").SetData(bG);
            cf.RootStorage.AddStream("H").SetData(bH);

            cf.Save("8_Streams.cfs");
            cf.Close();

            // Test Phase 2


            cfTest = new CompoundFile("8_Streams.cfs");

            testSt = cfTest.RootStorage.GetStream("B");
            Assert.IsNotNull(testSt);
            Assert.IsTrue(testSt.Size == bB.Length);
            Assert.IsTrue(Helpers.CompareBuffer(bB, testSt.GetData()));

            testSt = cfTest.RootStorage.GetStream("C");
            Assert.IsNotNull(testSt);
            Assert.IsTrue(testSt.Size == bC.Length);
            Assert.IsTrue(Helpers.CompareBuffer(bC, testSt.GetData()));

            testSt = cfTest.RootStorage.GetStream("D");
            Assert.IsNotNull(testSt);
            Assert.IsTrue(testSt.Size == bD.Length);
            Assert.IsTrue(Helpers.CompareBuffer(bD, testSt.GetData()));

            testSt = cfTest.RootStorage.GetStream("E");
            Assert.IsNotNull(testSt);
            Assert.IsTrue(testSt.Size == bE.Length);
            Assert.IsTrue(Helpers.CompareBuffer(bE, testSt.GetData()));

            testSt = cfTest.RootStorage.GetStream("F");
            Assert.IsNotNull(testSt);
            Assert.IsTrue(testSt.Size == bF.Length);
            Assert.IsTrue(Helpers.CompareBuffer(bF, testSt.GetData()));

            testSt = cfTest.RootStorage.GetStream("G");
            Assert.IsNotNull(testSt);
            Assert.IsTrue(testSt.Size == bG.Length);
            Assert.IsTrue(Helpers.CompareBuffer(bG, testSt.GetData()));

            testSt = cfTest.RootStorage.GetStream("H");
            Assert.IsNotNull(testSt);
            Assert.IsTrue(testSt.Size == bH.Length);
            Assert.IsTrue(Helpers.CompareBuffer(bH, testSt.GetData()));

            cfTest.Close();


            File.Copy("8_Streams.cfs", "6_Streams.cfs", true);
            File.Delete("8_Streams.cfs");

            //###########
            // 
#if !NETCOREAPP2_0
            Trace.Listeners.Add(new ConsoleTraceListener());
#endif
            // Phase 3
            cf = new CompoundFile("6_Streams.cfs", CfsUpdateMode.Update, CfsConfiguration.SectorRecycle | CfsConfiguration.EraseFreeSectors);
            cf.RootStorage.Delete("D");
            cf.RootStorage.Delete("G");
            cf.Commit();

            cf.Close();

            //Test Phase 3


            cfTest = new CompoundFile("6_Streams.cfs");


            var catched = false;

            try
            {
                testSt = cfTest.RootStorage.GetStream("D");

            }
            catch (Exception ex)
            {
                if (ex is CfItemNotFound)
                    catched = true;
            }

            Assert.IsTrue(catched);

            catched = false;

            try
            {
                testSt = cfTest.RootStorage.GetStream("G");
            }
            catch (Exception ex)
            {
                if (ex is CfItemNotFound)
                    catched = true;
            }

            Assert.IsTrue(catched);

            cfTest.Close();

            //##########

            // Phase 4

            File.Copy("6_Streams.cfs", "6_Streams_Shrinked.cfs", true);
            CompoundFile.ShrinkCompoundFile("6_Streams_Shrinked.cfs");

            // Test Phase 4

            Assert.IsTrue(new FileInfo("6_Streams_Shrinked.cfs").Length < new FileInfo("6_Streams.cfs").Length);

            cfTest = new CompoundFile("6_Streams_Shrinked.cfs");
            var va = delegate (CfItem item)
            {
                if (!item.IsStream) return;
                var ia = item as CfStream;
                Assert.IsNotNull(ia);
                Assert.IsTrue(ia.Size > 0);
                var d = ia.GetData();
                Assert.IsNotNull(d);
                Assert.IsTrue(d.Length > 0);
                Assert.IsTrue(d.Length == ia.Size);
            };

            cfTest.RootStorage.VisitEntries(va, true);
            cfTest.Close();

            //##########

            //Phase 5

            cf = new CompoundFile("6_Streams_Shrinked.cfs", CfsUpdateMode.Update, CfsConfiguration.SectorRecycle);
            cf.RootStorage.AddStream("ZZZ").SetData(bF);
            cf.RootStorage.GetStream("E").Append(bE2);
            cf.Commit();
            cf.Close();


            cf = new CompoundFile("6_Streams_Shrinked.cfs", CfsUpdateMode.Update, CfsConfiguration.SectorRecycle);
            cf.RootStorage.Clsid = new Guid("EEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE");
            cf.Commit();
            cf.Close();

            cf = new CompoundFile("6_Streams_Shrinked.cfs", CfsUpdateMode.Update, CfsConfiguration.SectorRecycle);
            cf.RootStorage.AddStorage("MyStorage").AddStream("ZIP").Append(bE);
            cf.Commit();
            cf.Close();

            cf = new CompoundFile("6_Streams_Shrinked.cfs", CfsUpdateMode.Update, CfsConfiguration.SectorRecycle);
            cf.RootStorage.AddStorage("AnotherStorage").AddStream("ANS").Append(bE);
            cf.RootStorage.Delete("MyStorage");


            cf.Commit();
            cf.Close();

            //Test Phase 5

            //#####

            //Phase 6

            cf = new CompoundFile("6_Streams_Shrinked.cfs", CfsUpdateMode.Update, CfsConfiguration.SectorRecycle);
            var root = cf.RootStorage;

            root.AddStorage("MiniStorage").AddStream("miniSt").Append(bMini);

            cf.RootStorage.GetStorage("MiniStorage").AddStream("miniSt2").Append(bMini);
            cf.Commit();
            cf.Close();

            cf = new CompoundFile("6_Streams_Shrinked.cfs", CfsUpdateMode.Update, CfsConfiguration.SectorRecycle);
            //cf.RootStorage.GetStorage("MiniStorage").Delete("miniSt");


            cf.RootStorage.GetStorage("MiniStorage").GetStream("miniSt2").Append(bE);

            cf.Commit();
            cf.Close();

            //Test Phase 6

            cfTest = new CompoundFile("6_Streams_Shrinked.cfs");
            var d2 = cfTest.RootStorage.GetStorage("MiniStorage").GetStream("miniSt2").GetData();
            Assert.IsTrue(d2.Length == (bE.Length + bMini.Length));

            var cnt = 1;
            var buf = new byte[cnt];
            cnt = cfTest.RootStorage.GetStorage("MiniStorage").GetStream("miniSt2").Read(buf, bMini.Length, cnt);

            Assert.IsTrue(cnt == 1);
            Assert.IsTrue(buf[0] == 0x1A);

            cnt = 1;
            cnt = cfTest.RootStorage.GetStorage("MiniStorage").GetStream("miniSt2").Read(buf, bMini.Length - 1, cnt);
            Assert.IsTrue(cnt == 1);
            Assert.IsTrue(buf[0] == 0xEE);

            try
            {
                cfTest.RootStorage.GetStorage("MiniStorage").GetStream("miniSt");
            }
            catch (Exception ex)
            {
                Assert.IsTrue(ex is CfItemNotFound);
            }

            cfTest.Close();

            //##############

            //Phase 7

            cf = new CompoundFile("6_Streams_Shrinked.cfs", CfsUpdateMode.Update, CfsConfiguration.SectorRecycle);

            cf.RootStorage.GetStorage("MiniStorage").GetStream("miniSt2").SetData(bA);
            cf.Commit();
            cf.Close();


            //Test Phase 7

            cf = new CompoundFile("6_Streams_Shrinked.cfs", CfsUpdateMode.Update, CfsConfiguration.SectorRecycle);
            d2 = cf.RootStorage.GetStorage("MiniStorage").GetStream("miniSt2").GetData();
            Assert.IsNotNull(d2);
            Assert.IsTrue(d2.Length == bA.Length);

            cf.Close();

            //##############

            cf = new CompoundFile("6_Streams_Shrinked.cfs", CfsUpdateMode.ReadOnly, CfsConfiguration.SectorRecycle);

            var myStream = cf.RootStorage.GetStream("C");
            var data = myStream.GetData();
            Console.WriteLine(data[0] + " : " + data[data.Length - 1]);

            myStream = cf.RootStorage.GetStream("B");
            data = myStream.GetData();
            Console.WriteLine(data[0] + " : " + data[data.Length - 1]);

            cf.Close();

            sw.Stop();
            Console.WriteLine(sw.ElapsedMilliseconds);

        }

        [TestMethod]
        public void Test_RETRIVE_ALL_NAMED_ENTRIES()
        {
            var f = new CompoundFile("MultipleStorage4.cfs");
            var result = f.GetAllNamedEntries("MyStream");

            Assert.IsTrue(result.Count == 3);
        }


        [TestMethod]
        public void Test_CORRUPTED_CYCLIC_FAT_CHECK()
        {
            CompoundFile f = null;
            try
            {
                f = new CompoundFile("CyclicFAT.cfs");

            }
            catch (Exception ex)
            {
                Assert.IsTrue(ex is CfCorruptedFileException);
            }
            finally
            {
                if (f != null)
                    f.Close();
            }
        }

        [TestMethod]
        public void Test_DIFAT_CHECK()
        {
            CompoundFile f = null;
            try
            {
                f = new CompoundFile();
                var st = f.RootStorage.AddStream("LargeStream");
                st.Append(Helpers.GetBuffer(20000000, 0x0A));       //Forcing creation of two DIFAT sectors
                var b1 = Helpers.GetBuffer(3, 0x0B);
                st.Append(b1);                                      //Forcing creation of two DIFAT sectors

                f.Save("$OpenMcdf$LargeFile.cfs");

                f.Close();

                var cnt = 3;
                f = new CompoundFile("$OpenMcdf$LargeFile.cfs");

                var b2 = new byte[cnt];
                cnt = f.RootStorage.GetStream("LargeStream").Read(b2, 20000000, cnt);
                f.Close();
                Assert.IsTrue(Helpers.CompareBuffer(b1, b2));
            }
            finally
            {
                if (f != null)
                    f.Close();

                if (File.Exists("$OpenMcdf$LargeFile.cfs"))
                    File.Delete("$OpenMcdf$LargeFile.cfs");
            }

        }

        [TestMethod]
        public void Test_ADD_LARGE_NUMBER_OF_ITEMS()
        {
            var itemNumber = 10000;

            CompoundFile f = null;
            var buffer = Helpers.GetBuffer(10, 0x0A);
            try
            {
                f = new CompoundFile();

                for (var i = 0; i < itemNumber; i++)
                {
                    var st = f.RootStorage.AddStream("Stream" + i);
                    st.Append(buffer);
                }


                if (File.Exists("$ItemsLargeNumber.cfs"))
                    File.Delete("$ItemsLargeNumber.cfs");

                f.Save("$ItemsLargeNumber.cfs");
                f.Close();

                f = new CompoundFile("$ItemsLargeNumber.cfs");
                var cfs = f.RootStorage.GetStream("Stream" + (itemNumber / 2));

                Assert.IsTrue(cfs != null, "Item is null");
                Assert.IsTrue(Helpers.CompareBuffer(cfs.GetData(), buffer), "Items are different");
                f.Close();
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.Message);
            }
        }

        [TestMethod]
        public void Test_FIX_BUG_16_CORRUPTED_AFTER_RESIZE()
        {

            const string filePath = @"BUG_16_.xls";

            var cf = new CompoundFile(filePath);

            var dirStream = cf.RootStorage.GetStorage("_VBA_PROJECT_CUR").GetStorage("VBA").GetStream("dir");

            var currentData = dirStream.GetData();

            Array.Resize(ref currentData, currentData.Length - 50);

            dirStream.SetData(currentData);

            cf.Save(filePath + ".edited");
            cf.Close();
        }


        [TestMethod]
        public void Test_FIX_BUG_17_CORRUPTED_PPT_FILE()
        {

            const string filePath = @"2_MB-W.ppt";

            using (var file = new CompoundFile(filePath))
            {
                //CFStorage dataSpaceInfo = file.RootStorage.GetStorage("\u0006DataSpaces").GetStorage("DataSpaceInfo");
                var dsiItem = file.GetAllNamedEntries("DataSpaceInfo").FirstOrDefault();
            }
        }

        [TestMethod]
        public void Test_FIX_BUG_24_CORRUPTED_THUMBS_DB_FILE()
        {
            try
            {
                using (var cf = new CompoundFile("_thumbs_bug_24.db"))
                {
                    cf.RootStorage.VisitEntries(item => Console.WriteLine(item.Name), recursive: false);
                }
            }
            catch (Exception exc)
            {
                Assert.IsInstanceOfType(exc, typeof(CfCorruptedFileException));
            }

            using (var cf = new CompoundFile("report.xls"))
            {
                cf.RootStorage.VisitEntries(item => Console.WriteLine(item.Name), recursive: false);
            }

        }


        [TestMethod]
        public void Test_FIX_BUG_28_CompoundFile_Delete_ChildElementMaintainsFiles()
        {
            using (var compoundFile = new CompoundFile())
            {
                var storage1 = compoundFile.RootStorage.AddStorage("A");
                var storage2 = compoundFile.RootStorage.AddStorage("B");
                var storage3 = compoundFile.RootStorage.AddStorage("C");
                storage1.AddStream("A.1");
                compoundFile.RootStorage.Delete("B");
                storage1 = compoundFile.RootStorage.GetStorage("A");
                storage1.GetStream("A.1");
            }
        }

        [TestMethod]
        public void Test_CORRUPTEDDOC_BUG36_SHOULD_THROW_CORRUPTED_FILE_EXCEPTION()
        {
            FileStream fs = null;
            try
            {
                fs = new FileStream("CorruptedDoc_bug36.doc", FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
                var file = new CompoundFile(fs, CfsUpdateMode.ReadOnly, CfsConfiguration.LeaveOpen);

            }
            catch (Exception ex)
            {
                Assert.IsTrue(fs.CanRead && fs.CanSeek && fs.CanWrite);
            }


        }

        [TestMethod]
        public void Test_ISSUE_2_WRONG_CUTOFF_SIZE()
        {
            FileStream fs = null;
            try
            {
                if (File.Exists("TEST_ISSUE_2"))
                {
                    File.Delete("TEST_ISSUE_2");
                }

                var cf = new CompoundFile(CfsVersion.Ver3, CfsConfiguration.Default);
                var s = cf.RootStorage.AddStream("miniToNormal");
                s.Append(Helpers.GetBuffer(4090, 0xAA));

                cf.Save("TEST_ISSUE_2");
                cf.Close();
                var cf2 = new CompoundFile("TEST_ISSUE_2", CfsUpdateMode.Update, CfsConfiguration.Default);
                cf2.RootStorage.GetStream("miniToNormal").Append(Helpers.GetBuffer(6, 0xBB));
                cf2.Commit();
                cf2.Close();
            }
            catch (Exception ex)
            {
                Assert.IsTrue(fs.CanRead && fs.CanSeek && fs.CanWrite);
            }
        }

        [TestMethod]
        public void Test_PR_13()
        {
            var cf = new CompoundFile("report.xls");
            var g = cf.GetGuidBySid(0);
            Assert.IsNotNull(g);
            g = cf.GetGuidForStream(3);
            Assert.IsNotNull(g);
            Assert.IsTrue(!string.IsNullOrEmpty(cf.GetNameDirEntry(2)));
            Assert.IsTrue(cf.GetNumDirectories() > 0);
        }
        //[TestMethod]
        //public void Test_CORRUPTED_CYCLIC_DIFAT_VALIDATION_CHECK()
        //{

        //    CompoundFile cf = null;
        //    try
        //    {
        //        cf = new CompoundFile("CiclycDFAT.cfs");
        //        CFStorage s = cf.RootStorage.GetStorage("MyStorage");
        //        CFStream st = s.GetStream("MyStream");
        //        Assert.IsTrue(st.Size > 0);
        //    }
        //    catch (Exception ex)
        //    {
        //        Assert.IsTrue(ex is CFCorruptedFileException);
        //    }
        //    finally
        //    {
        //        if (cf != null)
        //        {
        //            cf.Close();
        //        }
        //    }
        //}
        //[TestMethod]
        //public void Test_REM()
        //{
        //    var f = new CompoundFile();

        //    byte[] bB = Helpers.GetBuffer(5 * 1024, 0x0B); 
        //    f.RootStorage.AddStream("Test").AppendData(bB);
        //    f.Save("Astorage.cfs");
        //}

        //}


        [TestMethod]
        public void Test_COPY_ENTRIES_FROM_TO_STORAGE()
        {
            var cfDst = new CompoundFile();
            var cfSrc = new CompoundFile("MultipleStorage4.cfs");

            Copy(cfSrc.RootStorage, cfDst.RootStorage);

            cfDst.Save("MultipleStorage4Copy.cfs");

            cfDst.Close();
            cfSrc.Close();

        }

        #region Copy heper method
        /// <summary>
        /// Copies the given <paramref name="source"/> to the given <paramref name="destination"/>
        /// </summary>
        /// <param name="source"></param>
        /// <param name="destination"></param>
        public static void Copy(CfStorage source, CfStorage destination)
        {
            source.VisitEntries(action =>
            {
                if (action.IsStorage)
                {
                    var destionationStorage = destination.AddStorage(action.Name);
                    destionationStorage.Clsid = action.Clsid;
                    destionationStorage.CreationDate = action.CreationDate;
                    destionationStorage.ModifyDate = action.ModifyDate;
                    Copy(action as CfStorage, destionationStorage);
                }
                else
                {
                    var sourceStream = action as CfStream;
                    var destinationStream = destination.AddStream(action.Name);
                    if (sourceStream != null) destinationStream.SetData(sourceStream.GetData());
                }

            }, false);
        }
        #endregion
        private const int Mb = 1024 * 1024;
        [TestMethod]
        public void Test_FIX_BUG_GH_14()
        {
            var filename = "MyFile.dat";
            var storageName = "MyStorage";
            var streamName = "MyStream";
            var bufferSize = 800 * Mb;
            var iterationCount = 1;
            var streamCount = 3;

            var compoundFileInit = new CompoundFile(CfsVersion.Ver4, CfsConfiguration.Default);
            compoundFileInit.Save(filename);
            compoundFileInit.Close();

            var compoundFile = new CompoundFile(filename, CfsUpdateMode.Update, CfsConfiguration.Default);
            var st = compoundFile.RootStorage.AddStorage(storageName);
            byte b = 0X0A;

            for (var streamId = 0; streamId < streamCount; ++streamId)
            {
                var sm = st.AddStream(streamName + streamId);
                for (var iteration = 0; iteration < iterationCount; ++iteration)
                {
                    sm.Append(Helpers.GetBuffer(bufferSize, b));
                    compoundFile.Commit();
                }

                b++;
            }
            compoundFile.Close();

            compoundFile = new CompoundFile(filename, CfsUpdateMode.ReadOnly, CfsConfiguration.Default);
            var testBuffer = new byte[100];
            byte t = 0x0A;

            for (var streamId = 0; streamId < streamCount; ++streamId)
            {
                compoundFile.RootStorage.GetStorage(storageName).GetStream(streamName + streamId).Read(testBuffer, bufferSize / 2, 100);
                Assert.IsTrue(testBuffer.All(g => g == t));
                compoundFile.RootStorage.GetStorage(storageName).GetStream(streamName + streamId).Read(testBuffer, bufferSize - 101, 100);
                Assert.IsTrue(testBuffer.All(g => g == t));
                compoundFile.RootStorage.GetStorage(storageName).GetStream(streamName + streamId).Read(testBuffer, 0, 100);
                Assert.IsTrue(testBuffer.All(g => g == t));
                t++;
            }

            compoundFile.Close();
        }

        [TestMethod]
        public void Test_FIX_BUG_GH_15()
        {
            var filename = "MyFile.dat";
            var storageName = "MyStorage";
            var streamName = "MyStream";
            var bufferSize = 800 * Mb;
            var iterationCount = 1;
            var streamCount = 1;

            var compoundFile = new CompoundFile(CfsVersion.Ver4, CfsConfiguration.Default);
            var st = compoundFile.RootStorage.AddStorage(storageName);

            for (var streamId = 0; streamId < streamCount; ++streamId)
            {
                var sm = st.AddStream(streamName + streamId);
                for (var iteration = 0; iteration < iterationCount; ++iteration)
                {
                    var b = (byte)(0x0A + iteration);
                    sm.Append(Helpers.GetBuffer(bufferSize, b));
                }
            }
            compoundFile.Save(filename);
            compoundFile.Close();

            var readBuffer = new byte[15];
            compoundFile = new CompoundFile(filename);

            byte c = 0x0A;
            for (var i = 0; i < iterationCount; i++)
            {
                compoundFile.RootStorage.GetStorage(storageName).GetStream(streamName + 0).Read(readBuffer, (bufferSize + ((long)bufferSize * i)) - 15, 15);
                Assert.IsTrue(readBuffer.All(by => by == c));
                c++;
            }

            compoundFile.Close();
        }

        [TestMethod]
        public void Test_PR_GH_18()
        {
            try
            {
                var f = new CompoundFile("MultipleStorage4.cfs", CfsUpdateMode.Update, CfsConfiguration.Default);
                var st = f.RootStorage.GetStorage("MyStorage").GetStorage("AnotherStorage").GetStream("MyStream");
                st.Write(Helpers.GetBuffer(100, 0x02), 100);
                f.Commit(true);
                Assert.IsTrue(st.GetData().Count() == 31220);
                f.Close();
            }
            catch (Exception ex)
            {
                Assert.Fail("Release Memory flag caused error");
            }
        }

        [TestMethod]
        public void Test_FIX_GH_38()
        {
            CompoundFile f = null;
            try
            {
                f = new CompoundFile("empty_directory_chain.doc", CfsUpdateMode.Update, CfsConfiguration.Default);

                f.Close();
            }
            catch (Exception ex)
            {
                Assert.IsInstanceOfType(ex, typeof(CfCorruptedFileException));
                if (f != null)
                    f.Close();
            }
        }

        [TestMethod]
        public void Test_FIX_GH_38_B()
        {
            CompoundFile f = null;
            try
            {
                f = new CompoundFile("no_sectors.doc", CfsUpdateMode.Update, CfsConfiguration.Default);

                f.Close();
            }
            catch (Exception ex)
            {
                Assert.IsInstanceOfType(ex, typeof(CfException));
                if (f != null)
                    f.Close();
            }
        }

        [TestMethod]
        public void Test_FIX_GH_50()
        {
            try
            {
                var f = new CompoundFile("64-67.numberOfMiniFATSectors.docx", CfsUpdateMode.Update, CfsConfiguration.Default);
            }
            catch (Exception e)
            {
                Assert.IsTrue(e is CfFileFormatException);
            }
        }

        [TestMethod]
        public void Test_FIX_GH_83()
        {
            try
            {
                var bigDataBuffer = Helpers.GetBuffer(1024 * 1024 * 260);

                using (var fTest = new FileStream("BigFile.data", FileMode.Create))
                {

                    fTest.Write(bigDataBuffer, 0, 1024 * 1024 * 260);
                    fTest.Flush();
                    fTest.Close();

                    var f = new CompoundFile();
                    var cfStream = f.RootStorage.AddStream("NewStream");
                    using (var fs = new FileStream("BigFile.data", FileMode.Open))
                    {
                        cfStream.CopyFrom(fs);
                    }
                    f.Save("BigFile.cfs");
                    f.Close();

                }
            }
            catch (Exception e)
            {
                Assert.Fail();
            }
            finally
            {
                if (File.Exists("BigFile.data"))
                    File.Delete("BigFile.data");

                if (File.Exists("BigFile.cfs"))
                    File.Delete("BigFile.cfs");
            }
        }

        [TestMethod]
        [ExpectedException(typeof(CfCorruptedFileException))]
        public void Test_CorruptedSectorChain_Doc()
        {
            var f = new CompoundFile("corrupted-sector-chain.doc");

            f.Close();
        }

        [TestMethod]
        [ExpectedException(typeof(CfCorruptedFileException))]
        public void Test_CorruptedSectorChain_Cfs()
        {
            var f = new CompoundFile("corrupted-sector-chain.cfs");

            f.Close();
        }

        [TestMethod]
        public void Test_WRONG_CORRUPTED_EXCEPTION()
        {
            var cf = new CompoundFile();

            for (var i = 0; i < 100; i++)
            {
                cf.RootStorage.AddStream("Stream" + i).SetData(Helpers.GetBuffer(100000, 0xAA));
            }

            cf.RootStorage.AddStream("BigStream").SetData(Helpers.GetBuffer(5250000, 0xAA));

            using (var stream = new MemoryStream())
            {
                cf.Save(stream);
            }

            cf.Close();
        }

        [TestMethod]
        [ExpectedException(typeof(CfCorruptedFileException))]
        public void Test_CorruptedSectorChain_Doc2()
        {
            var f = new CompoundFile("corrupted-sector-chain-2.doc");

            f.Close();
        }

        //[TestMethod]
        //public void Test_CORRUPTED_CYCLIC_DIFAT_VALIDATION_CHECK()
        //{

        //    CompoundFile cf = null;
        //    try
        //    {
        //        cf = new CompoundFile("CiclycDFAT.cfs");
        //        CFStorage s = cf.RootStorage.GetStorage("MyStorage");
        //        CFStream st = s.GetStream("MyStream");
        //        Assert.IsTrue(st.Size > 0);
        //    }
        //    catch (Exception ex)
        //    {
        //        Assert.IsTrue(ex is CFCorruptedFileException);
        //    }
        //    finally
        //    {
        //        if (cf != null)
        //        {
        //            cf.Close();
        //        }
        //    }
        //}
        //[TestMethod]
        //public void Test_REM()
        //{
        //    var f = new CompoundFile();

        //    byte[] bB = Helpers.GetBuffer(5 * 1024, 0x0B); 
        //    f.RootStorage.AddStream("Test").AppendData(bB);
        //    f.Save("Astorage.cfs");
        //}

        [TestMethod]
        public void Test_FIX_BUG_90_CompoundFile_Delete_Storages()
        {
            using (var compoundFile = new CompoundFile())
            {
                var root = compoundFile.RootStorage;
                var storageNames = new HashSet<string>();

                // add 99 storages to root
                for (var i = 1; i <= 99; i++)
                {
                    var name = "Storage " + i;
                    root.AddStorage(name);
                    storageNames.Add(name);
                }

                // remove storages until tree becomes unbalanced and its Root changes
                var rootChild = root.DirEntry.Child;
                var newChild = rootChild;
                var j = 1;
                while (newChild == rootChild && j <= 99)
                {
                    var name = "Storage " + j;
                    root.Delete(name);
                    storageNames.Remove(name);
                    if (root.Children.Root != null)
                        newChild = ((DirectoryEntry)(root.Children.Root).Value)
                            .Sid; // stop as soon as root.Children has a new Root
                    j++;
                }

                // check if all remaining storages are still there
                foreach (var storageName in storageNames)
                {
                    Assert.IsTrue(root.TryGetStorage(storageName, out var storage)); // <- no problem here
                }

                // write CompundFile into MemoryStream... 
                using (var memStream = new MemoryStream())
                {
                    compoundFile.Save(memStream);

                    // ....and read new CompundFile from that stream
                    using (var newCf = new CompoundFile(memStream))
                    {
                        // check if all storages can be found in to copied CompundFile
                        foreach (var storageName in storageNames)
                        {
                            Assert.IsTrue(newCf.RootStorage.TryGetStorage(storageName, out var storage)); //<- we will see some missing storages here
                        }
                    }
                }
            }
        }
    }
}
