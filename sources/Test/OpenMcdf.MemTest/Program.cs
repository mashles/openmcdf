using System;
using System.Diagnostics;
using System.IO;

//This project is used for profiling memory and performances of OpenMCDF .

namespace OpenMcdf.MemTest
{
    class Program
    {
        static void Main(string[] args)
        {

            //TestMultipleStreamCommit();
            TestCode();
            //StressMemory();
            //DummyFile();
            //Console.WriteLine("CLOSED");
            //Console.ReadKey();
        }

        private static void TestCode()
        {
            const int nFactor = 1000;

            var bA = GetBuffer(20 * 1024 * nFactor, 0x0A);
            var bB = GetBuffer(5 * 1024, 0x0B);
            var bC = GetBuffer(5 * 1024, 0x0C);
            var bD = GetBuffer(5 * 1024, 0x0D);
            var bE = GetBuffer(8 * 1024 * nFactor + 1, 0x1A);
            var bF = GetBuffer(16 * 1024 * nFactor, 0x1B);
            var bG = GetBuffer(14 * 1024 * nFactor, 0x1C);
            var bH = GetBuffer(12 * 1024 * nFactor, 0x1D);
            var bE2 = GetBuffer(8 * 1024 * nFactor, 0x2A);
            var bMini = GetBuffer(1027, 0xEE);

            var sw = new Stopwatch();
            sw.Start();

            var cf = new CompoundFile(CfsVersion.Ver3, CfsConfiguration.SectorRecycle);
            cf.RootStorage.AddStream("A").SetData(bA);
            cf.Save("OneStream.cfs");

            cf.Close();

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

            File.Copy("8_Streams.cfs", "6_Streams.cfs", true);

            cf = new CompoundFile("6_Streams.cfs", CfsUpdateMode.Update, CfsConfiguration.SectorRecycle|CfsConfiguration.EraseFreeSectors);
            cf.RootStorage.Delete("D");
            cf.RootStorage.Delete("G");
            cf.Commit();

            cf.Close();

            File.Copy("6_Streams.cfs", "6_Streams_Shrinked.cfs", true);

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
            cf.RootStorage.AddStorage("MyStorage").AddStream("ANS").Append(bE);
            cf.Commit();
            cf.Close();

            cf = new CompoundFile("6_Streams_Shrinked.cfs", CfsUpdateMode.Update, CfsConfiguration.SectorRecycle);
            cf.RootStorage.AddStorage("AnotherStorage").AddStream("ANS").Append(bE);
            cf.RootStorage.Delete("MyStorage");
            cf.Commit();
            cf.Close();

            CompoundFile.ShrinkCompoundFile("6_Streams_Shrinked.cfs");

            cf = new CompoundFile("6_Streams_Shrinked.cfs", CfsUpdateMode.Update, CfsConfiguration.SectorRecycle);
            cf.RootStorage.AddStorage("MiniStorage").AddStream("miniSt").Append(bMini);
            cf.RootStorage.GetStorage("MiniStorage").AddStream("miniSt2").Append(bMini);
            cf.Commit();
            cf.Close();

            cf = new CompoundFile("6_Streams_Shrinked.cfs", CfsUpdateMode.Update, CfsConfiguration.SectorRecycle);
            cf.RootStorage.GetStorage("MiniStorage").Delete("miniSt");


            cf.RootStorage.GetStorage("MiniStorage").GetStream("miniSt2").Append(bE);
            cf.Commit();
            cf.Close();

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

            Console.ReadKey();
        }

        private static void StressMemory()
        {
            const int nLoop = 20;
            const int mbSize = 10;

            var b = GetBuffer(1024 * 1024 * mbSize); //2GB buffer
            var cmp = new byte[] { 0x0, 0x1, 0x2, 0x3, 0x4, 0x5, 0x6, 0x7 };

            var cf = new CompoundFile(CfsVersion.Ver4, CfsConfiguration.Default);
            var st = cf.RootStorage.AddStream("MySuperLargeStream");
            cf.Save("LARGE.cfs");
            cf.Close();

            //Console.WriteLine("Closed save");
            //Console.ReadKey();

            cf = new CompoundFile("LARGE.cfs", CfsUpdateMode.Update, CfsConfiguration.Default);
            var cfst = cf.RootStorage.GetStream("MySuperLargeStream");

            var sw = new Stopwatch();
            sw.Start();
            for (var i = 0; i < nLoop; i++)
            {

                cfst.Append(b);
                cf.Commit(true);

                Console.WriteLine("     Updated " + i);
                //Console.ReadKey();
            }

            cfst.Append(cmp);
            cf.Commit(true);
            sw.Stop();


            cf.Close();

            Console.WriteLine(sw.Elapsed.TotalMilliseconds);
            sw.Reset();

            //Console.WriteLine(sw.Elapsed.TotalMilliseconds);

            //Console.WriteLine("Closed Transacted");
            //Console.ReadKey();

            cf = new CompoundFile("LARGE.cfs");
            var count = 8;
            sw.Reset();
            sw.Start();
            var data = new byte[count];
            count = cf.RootStorage.GetStream("MySuperLargeStream").Read(data, b.Length * (long)nLoop, count);
            sw.Stop();
            Console.Write(count);
            cf.Close();

            Console.WriteLine("Closed Final " + sw.ElapsedMilliseconds);
            Console.ReadKey();

        }

        private static void DummyFile()
        {
            Console.WriteLine("Start");
            var fs = new FileStream("myDummyFile", FileMode.Create);
            fs.Close();

            var sw = new Stopwatch();

            var b = GetBuffer(1024 * 1024 * 50); //2GB buffer

            fs = new FileStream("myDummyFile", FileMode.Open);
            sw.Start();
            for (var i = 0; i < 42; i++)
            {

                fs.Seek(b.Length * i, SeekOrigin.Begin);
                fs.Write(b, 0, b.Length);

            }

            fs.Close();
            sw.Stop();
            Console.WriteLine("Stop - " + sw.ElapsedMilliseconds);
            sw.Reset();

            Console.ReadKey();
        }

        private static void AddNodes(string depth, CfStorage cfs)
        {

            var va = delegate(CfItem target)
            {

                var temp = target.Name + (target is CfStorage ? "" : " (" + target.Size + " bytes )");

                //Stream

                Console.WriteLine(depth + temp);

                if (target is CfStorage)
                {  //Storage

                    var newDepth = depth + "    ";

                    //Recursion into the storage
                    AddNodes(newDepth, (CfStorage)target);

                }
            };

            //Visit NON-recursively (first level only)
            cfs.VisitEntries(va, false);
        }

        public static void TestMultipleStreamCommit()
        {
            var srcFilename = Directory.GetCurrentDirectory() + @"\testfile\report.xls";
            var dstFilename = Directory.GetCurrentDirectory() + @"\testfile\reportOverwriteMultiple.xls";
            //Console.WriteLine(Directory.GetCurrentDirectory());
            //Console.ReadKey(); 
            File.Copy(srcFilename, dstFilename, true);

            var cf = new CompoundFile(dstFilename, CfsUpdateMode.Update, CfsConfiguration.SectorRecycle);

            var r = new Random();

            var start = DateTime.Now;

            for (var i = 0; i < 1000; i++)
            {
                var buffer = GetBuffer(r.Next(100, 3500), 0x0A);

                if (i > 0)
                {
                    if (r.Next(0, 100) > 50)
                    {
                        cf.RootStorage.Delete("MyNewStream" + (i - 1));
                    }
                }

                var addedStream = cf.RootStorage.AddStream("MyNewStream" + i);

                addedStream.SetData(buffer);

                // Random commit, not on single addition
                if (r.Next(0, 100) > 50)
                    cf.Commit();
            }

            cf.Close();

            var sp = (DateTime.Now - start);
            Console.WriteLine(sp.TotalMilliseconds);

        }

        private static byte[] GetBuffer(int count)
        {
            var r = new Random();
            var b = new byte[count];
            r.NextBytes(b);
            return b;
        }

        private static byte[] GetBuffer(int count, byte c)
        {
            var b = new byte[count];
            for (var i = 0; i < b.Length; i++)
            {
                b[i] = c;
            }

            return b;
        }

        private static bool CompareBuffer(byte[] b, byte[] p)
        {
            if (b == null && p == null)
                throw new Exception("Null buffers");

            if (b == null && p != null) return false;
            if (b != null && p == null) return false;

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
