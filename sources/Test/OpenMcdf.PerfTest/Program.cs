using System;
using System.IO;

namespace OpenMcdf.PerfTest
{
    class Program
    {
        static readonly int _maxStreamCount = 5000;
        static readonly string _fileName = "PerfLoad.cfs";

        static void Main(string[] args)
        {
            File.Delete(_fileName);
            if (!File.Exists(_fileName))
            {
                CreateFile(_fileName);
            }

            var cf = new CompoundFile(_fileName);
            var dt = DateTime.Now;
            var s = cf.RootStorage.GetStream("Test1");
            var ts = DateTime.Now.Subtract(dt);
            Console.WriteLine(ts.TotalMilliseconds.ToString());
            Console.Read();
        }

        private static void CreateFile(string fn)
        {
            var cf = new CompoundFile();
            for (var i = 0; i < _maxStreamCount; i++)
            {
                cf.RootStorage.AddStream("Test" + i).SetData(Helpers.GetBuffer(300));
            }
            cf.Save(_fileName);
            cf.Close();
        }
    }
}
