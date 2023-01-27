using System;
using System.IO;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Order;

namespace OpenMcdf.Benchmark
{
    [CoreJob]
    [CsvExporter]
    [HtmlExporter]
    [MarkdownExporter]
    //[DryCoreJob] // I always forget this attribute, so please leave it commented out 
    [MemoryDiagnoser]
    [Orderer(SummaryOrderPolicy.FastestToSlowest)]
    public class InMemory : IDisposable
    {
        private const int Kb = 1024;
        private const int Mb = Kb * Kb;
        private const string StorageName = "MyStorage";
        private const string StreamName = "MyStream";

        private byte[] _readBuffer;

        private Stream _stream;

        [Params(Kb / 2, Kb, 4 * Kb, 128 * Kb, 256 * Kb, 512 * Kb, Kb * Kb)]
        public int BufferSize { get; set; }

        [Params(Mb /*, 8 * Mb, 64 * Mb, 128 * Mb*/)]
        public int TotalStreamSize { get; set; }

        public void Dispose()
        {
            _stream?.Dispose();
        }

        [GlobalSetup]
        public void GlobalSetup()
        {
            _stream = new MemoryStream();
            _readBuffer = new byte[BufferSize];
            CreateFile(1);
        }

        [GlobalCleanup]
        public void GlobalCleanup()
        {
            _stream.Dispose();
            _stream = null;
            _readBuffer = null;
        }


        [Benchmark]
        public void Test()
        {
            //
            _stream.Seek(0L, SeekOrigin.Begin);
            //
            var compoundFile = new CompoundFile(_stream);
            var cfStream = compoundFile.RootStorage
                .GetStorage(StorageName)
                .GetStream(StreamName + 0);
            var streamSize = cfStream.Size;
            var position = 0L;
            while (true)
            {
                if (position >= streamSize) break;
                var read = cfStream
                    .Read(_readBuffer, position, _readBuffer.Length);
                position += read;
                if (read <= 0) break;
            }

            //compoundFile.Close();
        }

        private void CreateFile(int streamCount)
        {
            var iterationCount = TotalStreamSize / BufferSize;

            var buffer = new byte[BufferSize];
            Array.Fill(buffer, byte.MaxValue);
            const CfsConfiguration flags = CfsConfiguration.Default | CfsConfiguration.LeaveOpen;
            using (var compoundFile = new CompoundFile(CfsVersion.Ver4, flags))
            {
                var st = compoundFile.RootStorage.AddStorage(StorageName);
                for (var streamId = 0; streamId < streamCount; ++streamId)
                {
                    var sm = st.AddStream(StreamName + streamId);

                    for (var iteration = 0; iteration < iterationCount; ++iteration) sm.Append(buffer);
                }

                compoundFile.Save(_stream);
                compoundFile.Close();
            }
        }
    }
}