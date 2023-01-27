/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 * 
 * The Original Code is OpenMCDF - Compound Document Format library.
 * 
 * The Initial Developer of the Original Code is Federico Blaseotto.*/

#define FLAT_WRITE // No optimization on the number of write operations

using System;
using System.Collections.Generic;
using System.IO;
using RedBlackTree;

namespace OpenMcdf
{
    internal class CfItemComparer : IComparer<CfItem>
    {
        public int Compare(CfItem x, CfItem y)
        {
            // X CompareTo Y : X > Y --> 1 ; X < Y  --> -1
            return (x.DirEntry.CompareTo(y.DirEntry));

            //Compare X < Y --> -1
        }
    }

    /// <summary>
    /// Configuration parameters for the compund files.
    /// They can be OR-combined to configure 
    /// <see cref="T:OpenMcdf.CompoundFile">Compound file</see> behaviour.
    /// All flags are NOT set by Default.
    /// </summary>
    [Flags]
    public enum CfsConfiguration
    {
        /// <summary>
        /// Sector Recycling turn off, 
        /// free sectors erasing off, 
        /// format validation exception raised
        /// </summary>
        Default = 1,

        /// <summary>
        /// Sector recycling reduces data writing performances 
        /// but avoids space wasting in scenarios with frequently
        /// data manipulation of the same streams.
        /// </summary>
        SectorRecycle = 2,

        /// <summary>
        /// Free sectors are erased to avoid information leakage
        /// </summary>
        EraseFreeSectors = 4,

        /// <summary>
        /// No exception is raised when a validation error occurs.
        /// This can possibly lead to a security issue but gives 
        /// a chance to corrupted files to load.
        /// </summary>
        NoValidationException = 8,

        /// <summary>
        /// If this flag is set true,
        /// backing stream is kept open after CompoundFile disposal
        /// </summary>
        LeaveOpen = 16,
    }

    /// <summary>
    /// Binary File Format Version. Sector size  is 512 byte for version 3,
    /// 4096 for version 4
    /// </summary>
    public enum CfsVersion
    {
        /// <summary>
        /// Compound file version 3 - The default and most common version available. Sector size 512 bytes, 2GB max file size.
        /// </summary>
        Ver3 = 3,
        /// <summary>
        /// Compound file version 4 - Sector size is 4096 bytes. Using this version could bring some compatibility problem with existing applications.
        /// </summary>
        Ver4 = 4
    }

    /// <summary>
    /// Update mode of the compound file.
    /// Default is ReadOnly.
    /// </summary>
    public enum CfsUpdateMode
    {
        /// <summary>
        /// ReadOnly update mode prevents overwriting
        /// of the opened file. 
        /// Data changes are allowed but they have to be 
        /// persisted on a different file when required 
        /// using <see cref="M:OpenMcdf.CompoundFile.Save">method</see>
        /// </summary>
        ReadOnly,

        /// <summary>
        /// Update mode allows subsequent data changing operations
        /// to be persisted directly on the opened file or stream
        /// using the <see cref="M:OpenMcdf.CompoundFile.Commit">Commit</see>
        /// method when required. Warning: this option may cause existing data loss if misused.
        /// </summary>
        Update
    }

    /// <summary>
    /// Standard Microsoft&#169; Compound File implementation.
    /// It is also known as OLE/COM structured storage 
    /// and contains a hierarchy of storage and stream objects providing
    /// efficent storage of multiple kinds of documents in a single file.
    /// Version 3 and 4 of specifications are supported.
    /// </summary>
    public sealed class CompoundFile : IDisposable
    {
        /// <summary>
        /// Get the configuration parameters of the CompoundFile object.
        /// </summary>
        private CfsConfiguration Configuration { get; } = CfsConfiguration.Default;

        /// <summary>
        /// Returns the size of standard sectors switching on CFS version (3 or 4)
        /// </summary>
        /// <returns>Standard sector size</returns>
        private int GetSectorSize()
        {
            return 2 << (_header.SectorShift - 1);
        }

        /// <summary>
        /// Number of DIFAT entries in the header
        /// </summary>
        private const int HeaderDifatEntriesCount = 109;

        /// <summary>
        /// Number of FAT entries in a DIFAT Sector
        /// </summary>
        private readonly int _difatSectorFatEntriesCount;

        /// <summary>
        /// Sectors ID entries in a FAT Sector
        /// </summary>
        private readonly int _fatSectorEntriesCount;

        /// <summary>
        /// Sector ID Size (int)
        /// </summary>
        private const int SizeOfSid = 4;

        /// <summary>
        /// Flag for sector recycling.
        /// </summary>
        private readonly bool _sectorRecycle;

        /// <summary>
        /// Flag for unallocated sector zeroing out.
        /// </summary>
        private readonly bool _eraseFreeSectors;

        /// <summary>
        /// Initial capacity of the flushing queue used
        /// to optimize commit writing operations
        /// </summary>
        private const int FlushingQueueSize = 6000;

        /// <summary>
        /// Maximum size of the flushing buffer used
        /// to optimize commit writing operations
        /// </summary>
        private const int FlushingBufferMaxSize = 1024 * 1024 * 16;


        private SectorCollection _sectors = new SectorCollection();


        /// <summary>
        /// CompoundFile header
        /// </summary>
        private Header _header;

        /// <summary>
        /// Compound underlying stream. Null when new CF has been created.
        /// </summary>
        internal Stream SourceStream;


        /// <summary>
        /// Create a blank, version 3 compound file.
        /// Sector recycle is turned off to achieve the best reading/writing 
        /// performance in most common scenarios.
        /// </summary>
        /// <example>
        /// <code>
        /// 
        ///     byte[] b = new byte[10000];
        ///     for (int i = 0; i &lt; 10000; i++)
        ///     {
        ///         b[i % 120] = (byte)i;
        ///     }
        ///
        ///     CompoundFile cf = new CompoundFile();
        ///     CFStream myStream = cf.RootStorage.AddStream("MyStream");
        ///
        ///     Assert.IsNotNull(myStream);
        ///     myStream.SetData(b);
        ///     cf.Save("MyCompoundFile.cfs");
        ///     cf.Close();
        ///     
        /// </code>
        /// </example>
        public CompoundFile(): this(CfsVersion.Ver3,CfsConfiguration.Default)
        {

            //this.header = new Header();
            //this.sectorRecycle = false;

            ////this.sectors.OnVer3SizeLimitReached += new Ver3SizeLimitReached(OnSizeLimitReached);

            //DIFAT_SECTOR_FAT_ENTRIES_COUNT = (GetSectorSize() / 4) - 1;
            //FAT_SECTOR_ENTRIES_COUNT = (GetSectorSize() / 4);

            ////Root -- 
            //IDirectoryEntry de = DirectoryEntry.New("Root Entry", StgType.StgRoot, directoryEntries);
            //rootStorage = new CFStorage(this, de);
            //rootStorage.DirEntry.StgType = StgType.StgRoot;
            //rootStorage.DirEntry.StgColor = StgColor.Black;

            ////this.InsertNewDirectoryEntry(rootStorage.DirEntry);
        }

        void OnSizeLimitReached()
        {

            var rangeLockSector = new Sector(GetSectorSize(), SourceStream);
            _sectors.Add(rangeLockSector);

            rangeLockSector.Type = SectorType.RangeLockSector;

            _transactionLockAdded = true;
            _lockSectorId = rangeLockSector.Id;
        }


        /// <summary>
        /// Create a new, blank, compound file.
        /// </summary>
        /// <param name="cfsVersion">Use a specific Compound File Version to set 512 or 4096 bytes sectors</param>
        /// <param name="configFlags">Set <see cref="T:OpenMcdf.CFSConfiguration">configuration</see> parameters for the new compound file</param>
        /// <example>
        /// <code>
        /// 
        ///     byte[] b = new byte[10000];
        ///     for (int i = 0; i &lt; 10000; i++)
        ///     {
        ///         b[i % 120] = (byte)i;
        ///     }
        ///
        ///     CompoundFile cf = new CompoundFile(CFSVersion.Ver_4, CFSConfiguration.Default);
        ///     CFStream myStream = cf.RootStorage.AddStream("MyStream");
        ///
        ///     Assert.IsNotNull(myStream);
        ///     myStream.SetData(b);
        ///     cf.Save("MyCompoundFile.cfs");
        ///     cf.Close();
        ///     
        /// </code>
        /// </example>
        public CompoundFile(CfsVersion cfsVersion, CfsConfiguration configFlags)
        {
            Configuration = configFlags;

            var sectorRecycle = configFlags.HasFlag(CfsConfiguration.SectorRecycle);
            var eraseFreeSectors = configFlags.HasFlag(CfsConfiguration.EraseFreeSectors);

            _header = new Header((ushort)cfsVersion);

            if (cfsVersion == CfsVersion.Ver4)
                _sectors.OnVer3SizeLimitReached += OnSizeLimitReached;

            _sectorRecycle = sectorRecycle;


            _difatSectorFatEntriesCount = (GetSectorSize() / 4) - 1;
            _fatSectorEntriesCount = (GetSectorSize() / 4);

            //Root -- 
            var rootDir = DirectoryEntry.New("Root Entry", StgType.StgRoot, _directoryEntries);
            rootDir.StgColor = StgColor.Black;
            //this.InsertNewDirectoryEntry(rootDir);

            _rootStorage = new CfStorage(this, rootDir);


            //
        }


        /// <summary>
        /// Load an existing compound file.
        /// </summary>
        /// <param name="fileName">Compound file to read from</param>
        /// <example>
        /// <code>
        /// //A xls file should have a Workbook stream
        /// String filename = "report.xls";
        ///
        /// CompoundFile cf = new CompoundFile(filename);
        /// CFStream foundStream = cf.RootStorage.GetStream("Workbook");
        ///
        /// byte[] temp = foundStream.GetData();
        ///
        /// Assert.IsNotNull(temp);
        ///
        /// cf.Close();
        /// </code>
        /// </example>
        /// <remarks>
        /// File will be open in read-only mode: it has to be saved
        /// with a different filename. A wrapping implementation has to be provided 
        /// in order to remove/substitute an existing file. Version will be
        /// automatically recognized from the file. Sector recycle is turned off
        /// to achieve the best reading/writing performance in most common scenarios.
        /// </remarks>
        public CompoundFile(string fileName)
        {
            _sectorRecycle = false;
            _updateMode = CfsUpdateMode.ReadOnly;
            _eraseFreeSectors = false;

            LoadFile(fileName);

            _difatSectorFatEntriesCount = (GetSectorSize() / 4) - 1;
            _fatSectorEntriesCount = (GetSectorSize() / 4);
        }

        /// <summary>
        /// Load an existing compound file.
        /// </summary>
        /// <param name="fileName">Compound file to read from</param>
        /// <param name="sectorRecycle">If true, recycle unused sectors</param>
        /// <param name="updateMode">Select the update mode of the underlying data file</param>
        /// <param name="eraseFreeSectors">If true, overwrite with zeros unallocated sectors</param>
        /// <example>
        /// <code>
        /// String srcFilename = "data_YOU_CAN_CHANGE.xls";
        /// 
        /// CompoundFile cf = new CompoundFile(srcFilename, UpdateMode.Update, true, true);
        ///
        /// Random r = new Random();
        ///
        /// byte[] buffer = GetBuffer(r.Next(3, 4095), 0x0A);
        ///
        /// cf.RootStorage.AddStream("MyStream").SetData(buffer);
        /// 
        /// //This will persist data to the underlying media.
        /// cf.Commit();
        /// cf.Close();
        ///
        /// </code>
        /// </example>
        public CompoundFile(string fileName, CfsUpdateMode updateMode, CfsConfiguration configParameters)
        {
            Configuration = configParameters;
            _validationExceptionEnabled = !configParameters.HasFlag(CfsConfiguration.NoValidationException);
            _sectorRecycle = configParameters.HasFlag(CfsConfiguration.SectorRecycle);
            _updateMode = updateMode;
            _eraseFreeSectors = configParameters.HasFlag(CfsConfiguration.EraseFreeSectors);

            LoadFile(fileName);

            _difatSectorFatEntriesCount = (GetSectorSize() / 4) - 1;
            _fatSectorEntriesCount = (GetSectorSize() / 4);
        }

        private readonly bool _validationExceptionEnabled = true;

        public bool ValidationExceptionEnabled => _validationExceptionEnabled;


        /// <summary>
        /// Load an existing compound file.
        /// </summary>
        /// <param name="stream">A stream containing a compound file to read</param>
        /// <param name="sectorRecycle">If true, recycle unused sectors</param>
        /// <param name="updateMode">Select the update mode of the underlying data file</param>
        /// <param name="eraseFreeSectors">If true, overwrite with zeros unallocated sectors</param>
        /// <example>
        /// <code>
        /// 
        /// String filename = "reportREAD.xls";
        ///   
        /// FileStream fs = new FileStream(filename, FileMode.Open);
        /// CompoundFile cf = new CompoundFile(fs, UpdateMode.ReadOnly, false, false);
        /// CFStream foundStream = cf.RootStorage.GetStream("Workbook");
        ///
        /// byte[] temp = foundStream.GetData();
        ///
        /// Assert.IsNotNull(temp);
        ///
        /// cf.Close();
        ///
        /// </code>
        /// </example>
        /// <exception cref="T:OpenMcdf.CFException">Raised when trying to open a non-seekable stream</exception>
        /// <exception cref="T:OpenMcdf.CFException">Raised stream is null</exception>
        public CompoundFile(Stream stream, CfsUpdateMode updateMode, CfsConfiguration configParameters)
        {
            Configuration = configParameters;
            _validationExceptionEnabled = !configParameters.HasFlag(CfsConfiguration.NoValidationException);
            _sectorRecycle = configParameters.HasFlag(CfsConfiguration.SectorRecycle);
            _eraseFreeSectors = configParameters.HasFlag(CfsConfiguration.EraseFreeSectors);
            _closeStream = !configParameters.HasFlag(CfsConfiguration.LeaveOpen);

            _updateMode = updateMode;
            LoadStream(stream);

            _difatSectorFatEntriesCount = (GetSectorSize() / 4) - 1;
            _fatSectorEntriesCount = (GetSectorSize() / 4);
        }


        /// <summary>
        /// Load an existing compound file from a stream.
        /// </summary>
        /// <param name="stream">Streamed compound file</param>
        /// <example>
        /// <code>
        /// 
        /// String filename = "reportREAD.xls";
        ///   
        /// FileStream fs = new FileStream(filename, FileMode.Open);
        /// CompoundFile cf = new CompoundFile(fs);
        /// CFStream foundStream = cf.RootStorage.GetStream("Workbook");
        ///
        /// byte[] temp = foundStream.GetData();
        ///
        /// Assert.IsNotNull(temp);
        ///
        /// cf.Close();
        ///
        /// </code>
        /// </example>
        /// <exception cref="T:OpenMcdf.CFException">Raised when trying to open a non-seekable stream</exception>
        /// <exception cref="T:OpenMcdf.CFException">Raised stream is null</exception>
        public CompoundFile(Stream stream)
        {
            LoadStream(stream);

            _difatSectorFatEntriesCount = (GetSectorSize() / 4) - 1;
            _fatSectorEntriesCount = (GetSectorSize() / 4);
        }

        private readonly CfsUpdateMode _updateMode = CfsUpdateMode.ReadOnly;
        private string _fileName = string.Empty;


#if !FLAT_WRITE
        private byte[] buffer = new byte[FLUSHING_BUFFER_MAX_SIZE];
        private Queue<Sector> flushingQueue = new Queue<Sector>(FLUSHING_QUEUE_SIZE);
#endif


        /// <summary>
        /// Commit data changes since the previously commit operation
        /// to the underlying supporting stream or file on the disk.
        /// </summary>
        /// <param name="releaseMemory">If true, release loaded sectors to limit memory usage but reduces following read operations performance</param>
        /// <remarks>
        /// This method can be used only if 
        /// the supporting stream has been opened in 
        /// <see cref="T:OpenMcdf.UpdateMode">Update mode</see>.
        /// </remarks>
        public void Commit(bool releaseMemory = false)
        {
            if (_disposed)
                throw new CfDisposedException("Compound File closed: cannot commit data");

            if (_updateMode != CfsUpdateMode.Update)
                throw new CfInvalidOperation("Cannot commit data in Read-Only update mode");

            //try
            //{
#if !FLAT_WRITE

            int sId = -1;
            int sCount = 0;
            int bufOffset = 0;
#endif
            var sSize = GetSectorSize();

            if (_header.MajorVersion != (ushort)CfsVersion.Ver3)
                CheckForLockSector();

            SourceStream.Seek(0, SeekOrigin.Begin);
            SourceStream.Write(new byte[GetSectorSize()], 0, sSize);

            CommitDirectory();

            var gap = true;


            for (var i = 0; i < _sectors.Count; i++)
            {
#if FLAT_WRITE

                //Note:
                //Here sectors should not be loaded dynamically because
                //if they are null it means that no change has involved them;

                var s = _sectors[i];

                if (s is { DirtyFlag: true })
                {
                    if (gap)
                        SourceStream.Seek(sSize + i * (long)sSize, SeekOrigin.Begin);

                    SourceStream.Write(s.GetData(), 0, sSize);
                    SourceStream.Flush();
                    s.DirtyFlag = false;
                    gap = false;

                }
                else
                {
                    gap = true;
                }

                if (s == null || !releaseMemory) continue;
                s.ReleaseData();
                s = null;
                _sectors[i] = null;



#else
               

                Sector s = sectors[i] as Sector;


                if (s != null && s.DirtyFlag && flushingQueue.Count < (int)(buffer.Length / sSize))
                {
                    //First of a block of contiguous sectors, mark id, start enqueuing

                    if (gap)
                    {
                        sId = s.Id;
                        gap = false;
                    }

                    flushingQueue.Enqueue(s);


                }
                else
                {
                    //Found a gap, stop enqueuing, flush a write operation

                    gap = true;
                    sCount = flushingQueue.Count;

                    if (sCount == 0) continue;

                    bufOffset = 0;
                    while (flushingQueue.Count > 0)
                    {
                        Sector r = flushingQueue.Dequeue();
                        Buffer.BlockCopy(r.GetData(), 0, buffer, bufOffset, sSize);
                        r.DirtyFlag = false;

                        if (releaseMemory)
                        {
                            r.ReleaseData();
                        }

                        bufOffset += sSize;
                    }

                    sourceStream.Seek(((long)sSize + (long)sId * (long)sSize), SeekOrigin.Begin);
                    sourceStream.Write(buffer, 0, sCount * sSize);

               

                    //Console.WriteLine("W - " + (int)(sCount * sSize ));

                }
#endif
            }

#if !FLAT_WRITE
            sCount = flushingQueue.Count;
            bufOffset = 0;

            while (flushingQueue.Count > 0)
            {
                Sector r = flushingQueue.Dequeue();
                Buffer.BlockCopy(r.GetData(), 0, buffer, bufOffset, sSize);
                r.DirtyFlag = false;

                if (releaseMemory)
                {
                    r.ReleaseData();
                    r = null;
                }

                bufOffset += sSize;
            }

            if (sCount != 0)
            {
                sourceStream.Seek((long)sSize + (long)sId * (long)sSize, SeekOrigin.Begin);
                sourceStream.Write(buffer, 0, sCount * sSize);
                //Console.WriteLine("W - " + (int)(sCount * sSize));
            }

#endif

            // Seek to beginning position and save header (first 512 or 4096 bytes)
            SourceStream.Seek(0, SeekOrigin.Begin);
            _header.Write(SourceStream);

            SourceStream.SetLength((long)(_sectors.Count + 1) * sSize);
            SourceStream.Flush();

            if (releaseMemory)
                GC.Collect();

            //}
            //catch (Exception ex)
            //{
            //    throw new CFException("Internal error while committing data", ex);
            //}
        }

        /// <summary>
        /// Load compound file from an existing stream.
        /// </summary>
        /// <param name="stream">Stream to load compound file from</param>
        private void Load(Stream stream)
        {
            try
            {
                _header = new Header();
                _directoryEntries = new List<IDirectoryEntry>();

                SourceStream = stream;

                _header.Read(stream);

                var nSector = Ceiling(((stream.Length - GetSectorSize()) / (double)GetSectorSize()));

                if (stream.Length > 0x7FFFFF0)
                    _transactionLockAllocated = true;


                _sectors = new SectorCollection();
                //sectors = new ArrayList();
                for (var i = 0; i < nSector; i++)
                {
                    _sectors.Add(null);
                }

                LoadDirectories();

                _rootStorage
                    = new CfStorage(this, _directoryEntries[0]);
            }
            catch (Exception)
            {
                if (stream != null && _closeStream)
                    stream.Close();

                throw;
            }
        }

        private void LoadFile(string fileName)
        {

            _fileName = fileName;

            FileStream fs = null;

            try
            {
                if (_updateMode == CfsUpdateMode.ReadOnly)
                {
                    fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

                }
                else
                {
                    fs = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.Read);
                }

                Load(fs);

            }
            catch
            {
                if (fs != null)
                    fs.Close();

                throw;
            }
        }

        private void LoadStream(Stream stream)
        {
            if (stream == null)
                throw new CfException("Stream parameter cannot be null");

            if (!stream.CanSeek)
                throw new CfException("Cannot load a non-seekable Stream");


            stream.Seek(0, SeekOrigin.Begin);

            Load(stream);
        }

        /// <summary>
        /// Return true if this compound file has been 
        /// loaded from an existing file or stream
        /// </summary>
        public bool HasSourceStream => SourceStream != null;


        private void PersistMiniStreamToStream(List<Sector> miniSectorChain)
        {
            var miniStream
                = GetSectorChain(RootEntry.StartSetc, SectorType.Normal);

            var miniStreamView
                = new StreamView(
                    miniStream,
                    GetSectorSize(),
                    _rootStorage.Size,
                    null,
                    SourceStream);

            for (var i = 0; i < miniSectorChain.Count; i++)
            {
                var s = miniSectorChain[i];

                if (s.Id == -1)
                    throw new CfException("Invalid minisector index");

                // Ministream sectors already allocated
                miniStreamView.Seek(Sector.MinisectorSize * s.Id, SeekOrigin.Begin);
                miniStreamView.Write(s.GetData(), 0, Sector.MinisectorSize);
            }
        }

        /// <summary>
        /// Allocate space, setup sectors id and refresh header
        /// for the new or updated mini sector chain.
        /// </summary>
        /// <param name="sectorChain">The new MINI sector chain</param>
        private void AllocateMiniSectorChain(List<Sector> sectorChain)
        {
            var miniFat
                = GetSectorChain(_header.FirstMiniFatSectorId, SectorType.Normal);

            var miniStream
                = GetSectorChain(RootEntry.StartSetc, SectorType.Normal);

            var miniFatView
                = new StreamView(
                    miniFat,
                    GetSectorSize(),
                    _header.MiniFatSectorsNumber * Sector.MinisectorSize,
                    null,
                    SourceStream,
                    true
                    );

            var miniStreamView
                = new StreamView(
                    miniStream,
                    GetSectorSize(),
                    _rootStorage.Size,
                    null,
                    SourceStream);


            // Set updated/new sectors within the ministream
            // We are writing data in a NORMAL Sector chain.
            for (var i = 0; i < sectorChain.Count; i++)
            {
                var s = sectorChain[i];

                if (s.Id == -1)
                {
                    // Allocate, position ministream at the end of already allocated
                    // ministream's sectors

                    miniStreamView.Seek(_rootStorage.Size + Sector.MinisectorSize, SeekOrigin.Begin);
                    //miniStreamView.Write(s.GetData(), 0, Sector.MINISECTOR_SIZE);
                    s.Id = (int)(miniStreamView.Position - Sector.MinisectorSize) / Sector.MinisectorSize;

                    _rootStorage.DirEntry.Size = miniStreamView.Length;
                }
            }

            // Update miniFAT
            for (var i = 0; i < sectorChain.Count - 1; i++)
            {
                var currentId = sectorChain[i].Id;
                var nextId = sectorChain[i + 1].Id;

                miniFatView.Seek(currentId * 4, SeekOrigin.Begin);
                miniFatView.Write(BitConverter.GetBytes(nextId), 0, 4);
            }

            // Write End of Chain in MiniFAT
            miniFatView.Seek(sectorChain[sectorChain.Count - 1].Id * SizeOfSid, SeekOrigin.Begin);
            miniFatView.Write(BitConverter.GetBytes(Sector.Endofchain), 0, 4);

            // Update sector chains
            AllocateSectorChain(miniStreamView.BaseSectorChain);
            AllocateSectorChain(miniFatView.BaseSectorChain);

            //Update HEADER and root storage when ministream changes
            if (miniFat.Count > 0)
            {
                _rootStorage.DirEntry.StartSetc = miniStream[0].Id;
                _header.MiniFatSectorsNumber = (uint)miniFat.Count;
                _header.FirstMiniFatSectorId = miniFat[0].Id;
            }
        }

        internal void FreeData(CfStream stream)
        {
            if (stream.Size == 0)
                return;

            List<Sector> sectorChain = null;

            if (stream.Size < _header.MinSizeStandardStream)
            {
                sectorChain = GetSectorChain(stream.DirEntry.StartSetc, SectorType.Mini);
                FreeMiniChain(sectorChain, _eraseFreeSectors);
            }
            else
            {
                sectorChain = GetSectorChain(stream.DirEntry.StartSetc, SectorType.Normal);
                FreeChain(sectorChain, _eraseFreeSectors);
            }

            stream.DirEntry.StartSetc = Sector.Endofchain;
            stream.DirEntry.Size = 0;
        }

        private void FreeChain(List<Sector> sectorChain, bool zeroSector)
        {
            FreeChain(sectorChain, 0, zeroSector);
        }

        private void FreeChain(List<Sector> sectorChain, int nthSectorToRemove, bool zeroSector)
        {
            // Dummy zero buffer
            var zeroedSector = new byte[GetSectorSize()];

            var fat
                = GetSectorChain(-1, SectorType.Fat);

            var fatView
                = new StreamView(fat, GetSectorSize(), fat.Count * GetSectorSize(), null, SourceStream);

            // Zeroes out sector data (if required)-------------
            if (zeroSector)
            {
                for (var i = nthSectorToRemove; i < sectorChain.Count; i++)
                {
                    var s = sectorChain[i];
                    s.ZeroData();
                }
            }

            // Update FAT marking unallocated sectors ----------
            for (var i = nthSectorToRemove; i < sectorChain.Count; i++)
            {
                var currentId = sectorChain[i].Id;

                fatView.Seek(currentId * 4, SeekOrigin.Begin);
                fatView.Write(BitConverter.GetBytes(Sector.Freesect), 0, 4);
            }

            // Write new end of chain if partial free ----------
            if (nthSectorToRemove > 0 && sectorChain.Count > 0)
            {
                fatView.Seek(sectorChain[nthSectorToRemove - 1].Id * 4, SeekOrigin.Begin);
                fatView.Write(BitConverter.GetBytes(Sector.Endofchain), 0, 4);
            }
        }

        private void FreeMiniChain(List<Sector> sectorChain, bool zeroSector)
        {
            FreeMiniChain(sectorChain, 0, zeroSector);
        }

        private void FreeMiniChain(List<Sector> sectorChain, int nthSectorToRemove, bool zeroSector)
        {
            var zeroedMiniSector = new byte[Sector.MinisectorSize];

            var miniFat
                = GetSectorChain(_header.FirstMiniFatSectorId, SectorType.Normal);

            var miniStream
                = GetSectorChain(RootEntry.StartSetc, SectorType.Normal);

            var miniFatView
                = new StreamView(miniFat, GetSectorSize(), _header.MiniFatSectorsNumber * Sector.MinisectorSize, null, SourceStream);

            var miniStreamView
                = new StreamView(miniStream, GetSectorSize(), _rootStorage.Size, null, SourceStream);

            // Set updated/new sectors within the ministream ----------
            if (zeroSector)
            {
                for (var i = nthSectorToRemove; i < sectorChain.Count; i++)
                {
                    var s = sectorChain[i];

                    if (s.Id != -1)
                    {
                        // Overwrite
                        miniStreamView.Seek(Sector.MinisectorSize * s.Id, SeekOrigin.Begin);
                        miniStreamView.Write(zeroedMiniSector, 0, Sector.MinisectorSize);
                    }
                }
            }

            // Update miniFAT                ---------------------------------------
            for (var i = nthSectorToRemove; i < sectorChain.Count; i++)
            {
                var currentId = sectorChain[i].Id;

                miniFatView.Seek(currentId * 4, SeekOrigin.Begin);
                miniFatView.Write(BitConverter.GetBytes(Sector.Freesect), 0, 4);
            }

            // Write End of Chain in MiniFAT ---------------------------------------
            //miniFATView.Seek(sectorChain[(sectorChain.Count - 1) - nth_sector_to_remove].Id * SIZE_OF_SID, SeekOrigin.Begin);
            //miniFATView.Write(BitConverter.GetBytes(Sector.ENDOFCHAIN), 0, 4);

            // Write End of Chain in MiniFAT ---------------------------------------
            if (nthSectorToRemove > 0 && sectorChain.Count > 0)
            {
                miniFatView.Seek(sectorChain[nthSectorToRemove - 1].Id * 4, SeekOrigin.Begin);
                miniFatView.Write(BitConverter.GetBytes(Sector.Endofchain), 0, 4);
            }

            // Update sector chains           ---------------------------------------
            AllocateSectorChain(miniStreamView.BaseSectorChain);
            AllocateSectorChain(miniFatView.BaseSectorChain);

            //Update HEADER and root storage when ministream changes
            if (miniFat.Count > 0)
            {
                _rootStorage.DirEntry.StartSetc = miniStream[0].Id;
                _header.MiniFatSectorsNumber = (uint)miniFat.Count;
                _header.FirstMiniFatSectorId = miniFat[0].Id;
            }
        }

        /// <summary>
        /// Allocate space, setup sectors id in the FAT and refresh header
        /// for the new or updated sector chain (Normal or Mini sectors)
        /// </summary>
        /// <param name="sectorChain">The new or updated normal or mini sector chain</param>
        private void SetSectorChain(List<Sector> sectorChain)
        {
            if (sectorChain == null || sectorChain.Count == 0)
                return;

            var st = sectorChain[0].Type;

            if (st == SectorType.Normal)
            {
                AllocateSectorChain(sectorChain);
            }
            else if (st == SectorType.Mini)
            {
                AllocateMiniSectorChain(sectorChain);
            }
        }

        /// <summary>
        /// Allocate space, setup sectors id and refresh header
        /// for the new or updated sector chain.
        /// </summary>
        /// <param name="sectorChain">The new or updated generic sector chain</param>
        private void AllocateSectorChain(List<Sector> sectorChain)
        {

            foreach (var s in sectorChain)
            {
                if (s.Id == -1)
                {
                    _sectors.Add(s);
                    s.Id = _sectors.Count - 1;

                }
            }

            AllocateFatSectorChain(sectorChain);
        }

        private bool _transactionLockAdded;
        private int _lockSectorId = -1;
        private bool _transactionLockAllocated;

        /// <summary>
        /// Check for transaction lock sector addition and mark it in the FAT.
        /// </summary>
        private void CheckForLockSector()
        {
            //If transaction lock has been added and not yet allocated in the FAT...
            if (_transactionLockAdded && !_transactionLockAllocated)
            {
                var fatStream = new StreamView(GetFatSectorChain(), GetSectorSize(), SourceStream);

                fatStream.Seek(_lockSectorId * 4, SeekOrigin.Begin);
                fatStream.Write(BitConverter.GetBytes(Sector.Endofchain), 0, 4);

                _transactionLockAllocated = true;
            }

        }
        /// <summary>
        /// Allocate space, setup sectors id and refresh header
        /// for the new or updated FAT sector chain.
        /// </summary>
        /// <param name="sectorChain">The new or updated generic sector chain</param>
        private void AllocateFatSectorChain(List<Sector> sectorChain)
        {
            var fatSectors = GetSectorChain(-1, SectorType.Fat);

            var fatStream =
                new StreamView(
                    fatSectors,
                    GetSectorSize(),
                    _header.FatSectorsNumber * GetSectorSize(),
                    null,
                    SourceStream,
                    true
                    );

            // Write FAT chain values --

            for (var i = 0; i < sectorChain.Count - 1; i++)
            {

                var sN = sectorChain[i + 1];
                var sC = sectorChain[i];

                fatStream.Seek(sC.Id * 4, SeekOrigin.Begin);
                fatStream.Write(BitConverter.GetBytes(sN.Id), 0, 4);
            }

            fatStream.Seek(sectorChain[sectorChain.Count - 1].Id * 4, SeekOrigin.Begin);
            fatStream.Write(BitConverter.GetBytes(Sector.Endofchain), 0, 4);

            // Merge chain to CFS
            AllocateDifatSectorChain(fatStream.BaseSectorChain);
        }

        /// <summary>
        /// Setup the DIFAT sector chain
        /// </summary>
        /// <param name="faTsectorChain">A FAT sector chain</param>
        private void AllocateDifatSectorChain(List<Sector> faTsectorChain)
        {
            // Get initial sector's count
            _header.FatSectorsNumber = faTsectorChain.Count;

            // Allocate Sectors
            foreach (var s in faTsectorChain)
            {
                if (s.Id == -1)
                {
                    _sectors.Add(s);
                    s.Id = _sectors.Count - 1;
                    s.Type = SectorType.Fat;
                }
            }

            // Sector count...
            var nCurrentSectors = _sectors.Count;

            // Temp DIFAT count
            var nDifatSectors = (int)_header.DifatSectorsNumber;

            if (faTsectorChain.Count > HeaderDifatEntriesCount)
            {
                nDifatSectors = Ceiling((double)(faTsectorChain.Count - HeaderDifatEntriesCount) / _difatSectorFatEntriesCount);
                nDifatSectors = LowSaturation(nDifatSectors - (int)_header.DifatSectorsNumber); //required DIFAT
            }

            // ...sum with new required DIFAT sectors count
            nCurrentSectors += nDifatSectors;

            // ReCheck FAT bias
            while (_header.FatSectorsNumber * _fatSectorEntriesCount < nCurrentSectors)
            {
                var extraFatSector = new Sector(GetSectorSize(), SourceStream);
                _sectors.Add(extraFatSector);

                extraFatSector.Id = _sectors.Count - 1;
                extraFatSector.Type = SectorType.Fat;

                faTsectorChain.Add(extraFatSector);

                _header.FatSectorsNumber++;
                nCurrentSectors++;

                //... so, adding a FAT sector may induce DIFAT sectors to increase by one
                // and consequently this may induce ANOTHER FAT sector (TO-THINK: May this condition occure ?)
                if (nDifatSectors * _difatSectorFatEntriesCount <
                    (_header.FatSectorsNumber > HeaderDifatEntriesCount ?
                    _header.FatSectorsNumber - HeaderDifatEntriesCount :
                    0))
                {
                    nDifatSectors++;
                    nCurrentSectors++;
                }
            }


            var difatSectors =
                        GetSectorChain(-1, SectorType.Difat);

            var difatStream
                = new StreamView(difatSectors, GetSectorSize(), SourceStream);

            // Write DIFAT Sectors (if required)
            // Save room for the following chaining
            for (var i = 0; i < faTsectorChain.Count; i++)
            {
                if (i < HeaderDifatEntriesCount)
                {
                    _header.Difat[i] = faTsectorChain[i].Id;
                }
                else
                {
                    // room for DIFAT chaining at the end of any DIFAT sector (4 bytes)
                    if (i != HeaderDifatEntriesCount && (i - HeaderDifatEntriesCount) % _difatSectorFatEntriesCount == 0)
                    {
                        var temp = new byte[sizeof(int)];
                        difatStream.Write(temp, 0, sizeof(int));
                    }

                    difatStream.Write(BitConverter.GetBytes(faTsectorChain[i].Id), 0, sizeof(int));

                }
            }

            // Allocate room for DIFAT sectors
            for (var i = 0; i < difatStream.BaseSectorChain.Count; i++)
            {
                if (difatStream.BaseSectorChain[i].Id == -1)
                {
                    _sectors.Add(difatStream.BaseSectorChain[i]);
                    difatStream.BaseSectorChain[i].Id = _sectors.Count - 1;
                    difatStream.BaseSectorChain[i].Type = SectorType.Difat;
                }
            }

            _header.DifatSectorsNumber = (uint)nDifatSectors;


            // Chain first sector
            if (difatStream.BaseSectorChain != null && difatStream.BaseSectorChain.Count > 0)
            {
                _header.FirstDifatSectorId = difatStream.BaseSectorChain[0].Id;

                // Update header information
                _header.DifatSectorsNumber = (uint)difatStream.BaseSectorChain.Count;

                // Write chaining information at the end of DIFAT Sectors
                for (var i = 0; i < difatStream.BaseSectorChain.Count - 1; i++)
                {
                    Buffer.BlockCopy(
                        BitConverter.GetBytes(difatStream.BaseSectorChain[i + 1].Id),
                        0,
                        difatStream.BaseSectorChain[i].GetData(),
                        GetSectorSize() - sizeof(int),
                        4);
                }

                Buffer.BlockCopy(
                    BitConverter.GetBytes(Sector.Endofchain),
                    0,
                    difatStream.BaseSectorChain[difatStream.BaseSectorChain.Count - 1].GetData(),
                    GetSectorSize() - sizeof(int),
                    sizeof(int)
                    );
            }
            else
                _header.FirstDifatSectorId = Sector.Endofchain;

            // Mark DIFAT Sectors in FAT
            var fatSv =
                new StreamView(faTsectorChain, GetSectorSize(), _header.FatSectorsNumber * GetSectorSize(), null, SourceStream);

            for (var i = 0; i < _header.DifatSectorsNumber; i++)
            {
                fatSv.Seek(difatStream.BaseSectorChain[i].Id * 4, SeekOrigin.Begin);
                fatSv.Write(BitConverter.GetBytes(Sector.Difsect), 0, 4);
            }

            for (var i = 0; i < _header.FatSectorsNumber; i++)
            {
                fatSv.Seek(fatSv.BaseSectorChain[i].Id * 4, SeekOrigin.Begin);
                fatSv.Write(BitConverter.GetBytes(Sector.Fatsect), 0, 4);
            }

            //fatSv.Seek(fatSv.BaseSectorChain[fatSv.BaseSectorChain.Count - 1].Id * 4, SeekOrigin.Begin);
            //fatSv.Write(BitConverter.GetBytes(Sector.ENDOFCHAIN), 0, 4);

            _header.FatSectorsNumber = fatSv.BaseSectorChain.Count;
        }


        /// <summary>
        /// Get the DIFAT Sector chain
        /// </summary>
        /// <returns>A list of DIFAT sectors</returns>
        private List<Sector> GetDifatSectorChain()
        {
            var validationCount = 0;

            var result
                = new List<Sector>();

            var nextSecId
               = Sector.Endofchain;

            var processedSectors = new HashSet<int>();

            if (_header.DifatSectorsNumber != 0)
            {
                validationCount = (int)_header.DifatSectorsNumber;

                var s = _sectors[_header.FirstDifatSectorId];

                if (s == null) //Lazy loading
                {
                    s = new Sector(GetSectorSize(), SourceStream);
                    s.Type = SectorType.Difat;
                    s.Id = _header.FirstDifatSectorId;
                    _sectors[_header.FirstDifatSectorId] = s;
                }

                result.Add(s);

                while (true && validationCount >= 0)
                {
                    nextSecId = BitConverter.ToInt32(s.GetData(), GetSectorSize() - 4);
                    EnsureUniqueSectorIndex(nextSecId, processedSectors);

                    // Strictly speaking, the following condition is not correct from
                    // a specification point of view:
                    // only ENDOFCHAIN should break DIFAT chain but 
                    // a lot of existing compound files use FREESECT as DIFAT chain termination
                    if (nextSecId == Sector.Freesect || nextSecId == Sector.Endofchain) break;

                    validationCount--;

                    if (validationCount < 0)
                    {
                        if (_closeStream)
                            Close();

                        if (_validationExceptionEnabled)
                            throw new CfCorruptedFileException("DIFAT sectors count mismatched. Corrupted compound file");
                    }

                    s = _sectors[nextSecId];

                    if (s == null)
                    {
                        s = new Sector(GetSectorSize(), SourceStream);
                        s.Id = nextSecId;
                        _sectors[nextSecId] = s;
                    }

                    result.Add(s);
                }
            }

            return result;
        }

        private void EnsureUniqueSectorIndex(int nextSecId, HashSet<int> processedSectors)
        {
            if (processedSectors.Contains(nextSecId) && _validationExceptionEnabled)
            {
                throw new CfCorruptedFileException("The file is corrupted.");
            }

            processedSectors.Add(nextSecId);
        }

        /// <summary>
        /// Get the FAT sector chain
        /// </summary>
        /// <returns>List of FAT sectors</returns>
        private List<Sector> GetFatSectorChain()
        {
            var nHeaderFatEntry = 109; //Number of FAT sectors id in the header

            var result
               = new List<Sector>();

            var nextSecId
               = Sector.Endofchain;

            var difatSectors = GetDifatSectorChain();

            var idx = 0;

            // Read FAT entries from the header Fat entry array (max 109 entries)
            while (idx < _header.FatSectorsNumber && idx < nHeaderFatEntry)
            {
                nextSecId = _header.Difat[idx];
                var s = _sectors[nextSecId];

                if (s == null)
                {
                    s = new Sector(GetSectorSize(), SourceStream);
                    s.Id = nextSecId;
                    s.Type = SectorType.Fat;
                    _sectors[nextSecId] = s;
                }

                result.Add(s);

                idx++;
            }

            //Is there any DIFAT sector containing other FAT entries ?
            if (difatSectors.Count > 0)
            {
                var processedSectors = new HashSet<int>();
                var difatStream
                    = new StreamView
                        (
                        difatSectors,
                        GetSectorSize(),
                        _header.FatSectorsNumber > nHeaderFatEntry ?
                            (_header.FatSectorsNumber - nHeaderFatEntry) * 4 :
                            0,
                        null,
                            SourceStream
                        );

                var nextDifatSectorBuffer = new byte[4];



                var i = 0;

                while (result.Count < _header.FatSectorsNumber)
                {
                    difatStream.Read(nextDifatSectorBuffer, 0, 4);
                    nextSecId = BitConverter.ToInt32(nextDifatSectorBuffer, 0);

                    EnsureUniqueSectorIndex(nextSecId, processedSectors);

                    var s = _sectors[nextSecId];

                    if (s == null)
                    {
                        s = new Sector(GetSectorSize(), SourceStream);
                        s.Type = SectorType.Fat;
                        s.Id = nextSecId;
                        _sectors[nextSecId] = s;//UUU
                    }

                    result.Add(s);

                    //difatStream.Read(nextDIFATSectorBuffer, 0, 4);
                    //nextSecID = BitConverter.ToInt32(nextDIFATSectorBuffer, 0);


                    if (difatStream.Position == ((GetSectorSize() - 4) + i * GetSectorSize()))
                    {
                        // Skip DIFAT chain fields considering the possibility that the last FAT entry has been already read
                        difatStream.Read(nextDifatSectorBuffer, 0, 4);
                        if (BitConverter.ToInt32(nextDifatSectorBuffer, 0) == Sector.Endofchain)
                            break;
                        i++;
                    }
                }
            }

            return result;

        }

        /// <summary>
        /// Get a standard sector chain
        /// </summary>
        /// <param name="secId">First SecID of the required chain</param>
        /// <returns>A list of sectors</returns>
        private List<Sector> GetNormalSectorChain(int secId)
        {
            var result
                   = new List<Sector>();

            var nextSecId = secId;

            var fatSectors = GetFatSectorChain();
            var processedSectors = new HashSet<int>();

            var fatStream
                = new StreamView(fatSectors, GetSectorSize(), fatSectors.Count * GetSectorSize(), null, SourceStream);

            while (true)
            {
                if (nextSecId == Sector.Endofchain) break;

                if (nextSecId < 0)
                    throw new CfCorruptedFileException(string.Format("Next Sector ID reference is below zero. NextID : {0}", nextSecId));

                if (nextSecId >= _sectors.Count)
                    throw new CfCorruptedFileException(string.Format("Next Sector ID reference an out of range sector. NextID : {0} while sector count {1}", nextSecId, _sectors.Count));

                var s = _sectors[nextSecId];
                if (s == null)
                {
                    s = new Sector(GetSectorSize(), SourceStream);
                    s.Id = nextSecId;
                    s.Type = SectorType.Normal;
                    _sectors[nextSecId] = s;
                }

                result.Add(s);

                fatStream.Seek(nextSecId * 4, SeekOrigin.Begin);
                var next = fatStream.ReadInt32();

                EnsureUniqueSectorIndex(next, processedSectors);
                nextSecId = next;

            }


            return result;
        }

        /// <summary>
        /// Get a mini sector chain
        /// </summary>
        /// <param name="secId">First SecID of the required chain</param>
        /// <returns>A list of mini sectors (64 bytes)</returns>
        private List<Sector> GetMiniSectorChain(int secId)
        {
            var result
                  = new List<Sector>();

            if (secId != Sector.Endofchain)
            {
                var nextSecId = secId;

                var miniFat = GetNormalSectorChain(_header.FirstMiniFatSectorId);
                var miniStream = GetNormalSectorChain(RootEntry.StartSetc);

                var miniFatView
                    = new StreamView(miniFat, GetSectorSize(), _header.MiniFatSectorsNumber * Sector.MinisectorSize, null, SourceStream);

                var miniStreamView =
                    new StreamView(miniStream, GetSectorSize(), _rootStorage.Size, null, SourceStream);

                var miniFatReader = new BinaryReader(miniFatView);

                nextSecId = secId;

                var processedSectors = new HashSet<int>();

                while (true)
                {
                    if (nextSecId == Sector.Endofchain)
                        break;

                    var ms = new Sector(Sector.MinisectorSize, SourceStream);
                    var temp = new byte[Sector.MinisectorSize];

                    ms.Id = nextSecId;
                    ms.Type = SectorType.Mini;

                    miniStreamView.Seek(nextSecId * Sector.MinisectorSize, SeekOrigin.Begin);
                    miniStreamView.Read(ms.GetData(), 0, Sector.MinisectorSize);

                    result.Add(ms);

                    miniFatView.Seek(nextSecId * 4, SeekOrigin.Begin);
                    var next = miniFatReader.ReadInt32();

                    nextSecId = next;
                    EnsureUniqueSectorIndex(nextSecId, processedSectors);
                }
            }
            return result;
        }


        /// <summary>
        /// Get a sector chain from a compound file given the first sector ID
        /// and the required sector type.
        /// </summary>
        /// <param name="secId">First chain sector's id </param>
        /// <param name="chainType">Type of Sectors in the required chain (mini sectors, normal sectors or FAT)</param>
        /// <returns>A list of Sectors as the result of their concatenation</returns>
        internal List<Sector> GetSectorChain(int secId, SectorType chainType)
        {

            switch (chainType)
            {
                case SectorType.Difat:
                    return GetDifatSectorChain();

                case SectorType.Fat:
                    return GetFatSectorChain();

                case SectorType.Normal:
                    return GetNormalSectorChain(secId);

                case SectorType.Mini:
                    return GetMiniSectorChain(secId);

                default:
                    throw new CfException("Unsupproted chain type");
            }
        }

        private CfStorage _rootStorage;

        /// <summary>
        /// The entry point object that represents the 
        /// root of the structures tree to get or set storage or
        /// stream data.
        /// </summary>
        /// <example>
        /// <code>
        /// 
        ///    //Create a compound file
        ///    string FILENAME = "MyFileName.cfs";
        ///    CompoundFile ncf = new CompoundFile();
        ///
        ///    CFStorage l1 = ncf.RootStorage.AddStorage("Storage Level 1");
        ///
        ///    l1.AddStream("l1ns1");
        ///    l1.AddStream("l1ns2");
        ///    l1.AddStream("l1ns3");
        ///    CFStorage l2 = l1.AddStorage("Storage Level 2");
        ///    l2.AddStream("l2ns1");
        ///    l2.AddStream("l2ns2");
        ///
        ///    ncf.Save(FILENAME);
        ///    ncf.Close();
        /// </code>
        /// </example>
        public CfStorage RootStorage => _rootStorage;

        public CfsVersion Version => (CfsVersion)_header.MajorVersion;


        /// <summary>
        /// Reset a directory entry setting it to StgInvalid in the Directory.
        /// </summary>
        /// <param name="sid">Sid of the directory to invalidate</param>
        internal void ResetDirectoryEntry(int sid)
        {
            _directoryEntries[sid].SetEntryName(string.Empty);
            _directoryEntries[sid].Left = null;
            _directoryEntries[sid].Right = null;
            _directoryEntries[sid].Parent = null;
            _directoryEntries[sid].StgType = StgType.StgInvalid;
            _directoryEntries[sid].StartSetc = DirectoryEntry.Zero;
            _directoryEntries[sid].StorageClsid = Guid.Empty;
            _directoryEntries[sid].Size = 0;
            _directoryEntries[sid].StateBits = 0;
            _directoryEntries[sid].StgColor = StgColor.Red;
            _directoryEntries[sid].CreationDate = new byte[8];
            _directoryEntries[sid].ModifyDate = new byte[8];
        }



        //internal class NodeFactory : IRBTreeDeserializer<CFItem>
        //{

        //    public RBNode<CFItem> DeserizlizeFromValues()
        //    {
        //           RBNode<CFItem> node = new RBNode<CFItem>(value,(Color)value.DirEntry.StgColor,
        //    }
        //}

        internal static RbTree CreateNewTree()
        {
            var bst = new RbTree();
            //bst.NodeInserted += OnNodeInsert;
            //bst.NodeOperation += OnNodeOperation;
            //bst.NodeDeleted += new Action<RBNode<CFItem>>(OnNodeDeleted);
            //  bst.ValueAssignedAction += new Action<RBNode<CFItem>, CFItem>(OnValueAssigned);
            return bst;
        }

        //void OnValueAssigned(RBNode<CFItem> node, CFItem from)
        //{
        //    if (from.DirEntry != null && from.DirEntry.LeftSibling != DirectoryEntry.NOSTREAM)

        //    if (from.DirEntry != null && from.DirEntry.LeftSibling != DirectoryEntry.NOSTREAM)
        //        node.Value.DirEntry.LeftSibling = from.DirEntry.LeftSibling;

        //    if (from.DirEntry != null && from.DirEntry.RightSibling != DirectoryEntry.NOSTREAM)
        //        node.Value.DirEntry.RightSibling = from.DirEntry.RightSibling;
        //}


        internal RbTree GetChildrenTree(int sid)
        {
            var bst = new RbTree();


            // Load children from their original tree.
            DoLoadChildren(bst, _directoryEntries[sid]);
            //bst = DoLoadChildrenTrusted(directoryEntries[sid]);

            //bst.Print();
            //bst.Print();
            //Trace.WriteLine("#### After rethreading");

            return bst;
        }

        private RbTree DoLoadChildrenTrusted(IDirectoryEntry de)
        {
            RbTree bst = null;

            if (de.Child != DirectoryEntry.Nostream)
            {
                bst = new RbTree(_directoryEntries[de.Child]);
            }

            return bst;
        }


        private void DoLoadChildren(RbTree bst, IDirectoryEntry de)
        {

            if (de.Child != DirectoryEntry.Nostream)
            {
                if (_directoryEntries[de.Child].StgType == StgType.StgInvalid) return;

                LoadSiblings(bst, _directoryEntries[de.Child]);
                NullifyChildNodes(_directoryEntries[de.Child]);
                bst.Insert(_directoryEntries[de.Child]);
            }
        }

        private void NullifyChildNodes(IDirectoryEntry de)
        {
            de.Parent = null;
            de.Left = null;
            de.Right = null;
        }

        private readonly List<int> _levelSiDs = new List<int>();

        // Doubling methods allows iterative behavior while avoiding
        // to insert duplicate items
        private void LoadSiblings(RbTree bst, IDirectoryEntry de)
        {
            _levelSiDs.Clear();

            if (de.LeftSibling != DirectoryEntry.Nostream)
            {


                // If there're more left siblings load them...
                DoLoadSiblings(bst, _directoryEntries[de.LeftSibling]);
                //NullifyChildNodes(directoryEntries[de.LeftSibling]);
            }

            if (de.RightSibling != DirectoryEntry.Nostream)
            {
                _levelSiDs.Add(de.RightSibling);

                // If there're more right siblings load them...
                DoLoadSiblings(bst, _directoryEntries[de.RightSibling]);
                //NullifyChildNodes(directoryEntries[de.RightSibling]);
            }
        }

        private void DoLoadSiblings(RbTree bst, IDirectoryEntry de)
        {
            if (ValidateSibling(de.LeftSibling))
            {
                _levelSiDs.Add(de.LeftSibling);

                // If there're more left siblings load them...
                DoLoadSiblings(bst, _directoryEntries[de.LeftSibling]);
            }

            if (ValidateSibling(de.RightSibling))
            {
                _levelSiDs.Add(de.RightSibling);

                // If there're more right siblings load them...
                DoLoadSiblings(bst, _directoryEntries[de.RightSibling]);
            }

            NullifyChildNodes(de);
            bst.Insert(de);
        }

        private bool ValidateSibling(int sid)
        {
            if (sid != DirectoryEntry.Nostream)
            {
                // if this siblings id does not overflow current list
                if (sid >= _directoryEntries.Count)
                {
                    if (_validationExceptionEnabled)
                    {
                        //this.Close();
                        throw new CfCorruptedFileException("A Directory Entry references the non-existent sid number " + sid);
                    }

                    return false;
                }

                //if this sibling is valid...
                if (_directoryEntries[sid].StgType == StgType.StgInvalid)
                {
                    if (_validationExceptionEnabled)
                    {
                        //this.Close();
                        throw new CfCorruptedFileException("A Directory Entry has a valid reference to an Invalid Storage Type directory [" + sid + "]");
                    }

                    return false;
                }

                if (!Enum.IsDefined(typeof(StgType), _directoryEntries[sid].StgType))
                {

                    if (_validationExceptionEnabled)
                    {
                        //this.Close();
                        throw new CfCorruptedFileException("A Directory Entry has an invalid Storage Type");
                    }

                    return false;
                }

                if (_levelSiDs.Contains(sid))
                    throw new CfCorruptedFileException("Cyclic reference of directory item");

                return true; //No fault condition encountered for sid being validated
            }

            return false;
        }


        /// <summary>
        /// Load directory entries from compound file. Header and FAT MUST be already loaded.
        /// </summary>
        private void LoadDirectories()
        {
            var directoryChain
                = GetSectorChain(_header.FirstDirectorySectorId, SectorType.Normal);

            if (!(directoryChain.Count > 0))
                throw new CfCorruptedFileException("Directory sector chain MUST contain at least 1 sector");

            if (_header.FirstDirectorySectorId == Sector.Endofchain)
                _header.FirstDirectorySectorId = directoryChain[0].Id;

            var dirReader
                = new StreamView(directoryChain, GetSectorSize(), directoryChain.Count * GetSectorSize(), null, SourceStream);


            while (dirReader.Position < directoryChain.Count * GetSectorSize())
            {
                var de
                = DirectoryEntry.New(string.Empty, StgType.StgInvalid, _directoryEntries);

                //We are not inserting dirs. Do not use 'InsertNewDirectoryEntry'
                de.Read(dirReader, Version);

            }
        }



        /// <summary>
        ///  Commit directory entries change on the Current Source stream
        /// </summary>
        private void CommitDirectory()
        {
            const int directorySize = 128;

            var directorySectors
                = GetSectorChain(_header.FirstDirectorySectorId, SectorType.Normal);

            var sv = new StreamView(directorySectors, GetSectorSize(), 0, null, SourceStream);

            foreach (var di in _directoryEntries)
            {
                di.Write(sv);
            }

            var delta = _directoryEntries.Count;

            while (delta % (GetSectorSize() / directorySize) != 0)
            {
                var dummy = DirectoryEntry.New(string.Empty, StgType.StgInvalid, _directoryEntries);
                dummy.Write(sv);
                delta++;
            }

            foreach (var s in directorySectors)
            {
                s.Type = SectorType.Directory;
            }

            AllocateSectorChain(directorySectors);

            _header.FirstDirectorySectorId = directorySectors[0].Id;

            //Version 4 supports directory sectors count
            if (_header.MajorVersion == 3)
            {
                _header.DirectorySectorsNumber = 0;
            }
            else
            {
                _header.DirectorySectorsNumber = directorySectors.Count;
            }
        }


        /// <summary>
        /// Saves the in-memory image of Compound File to a file.
        /// </summary>
        /// <param name="fileName">File name to write the compound file to</param>
        /// <exception cref="T:OpenMcdf.CFException">Raised if destination file is not seekable</exception>

        public void Save(string fileName)
        {
            if (_disposed)
                throw new CfException("Compound File closed: cannot save data");

            FileStream fs = null;

            try
            {
                fs = new FileStream(fileName, FileMode.Create);
                Save(fs);
            }
            catch (Exception ex)
            {
                throw new CfException("Error saving file [" + fileName + "]", ex);
            }
            finally
            {
                if (fs != null)
                    fs.Flush();

                if (fs != null)
                    fs.Close();

            }
        }

        /// <summary>
        /// Saves the in-memory image of Compound File to a stream.
        /// </summary>        
        /// <remarks>
        /// Destination Stream must be seekable. Uncommitted data will be persisted to the destination stream.
        /// </remarks>
        /// <param name="stream">The stream to save compound File to</param>
        /// <exception cref="T:OpenMcdf.CFException">Raised if destination stream is not seekable</exception>
        /// <exception cref="T:OpenMcdf.CFDisposedException">Raised if Compound File Storage has been already disposed</exception>
        /// <example>
        /// <code>
        ///    MemoryStream ms = new MemoryStream(size);
        ///
        ///    CompoundFile cf = new CompoundFile();
        ///    CFStorage st = cf.RootStorage.AddStorage("MyStorage");
        ///    CFStream sm = st.AddStream("MyStream");
        ///
        ///    byte[] b = new byte[]{0x00,0x01,0x02,0x03};
        ///
        ///    sm.SetData(b);
        ///    cf.Save(ms);
        ///    cf.Close();
        /// </code>
        /// </example>
        public void Save(Stream stream)
        {
            if (_disposed)
                throw new CfDisposedException("Compound File closed: cannot save data");

            if (!stream.CanSeek)
                throw new CfException("Cannot save on a non-seekable stream");

            CheckForLockSector();
            var sSize = GetSectorSize();

            try
            {
                stream.Write((byte[])Array.CreateInstance(typeof(byte), sSize), 0, sSize);

                CommitDirectory();

                for (var i = 0; i < _sectors.Count; i++)
                {
                    var s = _sectors[i];

                    if (s == null)
                    {
                        // Load source (unmodified) sectors
                        // Here we have to ignore "Dirty flag" of 
                        // sectors because we are NOT modifying the source
                        // in a differential way but ALL sectors need to be 
                        // persisted on the destination stream
                        s = new Sector(sSize, SourceStream);
                        s.Id = i;

                        //sectors[i] = s;
                    }


                    stream.Write(s.GetData(), 0, sSize);

                    //s.ReleaseData();

                }

                stream.Seek(0, SeekOrigin.Begin);
                _header.Write(stream);
            }
            catch (Exception ex)
            {
                throw new CfException("Internal error while saving compound file to stream ", ex);
            }
        }


        /// <summary>
        /// Scan FAT o miniFAT for free sectors to reuse.
        /// </summary>
        /// <param name="sType">Type of sector to look for</param>
        /// <returns>A Queue of available sectors or minisectors already allocated</returns>
        internal Queue<Sector> FindFreeSectors(SectorType sType)
        {
            var freeList = new Queue<Sector>();

            if (sType == SectorType.Normal)
            {

                var fatChain = GetSectorChain(-1, SectorType.Fat);
                var fatStream = new StreamView(fatChain, GetSectorSize(), _header.FatSectorsNumber * GetSectorSize(), null, SourceStream);

                var idx = 0;

                while (idx < _sectors.Count)
                {
                    var id = fatStream.ReadInt32();

                    if (id == Sector.Freesect)
                    {
                        if (_sectors[idx] == null)
                        {
                            var s = new Sector(GetSectorSize(), SourceStream);
                            s.Id = idx;
                            _sectors[idx] = s;

                        }

                        freeList.Enqueue(_sectors[idx]);
                    }

                    idx++;
                }
            }
            else
            {
                var miniFat
                    = GetSectorChain(_header.FirstMiniFatSectorId, SectorType.Normal);

                var miniFatView
                    = new StreamView(miniFat, GetSectorSize(), _header.MiniFatSectorsNumber * Sector.MinisectorSize, null, SourceStream);

                var miniStream
                    = GetSectorChain(RootEntry.StartSetc, SectorType.Normal);

                var miniStreamView
                    = new StreamView(miniStream, GetSectorSize(), _rootStorage.Size, null, SourceStream);

                var idx = 0;

                var nMinisectors = (int)(miniStreamView.Length / Sector.MinisectorSize);

                while (idx < nMinisectors)
                {
                    //AssureLength(miniStreamView, (int)miniFATView.Length);

                    var nextId = miniFatView.ReadInt32();

                    if (nextId == Sector.Freesect)
                    {
                        var ms = new Sector(Sector.MinisectorSize, SourceStream);
                        var temp = new byte[Sector.MinisectorSize];

                        ms.Id = idx;
                        ms.Type = SectorType.Mini;

                        miniStreamView.Seek(ms.Id * Sector.MinisectorSize, SeekOrigin.Begin);
                        miniStreamView.Read(ms.GetData(), 0, Sector.MinisectorSize);

                        freeList.Enqueue(ms);
                    }

                    idx++;
                }
            }

            return freeList;
        }

        /// <summary>
        /// INTERNAL DEVELOPMENT. DO NOT CALL.
        /// <param name="directoryEntry"></param>
        /// <param name="buffer"></param>
        internal void AppendData(CfItem cfItem, byte[] buffer)
        {
            WriteData(cfItem, cfItem.Size, buffer);
        }

        /// <summary>
        /// Resize stream length
        /// </summary>
        /// <param name="cfItem"></param>
        /// <param name="length"></param>
        internal void SetStreamLength(CfItem cfItem, long length)
        {
            if (cfItem.Size == length)
                return;

            var newSectorType = SectorType.Normal;
            var newSectorSize = GetSectorSize();

            if (length < _header.MinSizeStandardStream)
            {
                newSectorType = SectorType.Mini;
                newSectorSize = Sector.MinisectorSize;
            }

            var oldSectorType = SectorType.Normal;
            var oldSectorSize = GetSectorSize();

            if (cfItem.Size < _header.MinSizeStandardStream)
            {
                oldSectorType = SectorType.Mini;
                oldSectorSize = Sector.MinisectorSize;
            }

            var oldSize = cfItem.Size;


            // Get Sector chain and delta size induced by client
            var sectorChain = GetSectorChain(cfItem.DirEntry.StartSetc, oldSectorType);
            var delta = length - cfItem.Size;

            // Check for transition ministream -> stream:
            // Only in this case we need to free old sectors,
            // otherwise they will be overwritten.

            var transitionToMini = false;
            var transitionToNormal = false;
            List<Sector> oldChain = null;

            if (cfItem.DirEntry.StartSetc != Sector.Endofchain)
            {
                if (
                    (length < _header.MinSizeStandardStream && cfItem.DirEntry.Size >= _header.MinSizeStandardStream)
                    || (length >= _header.MinSizeStandardStream && cfItem.DirEntry.Size < _header.MinSizeStandardStream)
                   )
                {
                    if (cfItem.DirEntry.Size < _header.MinSizeStandardStream)
                    {
                        transitionToNormal = true;
                        oldChain = sectorChain;
                    }
                    else
                    {
                        transitionToMini = true;
                        oldChain = sectorChain;
                    }

                    // No transition caused by size change

                }
            }


            Queue<Sector> freeList = null;
            StreamView sv = null;

            if (!transitionToMini && !transitionToNormal)   //############  NO TRANSITION
            {
                if (delta > 0) // Enlarging stream...
                {
                    if (_sectorRecycle)
                        freeList = FindFreeSectors(newSectorType); // Collect available free sectors

                    sv = new StreamView(sectorChain, newSectorSize, length, freeList, SourceStream);

                    //Set up  destination chain
                    SetSectorChain(sectorChain);
                }
                else if (delta < 0)  // Reducing size...
                {

                    var nSec = (int)Math.Floor(((double)(Math.Abs(delta)) / newSectorSize)); //number of sectors to mark as free

                    if (newSectorSize == Sector.MinisectorSize)
                        FreeMiniChain(sectorChain, nSec, _eraseFreeSectors);
                    else
                        FreeChain(sectorChain, nSec, _eraseFreeSectors);
                }

                if (sectorChain.Count > 0)
                {
                    cfItem.DirEntry.StartSetc = sectorChain[0].Id;
                    cfItem.DirEntry.Size = length;
                }
                else
                {
                    cfItem.DirEntry.StartSetc = Sector.Endofchain;
                    cfItem.DirEntry.Size = 0;
                }

            }
            else if (transitionToMini)                          //############## TRANSITION TO MINISTREAM
            {
                // Transition Normal chain -> Mini chain

                // Collect available MINI free sectors

                if (_sectorRecycle)
                    freeList = FindFreeSectors(SectorType.Mini);

                sv = new StreamView(oldChain, oldSectorSize, oldSize, null, SourceStream);

                // Reset start sector and size of dir entry
                cfItem.DirEntry.StartSetc = Sector.Endofchain;
                cfItem.DirEntry.Size = 0;

                var newChain = GetMiniSectorChain(Sector.Endofchain);
                var destSv = new StreamView(newChain, Sector.MinisectorSize, length, freeList, SourceStream);

                // Buffered trimmed copy from old (larger) to new (smaller)
                var cnt = 4096 < length ? 4096 : (int)length;

                var buf = new byte[4096];
                var toRead = length;

                //Copy old to new chain
                while (toRead > cnt)
                {
                    cnt = sv.Read(buf, 0, cnt);
                    toRead -= cnt;
                    destSv.Write(buf, 0, cnt);
                }

                sv.Read(buf, 0, (int)toRead);
                destSv.Write(buf, 0, (int)toRead);

                //Free old chain
                FreeChain(oldChain, _eraseFreeSectors);

                //Set up destination chain
                AllocateMiniSectorChain(destSv.BaseSectorChain);

                // Persist to normal strea
                PersistMiniStreamToStream(destSv.BaseSectorChain);

                //Update dir item
                if (destSv.BaseSectorChain.Count > 0)
                {
                    cfItem.DirEntry.StartSetc = destSv.BaseSectorChain[0].Id;
                    cfItem.DirEntry.Size = length;
                }
                else
                {
                    cfItem.DirEntry.StartSetc = Sector.Endofchain;
                    cfItem.DirEntry.Size = 0;
                }
            }
            else if (transitionToNormal)                        //############## TRANSITION TO NORMAL STREAM
            {
                // Transition Mini chain -> Normal chain

                if (_sectorRecycle)
                    freeList = FindFreeSectors(SectorType.Normal); // Collect available Normal free sectors

                sv = new StreamView(oldChain, oldSectorSize, oldSize, null, SourceStream);

                var newChain = GetNormalSectorChain(Sector.Endofchain);
                var destSv = new StreamView(newChain, GetSectorSize(), length, freeList, SourceStream);

                var cnt = 256 < length ? 256 : (int)length;

                var buf = new byte[256];
                var toRead = Math.Min(length, cfItem.Size);

                //Copy old to new chain
                while (toRead > cnt)
                {
                    cnt = sv.Read(buf, 0, cnt);
                    toRead -= cnt;
                    destSv.Write(buf, 0, cnt);
                }

                sv.Read(buf, 0, (int)toRead);
                destSv.Write(buf, 0, (int)toRead);

                //Free old mini chain
                var oldChainCount = oldChain.Count;
                FreeMiniChain(oldChain, _eraseFreeSectors);

                //Set up normal destination chain
                AllocateSectorChain(destSv.BaseSectorChain);

                //Update dir item
                if (destSv.BaseSectorChain.Count > 0)
                {
                    cfItem.DirEntry.StartSetc = destSv.BaseSectorChain[0].Id;
                    cfItem.DirEntry.Size = length;
                }
                else
                {
                    cfItem.DirEntry.StartSetc = Sector.Endofchain;
                    cfItem.DirEntry.Size = 0;
                }
            }
        }

        private void WriteData(CfItem cfItem, long position, byte[] buffer)
        {
            WriteData(cfItem, buffer, position, 0, buffer.Length);
        }

        private void WriteData(CfItem cfItem, long position, ReadOnlySpan<byte> buffer)
        {
            WriteData(cfItem, buffer, position, 0, buffer.Length);
        }

        internal void WriteData(CfItem cfItem, byte[] buffer, long position, int offset, int count)
        {

            if (buffer == null)
                throw new CfInvalidOperation("Parameter [buffer] cannot be null");

            if (cfItem.DirEntry == null)
                throw new CfException("Internal error [cfItem.DirEntry] cannot be null");

            if (buffer.Length == 0) return;

            // Get delta size induced by client
            var delta = (position + count) - cfItem.Size < 0 ? 0 : (position + count) - cfItem.Size;
            var newLength = cfItem.Size + delta;

            SetStreamLength(cfItem, newLength);

            // Calculate NEW sectors SIZE
            var st = SectorType.Normal;
            var sectorSize = GetSectorSize();

            if (cfItem.Size < _header.MinSizeStandardStream)
            {
                st = SectorType.Mini;
                sectorSize = Sector.MinisectorSize;
            }

            var sectorChain = GetSectorChain(cfItem.DirEntry.StartSetc, st);
            var sv = new StreamView(sectorChain, sectorSize, newLength, null, SourceStream);

            sv.Seek(position, SeekOrigin.Begin);
            sv.Write(buffer, offset, count);

            if (cfItem.Size < _header.MinSizeStandardStream)
            {
                PersistMiniStreamToStream(sv.BaseSectorChain);
                //SetSectorChain(sv.BaseSectorChain);
            }
        }
        
        internal void WriteData(CfItem cfItem, ReadOnlySpan<byte> buffer, long position, int offset, int count)
        {

            if (buffer == null)
                throw new CfInvalidOperation("Parameter [buffer] cannot be null");

            if (cfItem.DirEntry == null)
                throw new CfException("Internal error [cfItem.DirEntry] cannot be null");

            if (buffer.Length == 0) return;

            // Get delta size induced by client
            var delta = (position + count) - cfItem.Size < 0 ? 0 : (position + count) - cfItem.Size;
            var newLength = cfItem.Size + delta;

            SetStreamLength(cfItem, newLength);

            // Calculate NEW sectors SIZE
            var st = SectorType.Normal;
            var sectorSize = GetSectorSize();

            if (cfItem.Size < _header.MinSizeStandardStream)
            {
                st = SectorType.Mini;
                sectorSize = Sector.MinisectorSize;
            }

            var sectorChain = GetSectorChain(cfItem.DirEntry.StartSetc, st);
            var sv = new StreamView(sectorChain, sectorSize, newLength, null, SourceStream);

            sv.Seek(position, SeekOrigin.Begin);
            sv.Write(buffer);

            if (cfItem.Size < _header.MinSizeStandardStream)
            {
                PersistMiniStreamToStream(sv.BaseSectorChain);
                //SetSectorChain(sv.BaseSectorChain);
            }
        }

        internal void WriteData(CfItem cfItem, byte[] buffer)
        {
            WriteData(cfItem, 0, buffer);
        }
        
        internal void WriteData(CfItem cfItem, ReadOnlySpan<byte> buffer)
        {
            WriteData(cfItem, 0, buffer);
        }

        /// <summary>
        /// Check file size limit ( 2GB for version 3 )
        /// </summary>
        private void CheckFileLength()
        {
            throw new NotImplementedException();
        }


        internal int ReadData(CfStream cFStream, long position, byte[] buffer, int count)
        {
            if (count > buffer.Length)
                throw new ArgumentException("count parameter exceeds buffer size");

            var de = cFStream.DirEntry;

            count = (int)Math.Min(de.Size - position, count);

            StreamView sView = null;


            if (de.Size < _header.MinSizeStandardStream)
            {
                sView
                    = new StreamView(GetSectorChain(de.StartSetc, SectorType.Mini), Sector.MinisectorSize, de.Size, null, SourceStream);
            }
            else
            {

                sView = new StreamView(GetSectorChain(de.StartSetc, SectorType.Normal), GetSectorSize(), de.Size, null, SourceStream);
            }


            sView.Seek(position, SeekOrigin.Begin);
            var result = sView.Read(buffer, 0, count);

            return result;
        }

        internal int ReadData(CfStream cFStream, long position, byte[] buffer, int offset, int count)
        {

            var de = cFStream.DirEntry;

            count = (int)Math.Min(buffer.Length - offset, (long)count);

            StreamView sView = null;


            if (de.Size < _header.MinSizeStandardStream)
            {
                sView
                    = new StreamView(GetSectorChain(de.StartSetc, SectorType.Mini), Sector.MinisectorSize, de.Size, null, SourceStream);
            }
            else
            {

                sView = new StreamView(GetSectorChain(de.StartSetc, SectorType.Normal), GetSectorSize(), de.Size, null, SourceStream);
            }


            sView.Seek(position, SeekOrigin.Begin);
            var result = sView.Read(buffer, offset, count);

            return result;
        }


        internal byte[] GetData(CfStream cFStream)
        {

            if (_disposed)
                throw new CfDisposedException("Compound File closed: cannot access data");

            byte[] result = null;

            var de = cFStream.DirEntry;

            //IDirectoryEntry root = directoryEntries[0];

            if (de.Size < _header.MinSizeStandardStream)
            {

                var miniView
                    = new StreamView(GetSectorChain(de.StartSetc, SectorType.Mini), Sector.MinisectorSize, de.Size, null, SourceStream);

                var br = new BinaryReader(miniView);

                result = br.ReadBytes((int)de.Size);
                br.Close();

            }
            else
            {
                var sView
                    = new StreamView(GetSectorChain(de.StartSetc, SectorType.Normal), GetSectorSize(), de.Size, null, SourceStream);

                result = new byte[(int)de.Size];

                sView.Read(result, 0, result.Length);

            }

            return result;
        }
        public byte[] GetDataBySid(int sid)
        {
            if (_disposed)
                throw new CfDisposedException("Compound File closed: cannot access data");
            if (sid < 0)
                return null;
            byte[] result = null;
            try
            {
                var de = _directoryEntries[sid];
                if (de.Size < _header.MinSizeStandardStream)
                {
                    var miniView
                        = new StreamView(GetSectorChain(de.StartSetc, SectorType.Mini), Sector.MinisectorSize, de.Size, null, SourceStream);
                    var br = new BinaryReader(miniView);
                    result = br.ReadBytes((int)de.Size);
                    br.Close();
                }
                else
                {
                    var sView
                        = new StreamView(GetSectorChain(de.StartSetc, SectorType.Normal), GetSectorSize(), de.Size, null, SourceStream);
                    result = new byte[(int)de.Size];
                    sView.Read(result, 0, result.Length);
                }
            }
            catch
            {
                throw new CfException("Cannot get data for SID");
            }
            return result;
        }
        public Guid GetGuidBySid(int sid)
        {
            if (_disposed)
                throw new CfDisposedException("Compound File closed: cannot access data");
            if (sid < 0)
                throw new CfException("Invalid SID");
            var de = _directoryEntries[sid];
            return de.StorageClsid;
        }
        public Guid GetGuidForStream(int sid)
        {
            if (_disposed)
                throw new CfDisposedException("Compound File closed: cannot access data");
            if (sid < 0)
                throw new CfException("Invalid SID");
            var g = Guid.Empty;
            //find first storage containing a non-zero CLSID before SID in directory structure
            for (var i = sid - 1; i >= 0; i--)
            {
                if (_directoryEntries[i].StorageClsid != g && _directoryEntries[i].StgType == StgType.StgStorage)
                {
                    return _directoryEntries[i].StorageClsid;
                }
            }
            return g;
        }

        private static int Ceiling(double d)
        {
            return (int)Math.Ceiling(d);
        }

        private static int LowSaturation(int i)
        {
            return i > 0 ? i : 0;
        }


        internal void InvalidateDirectoryEntry(int sid)
        {
            if (sid >= _directoryEntries.Count)
                throw new CfException("Invalid SID of the directory entry to remove");

            ResetDirectoryEntry(sid);
        }

        internal void FreeAssociatedData(int sid)
        {
            // Clear the associated stream (or ministream) if required
            if (_directoryEntries[sid].Size > 0) //thanks to Mark Bosold for this !
            {
                if (_directoryEntries[sid].Size < _header.MinSizeStandardStream)
                {
                    var miniChain
                        = GetSectorChain(_directoryEntries[sid].StartSetc, SectorType.Mini);
                    FreeMiniChain(miniChain, _eraseFreeSectors);
                }
                else
                {
                    var chain
                        = GetSectorChain(_directoryEntries[sid].StartSetc, SectorType.Normal);
                    FreeChain(chain, _eraseFreeSectors);
                }
            }
        }

        private bool _closeStream = true;

        [Obsolete("Use flag LeaveOpen in CompoundFile constructor")]
        public void Close(bool closeStream = true)
        {
            _closeStream = closeStream;
            ((IDisposable)this).Dispose();
        }

        #region IDisposable Members

        private bool _disposed;//false

        void IDisposable.Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        #endregion

        private readonly object _lockObject = new object();

        /// <summary>
        /// When called from user code, release all resources, otherwise, in the case runtime called it,
        /// only unmanagd resources are released.
        /// </summary>
        /// <param name="disposing">If true, method has been called from User code, if false it's been called from .net runtime</param>
        private void Dispose(bool disposing)
        {
            try
            {
                if (!_disposed)
                {
                    lock (_lockObject)
                    {
                        if (disposing)
                        {
                            // Call from user code...

                            if (_sectors != null)
                            {
                                _sectors.Clear();
                                _sectors = null;
                            }

                            _rootStorage = null; // Some problem releasing resources...
                            _header = null;
                            _directoryEntries.Clear();
                            _directoryEntries = null;
                            _fileName = null;
                            //this.lockObject = null;
#if !FLAT_WRITE
                            this.buffer = null;
#endif
                        }

                        if (SourceStream != null && _closeStream && !Configuration.HasFlag(CfsConfiguration.LeaveOpen))
                            SourceStream.Close();
                    }
                }
            }
            finally
            {
                _disposed = true;
            }

        }

        internal bool IsClosed => _disposed;

        private List<IDirectoryEntry> _directoryEntries
            = new List<IDirectoryEntry>();

        internal IList<IDirectoryEntry> GetDirectories()
        {
            return _directoryEntries;
        }

        //internal List<IDirectoryEntry> DirectoryEntries
        //{
        //    get { return directoryEntries; }
        //}


        internal IDirectoryEntry RootEntry => _directoryEntries[0];

        private IList<IDirectoryEntry> FindDirectoryEntries(string entryName)
        {
            var result = new List<IDirectoryEntry>();

            foreach (var d in _directoryEntries)
            {
                if (d.GetEntryName() == entryName && d.StgType != StgType.StgInvalid)
                    result.Add(d);
            }

            return result;
        }



        /// <summary>
        /// Get a list of all entries with a given name contained in the document.
        /// </summary>
        /// <param name="entryName">Name of entries to retrive</param>
        /// <returns>A list of name-matching entries</returns>
        /// <remarks>This function is aimed to speed up entity lookup in 
        /// flat-structure files (only one or little more known entries)
        /// without the performance penalty related to entities hierarchy constraints.
        /// There is no implied hierarchy in the returned list.
        /// </remarks>
        public IList<CfItem> GetAllNamedEntries(string entryName)
        {
            var r = FindDirectoryEntries(entryName);
            var result = new List<CfItem>();

            foreach (var id in r)
            {
                if (id.GetEntryName() == entryName && id.StgType != StgType.StgInvalid)
                {
                    var i = id.StgType == StgType.StgStorage ? new CfStorage(this, id) : (CfItem)new CfStream(this, id);
                    result.Add(i);
                }
            }

            return result;
        }

        public int GetNumDirectories()
        {
            if (_disposed)
                throw new CfDisposedException("Compound File closed: cannot access data");
            return _directoryEntries.Count;
        }

        public string GetNameDirEntry(int id)
        {
            if (_disposed)
                throw new CfDisposedException("Compound File closed: cannot access data");
            if (id < 0)
                throw new CfException("Invalid Storage ID");
            return _directoryEntries[id].Name;
        }

        public StgType GetStorageType(int id)
        {
            if (_disposed)
                throw new CfDisposedException("Compound File closed: cannot access data");
            if (id < 0)
                throw new CfException("Invalid Storage ID");
            return _directoryEntries[id].StgType;
        }


        /// <summary>
        /// Compress free space by removing unallocated sectors from compound file
        /// effectively reducing stream or file size.
        /// </summary>
        /// <remarks>
        /// Current implementation supports compression only for ver. 3 compound files.
        /// </remarks>
        /// <example>
        /// <code>
        /// 
        ///  //This code has been extracted from unit test
        ///  
        ///    String FILENAME = "MultipleStorage3.cfs";
        ///
        ///    FileInfo srcFile = new FileInfo(FILENAME);
        ///
        ///    File.Copy(FILENAME, "MultipleStorage_Deleted_Compress.cfs", true);
        ///
        ///    CompoundFile cf = new CompoundFile("MultipleStorage_Deleted_Compress.cfs", UpdateMode.Update, true, true);
        ///
        ///    CFStorage st = cf.RootStorage.GetStorage("MyStorage");
        ///    st = st.GetStorage("AnotherStorage");
        ///    
        ///    Assert.IsNotNull(st);
        ///    st.Delete("Another2Stream"); //17Kb
        ///    cf.Commit();
        ///    cf.Close();
        ///
        ///    CompoundFile.ShrinkCompoundFile("MultipleStorage_Deleted_Compress.cfs");
        ///
        ///    FileInfo dstFile = new FileInfo("MultipleStorage_Deleted_Compress.cfs");
        ///
        ///    Assert.IsTrue(srcFile.Length > dstFile.Length);
        ///
        /// </code>
        /// </example>
        private static void ShrinkCompoundFile(Stream s)
        {
            var cf = new CompoundFile(s, CfsUpdateMode.ReadOnly, CfsConfiguration.LeaveOpen);

            if (cf._header.MajorVersion != (ushort)CfsVersion.Ver3)
                throw new CfException(
                    "Current implementation of free space compression does not support version 4 of Compound File Format");

            using var tempCf = new CompoundFile((CfsVersion)cf._header.MajorVersion, CfsConfiguration.Default);
            //Copy Root CLSID
            tempCf.RootStorage.Clsid = new Guid(cf.RootStorage.Clsid.ToByteArray());

            DoCompression(cf.RootStorage, tempCf.RootStorage);

            var tmpMs = new MemoryStream((int)cf.SourceStream.Length); //This could be a problem for v4

            tempCf.Save(tmpMs);

            // If we were based on a writable stream, we update
            // the stream and do reload from the compressed one...

            s.Seek(0, SeekOrigin.Begin);
            tmpMs.WriteTo(s);

            s.Seek(0, SeekOrigin.Begin);
            s.SetLength(tmpMs.Length);

            tmpMs.Close();
        }

        /// <summary>
        /// Remove unallocated sectors from compound file in order to reduce its size.
        /// </summary>
        /// <remarks>
        /// Current implementation supports compression only for ver. 3 compound files.
        /// </remarks>
        /// <example>
        /// <code>
        /// 
        ///  //This code has been extracted from unit test
        ///  
        ///    String FILENAME = "MultipleStorage3.cfs";
        ///
        ///    FileInfo srcFile = new FileInfo(FILENAME);
        ///
        ///    File.Copy(FILENAME, "MultipleStorage_Deleted_Compress.cfs", true);
        ///
        ///    CompoundFile cf = new CompoundFile("MultipleStorage_Deleted_Compress.cfs", UpdateMode.Update, true, true);
        ///
        ///    CFStorage st = cf.RootStorage.GetStorage("MyStorage");
        ///    st = st.GetStorage("AnotherStorage");
        ///    
        ///    Assert.IsNotNull(st);
        ///    st.Delete("Another2Stream"); //17Kb
        ///    cf.Commit();
        ///    cf.Close();
        ///
        ///    CompoundFile.ShrinkCompoundFile("MultipleStorage_Deleted_Compress.cfs");
        ///
        ///    FileInfo dstFile = new FileInfo("MultipleStorage_Deleted_Compress.cfs");
        ///
        ///    Assert.IsTrue(srcFile.Length > dstFile.Length);
        ///
        /// </code>
        /// </example>
        public static void ShrinkCompoundFile(string fileName)
        {
            var fs = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite);
            ShrinkCompoundFile(fs);
            fs.Close();
        }

        /// <summary>
        /// Recursively clones valid structures, avoiding to copy free sectors.
        /// </summary>
        /// <param name="currSrcStorage">Current source storage to clone</param>
        /// <param name="currDstStorage">Current cloned destination storage</param>
        private static void DoCompression(CfStorage currSrcStorage, CfStorage currDstStorage)
        {
            void Va(CfItem item)
            {
                if (item.IsStream)
                {
                    var itemAsStream = item as CfStream;
                    var st = currDstStorage.AddStream(itemAsStream.Name);
                    st.SetData(itemAsStream.GetData());
                }
                else if (item.IsStorage)
                {
                    var itemAsStorage = item as CfStorage;
                    var strg = currDstStorage.AddStorage(itemAsStorage.Name);
                    strg.Clsid = new Guid(itemAsStorage.Clsid.ToByteArray());
                    DoCompression(itemAsStorage, strg); // recursion, one level deeper
                }
            }

            currSrcStorage.VisitEntries(Va, false);
        }
    }
}
