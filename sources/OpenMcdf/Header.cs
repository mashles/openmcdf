/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 * 
 * The Original Code is OpenMCDF - Compound Document Format library.
 * 
 * The Initial Developer of the Original Code is Federico Blaseotto.*/


using System.IO;

namespace OpenMcdf
{
    internal class Header
    {
        //0 8 Compound document file identifier: D0H CFH 11H E0H A1H B1H 1AH E1H
        private byte[] _headerSignature
            = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };

        public byte[] HeaderSignature => _headerSignature;

        //8 16 Unique identifier (UID) of this file (not of interest in the following, may be all 0)
        private byte[] _clsid = new byte[16];

        public byte[] Clsid
        {
            get => _clsid;
            set => _clsid = value;
        }

        //24 2 Revision number of the file format (most used is 003EH)
        private ushort _minorVersion = 0x003E;

        public ushort MinorVersion => _minorVersion;

        //26 2 Version number of the file format (most used is 0003H)
        private ushort _majorVersion = 0x0003;

        public ushort MajorVersion => _majorVersion;

        //28 2 Byte order identifier (➜4.2): FEH FFH = Little-Endian FFH FEH = Big-Endian
        private ushort _byteOrder = 0xFFFE;

        public ushort ByteOrder => _byteOrder;

        //30 2 Size of a sector in the compound document file (➜3.1) in power-of-two (ssz), real sector
        //size is sec_size = 2ssz bytes (minimum value is 7 which means 128 bytes, most used 
        //value is 9 which means 512 bytes)
        private ushort _sectorShift = 9;

        public ushort SectorShift => _sectorShift;

        //32 2 Size of a short-sector in the short-stream container stream (➜6.1) in power-of-two (sssz),
        //real short-sector size is short_sec_size = 2sssz bytes (maximum value is sector size
        //ssz, see above, most used value is 6 which means 64 bytes)
        private ushort _miniSectorShift = 6;
        public ushort MiniSectorShift => _miniSectorShift;

        //34 10 Not used
        private byte[] _unUsed = new byte[6];

        public byte[] UnUsed => _unUsed;

        //44 4 Total number of sectors used Directory (➜5.2)
        private int _directorySectorsNumber;

        public int DirectorySectorsNumber
        {
            get => _directorySectorsNumber;
            set => _directorySectorsNumber = value;
        }

        //44 4 Total number of sectors used for the sector allocation table (➜5.2)
        private int _fatSectorsNumber;

        public int FatSectorsNumber
        {
            get => _fatSectorsNumber;
            set => _fatSectorsNumber = value;
        }

        //48 4 SecID of first sector of the directory stream (➜7)
        private int _firstDirectorySectorId = Sector.Endofchain;

        public int FirstDirectorySectorId
        {
            get => _firstDirectorySectorId;
            set => _firstDirectorySectorId = value;
        }

        //52 4 Not used
        private uint _unUsed2;

        public uint UnUsed2 => _unUsed2;

        //56 4 Minimum size of a standard stream (in bytes, minimum allowed and most used size is 4096
        //bytes), streams with an actual size smaller than (and not equal to) this value are stored as
        //short-streams (➜6)
        private uint _minSizeStandardStream = 4096;

        public uint MinSizeStandardStream
        {
            get => _minSizeStandardStream;
            set => _minSizeStandardStream = value;
        }

        //60 4 SecID of first sector of the short-sector allocation table (➜6.2), or –2 (End Of Chain
        //SecID, ➜3.1) if not extant
        private int _firstMiniFatSectorId = unchecked((int)0xFFFFFFFE);

        /// <summary>
        /// This integer field contains the starting sector number for the mini FAT
        /// </summary>
        public int FirstMiniFatSectorId
        {
            get => _firstMiniFatSectorId;
            set => _firstMiniFatSectorId = value;
        }

        //64 4 Total number of sectors used for the short-sector allocation table (➜6.2)
        private uint _miniFatSectorsNumber;

        public uint MiniFatSectorsNumber
        {
            get => _miniFatSectorsNumber;
            set => _miniFatSectorsNumber = value;
        }

        //68 4 SecID of first sector of the master sector allocation table (➜5.1), or –2 (End Of Chain
        //SecID, ➜3.1) if no additional sectors used
        private int _firstDifatSectorId = Sector.Endofchain;

        public int FirstDifatSectorId
        {
            get => _firstDifatSectorId;
            set => _firstDifatSectorId = value;
        }

        //72 4 Total number of sectors used for the master sector allocation table (➜5.1)
        private uint _difatSectorsNumber;

        public uint DifatSectorsNumber
        {
            get => _difatSectorsNumber;
            set => _difatSectorsNumber = value;
        }

        //76 436 First part of the master sector allocation table (➜5.1) containing 109 SecIDs
        private readonly int[] _difat = new int[109];

        public int[] Difat => _difat;


        public Header()
            : this(3)
        {

        }


        public Header(ushort version)
        {

            switch (version)
            {
                case 3:
                    _majorVersion = 3;
                    _sectorShift = 0x0009;
                    break;

                case 4:
                    _majorVersion = 4;
                    _sectorShift = 0x000C;
                    break;

                default:
                    throw new CfException("Invalid Compound File Format version");


            }

            for (var i = 0; i < 109; i++)
            {
                _difat[i] = Sector.Freesect;
            }


        }

        public void Write(Stream stream)
        {
            var rw = new StreamRw(stream);

            rw.Write(_headerSignature);
            rw.Write(_clsid);
            rw.Write(_minorVersion);
            rw.Write(_majorVersion);
            rw.Write(_byteOrder);
            rw.Write(_sectorShift);
            rw.Write(_miniSectorShift);
            rw.Write(_unUsed);
            rw.Write(_directorySectorsNumber);
            rw.Write(_fatSectorsNumber);
            rw.Write(_firstDirectorySectorId);
            rw.Write(_unUsed2);
            rw.Write(_minSizeStandardStream);
            rw.Write(_firstMiniFatSectorId);
            rw.Write(_miniFatSectorsNumber);
            rw.Write(_firstDifatSectorId);
            rw.Write(_difatSectorsNumber);

            foreach (var i in _difat)
            {
                rw.Write(i);
            }

            if (_majorVersion == 4)
            {
                var zeroHead = new byte[3584];
                rw.Write(zeroHead);
            }

            rw.Close();
        }

        public void Read(Stream stream)
        {
            var rw = new StreamRw(stream);

            _headerSignature = rw.ReadBytes(8);
            CheckSignature();
            _clsid = rw.ReadBytes(16);
            _minorVersion = rw.ReadUInt16();
            _majorVersion = rw.ReadUInt16();
            CheckVersion();
            _byteOrder = rw.ReadUInt16();
            _sectorShift = rw.ReadUInt16();
            _miniSectorShift = rw.ReadUInt16();
            _unUsed = rw.ReadBytes(6);
            _directorySectorsNumber = rw.ReadInt32();
            _fatSectorsNumber = rw.ReadInt32();
            _firstDirectorySectorId = rw.ReadInt32();
            _unUsed2 = rw.ReadUInt32();
            _minSizeStandardStream = rw.ReadUInt32();
            _firstMiniFatSectorId = rw.ReadInt32();
            _miniFatSectorsNumber = rw.ReadUInt32();
            _firstDifatSectorId = rw.ReadInt32();
            _difatSectorsNumber = rw.ReadUInt32();

            for (var i = 0; i < 109; i++)
            {
                Difat[i] = rw.ReadInt32();
            }

            rw.Close();
        }


        private void CheckVersion()
        {
            if (_majorVersion != 3 && _majorVersion != 4)
                throw new CfFileFormatException("Unsupported Binary File Format version: OpenMcdf only supports Compound Files with major version equal to 3 or 4 ");
        }

        /// <summary>
        /// Structured Storage signature
        /// </summary>
        private readonly byte[] _oleCfsSignature = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };

        private void CheckSignature()
        {
            for (var i = 0; i < _headerSignature.Length; i++)
            {
                if (_headerSignature[i] != _oleCfsSignature[i])
                    throw new CfFileFormatException("Invalid OLE structured storage file");
            }
        }
    }
}
