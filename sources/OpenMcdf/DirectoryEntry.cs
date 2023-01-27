/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 * 
 * The Original Code is OpenMCDF - Compound Document Format library.
 * 
 * The Initial Developer of the Original Code is Federico Blaseotto.*/

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using RedBlackTree;

namespace OpenMcdf
{
    public enum StgType
    {
        StgInvalid = 0,
        StgStorage = 1,
        StgStream = 2,
        StgLockbytes = 3,
        StgProperty = 4,
        StgRoot = 5
    }

    public enum StgColor
    {
        Red = 0,
        Black = 1
    }

    internal class DirectoryEntry : IDirectoryEntry
    {
        internal const int ThisIsGreater = 1;
        internal const int OtherIsGreater = -1;
        private readonly IList<IDirectoryEntry> _dirRepository;

        private int _sid = -1;
        public int Sid
        {
            get => _sid;
            set => _sid = value;
        }

        internal static int Nostream
            = unchecked((int)0xFFFFFFFF);

        internal static int Zero
            = 0;

        private DirectoryEntry(string name, StgType stgType, IList<IDirectoryEntry> dirRepository)
        {
            _dirRepository = dirRepository;

            _stgType = stgType;

            if (stgType == StgType.StgStorage)
            {
                _creationDate = BitConverter.GetBytes((DateTime.Now.ToFileTime()));
                StartSetc = Zero;
            }

            if (stgType == StgType.StgInvalid)
            {
                StartSetc = Zero;
            }

            if (name != string.Empty)
            {
                SetEntryName(name);
            }
        }

        private byte[] _entryName = new byte[64];

        public byte[] EntryName => _entryName;

        //set
        //{
        //    entryName = value;
        //}
        public string GetEntryName()
        {
            if (_entryName != null && _entryName.Length > 0)
            {
                return Encoding.Unicode.GetString(_entryName).Remove((NameLength - 1) / 2);
            }

            return string.Empty;
        }

        public void SetEntryName(string entryName)
        {
            if (entryName == string.Empty)
            {
                _entryName = new byte[64];
                NameLength = 0;
            }
            else
            {
                if (
                    entryName.Contains(@"\") ||
                    entryName.Contains(@"/") ||
                    entryName.Contains(@":") ||
                    entryName.Contains(@"!")

                    )
                    throw new CfException("Invalid character in entry: the characters '\\', '/', ':','!' cannot be used in entry name");

                if (entryName.Length > 31)
                    throw new CfException("Entry name MUST NOT exceed 31 characters");



                byte[] newName = null;
                var temp = Encoding.Unicode.GetBytes(entryName);
                newName = new byte[64];
                Buffer.BlockCopy(temp, 0, newName, 0, temp.Length);
                newName[temp.Length] = 0x00;
                newName[temp.Length + 1] = 0x00;

                _entryName = newName;
                NameLength = (ushort)(temp.Length + 2);
            }
        }

        public ushort NameLength { get; set; }

        private StgType _stgType = StgType.StgInvalid;
        public StgType StgType
        {
            get => _stgType;
            set => _stgType = value;
        }
        private StgColor _stgColor = StgColor.Red;

        public StgColor StgColor
        {
            get => _stgColor;
            set => _stgColor = value;
        }

        private int _leftSibling = Nostream;
        public int LeftSibling
        {
            get => _leftSibling;
            set => _leftSibling = value;
        }

        private int _rightSibling = Nostream;
        public int RightSibling
        {
            get => _rightSibling;
            set => _rightSibling = value;
        }

        private int _child = Nostream;
        public int Child
        {
            get => _child;
            set => _child = value;
        }

        private Guid _storageClsid
            = Guid.Empty;

        public Guid StorageClsid
        {
            get => _storageClsid;
            set => _storageClsid = value;
        }


        private int _stateBits;

        public int StateBits
        {
            get => _stateBits;
            set => _stateBits = value;
        }

        private byte[] _creationDate = new byte[8] { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };

        public byte[] CreationDate
        {
            get => _creationDate;
            set => _creationDate = value;
        }

        private byte[] _modifyDate = new byte[8] { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };

        public byte[] ModifyDate
        {
            get => _modifyDate;
            set => _modifyDate = value;
        }

        private int _startSetc = Sector.Endofchain;
        public int StartSetc
        {
            get => _startSetc;
            set => _startSetc = value;
        }
        private long _size;
        public long Size
        {
            get => _size;
            set => _size = value;
        }


        public int CompareTo(object obj)
        {

            var otherDir = obj as IDirectoryEntry;

            if (otherDir == null)
                throw new CfException("Invalid casting: compared object does not implement IDirectorEntry interface");

            if (NameLength > otherDir.NameLength)
            {
                return ThisIsGreater;
            }

            if (NameLength < otherDir.NameLength)
            {
                return OtherIsGreater;
            }

            var thisName = Encoding.Unicode.GetString(EntryName, 0, NameLength);
            var otherName = Encoding.Unicode.GetString(otherDir.EntryName, 0, otherDir.NameLength);

            for (var z = 0; z < thisName.Length; z++)
            {
                var thisChar = char.ToUpperInvariant(thisName[z]);
                var otherChar = char.ToUpperInvariant(otherName[z]);

                if (thisChar > otherChar)
                    return ThisIsGreater;
                if (thisChar < otherChar)
                    return OtherIsGreater;
            }

            return 0;

            //   return String.Compare(Encoding.Unicode.GetString(this.EntryName).ToUpper(), Encoding.Unicode.GetString(other.EntryName).ToUpper());
        }

        public override bool Equals(object obj)
        {
            return CompareTo(obj) == 0;
        }

        /// <summary>
        /// FNV hash, short for Fowler/Noll/Vo
        /// </summary>
        /// <param name="buffer"></param>
        /// <returns>(not warranted) unique hash for byte array</returns>
        private static ulong fnv_hash(byte[] buffer)
        {

            ulong h = 2166136261;
            int i;

            for (i = 0; i < buffer.Length; i++)
                h = (h * 16777619) ^ buffer[i];

            return h;
        }

        public override int GetHashCode()
        {
            return (int)fnv_hash(_entryName);
        }

        public void Write(Stream stream)
        {
            var rw = new StreamRw(stream);

            rw.Write(_entryName);
            rw.Write(NameLength);
            rw.Write((byte)_stgType);
            rw.Write((byte)_stgColor);
            rw.Write(_leftSibling);
            rw.Write(_rightSibling);
            rw.Write(_child);
            rw.Write(_storageClsid.ToByteArray());
            rw.Write(_stateBits);
            rw.Write(_creationDate);
            rw.Write(_modifyDate);
            rw.Write(_startSetc);
            rw.Write(_size);

            rw.Close();
        }

        //public Byte[] ToByteArray()
        //{
        //    MemoryStream ms
        //        = new MemoryStream(128);

        //    BinaryWriter bw = new BinaryWriter(ms);

        //    byte[] paddedName = new byte[64];
        //    Array.Copy(entryName, paddedName, entryName.Length);

        //    bw.Write(paddedName);
        //    bw.Write(nameLength);
        //    bw.Write((byte)stgType);
        //    bw.Write((byte)stgColor);
        //    bw.Write(leftSibling);
        //    bw.Write(rightSibling);
        //    bw.Write(child);
        //    bw.Write(storageCLSID.ToByteArray());
        //    bw.Write(stateBits);
        //    bw.Write(creationDate);
        //    bw.Write(modifyDate);
        //    bw.Write(startSetc);
        //    bw.Write(size);

        //    return ms.ToArray();
        //}

        public void Read(Stream stream, CfsVersion ver = CfsVersion.Ver3)
        {
            var rw = new StreamRw(stream);

            _entryName = rw.ReadBytes(64);
            NameLength = rw.ReadUInt16();
            _stgType = (StgType)rw.ReadByte();
            //rw.ReadByte();//Ignore color, only black tree
            _stgColor = (StgColor)rw.ReadByte();
            _leftSibling = rw.ReadInt32();
            _rightSibling = rw.ReadInt32();
            _child = rw.ReadInt32();

            // Thanks to bugaccount (BugTrack id 3519554)
            if (_stgType == StgType.StgInvalid)
            {
                _leftSibling = Nostream;
                _rightSibling = Nostream;
                _child = Nostream;
            }

            _storageClsid = new Guid(rw.ReadBytes(16));
            _stateBits = rw.ReadInt32();
            _creationDate = rw.ReadBytes(8);
            _modifyDate = rw.ReadBytes(8);
            _startSetc = rw.ReadInt32();

            if (ver == CfsVersion.Ver3)
            {
                // avoid dirty read for version 3 files (max size: 32bit integer)
                // where most significant bits are not initialized to zero

                _size = rw.ReadInt32();
                rw.ReadBytes(4); //discard most significant 4 (possibly) dirty bytes
            }
            else
            {
                _size = rw.ReadInt64();
            }
        }

        public string Name => GetEntryName();


        public IRbNode Left
        {
            get
            {
                if (_leftSibling == Nostream)
                    return null;

                return _dirRepository[_leftSibling];
            }
            set
            {
                _leftSibling = value != null ? ((IDirectoryEntry)value).Sid : Nostream;

                if (_leftSibling != Nostream)
                    _dirRepository[_leftSibling].Parent = this;
            }
        }

        public IRbNode Right
        {
            get
            {
                if (_rightSibling == Nostream)
                    return null;

                return _dirRepository[_rightSibling];
            }
            set
            {

                _rightSibling = value != null ? ((IDirectoryEntry)value).Sid : Nostream;

                if (_rightSibling != Nostream)
                    _dirRepository[_rightSibling].Parent = this;

            }
        }

        public Color Color
        {
            get => (Color)StgColor;
            set => StgColor = (StgColor)value;
        }

        private IDirectoryEntry _parent;

        public IRbNode Parent
        {
            get => _parent;
            set => _parent = value as IDirectoryEntry;
        }

        public IRbNode Grandparent()
        {
            return _parent != null ? _parent.Parent : null;
        }

        public IRbNode Sibling()
        {
            if (this == Parent.Left)
                return Parent.Right;
            return Parent.Left;
        }

        public IRbNode Uncle()
        {
            return _parent != null ? Parent.Sibling() : null;
        }

        internal static IDirectoryEntry New(string name, StgType stgType, IList<IDirectoryEntry> dirRepository)
        {
            DirectoryEntry de = null;
            if (dirRepository != null)
            {
                de = new DirectoryEntry(name, stgType, dirRepository);
                // No invalid directory entry found
                dirRepository.Add(de);
                de.Sid = dirRepository.Count - 1;
            }
            else
                throw new ArgumentNullException("dirRepository", "Directory repository cannot be null in New() method");

            return de;
        }

        internal static IDirectoryEntry Mock(string name, StgType stgType)
        {
            var de = new DirectoryEntry(name, stgType, null);

            return de;
        }

        internal static IDirectoryEntry TryNew(string name, StgType stgType, IList<IDirectoryEntry> dirRepository)
        {
            var de = new DirectoryEntry(name, stgType, dirRepository);

            // If we are not adding an invalid dirEntry as
            // in a normal loading from file (invalid dirs MAY pad a sector)
            if (de != null)
            {
                // Find first available invalid slot (if any) to reuse it
                for (var i = 0; i < dirRepository.Count; i++)
                {
                    if (dirRepository[i].StgType == StgType.StgInvalid)
                    {
                        dirRepository[i] = de;
                        de.Sid = i;
                        return de;
                    }
                }
            }

            // No invalid directory entry found
            dirRepository.Add(de);
            de.Sid = dirRepository.Count - 1;

            return de;
        }




        public override string ToString()
        {
            return Name + " [" + _sid + "]" + (_stgType == StgType.StgStream ? "Stream" : "Storage");
        }


        public void AssignValueTo(IRbNode other)
        {
            var d = other as DirectoryEntry;

            d.SetEntryName(GetEntryName());

            d._creationDate = new byte[_creationDate.Length];
            _creationDate.CopyTo(d._creationDate, 0);

            d._modifyDate = new byte[_modifyDate.Length];
            _modifyDate.CopyTo(d._modifyDate, 0);

            d._size = _size;
            d._startSetc = _startSetc;
            d._stateBits = _stateBits;
            d._stgType = _stgType;
            d._storageClsid = new Guid(_storageClsid.ToByteArray());
            d.Child = Child;
        }
    }
}
