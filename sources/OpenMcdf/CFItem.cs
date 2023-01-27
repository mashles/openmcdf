
/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 * 
 * The Original Code is OpenMCDF - Compound Document Format library.
 * 
 * The Initial Developer of the Original Code is Federico Blaseotto.*/

using System;

namespace OpenMcdf
{
    /// <summary>
    /// Abstract base class for Structured Storage entities.
    /// </summary>
    /// <example>
    /// <code>
    /// 
    /// const String STORAGE_NAME = "report.xls";
    /// CompoundFile cf = new CompoundFile(STORAGE_NAME);
    ///
    /// FileStream output = new FileStream("LogEntries.txt", FileMode.Create);
    /// TextWriter tw = new StreamWriter(output);
    ///
    /// // CFItem represents both storage and stream items
    /// VisitedEntryAction va = delegate(CFItem item)
    /// {
    ///      tw.WriteLine(item.Name);
    /// };
    ///
    /// cf.RootStorage.VisitEntries(va, true);
    ///
    /// tw.Close();
    /// 
    /// </code>
    /// </example>
    public abstract class CfItem : IComparable<CfItem>
    {
        private readonly CompoundFile _compoundFile;

        protected CompoundFile CompoundFile => _compoundFile;

        protected void CheckDisposed()
        {
            if (_compoundFile.IsClosed)
                throw new CfDisposedException("Owner Compound file has been closed and owned items have been invalidated");
        }

        protected CfItem()
        {
        }

        protected CfItem(CompoundFile compoundFile)
        {
            _compoundFile = compoundFile;
        }

        #region IDirectoryEntry Members

        private IDirectoryEntry _dirEntry;

        internal IDirectoryEntry DirEntry
        {
            get => _dirEntry;
            set => _dirEntry = value;
        }



        internal int CompareTo(CfItem other)
        {

            return _dirEntry.CompareTo(other.DirEntry);
        }


        #endregion

        #region IComparable Members

        public int CompareTo(object obj)
        {
            return _dirEntry.CompareTo(((CfItem)obj).DirEntry);
        }

        #endregion

        public static bool operator ==(CfItem leftItem, CfItem rightItem)
        {
            // If both are null, or both are same instance, return true.
            if (ReferenceEquals(leftItem, rightItem))
            {
                return true;
            }

            // If one is null, but not both, return false.
            if (((object)leftItem == null) || ((object)rightItem == null))
            {
                return false;
            }

            // Return true if the fields match:
            return leftItem.CompareTo(rightItem) == 0;
        }

        public static bool operator !=(CfItem leftItem, CfItem rightItem)
        {
            return !(leftItem == rightItem);
        }

        public override bool Equals(object obj)
        {
            return CompareTo(obj) == 0;
        }

        public override int GetHashCode()
        {
            return _dirEntry.GetEntryName().GetHashCode();
        }

        /// <summary>
        /// Get entity name
        /// </summary>
        public string Name
        {
            get
            {
                var n = _dirEntry.GetEntryName();
                if (n != null && n.Length > 0)
                {
                    return n.TrimEnd('\0');
                }

                return string.Empty;
            }

            
        }

        /// <summary>
        /// Size in bytes of the item. It has a valid value 
        /// only if entity is a stream, otherwise it is setted to zero.
        /// </summary>
        public long Size => _dirEntry.Size;


        /// <summary>
        /// Return true if item is Storage
        /// </summary>
        /// <remarks>
        /// This check doesn't use reflection or runtime type information
        /// and doesn't suffer related performance penalties.
        /// </remarks>
        public bool IsStorage => _dirEntry.StgType == StgType.StgStorage;

        /// <summary>
        /// Return true if item is a Stream
        /// </summary>
        /// <remarks>
        /// This check doesn't use reflection or runtime type information
        /// and doesn't suffer related performance penalties.
        /// </remarks>
        public bool IsStream => _dirEntry.StgType == StgType.StgStream;

        /// <summary>
        /// Return true if item is the Root Storage
        /// </summary>
        /// <remarks>
        /// This check doesn't use reflection or runtime type information
        /// and doesn't suffer related performance penalties.
        /// </remarks>
        public bool IsRoot => _dirEntry.StgType == StgType.StgRoot;

        /// <summary>
        /// Get/Set the Creation Date of the current item
        /// </summary>
        public DateTime CreationDate
        {
            get => DateTime.FromFileTime(BitConverter.ToInt64(_dirEntry.CreationDate, 0));

            set
            {
                if (_dirEntry.StgType != StgType.StgStream && _dirEntry.StgType != StgType.StgRoot)
                    _dirEntry.CreationDate = BitConverter.GetBytes((value.ToFileTime()));
                else
                    throw new CfException("Creation Date can only be set on storage entries");
            }
        }

        /// <summary>
        /// Get/Set the Modify Date of the current item
        /// </summary>
        public DateTime ModifyDate
        {
            get => DateTime.FromFileTime(BitConverter.ToInt64(_dirEntry.ModifyDate, 0));

            set
            {
                if (_dirEntry.StgType != StgType.StgStream && _dirEntry.StgType != StgType.StgRoot)
                    _dirEntry.ModifyDate = BitConverter.GetBytes((value.ToFileTime()));
                else
                    throw new CfException("Modify Date can only be set on storage entries");
            }
        }

        /// <summary>
        /// Get/Set Object class Guid for Root and Storage entries.
        /// </summary>
        public Guid Clsid
        {
            get => _dirEntry.StorageClsid;
            set
            {
                if (_dirEntry.StgType != StgType.StgStream)
                {
                    _dirEntry.StorageClsid = value;
                }
                else
                    throw new CfException("Object class GUID can only be set on Root and Storage entries");
            }
        }

        int IComparable<CfItem>.CompareTo(CfItem other)
        {
            return _dirEntry.CompareTo(other.DirEntry);
        }

        public override string ToString()
        {
            if (_dirEntry != null)
                return "[" + _dirEntry.LeftSibling + "," + _dirEntry.Sid + "," + _dirEntry.RightSibling + "]" + " " + _dirEntry.GetEntryName();
            return string.Empty;
        }
    }
}
