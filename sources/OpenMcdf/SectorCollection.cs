/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 * 
 * The Original Code is OpenMCDF - Compound Document Format library.
 * 
 * The Initial Developer of the Original Code is Federico Blaseotto.*/

using System;
using System.Collections;
using System.Collections.Generic;

namespace OpenMcdf
{
    /// <summary>
    /// Action to implement when transaction support - sector
    /// has to be written to the underlying stream (see specs).
    /// </summary>
    public delegate void Ver3SizeLimitReached();

    /// <summary>
    /// Ad-hoc Heap Friendly sector collection to avoid using 
    /// large array that may create some problem to GC collection 
    /// (see http://www.simple-talk.com/dotnet/.net-framework/the-dangers-of-the-large-object-heap/ )
    /// </summary>
    internal class SectorCollection : IList<Sector>
    {
        private const int MaxSectorV4CountLockRange = 524287; //0x7FFFFF00 for Version 4
        private const int SliceSize = 4096;

        private int _count;


        public event Ver3SizeLimitReached OnVer3SizeLimitReached;

        private readonly List<ArrayList> _largeArraySlices = new List<ArrayList>();

        private bool _sizeLimitReached;
        private void DoCheckSizeLimitReached()
        {
            if (OnVer3SizeLimitReached != null && !_sizeLimitReached && (_count - 1 > MaxSectorV4CountLockRange))
            {
                _sizeLimitReached = true;
                OnVer3SizeLimitReached();


            }
        }

        #region IList<T> Members

        public int IndexOf(Sector item)
        {
            throw new NotImplementedException();
        }

        public void Insert(int index, Sector item)
        {
            throw new NotImplementedException();
        }

        public void RemoveAt(int index)
        {
            throw new NotImplementedException();
        }

        public Sector this[int index]
        {
            get
            {
                var itemIndex = index / SliceSize;
                var itemOffset = index % SliceSize;

                if ((index > -1) && (index < _count))
                {
                    return (Sector)_largeArraySlices[itemIndex][itemOffset];
                }

                throw new CfException("Argument Out of Range, possibly corrupted file", new ArgumentOutOfRangeException("index", index, "Argument out of range"));

            }

            set
            {
                var itemIndex = index / SliceSize;
                var itemOffset = index % SliceSize;

                if (index > -1 && index < _count)
                {
                    _largeArraySlices[itemIndex][itemOffset] = value;
                }
                else
                    throw new ArgumentOutOfRangeException("index", index, "Argument out of range");
            }
        }

        #endregion

        #region ICollection<T> Members

        private int AddReturnCount(Sector item)
        {
            var itemIndex = _count / SliceSize;

            if (itemIndex < _largeArraySlices.Count)
            {
                _largeArraySlices[itemIndex].Add(item);
                _count++;
            }
            else
            {
                var ar = new ArrayList(SliceSize);
                ar.Add(item);
                _largeArraySlices.Add(ar);
                _count++;
            }

            return _count - 1;
        }

        public void Add(Sector item)
        {
            DoCheckSizeLimitReached();

            AddReturnCount(item);

        }

        public void Clear()
        {
            foreach (var slice in _largeArraySlices)
            {
                slice.Clear();
            }

            _largeArraySlices.Clear();

            _count = 0;
        }

        public bool Contains(Sector item)
        {
            throw new NotImplementedException();
        }

        public void CopyTo(Sector[] array, int arrayIndex)
        {
            throw new NotImplementedException();
        }

        public int Count => _count;

        public bool IsReadOnly => false;

        public bool Remove(Sector item)
        {
            throw new NotImplementedException();
        }

        #endregion

        #region IEnumerable<T> Members

        public IEnumerator<Sector> GetEnumerator()
        {

            for (var i = 0; i < _largeArraySlices.Count; i++)
            {
                for (var j = 0; j < _largeArraySlices[i].Count; j++)
                {
                    yield return (Sector)_largeArraySlices[i][j];

                }
            }
        }

        #endregion

        #region IEnumerable Members

        IEnumerator IEnumerable.GetEnumerator()
        {
            for (var i = 0; i < _largeArraySlices.Count; i++)
            {
                for (var j = 0; j < _largeArraySlices[i].Count; j++)
                {
                    yield return _largeArraySlices[i][j];
                }
            }
        }

        #endregion
    }
}
