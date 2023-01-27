/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 * 
 * The Original Code is OpenMCDF - Compound Document Format library.
 * 
 * The Initial Developer of the Original Code is Federico Blaseotto.*/


using System;
using System.IO;

namespace OpenMcdf
{
    internal enum SectorType
    {
        Normal, Mini, Fat, Difat, RangeLockSector, Directory
    }

    internal class Sector : IDisposable
    {
        public static int MinisectorSize = 64;

        public const int Freesect = unchecked((int)0xFFFFFFFF);
        public const int Endofchain = unchecked((int)0xFFFFFFFE);
        public const int Fatsect = unchecked((int)0xFFFFFFFD);
        public const int Difsect = unchecked((int)0xFFFFFFFC);

        private bool _dirtyFlag;

        public bool DirtyFlag
        {
            get => _dirtyFlag;
            set => _dirtyFlag = value;
        }

        public bool IsStreamed => (_stream != null && Size != MinisectorSize) ? (Id * Size) + Size < _stream.Length : false;

        private readonly Stream _stream;


        public Sector(int size, Stream stream)
        {
            Size = size;
            _stream = stream;
        }

        public Sector(int size, byte[] data)
        {
            Size = size;
            _data = data;
            _stream = null;
        }

        public Sector(int size)
        {
            Size = size;
            _data = null;
            _stream = null;
        }

        internal SectorType Type { get; set; }

        public int Id { get; set; } = -1;

        public int Size { get; private set; }

        private byte[] _data;

        public byte[] GetData()
        {
            if (_data != null) return _data;
            _data = new byte[Size];

            if (!IsStreamed) return _data;
            _stream.Seek(Size + Id * (long)Size, SeekOrigin.Begin);
            _stream.Read(_data, 0, Size);
            return _data;
        }

        //public void SetSectorData(byte[] b)
        //{
        //    this.data = b;
        //}

        //public void FillData(byte b)
        //{
        //    if (data != null)
        //    {
        //        for (int i = 0; i < data.Length; i++)
        //        {
        //            data[i] = b;
        //        }
        //    }
        //}

        public void ZeroData()
        {
            _data = new byte[Size];
            _dirtyFlag = true;
        }

        public void InitFatData()
        {
            _data = new byte[Size];
            
            for (var i = 0; i < Size; i++)
                _data[i] = 0xFF;

            _dirtyFlag = true;
        }

        internal void ReleaseData()
        {
            _data = null;
        }

        private readonly object _lockObject = new object();

        /// <summary>
        /// When called from user code, release all resources, otherwise, in the case runtime called it,
        /// only unmanaged resources are released.
        /// </summary>
        /// <param name="disposing">If true, method has been called from User code, if false it's been called from .net runtime</param>
        protected virtual void Dispose(bool disposing)
        {
            try
            {
                if (_disposed) return;
                lock (_lockObject)
                {
                    if (disposing)
                    {
                        // Call from user code...


                    }

                    _data = null;
                    _dirtyFlag = false;
                    Id = Endofchain;
                    Size = 0;

                }
            }
            finally
            {
                _disposed = true;
            }

        }

        #region IDisposable Members

        private bool _disposed;//false

        void IDisposable.Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        #endregion
    }




}
