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

namespace OpenMcdf
{
    /// <summary>
    /// Stream decorator for a Sector or miniSector chain
    /// </summary>
    internal class StreamView : Stream
    {
        private readonly int _sectorSize;

        private long _position;

        private readonly List<Sector> _sectorChain;
        private readonly Stream _stream;
        private readonly bool _isFatStream;
        private readonly List<Sector> _freeSectors = new List<Sector>();
        public IEnumerable<Sector> FreeSectors => _freeSectors;

        public StreamView(List<Sector> sectorChain, int sectorSize, Stream stream)
        {
            if (sectorChain == null)
                throw new CfException("Sector Chain cannot be null");

            if (sectorSize <= 0)
                throw new CfException("Sector size must be greater than zero");

            _sectorChain = sectorChain;
            _sectorSize = sectorSize;
            _stream = stream;
        }

        public StreamView(List<Sector> sectorChain, int sectorSize, long length, Queue<Sector> availableSectors, Stream stream, bool isFatStream = false)
            : this(sectorChain, sectorSize, stream)
        {
            _isFatStream = isFatStream;
            AdjustLength(length, availableSectors);
        }




        public List<Sector> BaseSectorChain => _sectorChain;

        public override bool CanRead => true;

        public override bool CanSeek => true;

        public override bool CanWrite => true;

        public override void Flush()
        {

        }

        private long _length;

        public override long Length => _length;

        public override long Position
        {
            get => _position;

            set
            {
                if (_position > _length - 1)
                    throw new ArgumentOutOfRangeException("value");

                _position = value;
            }
        }

        public override void Close()
        {
            base.Close();
        }

        private readonly byte[] _buf = new byte[4];

        public int ReadInt32()
        {
            Read(_buf, 0, 4);
            return (((_buf[0] | (_buf[1] << 8)) | (_buf[2] << 16)) | (_buf[3] << 24));
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            var nRead = 0;
            var nToRead = 0;

            // Don't try to read more bytes than this stream contains.
            var intMax = Math.Min(int.MaxValue, _length);
            count = Math.Min((int)(intMax), count);

            if (_sectorChain != null && _sectorChain.Count > 0)
            {
                // First sector
                var secIndex = (int)(_position / _sectorSize);

                // Bytes to read count is the min between request count
                // and sector border

                nToRead = Math.Min(
                    _sectorChain[0].Size - ((int)_position % _sectorSize),
                    count);

                if (secIndex < _sectorChain.Count)
                {
                    Buffer.BlockCopy(
                        _sectorChain[secIndex].GetData(),
                        (int)(_position % _sectorSize),
                        buffer,
                        offset,
                        nToRead
                        );
                }

                nRead += nToRead;

                secIndex++;

                // Central sectors
                while (nRead < (count - _sectorSize))
                {
                    nToRead = _sectorSize;

                    Buffer.BlockCopy(
                        _sectorChain[secIndex].GetData(),
                        0,
                        buffer,
                        offset + nRead,
                        nToRead
                        );

                    nRead += nToRead;
                    secIndex++;
                }

                // Last sector
                nToRead = count - nRead;

                if (nToRead != 0)
                {
					if (secIndex > _sectorChain.Count) throw new CfCorruptedFileException("The file is probably corrupted.");

                    Buffer.BlockCopy(
                        _sectorChain[secIndex].GetData(),
                        0,
                        buffer,
                        offset + nRead,
                        nToRead
                        );

                    nRead += nToRead;
                }

                _position += nRead;

                return nRead;

            }

            return 0;

        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            switch (origin)
            {
                case SeekOrigin.Begin:
                    _position = offset;
                    break;

                case SeekOrigin.Current:
                    _position += offset;
                    break;

                case SeekOrigin.End:
                    _position = Length - offset;
                    break;
            }

            if (_length <= _position) // Dont't adjust the length when position is inside the bounds of 0 and the current length.
                AdjustLength(_position);

            return _position;
        }

        private void AdjustLength(long value)
        {
            AdjustLength(value, null);
        }

        private void AdjustLength(long value, Queue<Sector> availableSectors)
        {
            _length = value;

            var delta = value - (_sectorChain.Count * (long)_sectorSize);

            if (delta > 0)
            {
                // enlargment required

                var nSec = (int)Math.Ceiling(((double)delta / _sectorSize));

                while (nSec > 0)
                {
                    Sector t = null;

                    if (availableSectors == null || availableSectors.Count == 0)
                    {
                        t = new Sector(_sectorSize, _stream);

                        if (_sectorSize == Sector.MinisectorSize)
                            t.Type = SectorType.Mini;
                    }
                    else
                    {
                        t = availableSectors.Dequeue();
                    }

                    if (_isFatStream)
                    {
                        t.InitFatData();
                    }
                    _sectorChain.Add(t);
                    nSec--;
                }

                //if (((int)delta % sectorSize) != 0)
                //{
                //    Sector t = new Sector(sectorSize);
                //    sectorChain.Add(t);
                //}
            }
            //else
            //{
            //    // FREE Sectors
            //    delta = Math.Abs(delta);

            //    int nSec = (int)Math.Floor(((double)delta / sectorSize));

            //    while (nSec > 0)
            //    {
            //        freeSectors.Add(sectorChain[sectorChain.Count - 1]);
            //        sectorChain.RemoveAt(sectorChain.Count - 1);
            //        nSec--;
            //    }
            //}
        }

        public override void SetLength(long value)
        {
            AdjustLength(value);
        }

        public void WriteInt32(int val)
        {
            var buffer = new byte[4];
            buffer[0] = (byte)val;
            buffer[1] = (byte)(val << 8);
            buffer[2] = (byte)(val << 16);
            buffer[3] = (byte)(val << 32);
            Write(buffer, 0, 4);
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            var byteWritten = 0;
            var roundByteWritten = 0;

            // Assure length
            if ((_position + count) > _length)
                AdjustLength((_position + count));

            if (_sectorChain != null)
            {
                // First sector
                var secOffset = (int)(_position / _sectorSize);
                var secShift = (int)(_position % _sectorSize);

                roundByteWritten = Math.Min(_sectorSize - (int)(_position % _sectorSize), count);

                if (secOffset < _sectorChain.Count)
                {
                    Buffer.BlockCopy(
                        buffer,
                        offset,
                        _sectorChain[secOffset].GetData(),
                        secShift,
                        roundByteWritten
                        );

                    _sectorChain[secOffset].DirtyFlag = true;
                }

                byteWritten += roundByteWritten;
                offset += roundByteWritten;
                secOffset++;

                // Central sectors
                while (byteWritten < (count - _sectorSize))
                {
                    roundByteWritten = _sectorSize;

                    Buffer.BlockCopy(
                        buffer,
                        offset,
                        _sectorChain[secOffset].GetData(),
                        0,
                        roundByteWritten
                        );

                    _sectorChain[secOffset].DirtyFlag = true;

                    byteWritten += roundByteWritten;
                    offset += roundByteWritten;
                    secOffset++;
                }

                // Last sector
                roundByteWritten = count - byteWritten;

                if (roundByteWritten != 0)
                {
                    Buffer.BlockCopy(
                        buffer,
                        offset,
                        _sectorChain[secOffset].GetData(),
                        0,
                        roundByteWritten
                        );

                    _sectorChain[secOffset].DirtyFlag = true;

                    offset += roundByteWritten;
                    byteWritten += roundByteWritten;
                }

                _position += count;

            }
        }
        
        public override void Write(ReadOnlySpan<byte> buffer)
        {
            var byteWritten = 0;
            var roundByteWritten = 0;
            var offset = 0;

            // Assure length
            if (_position + buffer.Length > _length)
                AdjustLength(_position + buffer.Length);

            if (_sectorChain == null) return;
            // First sector
            var secOffset = (int)(_position / _sectorSize);
            var secShift = (int)(_position % _sectorSize);

            roundByteWritten = Math.Min(_sectorSize - (int)(_position % _sectorSize), buffer.Length);

            if (secOffset < _sectorChain.Count)
            {
                buffer[offset..roundByteWritten].CopyTo(_sectorChain[secOffset].GetData().AsSpan()[secShift..(secShift + roundByteWritten)]);
                _sectorChain[secOffset].DirtyFlag = true;
            }

            byteWritten += roundByteWritten;
            offset += roundByteWritten;
            secOffset++;

            // Central sectors
            while (byteWritten < buffer.Length - _sectorSize)
            {
                roundByteWritten = _sectorSize;
                buffer[offset..(offset+roundByteWritten)].CopyTo(_sectorChain[secOffset].GetData().AsSpan()[..roundByteWritten]);
                _sectorChain[secOffset].DirtyFlag = true;

                byteWritten += roundByteWritten;
                offset += roundByteWritten;
                secOffset++;
            }

            // Last sector
            roundByteWritten = buffer.Length - byteWritten;

            if (roundByteWritten != 0)
            {
                buffer[offset..(offset+roundByteWritten)].CopyTo(_sectorChain[secOffset].GetData().AsSpan()[..roundByteWritten]);

                _sectorChain[secOffset].DirtyFlag = true;

                offset += roundByteWritten;
                byteWritten += roundByteWritten;
            }

            _position += buffer.Length;
        }
    }
}