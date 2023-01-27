using System;
using System.IO;

namespace OpenMcdf.Extensions
{
    public class StreamDecorator : Stream
    {
        private readonly CfStream _cfStream;
        private long _position;

        public StreamDecorator(CfStream cfstream)
        {
            _cfStream = cfstream;
        }

        public override bool CanRead => true;

        public override bool CanSeek => true;

        public override bool CanWrite => true;

        public override void Flush()
        {
            // nothing to do;
        }

        public override long Length => _cfStream.Size;

        public override long Position
        {
            get => _position;
            set => _position = value;
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            if (count > buffer.Length)
                throw new ArgumentException("Count parameter exceeds buffer size");

            if (buffer == null)
                throw new ArgumentNullException("Buffer cannot be null");

            if (offset < 0 || count < 0)
                throw new ArgumentOutOfRangeException("Offset and Count parameters must be non-negative numbers");

            if (_position >= _cfStream.Size)
                return 0;

            count = _cfStream.Read(buffer, _position, offset, count);
            _position += count;
            return count;
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
                    _position -= offset;
                    break;
                default:
                    throw new Exception("Invalid origin selected");
            }

            return _position;
        }

        public override void SetLength(long value)
        {
            _cfStream.Resize(value);
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            _cfStream.Write(buffer, _position, offset, count);
            _position += count;
        }

        public override void Close()
        {
            // Do nothing
        }
    }
}
