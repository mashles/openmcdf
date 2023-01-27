using System;
using Be.Windows.Forms;
using OpenMcdf;

namespace StructuredStorageExplorer
{
    public class StreamDataProvider : IByteProvider
    {
        /// <summary>
        /// Modifying stream
        /// </summary>
        readonly CfStream _modifiedStream;

        /// <summary>
        /// Contains information about changes.
        /// </summary>
        bool _hasChanges;

        /// <summary>
        /// Contains a byte collection.
        /// </summary>
        readonly ByteCollection _bytes;


        /// <summary>
        /// Initializes a new instance of the DynamicByteProvider class.
        /// </summary>
        /// <param name="bytes"></param>
        public StreamDataProvider(CfStream modifiedStream)
        {
            _bytes = new ByteCollection(modifiedStream.GetData());
            _modifiedStream = modifiedStream;
        }

        /// <summary>
        /// Raises the Changed event.
        /// </summary>
        void OnChanged(EventArgs e)
        {
            _hasChanges = true;

            if (Changed != null)
                Changed(this, e);
        }

        /// <summary>
        /// Raises the LengthChanged event.
        /// </summary>
        void OnLengthChanged(EventArgs e)
        {
            if (LengthChanged != null)
                LengthChanged(this, e);
        }

        /// <summary>
        /// Gets the byte collection.
        /// </summary>
        public ByteCollection Bytes => _bytes;

        #region IByteProvider Members
        /// <summary>
        /// True, when changes are done.
        /// </summary>
        public bool HasChanges()
        {
            return _hasChanges;
        }

        /// <summary>
        /// Applies changes.
        /// </summary>
        public void ApplyChanges()
        {
            _hasChanges = false;

            _modifiedStream.SetData(_bytes.ToArray());
        }

        /// <summary>
        /// Occurs, when the write buffer contains new changes.
        /// </summary>
        public event EventHandler Changed;

        /// <summary>
        /// Occurs, when InsertBytes or DeleteBytes method is called.
        /// </summary>
        public event EventHandler LengthChanged;


        /// <summary>
        /// Reads a byte from the byte collection.
        /// </summary>
        /// <param name="index">the index of the byte to read</param>
        /// <returns>the byte</returns>
        public byte ReadByte(long index)
        { return _bytes[(int)index]; }

        /// <summary>
        /// Write a byte into the byte collection.
        /// </summary>
        /// <param name="index">the index of the byte to write.</param>
        /// <param name="value">the byte</param>
        public void WriteByte(long index, byte value)
        {
            _bytes[(int)index] = value;
            OnChanged(EventArgs.Empty);
        }

        /// <summary>
        /// Deletes bytes from the byte collection.
        /// </summary>
        /// <param name="index">the start index of the bytes to delete.</param>
        /// <param name="length">the length of bytes to delete.</param>
        public void DeleteBytes(long index, long length)
        {
            var internalIndex = (int)Math.Max(0, index);
            var internalLength = (int)Math.Min((int)Length, length);
            _bytes.RemoveRange(internalIndex, internalLength);

            OnLengthChanged(EventArgs.Empty);
            OnChanged(EventArgs.Empty);
        }

        /// <summary>
        /// Inserts byte into the byte collection.
        /// </summary>
        /// <param name="index">the start index of the bytes in the byte collection</param>
        /// <param name="bs">the byte array to insert</param>
        public void InsertBytes(long index, byte[] bs)
        {
            _bytes.InsertRange((int)index, bs);

            OnLengthChanged(EventArgs.Empty);
            OnChanged(EventArgs.Empty);
        }

        /// <summary>
        /// Gets the length of the bytes in the byte collection.
        /// </summary>
        public long Length => _bytes.Count;

        /// <summary>
        /// Returns true
        /// </summary>
        public bool SupportsWriteByte()
        {
            return true;
        }

        /// <summary>
        /// Returns true
        /// </summary>
        public bool SupportsInsertBytes()
        {
            return true;
        }

        /// <summary>
        /// Returns true
        /// </summary>
        public bool SupportsDeleteBytes()
        {
            return true;
        }
        #endregion
    }
}
