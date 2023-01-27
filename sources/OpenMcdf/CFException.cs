/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 * 
 * The Original Code is OpenMCDF - Compound Document Format library.
 * 
 * The Initial Developer of the Original Code is Federico Blaseotto.*/

using System;
using System.Runtime.Serialization;

namespace OpenMcdf
{
    /// <summary>
    /// OpenMCDF base exception.
    /// </summary>
    [Serializable]
    public class CfException : Exception
    {
        public CfException()
        {
        }

        protected CfException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }

        public CfException(string message)
            : base(message, null)
        {

        }

        public CfException(string message, Exception innerException)
            : base(message, innerException)
        {

        }

    }

    /// <summary>
    /// Raised when a data setter/getter method is invoked
    /// on a stream or storage object after the disposal of the owner
    /// compound file object.
    /// </summary>
    [Serializable]
    public class CfDisposedException : CfException
    {
        public CfDisposedException()
        {
        }

        protected CfDisposedException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }

        public CfDisposedException(string message)
            : base(message, null)
        {

        }

        public CfDisposedException(string message, Exception innerException)
            : base(message, innerException)
        {

        }

    }

    /// <summary>
    /// Raised when opening a file with invalid header
    /// or not supported COM/OLE Structured storage version.
    /// </summary>
    [Serializable]
    public class CfFileFormatException : CfException
    {
        public CfFileFormatException()
        {
        }

        protected CfFileFormatException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }

        public CfFileFormatException(string message)
            : base(message, null)
        {

        }

        public CfFileFormatException(string message, Exception innerException)
            : base(message, innerException)
        {

        }

    }

    /// <summary>
    /// Raised when a named stream or a storage object
    /// are not found in a parent storage.
    /// </summary>
    [Serializable]
    public class CfItemNotFound : CfException
    {
        protected CfItemNotFound(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }

        public CfItemNotFound()
            : base("Entry not found")
        {
        }

        public CfItemNotFound(string message)
            : base(message, null)
        {

        }

        public CfItemNotFound(string message, Exception innerException)
            : base(message, innerException)
        {

        }

    }

    /// <summary>
    /// Raised when a method call is invalid for the current object state
    /// </summary>
    [Serializable]
    public class CfInvalidOperation : CfException
    {
         public CfInvalidOperation()
         {
        }

        protected CfInvalidOperation(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }

        public CfInvalidOperation(string message)
            : base(message, null)
        {

        }

        public CfInvalidOperation(string message, Exception innerException)
            : base(message, innerException)
        {

        }

    }

    /// <summary>
    /// Raised when trying to add a duplicated CFItem
    /// </summary>
    /// <remarks>
    /// Items are compared by name as indicated by specs.
    /// Two items with the same name CANNOT be added within 
    /// the same storage or sub-storage. 
    /// </remarks>
    [Serializable]
    public class CfDuplicatedItemException : CfException
    {
        public CfDuplicatedItemException()
        {
        }

        protected CfDuplicatedItemException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }

        public CfDuplicatedItemException(string message)
            : base(message, null)
        {

        }

        public CfDuplicatedItemException(string message, Exception innerException)
            : base(message, innerException)
        {

        }
    }

    /// <summary>
    /// Raised when trying to load a Compound File with invalid, corrupted or mismatched fields (4.1 - specifications) 
    /// </summary>
    /// <remarks>
    /// This exception is NOT raised when Compound file has been opened with NO_VALIDATION_EXCEPTION option.
    /// </remarks>
    [Serializable]
    public class CfCorruptedFileException : CfException
    {
        public CfCorruptedFileException()
        {
        }

        protected CfCorruptedFileException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }

        public CfCorruptedFileException(string message)
            : base(message, null)
        {

        }

        public CfCorruptedFileException(string message, Exception innerException)
            : base(message, innerException)
        {

        }

    }

}
