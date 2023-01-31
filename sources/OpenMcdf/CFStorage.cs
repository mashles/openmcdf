/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 * 
 * The Original Code is OpenMCDF - Compound Document Format library.
 * 
 * The Initial Developer of the Original Code is Federico Blaseotto.*/
using System;
using System.Collections.Generic;
using System.Linq;
using OpenMcdf.RBTree;

namespace OpenMcdf
{
    /// <summary>
    /// Action to apply to  visited items in the OLE structured storage
    /// </summary>
    /// <param name="item">Currently visited <see cref="T:OpenMcdf.CFItem">item</see></param>
    /// <example>
    /// <code>
    /// 
    /// //We assume that xls file should be a valid OLE compound file
    /// const String STORAGE_NAME = "report.xls";
    /// CompoundFile cf = new CompoundFile(STORAGE_NAME);
    ///
    /// FileStream output = new FileStream("LogEntries.txt", FileMode.Create);
    /// TextWriter tw = new StreamWriter(output);
    ///
    /// VisitedEntryAction va = delegate(CFItem item)
    /// {
    ///     tw.WriteLine(item.Name);
    /// };
    ///
    /// cf.RootStorage.VisitEntries(va, true);
    ///
    /// tw.Close();
    ///
    /// </code>
    /// </example>
    public delegate void VisitedEntryAction(CfItem item);

    /// <summary>
    /// Storage entity that acts like a logic container for streams
    /// or sub-storages in a compound file.
    /// </summary>
    public class CfStorage : CfItem
    {
        private OrderedMap<string, IDirectoryEntry> _children;
        internal OrderedMap<string, IDirectoryEntry> Children
        {
            get
            {
                // Lazy loading of children tree.
                if (_children != null) return _children;
                _children = LoadChildren(DirEntry.Sid) ?? CompoundFile.CreateNewTree();
                return _children;
            }
        }


        /// <summary>
        /// Create a CFStorage using an existing directory (previously loaded).
        /// </summary>
        /// <param name="compFile">The Storage Owner - CompoundFile</param>
        /// <param name="dirEntry">An existing Directory Entry</param>
        internal CfStorage(CompoundFile compFile, IDirectoryEntry dirEntry)
            : base(compFile)
        {
            if (dirEntry == null || dirEntry.Sid < 0)
                throw new CfException("Attempting to create a CFStorage using an uninitialized directory");

            DirEntry = dirEntry;
        }

        private OrderedMap<string, IDirectoryEntry> LoadChildren(int sid)
        {
            var childrenTree = CompoundFile.GetChildrenTree(sid);

            DirEntry.Child = childrenTree.Root != null ? childrenTree.Root.Value.Sid : DirectoryEntry.Nostream;

            return childrenTree;
        }

        /// <summary>
        /// Create a new child stream inside the current <see cref="T:OpenMcdf.CFStorage">storage</see>
        /// </summary>
        /// <param name="streamName">The new stream name</param>
        /// <returns>The new <see cref="T:OpenMcdf.CFStream">stream</see> reference</returns>
        /// <exception cref="T:OpenMcdf.CFDuplicatedItemException">Raised when adding an item with the same name of an existing one</exception>
        /// <exception cref="T:OpenMcdf.CFDisposedException">Raised when adding a stream to a closed compound file</exception>
        /// <exception cref="T:OpenMcdf.CFException">Raised when adding a stream with null or empty name</exception>
        /// <example>
        /// <code>
        /// 
        ///  String filename = "A_NEW_COMPOUND_FILE_YOU_CAN_WRITE_TO.cfs";
        ///
        ///  CompoundFile cf = new CompoundFile();
        ///
        ///  CFStorage st = cf.RootStorage.AddStorage("MyStorage");
        ///  CFStream sm = st.AddStream("MyStream");
        ///  byte[] b = Helpers.GetBuffer(220, 0x0A);
        ///  sm.SetData(b);
        ///
        ///  cf.Save(filename);
        ///  
        /// </code>
        /// </example>
        public CfStream AddStream(string streamName)
        {
            CheckDisposed();

            if (string.IsNullOrEmpty(streamName))
                throw new CfException("Stream name cannot be null or empty");

            var dirEntry = DirectoryEntry.TryNew(streamName, StgType.StgStream, CompoundFile.GetDirectories());

            // Add new Stream directory entry
            //cfo = new CFStream(this.CompoundFile, streamName);
            
            // Add object to Siblings tree
            Children.Add(dirEntry.Name, dirEntry);
            //... and set the root of the tree as new child of the current item directory entry
            if (Children.Root != null) DirEntry.Child = Children.Root.Value.Sid;
            return new CfStream(CompoundFile, dirEntry);
        }


        /// <summary>
        /// Get a named <see cref="T:OpenMcdf.CFStream">stream</see> contained in the current storage if existing.
        /// </summary>
        /// <param name="streamName">Name of the stream to look for</param>
        /// <returns>A stream reference if existing</returns>
        /// <exception cref="T:OpenMcdf.CFDisposedException">Raised if trying to delete item from a closed compound file</exception>
        /// <exception cref="T:OpenMcdf.CFItemNotFound">Raised if item to delete is not found</exception>
        /// <example>
        /// <code>
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
        public CfStream GetStream(string streamName)
        {
            CheckDisposed();

            if (Children.TryFindNode(streamName, out var node) && node.StgType == StgType.StgStream)
            {
                return new CfStream(CompoundFile, node);
            }

            throw new CfItemNotFound("Cannot find item [" + streamName + "] within the current storage");
        }

        /// <summary>
        /// Get a named <see cref="T:OpenMcdf.CFStream">stream</see> contained in the current storage if existing.
        /// </summary>
        /// <param name="streamName">Name of the stream to look for</param>
        /// <param name="cfStream">Found <see cref="T:OpenMcdf.CFStream"> if any</param>
        /// <returns><see cref="T:System.Boolean"> true if stream found, else false</returns>
        /// <example>
        /// <code>
        /// String filename = "report.xls";
        ///
        /// CompoundFile cf = new CompoundFile(filename);
        /// bool b = cf.RootStorage.TryGetStream("Workbook",out CFStream foundStream);
        ///
        /// byte[] temp = foundStream.GetData();
        ///
        /// Assert.IsNotNull(temp);
        /// Assert.IsTrue(b);
        ///
        /// cf.Close();
        /// </code>
        /// </example>
        public bool TryGetStream(string streamName, out CfStream cfStream)
        {
            cfStream = null;
            try
            {
                CheckDisposed();
                if (Children.TryFindNode(streamName, out var node) && node.StgType == StgType.StgStream)
                {
                    cfStream = new CfStream(CompoundFile, node);
                    return true;
                }
            }
            catch (CfDisposedException)
            {
                return false;
            }
            return false;
        }

        /// <summary>
        /// Get a named storage contained in the current one if existing.
        /// </summary>
        /// <param name="storageName">Name of the storage to look for</param>
        /// <returns>A storage reference if existing.</returns>
        /// <exception cref="T:OpenMcdf.CFDisposedException">Raised if trying to delete item from a closed compound file</exception>
        /// <exception cref="T:OpenMcdf.CFItemNotFound">Raised if item to delete is not found</exception>
        /// <example>
        /// <code>
        /// 
        /// String FILENAME = "MultipleStorage2.cfs";
        /// CompoundFile cf = new CompoundFile(FILENAME, UpdateMode.ReadOnly, false, false);
        ///
        /// CFStorage st = cf.RootStorage.GetStorage("MyStorage");
        ///
        /// Assert.IsNotNull(st);
        /// cf.Close();
        /// </code>
        /// </example>
        public CfStorage GetStorage(string storageName)
        {
            CheckDisposed();

            if (Children.TryFindNode(storageName, out var outDe) && outDe.StgType == StgType.StgStorage)
            {
                return new CfStorage(CompoundFile, outDe);
            }
            throw new CfItemNotFound("Cannot find item [" + storageName + "] within the current storage");
        }

        /// <summary>
        /// Get a named storage contained in the current one if existing.
        /// </summary>
        /// <param name="storageName">Name of the storage to look for</param>
        /// <param name="cfStorage">A storage reference if found else null</param>
        /// <returns><see cref="T:System.Boolean"> true if storage found, else false</returns>
        /// <example>
        /// <code>
        /// 
        /// String FILENAME = "MultipleStorage2.cfs";
        /// CompoundFile cf = new CompoundFile(FILENAME, UpdateMode.ReadOnly, false, false);
        ///
        /// bool b = cf.RootStorage.TryGetStorage("MyStorage",out CFStorage st);
        ///
        /// Assert.IsNotNull(st);
        /// Assert.IsTrue(b);
        /// 
        /// cf.Close();
        /// </code>
        /// </example>
        public bool TryGetStorage(string storageName, out CfStorage cfStorage)
        {
            cfStorage = null;
            try
            {
                CheckDisposed();
                if (Children.TryFindNode(storageName, out var outDe) && outDe.StgType == StgType.StgStorage)
                {
                    cfStorage = new CfStorage(CompoundFile, outDe);
                    return true;
                }

            }
            catch (CfDisposedException)
            {
                return false;
            }
            return false;
        }


        /// <summary>
        /// Create new child storage directory inside the current storage.
        /// </summary>
        /// <param name="storageName">The new storage name</param>
        /// <returns>Reference to the new <see cref="T:OpenMcdf.CFStorage">storage</see></returns>
        /// <exception cref="T:OpenMcdf.CFDuplicatedItemException">Raised when adding an item with the same name of an existing one</exception>
        /// <exception cref="T:OpenMcdf.CFDisposedException">Raised when adding a storage to a closed compound file</exception>
        /// <exception cref="T:OpenMcdf.CFException">Raised when adding a storage with null or empty name</exception>
        /// <example>
        /// <code>
        /// 
        ///  String filename = "A_NEW_COMPOUND_FILE_YOU_CAN_WRITE_TO.cfs";
        ///
        ///  CompoundFile cf = new CompoundFile();
        ///
        ///  CFStorage st = cf.RootStorage.AddStorage("MyStorage");
        ///  CFStream sm = st.AddStream("MyStream");
        ///  byte[] b = Helpers.GetBuffer(220, 0x0A);
        ///  sm.SetData(b);
        ///
        ///  cf.Save(filename);
        ///  
        /// </code>
        /// </example>
        public CfStorage AddStorage(string storageName)
        {
            CheckDisposed();

            if (string.IsNullOrEmpty(storageName))
                throw new CfException("Stream name cannot be null or empty");

            // Add new Storage directory entry
            var cfo
                = DirectoryEntry.New(storageName, StgType.StgStorage, CompoundFile.GetDirectories());

            //this.CompoundFile.InsertNewDirectoryEntry(cfo);

            try
            {
                // Add object to Siblings tree
                Children.Add(storageName, cfo);
            }
            catch (RbTreeDuplicatedItemException)
            {
                CompoundFile.ResetDirectoryEntry(cfo.Sid);
                throw new CfDuplicatedItemException("An entry with name '" + storageName + "' is already present in storage '" + Name + "' ");
            }

            var childrenRoot = Children.Root;
            if (childrenRoot != null) DirEntry.Child = childrenRoot.Value.Sid;

            return new CfStorage(CompoundFile, cfo);
        }

        /// <summary>
        /// Visit all entities contained in the storage applying a user provided action
        /// </summary>
        /// <exception cref="T:OpenMcdf.CFDisposedException">Raised when visiting items of a closed compound file</exception>
        /// <param name="action">User <see cref="T:OpenMcdf.VisitedEntryAction">action</see> to apply to visited entities</param>
        /// <param name="recursive"> Visiting recursion level. True means substorages are visited recursively, false indicates that only the direct children of this storage are visited</param>
        /// <example>
        /// <code>
        /// const String STORAGE_NAME = "report.xls";
        /// CompoundFile cf = new CompoundFile(STORAGE_NAME);
        ///
        /// FileStream output = new FileStream("LogEntries.txt", FileMode.Create);
        /// TextWriter tw = new StreamWriter(output);
        ///
        /// VisitedEntryAction va = delegate(CFItem item)
        /// {
        ///     tw.WriteLine(item.Name);
        /// };
        ///
        /// cf.RootStorage.VisitEntries(va, true);
        ///
        /// tw.Close();
        /// </code>
        /// </example>
        public void VisitEntries(Action<CfItem> action, bool recursive)
        {
            CheckDisposed();

            if (action == null) return;
            var subStorages
                = new List<IDirectoryEntry>();

            void InternalAction(IDirectoryEntry target)
            {
                if (target.StgType == StgType.StgStream)
                    action(new CfStream(CompoundFile, target));
                else
                    action(new CfStorage(CompoundFile, target));

                if (target.Child != DirectoryEntry.Nostream) subStorages.Add(target);
            }

            foreach (var child in Children)
            {
                InternalAction(child.Value);
            }

            if (!recursive || subStorages.Count <= 0) return;
            foreach (var d in subStorages)
            {
                new CfStorage(CompoundFile, d).VisitEntries(action, true);
            }
        }

        /// <summary>
        /// Remove an entry from the current storage and compound file.
        /// </summary>
        /// <param name="entryName">The name of the entry in the current storage to delete</param>
        /// <exception cref="ArgumentOutOfRangeException"></exception>
        /// <example>
        /// <code>
        /// cf = new CompoundFile("A_FILE_YOU_CAN_CHANGE.cfs", UpdateMode.Update, true, false);
        /// cf.RootStorage.Delete("AStream"); // AStream item is assumed to exist.
        /// cf.Commit(true);
        /// cf.Close();
        /// </code>
        /// </example>
        /// <exception cref="T:OpenMcdf.CFDisposedException">Raised if trying to delete item from a closed compound file</exception>
        /// <exception cref="T:OpenMcdf.CFItemNotFound">Raised if item to delete is not found</exception>
        /// <exception cref="T:OpenMcdf.CFException">Raised if trying to delete root storage</exception>
        public void Delete(string entryName)
        {
            CheckDisposed();
            var foundObj = Children.FindNode(entryName);
            if (foundObj == null)
                throw new CfItemNotFound("Entry named [" + entryName + "] was not found");

            //if (foundObj.GetType() != typeCheck)
            //    throw new CFException("Entry named [" + entryName + "] has not the correct type");

            if (foundObj.Value.StgType == StgType.StgRoot)
                throw new CfException("Root storage cannot be removed");

            switch (foundObj.Value.StgType)
            {
                case StgType.StgStorage:
                    var temp = new CfStorage(CompoundFile, foundObj.Value);
                    // This is a storage. we have to remove children items first
                    foreach (var ded in temp.Children.Select(de => de.Value).ToList())
                    {
                        temp.Delete(ded.Name);
                    }
                    // ...then we Remove storage item from children tree...
                    var dirEntryStg = foundObj.Value;
                    Children.RemoveNode(foundObj);

                    // ...after which we need to rethread the root of siblings tree...
                    DirEntry.Child = Children.Root?.Value.Sid ?? DirectoryEntry.Nostream;
                    CompoundFile.InvalidateDirectoryEntry(dirEntryStg.Sid);
                    break;
                
                case StgType.StgStream:
                    // Free directory associated data stream. 
                    CompoundFile.FreeAssociatedData(foundObj.Value.Sid);
                    // Remove item from children tree
                    var dirEntryStm = foundObj.Value;
                    Children.RemoveNode(foundObj);
                    // Rethread the root of siblings tree...
                    DirEntry.Child = Children.Root?.Value.Sid ?? DirectoryEntry.Nostream;
                    CompoundFile.InvalidateDirectoryEntry(dirEntryStm.Sid);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(foundObj.Value.StgType));
            }
        }


        /// <summary>
        /// Rename a Stream or Storage item in the current storage
        /// </summary>
        /// <param name="oldItemName">The item old name to lookup</param>
        /// <param name="newItemName">The new name to assign</param>
        public void RenameItem(string oldItemName, string newItemName)
        {
            if (Children.TryFindNode(oldItemName, out var item))
            {
                item.SetEntryName(newItemName);
            }
            else throw new CfItemNotFound("Item " + oldItemName + " not found in Storage");

            _children = null;
            _children = LoadChildren(DirEntry.Sid) ?? CompoundFile.CreateNewTree(); //Rethread
        }
    }
}
