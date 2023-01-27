/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 * 
 * The Original Code is OpenMCDF - Compound Document Format library.
 * 
 * The Initial Developer of the Original Code is Federico Blaseotto.*/


using System;
using System.IO;
using RedBlackTree;

namespace OpenMcdf
{
    internal interface IDirectoryEntry : IComparable, IRbNode
    {
        int Child { get; set; }
        byte[] CreationDate { get; set; }
        byte[] EntryName { get; }
        string GetEntryName();
        int LeftSibling { get; set; }
        byte[] ModifyDate { get; set; }
        string Name { get; }
        ushort NameLength { get; set; }
        void Read(Stream stream, CfsVersion ver = CfsVersion.Ver3);
        int RightSibling { get; set; }
        void SetEntryName(string entryName);
        int Sid { get; set; }
        long Size { get; set; }
        int StartSetc { get; set; }
        int StateBits { get; set; }
        StgColor StgColor { get; set; }
        StgType StgType { get; set; }
        Guid StorageClsid { get; set; }
        void Write(Stream stream);
    }
}
