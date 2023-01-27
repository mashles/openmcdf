﻿using System;
using System.IO;
using System.Text;

namespace OpenMcdf.Extensions.OLEProperties
{
    public class DictionaryEntry
    {
        private const int CpWinunicode = 0x04B0;

        readonly int _codePage;

        public DictionaryEntry(int codePage)
        {
            _codePage = codePage;
        }

        public uint PropertyIdentifier { get; set; }
        public int Length { get; set; }
        public string Name => GetName();

        private byte[] _nameBytes;

        public void Read(BinaryReader br)
        {
            PropertyIdentifier = br.ReadUInt32();
            Length = br.ReadInt32();

            if (_codePage != CpWinunicode)
            {
                _nameBytes = br.ReadBytes(Length);
            }
            else
            {
                _nameBytes = br.ReadBytes(Length << 2);

                var m = Length % 4;
                if (m > 0)
                    br.ReadBytes(m);
            }
        }

        public void Write(BinaryWriter bw)
        {
            bw.Write(PropertyIdentifier);
            bw.Write(Length);
            bw.Write(_nameBytes);

            //if (codePage == CP_WINUNICODE)
            //    int m = Length % 4;

            //if (m > 0)
            //    for (int i = 0; i < m; i++)
            //        bw.Write((byte)m);
        }

        private string GetName()
        {
            return Encoding.GetEncoding(_codePage).GetString(_nameBytes);
        }


    }
}
