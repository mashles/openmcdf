﻿using System.IO;
using OpenMcdf.Extensions.OLEProperties.Interfaces;

namespace OpenMcdf.Extensions.OLEProperties
{
    public class PropertyIdentifierAndOffset: IBinarySerializable
    {
        public uint PropertyIdentifier { get; set; }
        public uint Offset { get; set; }

        public void Read(BinaryReader br)
        {
            PropertyIdentifier = br.ReadUInt32();
            Offset = br.ReadUInt32();
        }

        public void Write(BinaryWriter bw)
        {
            bw.Write(PropertyIdentifier);
            bw.Write(Offset);
        }
    }
}