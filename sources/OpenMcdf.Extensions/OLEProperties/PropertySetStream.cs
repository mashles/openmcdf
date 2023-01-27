﻿using System;
using System.Collections.Generic;
using System.IO;
using OpenMcdf.Extensions.OLEProperties.Interfaces;

namespace OpenMcdf.Extensions.OLEProperties
{
    internal class PropertySetStream
    {


        public ushort ByteOrder { get; set; }
        public ushort Version { get; set; }
        public uint SystemIdentifier { get; set; }
        public Guid Clsid { get; set; }
        public uint NumPropertySets { get; set; }
        public Guid Fmtid0 { get; set; }
        public uint Offset0 { get; set; }
        public Guid Fmtid1 { get; set; }
        public uint Offset1 { get; set; }
        public PropertySet PropertySet0 { get; set; }
        public PropertySet PropertySet1 { get; set; }

        //private SummaryInfoMap map;

        public void Read(BinaryReader br)
        {
            ByteOrder = br.ReadUInt16();
            Version = br.ReadUInt16();
            SystemIdentifier = br.ReadUInt32();
            Clsid = new Guid(br.ReadBytes(16));
            NumPropertySets = br.ReadUInt32();
            Fmtid0 = new Guid(br.ReadBytes(16));
            Offset0 = br.ReadUInt32();

            if (NumPropertySets == 2)
            {
                Fmtid1 = new Guid(br.ReadBytes(16));
                Offset1 = br.ReadUInt32();
            }

            PropertySet0 = new PropertySet();
            PropertySet0.Size = br.ReadUInt32();
            PropertySet0.NumProperties = br.ReadUInt32();


            // Read property offsets (P0)
            for (var i = 0; i < PropertySet0.NumProperties; i++)
            {
                var pio = new PropertyIdentifierAndOffset();
                pio.PropertyIdentifier = br.ReadUInt32();
                pio.Offset = br.ReadUInt32();
                PropertySet0.PropertyIdentifierAndOffsets.Add(pio);
            }

            PropertySet0.LoadContext((int)Offset0, br);  //Read CodePage, Locale

            // Read properties (P0)
            for (var i = 0; i < PropertySet0.NumProperties; i++)
            {
                br.BaseStream.Seek(Offset0 + PropertySet0.PropertyIdentifierAndOffsets[i].Offset, SeekOrigin.Begin);
                PropertySet0.Properties.Add(ReadProperty(PropertySet0.PropertyIdentifierAndOffsets[i].PropertyIdentifier, PropertySet0.PropertyContext.CodePage, br));
            }

            if (NumPropertySets == 2)
            {
                br.BaseStream.Seek(Offset1, SeekOrigin.Begin);
                PropertySet1 = new PropertySet();
                PropertySet1.Size = br.ReadUInt32();
                PropertySet1.NumProperties = br.ReadUInt32();

                // Read property offsets
                for (var i = 0; i < PropertySet1.NumProperties; i++)
                {
                    var pio = new PropertyIdentifierAndOffset();
                    pio.PropertyIdentifier = br.ReadUInt32();
                    pio.Offset = br.ReadUInt32();
                    PropertySet1.PropertyIdentifierAndOffsets.Add(pio);
                }

                PropertySet1.LoadContext((int)Offset1, br);

                // Read properties
                for (var i = 0; i < PropertySet1.NumProperties; i++)
                {
                    br.BaseStream.Seek(Offset1 + PropertySet1.PropertyIdentifierAndOffsets[i].Offset, SeekOrigin.Begin);
                    PropertySet1.Properties.Add(ReadProperty(PropertySet1.PropertyIdentifierAndOffsets[i].PropertyIdentifier, PropertySet1.PropertyContext.CodePage, br));
                }
            }
        }

        private class OffsetContainer
        {
            public int OffsetPs { get; set; }

            public List<long> PropertyIdentifierOffsets { get; set; }
            public List<long> PropertyOffsets { get; set; }

            public OffsetContainer()
            {
                PropertyOffsets = new List<long>();
                PropertyIdentifierOffsets = new List<long>();
            }
        }

        public void Write(BinaryWriter bw)
        {
            var oc0 = new OffsetContainer();
            var oc1 = new OffsetContainer();

            bw.Write(ByteOrder);
            bw.Write(Version);
            bw.Write(SystemIdentifier);
            bw.Write(Clsid.ToByteArray());
            bw.Write(NumPropertySets);
            bw.Write(Fmtid0.ToByteArray());
            bw.Write(Offset0);

            if (NumPropertySets == 2)
            {
                bw.Write(Fmtid1.ToByteArray());
                bw.Write(Offset1);
            }


            oc0.OffsetPs = (int)bw.BaseStream.Position;
            bw.Write(PropertySet0.Size);
            bw.Write(PropertySet0.NumProperties);

            // w property offsets
            for (var i = 0; i < PropertySet0.NumProperties; i++)
            {
                oc0.PropertyIdentifierOffsets.Add(bw.BaseStream.Position); //Offset of 4 to Offset value
                PropertySet0.PropertyIdentifierAndOffsets[i].Write(bw);
            }

            for (var i = 0; i < PropertySet0.NumProperties; i++)
            {
                oc0.PropertyOffsets.Add(bw.BaseStream.Position);
                PropertySet0.Properties[i].Write(bw);
            }

            var padding0 = bw.BaseStream.Position % 4;

            //if (padding0 > 0)
            //{
            //    for (int p = 0; p < padding0; p++)
            //        bw.Write((byte)0);
            //}

            var size0 = (int)(bw.BaseStream.Position - oc0.OffsetPs);



            if (NumPropertySets == 2)
            {


                oc1.OffsetPs = (int)bw.BaseStream.Position;

                bw.Write(PropertySet1.Size);
                bw.Write(PropertySet1.NumProperties);

                // w property offsets
                for (var i = 0; i < PropertySet1.PropertyIdentifierAndOffsets.Count; i++)
                {
                    oc1.PropertyIdentifierOffsets.Add(bw.BaseStream.Position); //Offset of 4 to Offset value
                    PropertySet1.PropertyIdentifierAndOffsets[i].Write(bw);
                }

                for (var i = 0; i < PropertySet1.NumProperties; i++)
                {
                    oc1.PropertyOffsets.Add(bw.BaseStream.Position);
                    PropertySet1.Properties[i].Write(bw);
                }

                var size1 = (int)(bw.BaseStream.Position - oc1.OffsetPs);

                bw.Seek(oc1.OffsetPs + 4, SeekOrigin.Begin);
                bw.Write(size1);
            }

            bw.Seek(oc0.OffsetPs, SeekOrigin.Begin);
            bw.Write(size0);

            var shiftO1 = 2 + 2 + 4 + 16 + 4 + 16; //OFFSET0
            bw.Seek(shiftO1, SeekOrigin.Begin);
            bw.Write(oc0.OffsetPs);

            if (NumPropertySets == 2)
            {
                bw.Seek(shiftO1 + 4 + 16, SeekOrigin.Begin);
                bw.Write(oc1.OffsetPs);
            }

            //-----------

            for (var i = 0; i < PropertySet0.PropertyIdentifierAndOffsets.Count; i++)
            {
                bw.Seek((int)oc0.PropertyIdentifierOffsets[i] + 4, SeekOrigin.Begin); //Offset of 4 to Offset value
                bw.Write((int)(oc0.PropertyOffsets[i] - oc0.OffsetPs));
            }



            if (PropertySet1 != null)
            {
                for (var i = 0; i < PropertySet1.PropertyIdentifierAndOffsets.Count; i++)
                {
                    bw.Seek((int)oc1.PropertyIdentifierOffsets[i], SeekOrigin.Begin); //Offset of 4 to Offset value
                    bw.Write(oc1.PropertyOffsets[i] - oc1.OffsetPs);
                }
            }
        }



        private IProperty ReadProperty(uint propertyIdentifier, int codePage, BinaryReader br)
        {
            if (propertyIdentifier != 0)
            {
                var vType = (VtPropertyType)br.ReadUInt16();
                br.ReadUInt16(); // Ushort Padding

                var pr = PropertyFactory.Instance.NewProperty(vType, codePage);
                pr.Read(br);

                return pr;
            }

            IDictionaryProperty dictionaryProperty = new DictionaryProperty(codePage);
            dictionaryProperty.Read(br);
            return dictionaryProperty;
        }
    }
}