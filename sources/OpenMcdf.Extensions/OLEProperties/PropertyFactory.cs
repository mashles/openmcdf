using System;
using System.IO;
using System.Text;
using System.Threading;
using OpenMcdf.Extensions.OLEProperties.Interfaces;

namespace OpenMcdf.Extensions.OLEProperties
{
    internal class PropertyFactory
    {
        private static readonly ThreadLocal<PropertyFactory> _instance
            = new ThreadLocal<PropertyFactory>(() => { return new PropertyFactory(); });

        public static PropertyFactory Instance
        {
            get
            {

#if NETSTANDARD2_0_OR_GREATER
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
#endif
                return _instance.Value;
            }
        }

        private PropertyFactory()
        {

        }

        public ITypedPropertyValue NewProperty(VtPropertyType vType, int codePage, bool isVariant = false)
        {
            ITypedPropertyValue pr = null;

            switch ((VtPropertyType)((ushort)vType & 0x00FF))
            {
                case VtPropertyType.VtI1:
                    pr = new VtI1Property(vType, isVariant);
                    break;
                case VtPropertyType.VtI2:
                    pr = new VtI2Property(vType, isVariant);
                    break;
                case VtPropertyType.VtI4:
                    pr = new VtI4Property(vType, isVariant);
                    break;
                case VtPropertyType.VtR4:
                    pr = new VtR4Property(vType, isVariant);
                    break;
                case VtPropertyType.VtR8:
                    pr = new VtR8Property(vType, isVariant);
                    break;
                case VtPropertyType.VtCy:
                    pr = new VtCyProperty(vType, isVariant);
                    break;
                case VtPropertyType.VtDate:
                    pr = new VtDateProperty(vType, isVariant);
                    break;
                case VtPropertyType.VtInt:
                    pr = new VtIntProperty(vType, isVariant);
                    break;
                case VtPropertyType.VtUint:
                    pr = new VtUintProperty(vType, isVariant);
                    break;
                case VtPropertyType.VtUi1:
                    pr = new VtUi1Property(vType, isVariant);
                    break;
                case VtPropertyType.VtUi2:
                    pr = new VtUi2Property(vType, isVariant);
                    break;
                case VtPropertyType.VtUi4:
                    pr = new VtUi4Property(vType, isVariant);
                    break;
                case VtPropertyType.VtUi8:
                    pr = new VtUi8Property(vType, isVariant);
                    break;
                case VtPropertyType.VtBstr:
                case VtPropertyType.VtLpstr:
                    pr = new VtLpstrProperty(vType, codePage, isVariant);
                    break;
                case VtPropertyType.VtLpwstr:
                    pr = new VtLpwstrProperty(vType, codePage, isVariant);
                    break;
                case VtPropertyType.VtFiletime:
                    pr = new VtFiletimeProperty(vType, isVariant);
                    break;
                case VtPropertyType.VtDecimal:
                    pr = new VtDecimalProperty(vType, isVariant);
                    break;
                case VtPropertyType.VtBool:
                    pr = new VtBoolProperty(vType, isVariant);
                    break;
                case VtPropertyType.VtEmpty:
                    pr = new VtEmptyProperty(vType, isVariant);
                    break;
                case VtPropertyType.VtVariantVector:
                    pr = new VtVariantVector(vType, codePage, isVariant);
                    break;
                case VtPropertyType.VtCf:
                    pr = new VtCfProperty(vType, isVariant);
                    break;
                case VtPropertyType.VtBlobObject:
                case VtPropertyType.VtBlob:
                    pr = new VtBlobProperty(vType, isVariant);
                    break;
                default:
                    throw new Exception("Unrecognized property type");
            }

            return pr;
        }

        #region Property implementations

        private class VtEmptyProperty : TypedPropertyValue<object>
        {
            public VtEmptyProperty(VtPropertyType vType, bool isVariant) : base(vType, isVariant)
            {

            }

            public override object ReadScalarValue(BinaryReader br)
            {
                return null;
            }

            public override void WriteScalarValue(BinaryWriter bw, object pValue)
            {
            }
        }
        private class VtI1Property : TypedPropertyValue<sbyte>
        {
            public VtI1Property(VtPropertyType vType, bool isVariant) : base(vType, isVariant)
            {

            }

            public override sbyte ReadScalarValue(BinaryReader br)
            {
                var r = br.ReadSByte();
                return r;
            }

            public override void WriteScalarValue(BinaryWriter bw, sbyte pValue)
            {
                bw.Write(pValue);
            }
        }

        private class VtUi1Property : TypedPropertyValue<byte>
        {
            public VtUi1Property(VtPropertyType vType, bool isVariant) : base(vType, isVariant)
            {

            }

            public override byte ReadScalarValue(BinaryReader br)
            {
                var r = br.ReadByte();
                return r;
            }

            public override void WriteScalarValue(BinaryWriter bw, byte pValue)
            {
                bw.Write(pValue);
            }
        }

        private class VtUi4Property : TypedPropertyValue<uint>
        {
            public VtUi4Property(VtPropertyType vType, bool isVariant) : base(vType, isVariant)
            {

            }

            public override uint ReadScalarValue(BinaryReader br)
            {
                var r = br.ReadUInt32();
                return r;
            }

            public override void WriteScalarValue(BinaryWriter bw, uint pValue)
            {
                bw.Write(pValue);
            }
        }

        private class VtUi8Property : TypedPropertyValue<ulong>
        {
            public VtUi8Property(VtPropertyType vType, bool isVariant) : base(vType, isVariant)
            {

            }

            public override ulong ReadScalarValue(BinaryReader br)
            {
                var r = br.ReadUInt64();
                return r;
            }

            public override void WriteScalarValue(BinaryWriter bw, ulong pValue)
            {
                bw.Write(pValue);
            }
        }

        private class VtI2Property : TypedPropertyValue<short>
        {
            public VtI2Property(VtPropertyType vType, bool isVariant) : base(vType, isVariant)
            {

            }

            public override short ReadScalarValue(BinaryReader br)
            {
                var r = br.ReadInt16();
                return r;
            }

            public override void WriteScalarValue(BinaryWriter bw, short pValue)
            {
                bw.Write(pValue);
            }
        }



        private class VtUi2Property : TypedPropertyValue<ushort>
        {
            public VtUi2Property(VtPropertyType vType, bool isVariant) : base(vType, isVariant)
            {

            }

            public override ushort ReadScalarValue(BinaryReader br)
            {
                var r = br.ReadUInt16();
                return r;
            }

            public override void WriteScalarValue(BinaryWriter bw, ushort pValue)
            {
                bw.Write(pValue);
            }
        }

        private class VtI4Property : TypedPropertyValue<int>
        {
            public VtI4Property(VtPropertyType vType, bool isVariant) : base(vType, isVariant)
            {
            }

            public override int ReadScalarValue(BinaryReader br)
            {
                var r = br.ReadInt32();
                return r;
            }

            public override void WriteScalarValue(BinaryWriter bw, int pValue)
            {
                bw.Write(pValue);
            }
        }

        private class VtI8Property : TypedPropertyValue<long>
        {
            public VtI8Property(VtPropertyType vType, bool isVariant) : base(vType, isVariant)
            {
            }

            public override long ReadScalarValue(BinaryReader br)
            {
                var r = br.ReadInt64();
                return r;
            }

            public override void WriteScalarValue(BinaryWriter bw, long pValue)
            {
                bw.Write(pValue);
            }
        }

        private class VtIntProperty : TypedPropertyValue<int>
        {
            public VtIntProperty(VtPropertyType vType, bool isVariant) : base(vType, isVariant)
            {
            }

            public override int ReadScalarValue(BinaryReader br)
            {
                var r = br.ReadInt32();
                return r;
            }

            public override void WriteScalarValue(BinaryWriter bw, int pValue)
            {
                bw.Write(pValue);
            }
        }

        private class VtUintProperty : TypedPropertyValue<uint>
        {
            public VtUintProperty(VtPropertyType vType, bool isVariant) : base(vType, isVariant)
            {
            }

            public override uint ReadScalarValue(BinaryReader br)
            {
                var r = br.ReadUInt32();
                return r;
            }

            public override void WriteScalarValue(BinaryWriter bw, uint pValue)
            {
                bw.Write(pValue);
            }
        }


        private class VtR4Property : TypedPropertyValue<float>
        {
            public VtR4Property(VtPropertyType vType, bool isVariant) : base(vType, isVariant)
            {

            }

            public override float ReadScalarValue(BinaryReader br)
            {
                var r = br.ReadSingle();
                return r;
            }

            public override void WriteScalarValue(BinaryWriter bw, float pValue)
            {
                bw.Write(pValue);
            }
        }

        private class VtR8Property : TypedPropertyValue<double>
        {
            public VtR8Property(VtPropertyType vType, bool isVariant) : base(vType, isVariant)
            {

            }

            public override double ReadScalarValue(BinaryReader br)
            {
                var r = br.ReadDouble();
                return r;
            }

            public override void WriteScalarValue(BinaryWriter bw, double pValue)
            {
                bw.Write(pValue);
            }
        }

        private class VtCyProperty : TypedPropertyValue<long>
        {
            public VtCyProperty(VtPropertyType vType, bool isVariant) : base(vType, isVariant)
            {
            }

            public override long ReadScalarValue(BinaryReader br)
            {
                var temp = br.ReadInt64();

                var tmp = (temp /= 10000);

                return (tmp);
            }

            public override void WriteScalarValue(BinaryWriter bw, long pValue)
            {
                bw.Write(pValue * 10000);
            }
        }

        private class VtDateProperty : TypedPropertyValue<DateTime>
        {
            public VtDateProperty(VtPropertyType vType, bool isVariant) : base(vType, isVariant)
            {

            }

            public override DateTime ReadScalarValue(BinaryReader br)
            {
                var temp = br.ReadDouble();

                return DateTime.FromOADate(temp);
            }

            public override void WriteScalarValue(BinaryWriter bw, DateTime pValue)
            {
                bw.Write(pValue.ToOADate());
            }
        }

        private class VtLpstrProperty : TypedPropertyValue<string>
        {

            private byte[] _data;
            private readonly int _codePage;

            public VtLpstrProperty(VtPropertyType vType, int codePage, bool isVariant) : base(vType, isVariant)
            {
                _codePage = codePage;
            }

            public override string ReadScalarValue(BinaryReader br)
            {
                var size = br.ReadUInt32();
                _data = br.ReadBytes((int)size);

                return Encoding.GetEncoding(_codePage).GetString(_data);
            }

            public override void WriteScalarValue(BinaryWriter bw, string pValue)
            {

                _data = Encoding.GetEncoding(_codePage).GetBytes(pValue);

                bw.Write((uint)_data.Length);
                bw.Write(_data);
            }
        }

        private class VtLpwstrProperty : TypedPropertyValue<string>
        {

            private byte[] _data;
            private int _codePage;

            public VtLpwstrProperty(VtPropertyType vType, int codePage, bool isVariant) : base(vType, isVariant)
            {
                _codePage = codePage;
            }

            public override string ReadScalarValue(BinaryReader br)
            {
                var nChars = br.ReadUInt32();
                _data = br.ReadBytes((int)(nChars * 2));  //WChar
                return Encoding.Unicode.GetString(_data);
            }

            public override void WriteScalarValue(BinaryWriter bw, string pValue)
            {
                _data = Encoding.Unicode.GetBytes(pValue);
                bw.Write((uint)_data.Length >> 2);
                bw.Write(_data);
            }
        }

        private class VtFiletimeProperty : TypedPropertyValue<DateTime>
        {

            public VtFiletimeProperty(VtPropertyType vType, bool isVariant) : base(vType, isVariant)
            {

            }

            public override DateTime ReadScalarValue(BinaryReader br)
            {
                var tmp = br.ReadInt64();

                return DateTime.FromFileTime(tmp);
            }

            public override void WriteScalarValue(BinaryWriter bw, DateTime pValue)
            {
                bw.Write((pValue).ToFileTime());

            }
        }

        private class VtDecimalProperty : TypedPropertyValue<decimal>
        {

            public VtDecimalProperty(VtPropertyType vType, bool isVariant) : base(vType, isVariant)
            {

            }

            public override decimal ReadScalarValue(BinaryReader br)
            {
                decimal d;


                br.ReadInt16(); // wReserved
                var scale = br.ReadByte();
                var sign = br.ReadByte();

                var u = br.ReadUInt32();
                d = Convert.ToDecimal(Math.Pow(2, 64)) * u;
                d += br.ReadUInt64();

                if (sign != 0)
                    d = -d;
                d /= (10 << scale);

                PropertyValue = d;
                return d;
            }

            public override void WriteScalarValue(BinaryWriter bw, decimal pValue)
            {
                var parts = decimal.GetBits(pValue);

                var sign = (parts[3] & 0x80000000) != 0;
                var scale = (byte)((parts[3] >> 16) & 0x7F);


                bw.Write((short)0);
                bw.Write(scale);
                bw.Write(sign ? (byte)0 : (byte)1);

                bw.Write(parts[2]);
                bw.Write(parts[1]);
                bw.Write(parts[0]);
            }
        }

        private class VtBoolProperty : TypedPropertyValue<bool>
        {
            public VtBoolProperty(VtPropertyType vType, bool isVariant) : base(vType, isVariant)
            {

            }

            public override bool ReadScalarValue(BinaryReader br)
            {

                PropertyValue = br.ReadUInt16() == 0xFFFF ? true : false;
                return (bool)PropertyValue;
                //br.ReadUInt16();//padding
            }

            public override void WriteScalarValue(BinaryWriter bw, bool pValue)
            {
                bw.Write(pValue ? (ushort)0xFFFF : (ushort)0);

            }

        }

        private class VtCfProperty : TypedPropertyValue<object>
        {
            public VtCfProperty(VtPropertyType vType, bool isVariant) : base(vType, isVariant)
            {

            }

            public override object ReadScalarValue(BinaryReader br)
            {

                var size = br.ReadInt32();
                var data = br.ReadBytes(size);
                return data;
                //br.ReadUInt16();//padding
            }

            public override void WriteScalarValue(BinaryWriter bw, object pValue)
            {
                var r = pValue as byte[];
                if (r != null)
                    bw.Write(r);
            }

        }

        private class VtBlobProperty : TypedPropertyValue<object>
        {
            public VtBlobProperty(VtPropertyType vType, bool isVariant) : base(vType, isVariant)
            {

            }

            public override object ReadScalarValue(BinaryReader br)
            {
                var size = br.ReadInt32();
                var data = br.ReadBytes(size);
                return data;
            }

            public override void WriteScalarValue(BinaryWriter bw, object pValue)
            {
                var r = pValue as byte[];
                if (r != null)
                    bw.Write(r);

            }

        }

        private class VtVariantVector : TypedPropertyValue<object>
        {
            private readonly int _codePage;

            public VtVariantVector(VtPropertyType vType, int codePage, bool isVariant) : base(vType, isVariant)
            {
                _codePage = codePage;
            }

            public override object ReadScalarValue(BinaryReader br)
            {
                var vType = (VtPropertyType)br.ReadUInt16();
                br.ReadUInt16(); // Ushort Padding

                var p = Instance.NewProperty(vType, _codePage, true);
                p.Read(br);
                return p;
            }

            public override void WriteScalarValue(BinaryWriter bw, object pValue)
            {
                var p = (ITypedPropertyValue)pValue;

                p.Write(bw);
            }
        }

#endregion

    }
}
