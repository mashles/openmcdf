using System.Collections.Generic;
using System.IO;
using OpenMcdf.Extensions.OLEProperties.Interfaces;

namespace OpenMcdf.Extensions.OLEProperties
{
    internal abstract class TypedPropertyValue<T> : ITypedPropertyValue
    {
        private readonly bool _isVariant;
        private readonly PropertyDimensions _dim = PropertyDimensions.IsScalar;

        private readonly VtPropertyType _vtType;

        public PropertyType PropertyType => PropertyType.TypedPropertyValue;

        public VtPropertyType VtType => _vtType;

        protected object PropertyValue;

        public TypedPropertyValue(VtPropertyType vtType, bool isVariant = false)
        {
            _vtType = vtType;
            _dim = CheckPropertyDimensions(vtType);
            _isVariant = isVariant;
        }

        public PropertyDimensions PropertyDimensions => _dim;

        public bool IsVariant => _isVariant;

        private PropertyDimensions CheckPropertyDimensions(VtPropertyType vtType)
        {
            if ((((ushort)vtType) & 0x1000) != 0)
                return PropertyDimensions.IsVector;
            if ((((ushort)vtType) & 0x2000) != 0)
                return PropertyDimensions.IsArray;
            return PropertyDimensions.IsScalar;
        }

        public virtual object Value
        {
            get => PropertyValue;

            set => PropertyValue = value;
        }

        public abstract T ReadScalarValue(BinaryReader br);


        public void Read(BinaryReader br)
        {
            var currentPos = br.BaseStream.Position;
            var size = 0;
            var m = 0;

            switch (PropertyDimensions)
            {
                case PropertyDimensions.IsScalar:
                    PropertyValue = ReadScalarValue(br);
                    size = (int)(br.BaseStream.Position - currentPos);

                    m = size % 4;

                    if (m > 0 && !IsVariant)
                        br.ReadBytes(m); // padding
                    break;

                case PropertyDimensions.IsVector:
                    var nItems = br.ReadUInt32();

                    var res = new List<T>();


                    for (var i = 0; i < nItems; i++)
                    {
                        var s = ReadScalarValue(br);

                        res.Add(s);
                    }

                    PropertyValue = res;
                    size = (int)(br.BaseStream.Position - currentPos);

                    m = size % 4;
                    if (m > 0 && !IsVariant)
                        br.ReadBytes(m); // padding
                    break;
            }
        }

        public abstract void WriteScalarValue(BinaryWriter bw, T pValue);

        public void Write(BinaryWriter bw)
        {
            var currentPos = bw.BaseStream.Position;
            var size = 0;
            var m = 0;
            var needsPadding = HasPadding();

            switch (PropertyDimensions)
            {
                case PropertyDimensions.IsScalar:

                    bw.Write((ushort)_vtType);
                    bw.Write((ushort)0);

                    WriteScalarValue(bw, (T)PropertyValue);
                    size = (int)(bw.BaseStream.Position - currentPos);
                    m = size % 4;

                    if (m > 0 && needsPadding)
                        for (var i = 0; i < m; i++)  // padding
                            bw.Write((byte)0);
                    break;

                case PropertyDimensions.IsVector:

                    bw.Write((ushort)_vtType);
                    bw.Write((ushort)0);
                    bw.Write((uint)((List<T>)PropertyValue).Count);

                    for (var i = 0; i < ((List<T>)PropertyValue).Count; i++)
                    {
                        WriteScalarValue(bw, ((List<T>)PropertyValue)[i]);
                    }

                    size = (int)(bw.BaseStream.Position - currentPos);
                    m = size % 4;

                    if (m > 0 && needsPadding)
                        for (var i = 0; i < m; i++)  // padding
                            bw.Write((byte)0);
                    break;
            }
        }

        private bool HasPadding()
        {

            var vt = (VtPropertyType)((ushort)VtType & 0x00FF);

            switch (vt)
            {
                case VtPropertyType.VtLpstr:
                    if (IsVariant) return false;
                    if (_dim == PropertyDimensions.IsVector) return false;
                    break;
                case VtPropertyType.VtVariantVector:
                    if (_dim == PropertyDimensions.IsVector) return false;
                    break;
                default:
                    return true;
            }

            return true;
        }
    }
}