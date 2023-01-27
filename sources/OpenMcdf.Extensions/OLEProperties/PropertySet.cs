using System.Collections.Generic;
using System.IO;
using System.Linq;
using OpenMcdf.Extensions.OLEProperties.Interfaces;

namespace OpenMcdf.Extensions.OLEProperties
{
    internal class PropertySet
    {

        public PropertyContext PropertyContext
        {
            get;  set;
        }

        public uint Size { get; set; }

        public uint NumProperties { get; set; }

        List<PropertyIdentifierAndOffset> _propertyIdentifierAndOffsets
            = new List<PropertyIdentifierAndOffset>();

        public List<PropertyIdentifierAndOffset> PropertyIdentifierAndOffsets
        {
            get => _propertyIdentifierAndOffsets;
            set => _propertyIdentifierAndOffsets = value;
        }

        List<IProperty> _properties = new List<IProperty>();
        public List<IProperty> Properties
        {
            get => _properties;
            set => _properties = value;
        }

        public void LoadContext(int propertySetOffset, BinaryReader br)
        {
            var currPos = br.BaseStream.Position;

            PropertyContext = new PropertyContext();
            var codePageOffset = (int)(propertySetOffset + PropertyIdentifierAndOffsets.Where(pio => pio.PropertyIdentifier == 1).First().Offset);
            br.BaseStream.Seek(codePageOffset, SeekOrigin.Begin);

            var vType = (VtPropertyType)br.ReadUInt16();
            br.ReadUInt16(); // Ushort Padding
            PropertyContext.CodePage = (ushort)br.ReadInt16();

            br.BaseStream.Position = currPos;
        }

    }
}
