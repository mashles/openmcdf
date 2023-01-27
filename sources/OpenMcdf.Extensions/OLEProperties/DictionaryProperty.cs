using System.Collections.Generic;
using System.IO;
using System.Text;
using OpenMcdf.Extensions.OLEProperties.Interfaces;

namespace OpenMcdf.Extensions.OLEProperties
{
    public class DictionaryProperty : IDictionaryProperty
    {
        private readonly int _codePage;

        public DictionaryProperty(int codePage)
        {
            _codePage = codePage;
            _entries = new Dictionary<uint, string>();

        }

        public PropertyType PropertyType => PropertyType.DictionaryProperty;

        private Dictionary<uint, string> _entries;

        public object Value
        {
            get => _entries;
            set => _entries = (Dictionary<uint, string>)value;
        }

        public void Read(BinaryReader br)
        {
            var curPos = br.BaseStream.Position;

            var numEntries = br.ReadUInt32();

            for (uint i = 0; i < numEntries; i++)
            {
                var de = new DictionaryEntry(_codePage);

                de.Read(br);
                _entries.Add(de.PropertyIdentifier, de.Name);
            }

            var m = (int)(br.BaseStream.Position - curPos) % 4;

            if (m > 0)
            {
                for(var i = 0; i < m; i++)
                {
                    br.ReadByte();
                }
            }

        }

        public void Write(BinaryWriter bw)
        {
            bw.Write(_entries.Count);

            foreach (var kv in _entries)
            {
                bw.Write(kv.Key);
                var s = kv.Value;
                if (!s.EndsWith("\0"))
                    s += "\0";
                bw.Write(Encoding.GetEncoding(_codePage).GetBytes(s));
            }

        }
    }
}
