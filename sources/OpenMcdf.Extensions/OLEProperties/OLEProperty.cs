
namespace OpenMcdf.Extensions.OLEProperties
{
    public class OleProperty
    {
        private readonly OlePropertiesContainer _container;

        internal OleProperty(OlePropertiesContainer container)
        {
            _container = container;
        }

        public string PropertyName => DecodePropertyIdentifier();

        private string DecodePropertyIdentifier()
        {
            return PropertyIdentifier.GetDescription(_container.ContainerType, _container.PropertyNames);
        }

        //public string Description { get { return description; }
        public uint PropertyIdentifier { get; internal set; }

        public VtPropertyType VtType
        {
            get;
            internal set;
        }

        public object Value
        {
            get;
            set;
        }

        public override bool Equals(object obj)
        {
            var other = obj as OleProperty;
            if (other == null) return false;

            return other.PropertyIdentifier == PropertyIdentifier;
        }

        public override int GetHashCode()
        {
            return (int)PropertyIdentifier;
        }

    }
}