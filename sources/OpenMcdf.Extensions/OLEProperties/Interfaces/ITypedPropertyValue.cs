namespace OpenMcdf.Extensions.OLEProperties.Interfaces
{
    public interface ITypedPropertyValue : IProperty
    {
        VtPropertyType VtType
        {
            get;
            //set;
        }

        PropertyDimensions PropertyDimensions
        {
            get;
        }

        bool IsVariant
        {
            get; 
        }
    }
}
