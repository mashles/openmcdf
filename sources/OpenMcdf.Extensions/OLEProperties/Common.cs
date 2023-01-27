using System;

namespace OpenMcdf.Extensions.OLEProperties
{
    public enum Behavior
    {
        CaseSensitive, CaseInsensitive
    }

    public class PropertyContext
    {

        public int CodePage { get; set; }
        public Behavior Behavior { get; set; }
        public uint Locale { get; set; }
    }

    public static class WellKnownFmtid
    {
        public static string FmtidSummaryInformation = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}";
        public static string FmtidDocSummaryInformation = "{D5CDD502-2E9C-101B-9397-08002B2CF9AE}";
        public static string FmtidUserDefinedProperties = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
        public static string FmtidGlobalInfo = "{56616F00-C154-11CE-8553-00AA00A1F95B}";
        public static string FmtidImageContents = "{56616400-C154-11CE-8553-00AA00A1F95B}";
        public static string FmtidImageInfo = "{56616500-C154-11CE-8553-00AA00A1F95B}";
    }

    public enum PropertyDimensions
    {
        IsScalar, IsVector, IsArray
    }

    public enum PropertyType
    {
        TypedPropertyValue = 0, DictionaryProperty = 1
    }
}
