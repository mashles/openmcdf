using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace OpenMcdf;

public class DirectoryComparer : IComparer<string>
{
    private const int ThisIsGreater = 1;
    private const int OtherIsGreater = -1;
    public int Compare(string thisDir, string otherDir)
    {
        if (thisDir.Length > otherDir.Length)
        {
            return ThisIsGreater;
        }

        if (thisDir.Length < otherDir.Length)
        {
            return OtherIsGreater;
        }

        for (var z = 0; z < thisDir.Length; z++)
        {
            var thisChar = char.ToUpperInvariant(thisDir[z]);
            var otherChar = char.ToUpperInvariant(otherDir[z]);

            if (thisChar > otherChar)
                return ThisIsGreater;
            if (thisChar < otherChar)
                return OtherIsGreater;
        }

        return 0;
    }
}