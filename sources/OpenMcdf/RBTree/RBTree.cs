#define ASSERT

using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

// ------------------------------------------------------------- 
// This is a porting from java code, under MIT license of       |
// the beautiful Red-Black Tree implementation you can find at  |
// http://en.literateprograms.org/Red-black_tree_(Java)#chunk   |
// Many Thanks to original Implementors.                        |
// -------------------------------------------------------------

namespace OpenMcdf.RBTree
{
    public class RbTreeException : Exception
    {
        public RbTreeException(string msg)
            : base(msg)
        {
        }
    }
    public class RbTreeDuplicatedItemException : RbTreeException
    {
        public RbTreeDuplicatedItemException(string msg)
            : base(msg)
        {
        }
    }

    public enum Color { Red = 0, Black = 1 }

    /// <summary>
    /// Red Black Node interface
    /// </summary>
    public interface IRbNode : IComparable
    {

        IRbNode Left
        {
            get;
            set;
        }

        IRbNode Right
        {
            get;
            set;
        }


        Color Color

        { get; set; }



        IRbNode Parent { get; set; }


        IRbNode Grandparent();


        IRbNode Sibling();
        //        {
        //#if ASSERT
        //            Debug.Assert(Parent != null); // Root node has no sibling
        //#endif
        //            if (this == Parent.Left)
        //                return Parent.Right;
        //            else
        //                return Parent.Left;
        //        }

        IRbNode Uncle();
        //        {
        //#if ASSERT
        //            Debug.Assert(Parent != null); // Root node has no uncle
        //            Debug.Assert(Parent.Parent != null); // Children of root have no uncle
        //#endif
        //            return Parent.Sibling();
        //        }
        //    }

        void AssignValueTo(IRbNode other);
    }
    
    internal enum NodeOperation
    {
        LeftAssigned, RightAssigned, ColorAssigned, ParentAssigned,
        ValueAssigned
    }


}
